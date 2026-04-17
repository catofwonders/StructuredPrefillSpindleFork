declare const spindle: import('lumiverse-spindle-types').SpindleAPI

// ─── Types ───────────────────────────────────────────────────────────────────

interface Settings {
  enabled: boolean
  hide_prefill_in_display: boolean
  newline_token: string
  min_chars_after_prefix: number
  continue_overlap_chars: number
  anti_slop_ban_list: string
  prefill_gen_enabled: boolean
  prefill_gen_extra_prompt: string
  prefill_gen_extra_prompt_role: string
  prefill_gen_connection_id: string
  prefill_gen_max_tokens: number
  prefill_gen_stop: string
  prefill_gen_keep_matched_stop_string: boolean
  prefill_gen_timeout_ms: number
}

interface LlmMessageDTO {
  role: 'system' | 'user' | 'assistant'
  content: string
  name?: string
}

interface RuntimeState {
  active: boolean
  lastInjectedAt: number
  accumulatedStreamText: string
  lastAppliedText: string
  expectedPrefill: string
  newlineToken: string
  patternMode: 'default' | 'anthropic'
  knownNames: string[]
  activeChatId: string
  activeGenerationId: string
  activeConnectionId: string
  activeInjectionMode: 'json_schema' | 'json_object' | 'prompt_only' | 'gemini' | 'anthropic_tool'
  retryMessages: LlmMessageDTO[]
  continue: {
    active: boolean
    baseText: string
    displayBase: string
  }
  hidePrefillLiteral: string
  hidePrefillRegex: RegExp | null
  streamGuard: {
    startedAt: number
    lastRawLen: number
    lastDecodedLen: number
    lastProgressAt: number
    suspiciousStreak: number
    stopRequested: boolean
  }
}

// ─── Constants ───────────────────────────────────────────────────────────────

const EXT_NAME = 'StructuredPrefill'

const DEFAULT_SETTINGS: Settings = {
  enabled: true,
  hide_prefill_in_display: true,
  newline_token: '\\n',
  min_chars_after_prefix: 80,
  continue_overlap_chars: 14,
  anti_slop_ban_list: '',
  prefill_gen_enabled: false,
  prefill_gen_extra_prompt: '',
  prefill_gen_extra_prompt_role: 'system',
  prefill_gen_connection_id: '',
  prefill_gen_max_tokens: 15,
  prefill_gen_stop: '',
  prefill_gen_keep_matched_stop_string: false,
  prefill_gen_timeout_ms: 120000,
}

// ─── State ───────────────────────────────────────────────────────────────────

let settings: Settings = { ...DEFAULT_SETTINGS }
let currentUserId: string | undefined

const state: RuntimeState = {
  active: false,
  lastInjectedAt: 0,
  accumulatedStreamText: '',
  lastAppliedText: '',
  expectedPrefill: '',
  newlineToken: '',
  patternMode: 'default',
  knownNames: [],
  activeChatId: '',
  activeGenerationId: '',
  activeConnectionId: '',
  activeInjectionMode: 'json_schema',
  retryMessages: [],
  continue: {
    active: false,
    baseText: '',
    displayBase: '',
  },
  hidePrefillLiteral: '',
  hidePrefillRegex: null,
  streamGuard: {
    startedAt: 0,
    lastRawLen: 0,
    lastDecodedLen: 0,
    lastProgressAt: 0,
    suspiciousStreak: 0,
    stopRequested: false,
  },
}

// ─── Provider Compatibility Cache (three-tier ladder) ───────────────────────
//
// Some OpenAI-compatible proxies reject `response_format: json_schema` (and
// sometimes even `response_format: json_object`) silently — they return empty
// content with no error. We maintain a three-tier ladder:
//
//   tier 1: json_schema    — full regex-locked structured output (real OpenAI)
//   tier 2: json_object    — JSON-shape-only, works on many proxies
//   tier 3: prompt_only    — no response_format at all, just a system nudge
//                            telling the model to wrap its reply in JSON
//
// We probe tier 1 first on unknown connections. If a tier returns empty, we
// auto-retry via spindle.generate.raw() with the next tier, cache the working
// tier, and apply its output. Future messages on that connection skip straight
// to the cached tier.
//
// Storage: per-user file `json_schema_blocklist.json` — kept as the historical
// filename so old installs migrate cleanly. Stored as a JSON object mapping
// connectionId -> tier.

type CompatTier = 'json_schema' | 'json_object' | 'prompt_only'
const TIER_ORDER: CompatTier[] = ['json_schema', 'json_object', 'prompt_only']

function nextTier(current: CompatTier): CompatTier | null {
  const idx = TIER_ORDER.indexOf(current)
  if (idx < 0 || idx >= TIER_ORDER.length - 1) return null
  return TIER_ORDER[idx + 1]
}

const compatCache: Map<string, CompatTier> = new Map()
let compatCacheLoaded = false
const BLOCKLIST_FILE = 'json_schema_blocklist.json'

async function loadCompatCache(): Promise<void> {
  if (compatCacheLoaded) return
  try {
    // Migration-tolerant load: accept either old-format string[] (legacy) or
    // new-format { connectionId: tier }. Old-format entries become 'json_object'.
    const stored = await spindle.userStorage.getJson<unknown>(BLOCKLIST_FILE, { fallback: {}, userId: currentUserId })
    compatCache.clear()
    if (Array.isArray(stored)) {
      for (const id of stored) {
        if (typeof id === 'string' && id) compatCache.set(id, 'json_object')
      }
    } else if (stored && typeof stored === 'object') {
      for (const [id, tier] of Object.entries(stored as Record<string, unknown>)) {
        if (typeof id === 'string' && id && (tier === 'json_schema' || tier === 'json_object' || tier === 'prompt_only')) {
          compatCache.set(id, tier as CompatTier)
        }
      }
    }
    compatCacheLoaded = true
    if (compatCache.size > 0) {
      spindle.log.info(`[SP] Loaded ${compatCache.size} connection compatibility entries`)
    }
  } catch (err) {
    spindle.log.warn(`[SP] Could not load compat cache: ${err}`)
    compatCacheLoaded = true
  }
}

async function persistCompatCache(): Promise<void> {
  try {
    const obj: Record<string, CompatTier> = {}
    for (const [k, v] of compatCache.entries()) obj[k] = v
    await spindle.userStorage.setJson(BLOCKLIST_FILE, obj, { indent: 2, userId: currentUserId })
  } catch (err) {
    spindle.log.warn(`[SP] Could not persist compat cache: ${err}`)
  }
}

async function setCompatTier(connectionId: string, tier: CompatTier): Promise<void> {
  if (!connectionId) return
  if (compatCache.get(connectionId) === tier) return
  compatCache.set(connectionId, tier)
  spindle.log.info(`[SP] Set compat tier for ${connectionId}: ${tier}`)
  await persistCompatCache()
}

function getCompatTier(connectionId: string): CompatTier {
  if (!connectionId) return 'json_schema'
  return compatCache.get(connectionId) ?? 'json_schema'
}

// ─── Settings I/O ────────────────────────────────────────────────────────────

let settingsLoaded = false

async function loadSettings(): Promise<void> {
  if (settingsLoaded) return
  try {
    // Use userStorage (per-user) instead of storage (shared) for operator-scoped compatibility
    settings = await spindle.userStorage.getJson<Settings>('settings.json', {
      fallback: { ...DEFAULT_SETTINGS },
      userId: currentUserId,
    })
    // Merge missing keys from defaults
    for (const [key, value] of Object.entries(DEFAULT_SETTINGS)) {
      if ((settings as any)[key] == null) {
        ;(settings as any)[key] = value
      }
    }
    settingsLoaded = true
  } catch {
    settings = { ...DEFAULT_SETTINGS }
  }
  // Load the compatibility cache alongside settings
  await loadCompatCache()
}

async function saveSettings(): Promise<void> {
  try {
    await spindle.userStorage.setJson('settings.json', settings, { indent: 2, userId: currentUserId })
  } catch (err) {
    spindle.log.error(`[SP] Failed to save settings: ${err}`)
  }
}

// ─── Pure Utility Functions ──────────────────────────────────────────────────

function escapeRegExp(str: string): string {
  return String(str).replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
}

function escapePrefixLiteral(str: string): string {
  return escapeRegExp(str).replace(/"/g, '(?:\\\\)*"')
}

function clampInt(value: unknown, min: number, max: number, fallback: number): number {
  const num = Number(value)
  if (!Number.isFinite(num)) return fallback
  const int = Math.trunc(num)
  return Math.min(max, Math.max(min, int))
}

function normalizeNewlines(text: string): string {
  return String(text ?? '').replace(/\r\n/g, '\n').replace(/\r/g, '\n')
}

function prefixHasSlots(prefix: string): boolean {
  return /\[\[[^\]]+?\]\]/.test(String(prefix ?? ''))
}

function looksLikeStructuredJsonBlob(text: string): boolean {
  const s = String(text ?? '')
  if (!s.includes('{') || !s.includes('"')) return false
  return /\{\s*"(?:response|value|prefix|content)"\s*:/.test(s)
}

// ─── Newline Token Handling ──────────────────────────────────────────────────

function chooseNewlineToken(prefix: string, preferredToken: string): string {
  const token = String(preferredToken ?? '\\n')
  if (!token) return '\\n'
  // If the prefix already contains the literal token string, we need a different one
  // to avoid ambiguity. Try a few candidates.
  const candidates = [token, '<NL>', '⏎', '\\n', '<NEWLINE>']
  for (const candidate of candidates) {
    if (candidate && !prefix.includes(candidate)) return candidate
  }
  return token
}

function encodeNewlines(text: string, newlineToken: string): string {
  if (!newlineToken) return text
  return String(text ?? '').replace(/\n/g, newlineToken)
}

function decodeNewlines(text: string, newlineToken: string): string {
  if (!newlineToken) return text
  const escaped = escapeRegExp(newlineToken)
  return String(text ?? '').replace(new RegExp(escaped, 'g'), '\n')
}

// ─── Quote Handling ──────────────────────────────────────────────────────────

function curlyQuoteLiteralsOutsideSlots(template: string): string {
  // Convert literal " to curly quotes outside [[...]] slot markers.
  // This prevents JSON string termination issues in structured output.
  const parts = String(template ?? '').split(/(\[\[[^\]]*\]\])/g)
  return parts
    .map((part, i) => {
      if (i % 2 === 1) return part // inside [[...]] — leave as-is
      return part.replace(/"/g, '\u201C').replace(/'/g, '\u2018') // " → " and ' → '
    })
    .join('')
}

function straightenCurlyQuotes(s: string): string {
  return String(s ?? '')
    .replace(/[\u201C\u201D]/g, '"')
    .replace(/[\u2018\u2019]/g, "'")
}

// ─── Provider Compatibility ──────────────────────────────────────────────────

type ProviderMode = 'openai' | 'gemini' | 'anthropic' | 'unsupported'

// Providers whose `response_format: json_schema` either isn't supported or is
// silently downgraded to plain JSON mode (no regex enforcement). These fall
// back to 'unsupported' rather than pretending the prefill is locked.
const DOWNGRADES_JSON_SCHEMA = new Set([
  'ai21', 'deepseek', 'moonshot', 'zai', 'siliconflow',
])

function getProviderMode(source: string, modelId: string): ProviderMode {
  const src = String(source ?? '').toLowerCase()
  const model = String(modelId ?? '').toLowerCase()

  if (!src) return 'unsupported'

  // Anthropic / Claude — tool-call forced output
  if (src === 'claude' || src === 'anthropic') return 'anthropic'

  // Google Gemini — responseMimeType + responseSchema
  if (src === 'google' || src === 'gemini' || src === 'makersuite' || src === 'google-ai-studio') {
    return 'gemini'
  }

  // OpenRouter routes by model — detect Claude / Gemini under the hood
  if (src === 'openrouter') {
    if (model.includes('claude') || model.includes('anthropic')) return 'anthropic'
    if (model.includes('gemini') || model.startsWith('google/')) return 'gemini'
    return 'openai' // OpenRouter OpenAI-compatible for everything else
  }

  if (DOWNGRADES_JSON_SCHEMA.has(src)) return 'unsupported'

  // Default: assume OpenAI-compatible
  return 'openai'
}

// Legacy regex-pattern-mode selector. Anthropic rejects non-ASCII in regex
// patterns, so we still flag that case for the schema builder. Gemini/other
// providers use 'default'. Only consulted for the OpenAI regex pattern builder.
function getPatternModeForRequest(providerMode: ProviderMode): 'default' | 'anthropic' {
  return providerMode === 'anthropic' ? 'anthropic' : 'default'
}

// ─── Template Parsing ────────────────────────────────────────────────────────

function splitHidePrefillTemplate(prefixTemplate: string): { hideTemplate: string; hasKeepMarker: boolean } {
  const normalized = normalizeNewlines(prefixTemplate)
  if (!normalized) return { hideTemplate: '', hasKeepMarker: false }

  const markerRe = /\[\[\s*keep\s*\]\]/i
  const m = markerRe.exec(normalized)
  if (!m) return { hideTemplate: normalized, hasKeepMarker: false }

  return {
    hideTemplate: normalized.slice(0, m.index),
    hasKeepMarker: true,
  }
}

function splitEndPrefillTemplate(prefixTemplate: string): { template: string; hasEndMarker: boolean } {
  const normalized = normalizeNewlines(prefixTemplate)
  if (!normalized) return { template: '', hasEndMarker: false }

  const markerRe = /\[\[\s*(end|stop|eos)\s*\]\]/i
  const m = markerRe.exec(normalized)
  if (!m) return { template: normalized, hasEndMarker: false }

  return {
    template: normalized.slice(0, m.index),
    hasEndMarker: true,
  }
}

function templateHasPrefillGenSlot(template: string): boolean {
  return /\[\[\s*pg\s*\]\]/i.test(String(template ?? ''))
}

// ─── Regex Pattern Building ──────────────────────────────────────────────────

function anyCharIncludingNewlineExpr(): string {
  // [\\s\\S] equivalent that avoids \\s for Anthropic compatibility
  return '[\\t\\n\\r \\x21-\\x7E\\x80-\\uFFFF]'
}

function buildSlotRegex(slotContent: string): string {
  const inner = String(slotContent ?? '').trim()
  if (!inner) return '(?:.*?)'

  // [[re:...]] — raw regex passthrough
  const reMatch = /^re:\s*([\s\S]+)$/i.exec(inner)
  if (reMatch) return reMatch[1]

  // [[name]] — match any known character name
  if (/^name$/i.test(inner)) {
    if (state.knownNames.length === 0) return '(?:.*?)'
    const escaped = state.knownNames.map(n => escapePrefixLiteral(n))
    return `(?:${escaped.join('|')})`
  }

  // [[any]] — match anything (non-greedy)
  if (/^any$/i.test(inner)) return '(?:.*?)'

  // [[pg]] — already replaced before this point; treat as empty
  if (/^pg$/i.test(inner)) return ''

  // [[keep]] / [[end]] / [[stop]] / [[eos]] — structural markers, empty in regex
  if (/^(keep|end|stop|eos)$/i.test(inner)) return ''

  // Unknown slot — treat as literal
  return escapePrefixLiteral(`[[${inner}]]`)
}

function buildPrefixRegexFromWireTemplate(wireTemplate: string): string {
  const template = String(wireTemplate ?? '')
  if (!template) return ''

  // Split on [[...]] slot markers
  const parts = template.split(/(\[\[[^\]]*?\]\])/g)
  let regex = ''

  for (let i = 0; i < parts.length; i++) {
    const part = parts[i]
    if (i % 2 === 0) {
      // Literal text segment
      regex += escapePrefixLiteral(part)
    } else {
      // Slot: [[...]]
      const slotInner = part.slice(2, -2)
      regex += buildSlotRegex(slotInner)
    }
  }

  return regex
}

function buildAntiSlopContinuation(banList: string): string {
  const raw = String(banList ?? '').trim()
  if (!raw) return ''

  const words = raw
    .split(/\r?\n/g)
    .map(w => w.trim())
    .filter(Boolean)

  if (words.length === 0) return ''

  // Build a character class that forbids the first character of each banned word,
  // then use a negative lookahead for the full words.
  // This is a simplified approach — the original ST version is more complex.
  const escaped = words.map(w => escapeRegExp(w))
  const lookaheads = escaped.map(w => `(?!${w})`).join('')
  const anyChar = anyCharIncludingNewlineExpr()
  return `(?:${lookaheads}${anyChar})`
}

function buildJsonSchemaForPrefillValuePattern(
  prefix: string,
  minCharsAfterPrefix: number,
  joinSuffixRegex: string = '',
  opts: { mustEndAfterTemplate?: boolean } = {},
): Record<string, any> {
  const mustEndAfterTemplate = !!opts.mustEndAfterTemplate
  const minChars = mustEndAfterTemplate ? 0 : clampInt(minCharsAfterPrefix, 1, 10000, 1)
  const newlineToken = state.newlineToken || '\\n'
  const wirePrefix = encodeNewlines(prefix, newlineToken)

  let prefixRegex = buildPrefixRegexFromWireTemplate(wirePrefix)

  // Allow both encoded newline token and real \n in the pattern
  if (newlineToken) {
    const escapedToken = escapeRegExp(newlineToken)
    prefixRegex = prefixRegex.split(escapedToken).join(`(?:${escapedToken}|\\n)`)
  }

  // Anthropic rejects non-ASCII in regex patterns
  if (state.patternMode === 'anthropic') {
    // eslint-disable-next-line no-control-regex
    prefixRegex = prefixRegex.replace(/[^\x00-\x7F]/g, '.')
  }

  if (joinSuffixRegex) {
    prefixRegex += joinSuffixRegex
  }

  const defaultAnyChar = anyCharIncludingNewlineExpr()
  const antiSlopExpr = buildAntiSlopContinuation(settings.anti_slop_ban_list)
  const anyChar = antiSlopExpr || defaultAnyChar

  let pattern = ''
  if (mustEndAfterTemplate) {
    let trailing = `[\\t \\r\\n]*`
    if (newlineToken) {
      const tokenIsAscii = /^[\x00-\x7F]*$/.test(String(newlineToken))
      if (tokenIsAscii) {
        const escapedToken = escapeRegExp(newlineToken)
        trailing = `[\\t ]*(?:${escapedToken}|\\n)?(?:[\\t ]*(?:${escapedToken}|\\n)[\\t ]*)*`
      }
    }
    pattern = `^(?:${prefixRegex})${trailing}$`
  } else if (state.patternMode === 'anthropic') {
    pattern = `^(?:${prefixRegex})${anyChar}+$`
  } else {
    pattern = `^(?:${prefixRegex})${anyChar}{${minChars},}$`
  }

  // Validate the pattern locally
  try {
    new RegExp(pattern)
  } catch (err) {
    spindle.log.warn(`Invalid injected regex pattern; falling back to minimal-safe pattern: ${err}`)
    if (mustEndAfterTemplate) {
      pattern = `^(?:${prefixRegex})[\\t \\r\\n]*$`
    } else {
      pattern =
        state.patternMode === 'anthropic'
          ? `^(?:${prefixRegex})${anyChar}+$`
          : `^(?:${prefixRegex})${anyChar}{${minChars},}$`
    }
  }

  return {
    name: 'response',
    strict: true,
    schema: {
      type: 'object',
      properties: {
        response: {
          type: 'string',
          pattern: pattern,
        },
      },
      required: ['response'],
      additionalProperties: false,
    },
  }
}

// Build a plain JSON Schema (no OpenAI wrapper) for Gemini's responseSchema and
// Anthropic's tool input_schema. Gemini/Anthropic don't honor regex `pattern`
// constraints, so we only ship the shape enforcement here.
function buildPlainJsonSchemaForPrefill(): Record<string, any> {
  return {
    type: 'object',
    properties: {
      response: {
        type: 'string',
        description: 'The model response text.',
      },
    },
    required: ['response'],
    additionalProperties: false,
  }
}

// ─── JSON Extraction & Unwrapping ────────────────────────────────────────────

function tryExtractJsonStringField(rawText: string, fieldName: string): string | null {
  if (typeof rawText !== 'string' || rawText.length === 0) return null

  const safeField = String(fieldName ?? '').replace(/[^a-zA-Z0-9_]/g, '')
  if (!safeField) return null

  const match = new RegExp(`"${safeField}"\\s*:\\s*"`, 'm').exec(rawText)
  if (!match) return null

  let index = match.index + match[0].length
  let out = ''
  let escaped = false

  while (index < rawText.length) {
    const ch = rawText[index]
    if (escaped) {
      switch (ch) {
        case '"': out += '"'; break
        case '\\': out += '\\'; break
        case '/': out += '/'; break
        case 'b': out += '\b'; break
        case 'f': out += '\f'; break
        case 'n': out += '\n'; break
        case 'r': out += '\r'; break
        case 't': out += '\t'; break
        case 'u': {
          const hex = rawText.slice(index + 1, index + 5)
          if (/^[0-9a-fA-F]{4}$/.test(hex)) {
            out += String.fromCharCode(parseInt(hex, 16))
            index += 4
          } else {
            out += '\\u'
          }
          break
        }
        default: out += '\\' + ch
      }
      escaped = false
    } else if (ch === '\\') {
      escaped = true
    } else if (ch === '"') {
      return out
    } else {
      out += ch
    }
    index++
  }

  // Unterminated string — return what we have (useful for streaming)
  return out.length > 0 ? out : null
}

function tryExtractJsonStringFieldLoose(rawText: string, fieldName: string): string | null {
  if (typeof rawText !== 'string' || rawText.length === 0) return null

  const safeField = String(fieldName ?? '').replace(/[^a-zA-Z0-9_]/g, '')
  if (!safeField) return null

  const match = new RegExp(`"${safeField}"\\s*:\\s*"`, 'm').exec(rawText)
  if (!match) return null

  let index = match.index + match[0].length
  let out = ''
  let escaped = false
  let depth = 0

  while (index < rawText.length) {
    const ch = rawText[index]
    if (escaped) {
      switch (ch) {
        case '"': out += '"'; break
        case '\\': out += '\\'; break
        case '/': out += '/'; break
        case 'b': out += '\b'; break
        case 'f': out += '\f'; break
        case 'n': out += '\n'; break
        case 'r': out += '\r'; break
        case 't': out += '\t'; break
        case 'u': {
          const hex = rawText.slice(index + 1, index + 5)
          if (/^[0-9a-fA-F]{4}$/.test(hex)) {
            out += String.fromCharCode(parseInt(hex, 16))
            index += 4
          } else {
            out += '\\u'
          }
          break
        }
        default: out += '\\' + ch
      }
      escaped = false
    } else if (ch === '\\') {
      escaped = true
    } else if (ch === '"') {
      // Loose mode: if the next non-whitespace char is } or , we're done.
      // Otherwise treat the quote as part of the content (unescaped quote in model output).
      const rest = rawText.slice(index + 1).trimStart()
      if (!rest || rest[0] === '}' || rest[0] === ',') {
        return out
      }
      out += ch
    } else {
      out += ch
    }
    index++
  }

  return out.length > 0 ? out : null
}

function tryUnwrapStructuredOutput(text: string): string | null {
  if (typeof text !== 'string' || text.length === 0) return null

  const decode = (s: string) => straightenCurlyQuotes(decodeNewlines(s, state.newlineToken))

  const applyContinueJoin = (decodedValue: string): string => {
    const s = String(decodedValue ?? '')
    if (!state.continue.active) return s
    const base = String(state.continue.baseText ?? '')
    if (!base) return s

    // If the decoded value starts with the base, it's the full message already
    if (s.length >= base.length && s.startsWith(base)) return s

    return base + s
  }

  // Try full JSON parse first
  try {
    const parsed = JSON.parse(text)
    if (parsed && typeof parsed === 'object') {
      if (typeof parsed.response === 'string') {
        const decoded = decode(parsed.response)
        return state.continue.active ? applyContinueJoin(decoded) : decoded
      }
      if (typeof parsed.value === 'string') {
        const decoded = decode(parsed.value)
        return state.continue.active ? applyContinueJoin(decoded) : decoded
      }
      if (typeof parsed.prefix === 'string' || typeof parsed.content === 'string') {
        const prefix = typeof parsed.prefix === 'string' ? decode(parsed.prefix) : ''
        const content = typeof parsed.content === 'string' ? decode(parsed.content) : ''
        const joined = prefix + content
        return state.continue.active ? applyContinueJoin(joined) : joined
      }
    }
  } catch {
    // JSON.parse failed — try loose extraction
    const looseResponse = tryExtractJsonStringFieldLoose(text, 'response')
    if (typeof looseResponse === 'string') {
      const decoded = decode(looseResponse)
      return state.continue.active ? applyContinueJoin(decoded) : decoded
    }
    const looseValue = tryExtractJsonStringFieldLoose(text, 'value')
    if (typeof looseValue === 'string') {
      const decoded = decode(looseValue)
      return state.continue.active ? applyContinueJoin(decoded) : decoded
    }
  }

  // Fallback to partial extraction (streaming)
  const response = tryExtractJsonStringField(text, 'response')
  if (typeof response === 'string') {
    const decoded = decode(response)
    return state.continue.active ? applyContinueJoin(decoded) : decoded
  }

  const legacy = tryExtractJsonStringField(text, 'value')
  if (typeof legacy === 'string') {
    const decoded = decode(legacy)
    return state.continue.active ? applyContinueJoin(decoded) : decoded
  }

  return null
}

// ─── Hide-Prefill Stripping ──────────────────────────────────────────────────

function clearHidePrefillState(): void {
  state.hidePrefillLiteral = ''
  state.hidePrefillRegex = null
}

function buildPrefillStripper(prefixTemplate: string): void {
  const { hideTemplate } = splitHidePrefillTemplate(prefixTemplate)
  if (!hideTemplate) return

  if (!prefixHasSlots(hideTemplate)) {
    state.hidePrefillLiteral = hideTemplate
    state.hidePrefillRegex = null
    return
  }

  const prefixRegex = buildPrefixRegexFromWireTemplate(hideTemplate)
  try {
    state.hidePrefillRegex = new RegExp(`^((?:${prefixRegex}))`)
  } catch (err) {
    spindle.log.warn(`Failed to build hide-prefill regex; falling back to literal: ${err}`)
    state.hidePrefillRegex = null
    state.hidePrefillLiteral = hideTemplate
  }
}

function stripHidePrefill(text: string): string {
  if (!settings.hide_prefill_in_display) return text
  if (!state.hidePrefillLiteral && !state.hidePrefillRegex) return text

  const s = String(text ?? '')

  if (state.hidePrefillRegex) {
    const m = state.hidePrefillRegex.exec(s)
    if (m) return s.slice(m[0].length)
  }

  if (state.hidePrefillLiteral && s.startsWith(state.hidePrefillLiteral)) {
    return s.slice(state.hidePrefillLiteral.length)
  }

  return s
}

// ─── Continue State ──────────────────────────────────────────────────────────

function clearContinueState(): void {
  state.continue.active = false
  state.continue.baseText = ''
  state.continue.displayBase = ''
}

// ─── Stream Guard ────────────────────────────────────────────────────────────

function resetStreamGuard(): void {
  state.streamGuard = {
    startedAt: Date.now(),
    lastRawLen: 0,
    lastDecodedLen: 0,
    lastProgressAt: Date.now(),
    suspiciousStreak: 0,
    stopRequested: false,
  }
}

function isSuspiciousPaddingDelta(delta: string): boolean {
  const s = String(delta ?? '')
  if (s.length < 120) return false

  const backslashes = (s.match(/\\/g) || []).length
  const commas = (s.match(/,/g) || []).length
  const punct = (s.match(/[\\,.;:!?'"()\[\]{}<>/~`|_-]/g) || []).length

  const punctRatio = punct / Math.max(1, s.length)
  const slashRatio = backslashes / Math.max(1, s.length)

  if (slashRatio > 0.25) return true
  if (commas > 80 && punctRatio > 0.45) return true
  if (punctRatio > 0.6) return true
  if (!s.includes(' ') && s.length > 260 && /[A-Za-z]/.test(s) && commas > 40) return true
  if (/^(.)\1{250,}$/s.test(s)) return true

  return false
}

function maybeAbortOnStreamLoop(rawText: string, decodedFullText: string | null): void {
  if (!state.active) return
  if (state.streamGuard.stopRequested) return

  const rawLen = String(rawText ?? '').length
  const decodedLen = String(decodedFullText ?? '').length
  const now = Date.now()
  const guard = state.streamGuard

  if (decodedLen !== guard.lastDecodedLen) {
    const prevLen = guard.lastDecodedLen
    guard.lastDecodedLen = decodedLen
    guard.lastProgressAt = now

    if (decodedLen > prevLen) {
      const delta = String(decodedFullText ?? '').slice(prevLen)
      if (isSuspiciousPaddingDelta(delta)) {
        guard.suspiciousStreak++
      } else {
        guard.suspiciousStreak = Math.max(0, guard.suspiciousStreak - 1)
      }
    }
  }

  if (rawLen > guard.lastRawLen) guard.lastRawLen = rawLen

  const sinceStart = now - (guard.startedAt || now)
  const sinceProgress = now - (guard.lastProgressAt || now)

  if (sinceStart > 5000 && sinceProgress > 15000 && rawLen > 5000) {
    guard.stopRequested = true
    spindle.toast.error('Detected stalled/looping structured output. Consider stopping generation.', {
      title: EXT_NAME,
      duration: 9000,
    })
    return
  }

  if (sinceStart > 1500 && guard.suspiciousStreak >= 4 && rawLen > 1500) {
    guard.stopRequested = true
    spindle.toast.error('Detected runaway padding in structured output. Consider stopping generation.', {
      title: EXT_NAME,
      duration: 9000,
    })
  }
}

// ─── Prefill Generator ([[pg]]) ──────────────────────────────────────────────

function normalizePrefillGenExtraPromptRole(raw: string): 'system' | 'user' | 'assistant' {
  const role = String(raw ?? 'system').trim().toLowerCase()
  if (role === 'user' || role === 'assistant') return role
  return 'system'
}

async function runPrefillGenerator(messages: LlmMessageDTO[], tailIndex: number): Promise<string> {
  if (!settings.prefill_gen_enabled || !settings.prefill_gen_connection_id) return ''

  const baseMessages = messages
    .filter((_, i) => i !== tailIndex)
    .map(m => ({ ...m }))

  if (baseMessages.length === 0) {
    throw new Error('Prefill generator: no messages to generate from')
  }

  // Some providers reject trailing assistant messages
  const last = baseMessages[baseMessages.length - 1]
  if (last?.role === 'assistant') {
    baseMessages[baseMessages.length - 1] = { ...last, role: 'user' as const }
  }

  // Insert extra prompt if configured
  const extraPrompt = String(settings.prefill_gen_extra_prompt ?? '').trim()
  const extraRole = normalizePrefillGenExtraPromptRole(settings.prefill_gen_extra_prompt_role)
  if (extraPrompt) {
    if (extraRole === 'system') {
      let insertAt = 0
      while (insertAt < baseMessages.length && baseMessages[insertAt]?.role === 'system') insertAt++
      baseMessages.splice(insertAt, 0, { role: 'system', content: extraPrompt })
    } else {
      baseMessages.push({ role: extraRole, content: extraPrompt })
    }
  }

  const stopStrings = String(settings.prefill_gen_stop ?? '')
    .split(/\r?\n/g)
    .map(s => s.trim())
    .filter(Boolean)

  try {
    const result = await (spindle.generate.raw as any)({
      messages: baseMessages,
      parameters: {
        max_tokens: clampInt(settings.prefill_gen_max_tokens, 1, 20000, 15),
        temperature: 1,
        top_p: 1,
        stream: false,
        ...(stopStrings.length > 0 ? { stop: stopStrings } : {}),
      },
      connection_id: settings.prefill_gen_connection_id,
      userId: currentUserId,
      user_id: currentUserId,
    })

    return String(result?.content ?? '')
  } catch (err: any) {
    spindle.toast.error(String(err?.message ?? 'Prefill generator failed'), {
      title: EXT_NAME,
      duration: 9000,
    })
    return ''
  }
}

// ─── Interceptor (Core Hook) ─────────────────────────────────────────────────

spindle.registerInterceptor(async (messages: LlmMessageDTO[], context: any) => {
  // Ensure settings are loaded (interceptor can fire before frontend connects)
  if (!settingsLoaded) await loadSettings()
  
  if (!settings.enabled) return messages

  // Extract generation metadata from context
  // Context fields confirmed: chatId, connectionId, personaId, generationType, activatedWorldInfo
  const generationType = String(context?.generationType ?? '').toLowerCase()
  const connectionId = String(context?.connectionId ?? '')
  const chatId = String(context?.chatId ?? '')

  // Diagnostic: log context on every interceptor call
  spindle.log.info(`[SP] Interceptor fired: type=${generationType} connectionId=${connectionId} chatId=${chatId || '(empty)'}`)

  // Skip impersonate and quiet generations
  if (generationType === 'impersonate' || generationType === 'quiet') return messages

  // Resolve provider and model from the connection profile
  let provider = ''
  let modelId = ''
  if (connectionId) {
    try {
      const conn = await spindle.connections.get(connectionId, currentUserId)
      if (conn) {
        provider = String(conn.provider ?? '').toLowerCase()
        modelId = String(conn.model ?? '')
        spindle.log.info(`[SP] Resolved connection: provider=${provider} model=${modelId}`)
      }
    } catch (err) {
      spindle.log.warn(`[SP] Could not resolve connection ${connectionId}: ${err}`)
    }
  }

  // Resolve provider mode — decides how we inject structured output
  const providerMode = getProviderMode(provider, modelId)
  if (providerMode === 'unsupported') {
    spindle.log.info(`[SP] Provider "${provider}" not compatible with structured prefill, skipping`)
    return messages
  }

  const isContinue = generationType === 'continue'

  if (!Array.isArray(messages) || messages.length === 0) return messages

  state.patternMode = getPatternModeForRequest(providerMode)

  // Collect known character names for [[name]] placeholder
  // Lumiverse context doesn't provide user/char names directly,
  // but we can try to get them from the persona and chat
  const names = new Set<string>()
  // TODO: Could resolve names from spindle.personas.getActive() and chat character
  // For now, [[name]] slot will match any text
  state.knownNames = [...names]

  // Find the tail assistant message (prefill)
  let tailIndex = messages.length - 1
  while (tailIndex >= 0 && messages[tailIndex]?.role === 'system') tailIndex--
  const tail = tailIndex >= 0 ? messages[tailIndex] : null

  if (!tail || tail.role !== 'assistant' || typeof tail.content !== 'string') return messages
  let tailContent = tail.content
  if (!tailContent) return messages

  // Strip character name prefix (e.g. "CharName: ") from the tail
  const sortedNames = [...state.knownNames].sort((a, b) => b.length - a.length)
  for (const name of sortedNames) {
    if (name && tailContent.startsWith(name + ': ')) {
      tailContent = tailContent.slice(name.length + 2)
      break
    }
  }

  resetStreamGuard()
  clearHidePrefillState()
  clearContinueState()

  let schemaPrefix = ''
  let joinSuffixRegex = ''
  let mustEndAfterTemplate = false
  let prefillTemplate = String(tailContent)

  // Make a mutable copy of messages
  const modifiedMessages = messages.map(m => ({ ...m }))

  if (isContinue) {
    // Continue mode — simplified for initial port
    // In full implementation, would read base text from chat via chat_mutation
    state.continue.active = true
    state.continue.baseText = tailContent
    schemaPrefix = ''
    state.newlineToken = chooseNewlineToken(schemaPrefix || tailContent, settings.newline_token)
  } else {
    // Normal generation — replace assistant prefill with structured output constraint

    // Handle [[pg]] prefill generator
    if (templateHasPrefillGenSlot(prefillTemplate)) {
      try {
        const generated = await runPrefillGenerator(modifiedMessages, tailIndex)
        prefillTemplate = prefillTemplate.replace(/\[\[\s*pg\s*\]\]/gi, generated)
      } catch (err) {
        spindle.log.warn(`Prefill generator failed: ${err}`)
        prefillTemplate = prefillTemplate.replace(/\[\[\s*pg\s*\]\]/gi, '')
      }
    }

    // Remove the assistant prefill message
    modifiedMessages.splice(tailIndex, 1)

    // Clean up legacy slot markers
    prefillTemplate = prefillTemplate.replace(/\[\[\s*sp\s*:[^\]]*\]\]/gi, '')

    // Convert literal quotes to curly quotes for JSON robustness
    prefillTemplate = curlyQuoteLiteralsOutsideSlots(prefillTemplate)

    // Handle [[end]] marker
    const endSplit = splitEndPrefillTemplate(prefillTemplate)
    prefillTemplate = endSplit.template
    mustEndAfterTemplate = endSplit.hasEndMarker

    schemaPrefix = prefillTemplate
    state.newlineToken = chooseNewlineToken(schemaPrefix, settings.newline_token)

    if (settings.hide_prefill_in_display) {
      buildPrefillStripper(straightenCurlyQuotes(schemaPrefix))
    }
  }

  // Safety: don't send if the last message is assistant-role (many providers reject this with output formats)
  if (modifiedMessages.length > 0) {
    const lastMsg = modifiedMessages[modifiedMessages.length - 1]
    if (!isContinue && lastMsg?.role === 'assistant') return messages // bail
    if (state.patternMode === 'anthropic' && lastMsg?.role !== 'user' && lastMsg?.role !== 'system') return messages
  }

  const minCharsSetting = clampInt(settings.min_chars_after_prefix, 1, 10000, 80)
  const minCharsAfterPrefix = isContinue ? 1 : (mustEndAfterTemplate ? 0 : minCharsSetting)
  const jsonSchema = buildJsonSchemaForPrefillValuePattern(schemaPrefix, minCharsAfterPrefix, joinSuffixRegex, { mustEndAfterTemplate })

  // Activate runtime state
  state.active = true
  state.lastInjectedAt = Date.now()
  state.expectedPrefill = straightenCurlyQuotes(String(schemaPrefix ?? ''))
  state.accumulatedStreamText = ''
  state.lastAppliedText = ''
  state.activeChatId = chatId
  state.activeGenerationId = ''
  state.activeConnectionId = connectionId
  state.retryMessages = modifiedMessages.map(m => ({ ...m }))

  // Attach a per-chat stream observer to decode the JSON response as it
  // streams in, apply the final text via chat_mutation, and clean up.
  attachStreamObserver(chatId)

  spindle.log.info(`[SP] Injecting structured prefill: provider=${provider} model=${modelId} providerMode=${providerMode}`)

  try {
    const injectedPattern = jsonSchema?.schema?.properties?.response?.pattern
    if (typeof injectedPattern === 'string' && injectedPattern.length > 0) {
      spindle.log.info(`[SP] Schema pattern (${injectedPattern.length} chars): ${injectedPattern.slice(0, 200)}${injectedPattern.length > 200 ? '...' : ''}`)
    }
  } catch { /* ignore */ }

  // Build provider-specific parameters.
  // OpenAI:    response_format: { type: 'json_schema', json_schema: {...} }  ← regex-enforced
  // OpenAI*:   response_format: { type: 'json_object' } + system nudge        ← fallback for proxies
  // Gemini:    responseMimeType + responseSchema                              ← shape-only
  // Anthropic: tools + tool_choice forced tool call                           ← shape-only
  //
  // For OpenAI-compatible providers, if we've previously detected that this
  // connection's proxy rejects json_schema (returned empty content), we skip
  // straight to json_object mode. Otherwise we try json_schema first and let
  // the onEnd handler detect failures and auto-retry.
  let parameters: Record<string, unknown>

  if (providerMode === 'anthropic') {
    state.activeInjectionMode = 'anthropic_tool'
    const toolSchema = buildPlainJsonSchemaForPrefill()
    parameters = {
      tools: [{
        name: 'prefill_response',
        description: 'Return the model response text as a structured tool call.',
        input_schema: toolSchema,
      }],
      tool_choice: { type: 'tool', name: 'prefill_response' },
    }
  } else if (providerMode === 'gemini') {
    state.activeInjectionMode = 'gemini'
    parameters = {
      responseMimeType: 'application/json',
      responseSchema: buildPlainJsonSchemaForPrefill(),
    }
  } else {
    // openai and openai-compatible — pick a tier from the compat cache.
    // Tier 1: json_schema (regex-locked)
    // Tier 2: json_object (shape only)
    // Tier 3: prompt_only (no response_format; system nudge only)
    const tier = getCompatTier(connectionId)
    state.activeInjectionMode = tier

    if (tier === 'json_schema') {
      parameters = {
        response_format: {
          type: 'json_schema',
          json_schema: jsonSchema,
        },
      }
    } else if (tier === 'json_object') {
      spindle.log.info(`[SP] Using cached tier json_object for ${connectionId}`)
      parameters = {
        response_format: { type: 'json_object' },
      }
      // Nudge the model since json_object has no schema
      modifiedMessages.unshift({
        role: 'system',
        content: 'You MUST respond with a JSON object of the form {"response": "your full reply as a single string"}. No other keys. No prose outside the JSON.',
      })
    } else {
      // prompt_only — no response_format at all, just a strong system nudge.
      // This works on proxies that reject response_format entirely.
      spindle.log.info(`[SP] Using cached tier prompt_only for ${connectionId}`)
      parameters = {}
      modifiedMessages.unshift({
        role: 'system',
        content: 'CRITICAL: Your entire response must be a single JSON object of the form {"response":"your full reply as a single string, with all newlines escaped as \\n"}. Do not include any text before or after the JSON. Do not use markdown code fences. Start your reply with the literal character { and end with }.',
      })
    }
  }

  // Notify frontend that structured prefill is active
  spindle.sendToFrontend({
    type: 'sp_activated',
    chatId,
    hidePrefill: settings.hide_prefill_in_display,
    providerMode,
    injectionMode: state.activeInjectionMode,
  }, currentUserId)

  // Return InterceptorResultDTO — messages + parameters.
  // Requires the "generation_parameters" permission; without it parameters are
  // silently stripped and the extension degrades to a pure message-only pass.
  spindle.log.info(`[SP] Returning InterceptorResultDTO with injection mode=${state.activeInjectionMode}, parameter keys=[${Object.keys(parameters).join(', ')}]`)
  return {
    messages: modifiedMessages,
    parameters,
  }
}, 50) // High priority — run early

// ─── Stream Observation (per-chat) ───────────────────────────────────────────
//
// The old approach used three global event listeners (STREAM_TOKEN_RECEIVED,
// GENERATION_ENDED, GENERATION_STOPPED) with manual chatId filtering. The new
// Lumiverse API gives us `spindle.generate.observe(chatId)` which:
//   - automatically filters events to a single chat
//   - accumulates streamed content into observer.content
//   - exposes onToken/onEnd/onStop/onStart with typed payloads
//   - must be .dispose()d when done to free resources
//
// One observer per chatId is kept in a registry so we can dispose the previous
// observer if a new generation starts on the same chat before the last one
// fully cleans up.

type ActiveObserver = ReturnType<typeof spindle.generate.observe>
const activeObservers: Map<string, ActiveObserver> = new Map()

function disposeObserverFor(chatId: string): void {
  const existing = activeObservers.get(chatId)
  if (existing) {
    try { existing.dispose() } catch { /* ignore */ }
    activeObservers.delete(chatId)
  }
}

function attachStreamObserver(chatId: string): void {
  if (!chatId) return
  disposeObserverFor(chatId)

  const observer = spindle.generate.observe(chatId)
  activeObservers.set(chatId, observer)

  observer.onStart((info: any) => {
    if (info?.generationId) state.activeGenerationId = String(info.generationId)
  })

  observer.onToken((tokenPayload: any) => {
    if (!state.active || !settings.enabled) return
    // Reasoning tokens (chain-of-thought) aren't part of the structured output
    // body — skip them so they don't poison the JSON accumulator.
    if ((tokenPayload as any)?.type === 'reasoning') return

    // Diagnostic: confirm at least one content token arrived
    if (state.accumulatedStreamText.length === 0) {
      const first = String(tokenPayload?.token ?? '').slice(0, 80)
      spindle.log.info(`[SP] First stream token received: ${JSON.stringify(first)}`)
    }

    state.accumulatedStreamText += String(tokenPayload?.token ?? '')
    const rawText = state.accumulatedStreamText
    const unwrapped = tryUnwrapStructuredOutput(rawText)

    if (typeof unwrapped === 'string') {
      const displayText = stripHidePrefill(unwrapped)
      state.lastAppliedText = unwrapped

      spindle.sendToFrontend({
        type: 'sp_stream_update',
        chatId: state.activeChatId,
        decodedText: displayText,
        rawIsJson: true,
      }, currentUserId)
    }

    maybeAbortOnStreamLoop(rawText, unwrapped)
  })

  observer.onEnd(async (result: any) => {
    try {
      if (!state.active) return
      if (result?.error) {
        spindle.log.warn(`[SP] Generation ended with error: ${result.error}`)
      }

      // Prefer the observer's accumulated content (more reliable than the
      // event payload's content field for some providers).
      let rawText = String(result?.content ?? observer.content ?? state.accumulatedStreamText ?? '')

      // ─── Diagnostic: what did the provider actually return? ───
      const preview = rawText.length > 300 ? rawText.slice(0, 300) + '...' : rawText
      spindle.log.info(`[SP] Raw response (${rawText.length} chars): ${JSON.stringify(preview)}`)
      spindle.log.info(`[SP] Tokens accumulated during stream: ${state.accumulatedStreamText.length} chars`)
      spindle.log.info(`[SP] Observer.content length: ${String(observer.content ?? '').length} chars`)
      spindle.log.info(`[SP] result.content length: ${String(result?.content ?? '').length} chars`)

      // ─── Auto-fallback: climb the compatibility tier ladder ───
      // Failure mode 1: empty content (proxy rejected the request outright)
      // Failure mode 2: content came through but isn't our {"response":"..."}
      //                 shape — means the proxy silently stripped the parameter
      //                 and the model just chatted normally
      //
      // Prompt-only tier is treated as the terminal tier: if we're on it and
      // the model returned plain prose instead of JSON, we accept the plain
      // prose (the proxy stripped no parameter; it's the model not following
      // the system-nudge). No more tiers to try.
      const isOpenAITier =
        state.activeInjectionMode === 'json_schema' ||
        state.activeInjectionMode === 'json_object' ||
        state.activeInjectionMode === 'prompt_only'

      let tierAtStart: CompatTier | null = null
      if (isOpenAITier) tierAtStart = state.activeInjectionMode as CompatTier

      function looksLikeOurJson(text: string): boolean {
        if (!text) return false
        // Quick shape check — does it look like {"response":"..."}?
        // We use the unwrapper since it handles partial/loose JSON too.
        const out = tryUnwrapStructuredOutput(text)
        return typeof out === 'string'
      }

      const initialLooksRight = looksLikeOurJson(rawText)
      const initialIsFailure = rawText.length === 0 || !initialLooksRight

      const shouldRetry =
        !!tierAtStart &&
        initialIsFailure &&
        !!state.activeConnectionId &&
        tierAtStart !== 'prompt_only'   // don't retry past the terminal tier

      if (initialIsFailure && !shouldRetry) {
        if (rawText.length > 0 && !initialLooksRight) {
          spindle.log.warn(`[SP] Got ${rawText.length} chars of plain text on prompt_only tier — applying raw (the model didn't follow the JSON instruction)`)
        }
      }

      if (shouldRetry && tierAtStart) {
        if (rawText.length > 0) {
          spindle.log.warn(`[SP] Got ${rawText.length} chars but not in our JSON shape — proxy likely stripped ${tierAtStart} parameter. Falling through the ladder.`)
        }

        let currentTier: CompatTier | null = tierAtStart
        let toastShown = false

        while (currentTier) {
          const nt = nextTier(currentTier)
          if (!nt) break
          currentTier = nt

          spindle.log.warn(`[SP] Trying fallback tier: ${currentTier} on connection ${state.activeConnectionId}`)

          if (!toastShown) {
            try {
              spindle.toast.warning(
                `This connection's ${tierAtStart} mode isn't working — falling back to ${currentTier}. Future messages will use the working mode automatically.`,
                { title: EXT_NAME, duration: 9000 }
              )
            } catch { /* toast best-effort */ }
            toastShown = true
          }

          // Build parameters + messages for this tier
          const retryMessages = state.retryMessages.slice()
          let retryParams: Record<string, unknown> = {}

          if (currentTier === 'json_object') {
            retryParams = { response_format: { type: 'json_object' } }
            retryMessages.unshift({
              role: 'system',
              content: 'You MUST respond with a JSON object of the form {"response": "your full reply as a single string"}. No other keys. No prose outside the JSON.',
            })
          } else if (currentTier === 'prompt_only') {
            retryParams = {}
            retryMessages.unshift({
              role: 'system',
              content: 'CRITICAL: Your entire response must be a single JSON object of the form {"response":"your full reply as a single string, with all newlines escaped as \\n"}. Do not include any text before or after the JSON. Do not use markdown code fences. Start your reply with the literal character { and end with }.',
            })
          }

          try {
            spindle.log.info(`[SP] Retry (${currentTier}) sending with currentUserId=${String(currentUserId ?? '(empty)')}`)

            // Resolve the connection's model and provider so we can pass them
            // explicitly — `generate.raw` doesn't auto-resolve from connection_id
            // the way the main generation pipeline does.
            let retryProvider = ''
            let retryModel = ''
            try {
              const conn = await spindle.connections.get(state.activeConnectionId, currentUserId)
              if (conn) {
                retryProvider = String(conn.provider ?? '')
                retryModel = String(conn.model ?? '')
                spindle.log.info(`[SP] Retry resolved model=${retryModel} provider=${retryProvider}`)
              }
            } catch (err) {
              spindle.log.warn(`[SP] Retry could not resolve connection: ${err}`)
            }

            const retryRequest: Record<string, unknown> = {
              messages: retryMessages as any,
              parameters: retryParams,
              connection_id: state.activeConnectionId,
              userId: currentUserId,
              user_id: currentUserId,
            }
            // Cover both common field names — runtime ignores the unused one
            if (retryModel) {
              retryRequest.model = retryModel
              ;(retryRequest.parameters as any).model = retryModel
            }
            if (retryProvider) {
              retryRequest.provider = retryProvider
            }

            const retryResult = await (spindle.generate.raw as any)(retryRequest)
            const retryText = String((retryResult as any)?.content ?? '')
            const retryLooksRight = looksLikeOurJson(retryText)
            spindle.log.info(`[SP] Retry (${currentTier}) returned ${retryText.length} chars, parses as JSON: ${retryLooksRight}`)

            if (retryText.length > 0 && (retryLooksRight || currentTier === 'prompt_only')) {
              // Success — take this result. prompt_only accepts plain prose
              // as best-effort since it has no structured-output guarantee.
              rawText = retryText
              await setCompatTier(state.activeConnectionId, currentTier)
              break
            }
          } catch (err) {
            spindle.log.warn(`[SP] Retry (${currentTier}) failed: ${err}`)
          }
        }

        if (rawText.length === 0) {
          spindle.log.warn(`[SP] All compatibility tiers exhausted — connection may be broken`)
          if (currentTier) await setCompatTier(state.activeConnectionId, currentTier)
        }
      }

      const unwrapped = tryUnwrapStructuredOutput(rawText)
      if (typeof unwrapped === 'string') {
        spindle.log.info(`[SP] JSON unwrap succeeded: extracted ${unwrapped.length} chars from response field`)
      } else {
        spindle.log.warn(`[SP] JSON unwrap FAILED — raw content does not parse as {"response": "..."} — falling back to raw text`)
      }

      const finalText = typeof unwrapped === 'string' ? unwrapped : rawText
      const displayText = stripHidePrefill(finalText)
      state.lastAppliedText = finalText

      spindle.log.info(`[SP] Final decoded text: ${finalText.length} chars, display: ${displayText.length} chars`)

      if (result?.messageId && state.activeChatId) {
        try {
          const granted = await spindle.permissions.getGranted()
          if (granted.includes('chat_mutation')) {
            // First update the stored content (cheap — in case delete/append fails we still
            // have the decoded text on disk).
            try {
              await (spindle as any).chat.updateMessage(state.activeChatId, result.messageId, { content: finalText })
              spindle.log.info(`[SP] Stored decoded text on message ${result.messageId}`)
            } catch (err) {
              spindle.log.warn(`[SP] updateMessage failed (continuing to delete/replace): ${err}`)
            }

            // Now delete + re-append so Lumiverse re-renders from scratch instead of
            // showing the cached raw JSON stream. This fixes the "JSON wrapper still
            // visible in chat" issue on prompt_only tier where Lumiverse doesn't know
            // the stream was JSON-wrapped.
            try {
              await (spindle as any).chat.deleteMessage(state.activeChatId, result.messageId)
              const appended = await (spindle as any).chat.appendMessage(state.activeChatId, {
                role: 'assistant',
                content: finalText,
                metadata: { source: 'structured_prefill_rerender', originalMessageId: result.messageId },
              })
              spindle.log.info(`[SP] Replaced message ${result.messageId} with fresh render (new id: ${appended?.id ?? '?'})`)
            } catch (err) {
              spindle.log.warn(`[SP] Delete+append failed — message kept with raw content: ${err}`)
            }
          } else {
            spindle.log.info(`[SP] chat_mutation not granted — decoded text sent to frontend only`)
          }
        } catch (err) {
          spindle.log.warn(`[SP] Failed to apply decoded text: ${err}`)
        }
      }

      spindle.sendToFrontend({
        type: 'sp_generation_complete',
        chatId: state.activeChatId,
        decodedText: displayText,
        fullText: finalText,
      }, currentUserId)
    } finally {
      state.active = false
      clearHidePrefillState()
      clearContinueState()
      resetStreamGuard()
      state.accumulatedStreamText = ''
      state.retryMessages = []
      disposeObserverFor(chatId)
    }
  })

  observer.onStop((result: any) => {
    try {
      if (!state.active) return
      const rawText = String(result?.content ?? observer.content ?? state.accumulatedStreamText ?? '')
      const unwrapped = tryUnwrapStructuredOutput(rawText)
      const finalText = typeof unwrapped === 'string' ? unwrapped : state.lastAppliedText || rawText
      const displayText = stripHidePrefill(finalText)

      spindle.sendToFrontend({
        type: 'sp_generation_complete',
        chatId: state.activeChatId,
        decodedText: displayText,
        fullText: finalText,
      }, currentUserId)
    } finally {
      state.active = false
      clearHidePrefillState()
      clearContinueState()
      resetStreamGuard()
      state.accumulatedStreamText = ''
      state.retryMessages = []
      disposeObserverFor(chatId)
    }
  })
}

// ─── Frontend Message Handler ────────────────────────────────────────────────

spindle.onFrontendMessage(async (payload: any, userId) => {
  // Capture userId for operator-scoped extensions (lesson from Mode Toggles port)
  if (userId && userId !== currentUserId) {
    spindle.log.info(`[SP] Captured userId from frontend: ${userId}`)
    currentUserId = userId
  }

  // Lazy-load settings on first message so userId is known (pattern from LumiScript)
  if (!settingsLoaded) await loadSettings()

  switch (payload?.type) {
    case 'get_settings': {
      spindle.sendToFrontend({ type: 'settings', settings }, userId)
      break
    }

    case 'update_settings': {
      const updates = payload.settings
      if (updates && typeof updates === 'object') {
        Object.assign(settings, updates)
        await saveSettings()
        spindle.sendToFrontend({ type: 'settings', settings }, userId)
      }
      break
    }

    case 'get_connections': {
      try {
        // Pass userId for operator-scoped extensions
        const connections = await spindle.connections.list(currentUserId)
        spindle.sendToFrontend({ type: 'connections', connections }, userId)
      } catch (err) {
        spindle.log.warn(`[SP] Failed to list connections: ${err}`)
        spindle.sendToFrontend({ type: 'connections', connections: [] }, userId)
      }
      break
    }

    case 'get_state': {
      const compatEntries: Record<string, CompatTier> = {}
      for (const [k, v] of compatCache.entries()) compatEntries[k] = v
      spindle.sendToFrontend({
        type: 'state',
        active: state.active,
        patternMode: state.patternMode,
        compatCache: compatEntries,
      }, userId)
      break
    }

    case 'clear_blocklist': {
      compatCache.clear()
      try {
        await spindle.userStorage.setJson(BLOCKLIST_FILE, {}, { indent: 2, userId: currentUserId })
        spindle.log.info(`[SP] Cleared connection compatibility cache`)
      } catch (err) {
        spindle.log.warn(`[SP] Could not clear cache: ${err}`)
      }
      spindle.sendToFrontend({ type: 'blocklist_cleared' }, userId)
      break
    }

    case 'remove_from_blocklist': {
      const connId = String(payload?.connectionId ?? '')
      if (connId && compatCache.has(connId)) {
        compatCache.delete(connId)
        await persistCompatCache()
        spindle.log.info(`[SP] Removed connection ${connId} from compat cache`)
      }
      const compatEntries: Record<string, CompatTier> = {}
      for (const [k, v] of compatCache.entries()) compatEntries[k] = v
      spindle.sendToFrontend({ type: 'blocklist', compatCache: compatEntries }, userId)
      break
    }
  }
})

// ─── Initialization ──────────────────────────────────────────────────────────

;(async () => {
  // Settings are loaded lazily in onFrontendMessage once userId is known.
  // This avoids read/write path mismatches for operator-scoped extensions.
  
  // Log available APIs for debugging
  try {
    const granted = await spindle.permissions.getGranted()
    spindle.log.info(`[SP] Granted permissions: [${granted.join(', ')}]`)
  } catch (e) {
    spindle.log.warn(`[SP] Could not check permissions: ${e}`)
  }

  spindle.log.info(`[SP] ${EXT_NAME} backend ready`)
})()
