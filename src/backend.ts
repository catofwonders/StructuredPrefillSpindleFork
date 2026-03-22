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

// ─── Settings I/O ────────────────────────────────────────────────────────────

async function loadSettings(): Promise<void> {
  try {
    settings = await spindle.storage.getJson<Settings>('settings.json', {
      fallback: { ...DEFAULT_SETTINGS },
    })
    // Merge missing keys from defaults
    for (const [key, value] of Object.entries(DEFAULT_SETTINGS)) {
      if ((settings as any)[key] == null) {
        ;(settings as any)[key] = value
      }
    }
  } catch {
    settings = { ...DEFAULT_SETTINGS }
  }
}

async function saveSettings(): Promise<void> {
  try {
    await spindle.storage.setJson('settings.json', settings, { indent: 2 })
  } catch (err) {
    spindle.log.error(`Failed to save settings: ${err}`)
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

function supportsStructuredPrefillForSource(source: string): boolean {
  const src = String(source ?? '').toLowerCase()
  const incompatible = new Set([
    'claude', 'anthropic', // Tool-based, not OpenAI response format
    'ai21', 'deepseek', 'moonshot', 'zai', 'siliconflow', // Map json_schema to JSON mode
    '',
  ])
  return !incompatible.has(src)
}

function getPatternModeForRequest(source: string, modelId: string): 'default' | 'anthropic' {
  const src = String(source ?? '').toLowerCase()
  const model = String(modelId ?? '').toLowerCase()

  // OpenRouter routes to Anthropic models — use Anthropic-safe regex
  if (src === 'openrouter' && (model.includes('claude') || model.includes('anthropic'))) {
    return 'anthropic'
  }
  if (src === 'claude' || src === 'anthropic') return 'anthropic'
  return 'default'
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
    description: '',
    strict: true,
    value: {
      type: 'object',
      properties: {
        response: {
          type: 'string',
          description: '',
          pattern: pattern,
        },
      },
      required: ['response'],
      additionalProperties: false,
    },
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
    const result = await spindle.generate.raw({
      messages: baseMessages,
      parameters: {
        max_tokens: clampInt(settings.prefill_gen_max_tokens, 1, 20000, 15),
        temperature: 1,
        top_p: 1,
        stream: false,
        ...(stopStrings.length > 0 ? { stop: stopStrings } : {}),
      },
      connection_id: settings.prefill_gen_connection_id,
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
  if (!settings.enabled) return messages

  // Extract generation metadata from context — try multiple possible property names
  // (Lumiverse may use different names than what we expect)
  const generationType = String(context?.generationType ?? context?.type ?? context?.generation_type ?? '').toLowerCase()
  const provider = String(context?.provider ?? context?.chat_completion_source ?? context?.source ?? '').toLowerCase()
  const modelId = String(context?.model ?? context?.modelId ?? '')
  const chatId = String(context?.chatId ?? context?.chat_id ?? '')

  // Diagnostic: log what we got from context so we can debug if things aren't working
  spindle.log.info(`[SP] Interceptor fired: type=${generationType} provider=${provider} model=${modelId} chatId=${chatId || '(empty)'}`)
  
  // Log context keys on first run so we know what's available
  if (context && typeof context === 'object') {
    try {
      const keys = Object.keys(context).join(', ')
      spindle.log.info(`[SP] Context keys: ${keys}`)
    } catch { /* ignore */ }
  }

  // Skip impersonate and quiet generations
  if (generationType === 'impersonate' || generationType === 'quiet') return messages

  // Check provider compatibility
  if (!supportsStructuredPrefillForSource(provider)) return messages

  // Don't conflict with existing json_schema or tools
  if (context?.json_schema || context?.response_format) return messages
  if (Array.isArray(context?.tools) && context.tools.length > 0) return messages

  const isContinue = generationType === 'continue'

  if (!Array.isArray(messages) || messages.length === 0) return messages

  state.patternMode = getPatternModeForRequest(provider, modelId)

  // Collect known character names for [[name]] placeholder
  const names = new Set<string>()
  const userName = String(context?.user_name ?? '').trim()
  const charName = String(context?.char_name ?? '').trim()
  if (userName) names.add(userName)
  if (charName) names.add(charName)
  if (Array.isArray(context?.group_names)) {
    for (const gn of context.group_names) {
      const n = String(gn ?? '').trim()
      if (n) names.add(n)
    }
  }
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

  // ─── Inject json_schema into the generation context ───
  // We try multiple strategies since we're not sure which property Lumiverse checks.
  // The Generation docs show `parameters.response_format` works for spindle.generate.raw(),
  // so the interceptor context likely uses a similar structure.
  if (context && typeof context === 'object') {
    // Strategy 1: OpenAI-compatible response_format (most likely)
    context.response_format = {
      type: 'json_schema',
      json_schema: jsonSchema,
    }
    // Strategy 2: Top-level json_schema (SillyTavern style)
    context.json_schema = jsonSchema
    // Strategy 3: Nested in parameters (mirrors spindle.generate.raw behavior)
    if (!context.parameters) context.parameters = {}
    context.parameters.response_format = {
      type: 'json_schema',
      json_schema: jsonSchema,
    }
    
    spindle.log.info(`[SP] Injected json_schema via context.response_format, context.json_schema, and context.parameters.response_format`)
  } else {
    spindle.log.warn(`[SP] Context is not an object — cannot inject json_schema!`)
    return messages // bail, can't do anything without context mutation
  }

  // Activate runtime state
  state.active = true
  state.lastInjectedAt = Date.now()
  state.expectedPrefill = straightenCurlyQuotes(String(schemaPrefix ?? ''))
  state.accumulatedStreamText = ''
  state.lastAppliedText = ''
  state.activeChatId = chatId
  state.activeGenerationId = ''

  spindle.log.info(`Injecting structured prefill: provider=${provider} model=${modelId} mode=${state.patternMode}`)

  try {
    const injectedPattern = jsonSchema?.value?.properties?.response?.pattern
    if (typeof injectedPattern === 'string' && injectedPattern.length > 0) {
      spindle.log.info(`Schema pattern (${injectedPattern.length} chars): ${injectedPattern.slice(0, 200)}${injectedPattern.length > 200 ? '...' : ''}`)
    }
  } catch { /* ignore */ }

  // Notify frontend that structured prefill is active
  spindle.sendToFrontend({
    type: 'sp_activated',
    chatId,
    hidePrefill: settings.hide_prefill_in_display,
  })

  return modifiedMessages
}, 50) // High priority — run early

// ─── Stream Token Handler ────────────────────────────────────────────────────

spindle.on('STREAM_TOKEN_RECEIVED', (payload) => {
  if (!state.active || !settings.enabled) return

  const { generationId, chatId, token } = payload ?? {}
  if (chatId && state.activeChatId && chatId !== state.activeChatId) return

  if (generationId) state.activeGenerationId = String(generationId)

  // Accumulate the stream
  state.accumulatedStreamText += String(token ?? '')

  const rawText = state.accumulatedStreamText
  const unwrapped = tryUnwrapStructuredOutput(rawText)

  if (typeof unwrapped === 'string') {
    const displayText = stripHidePrefill(unwrapped)
    state.lastAppliedText = unwrapped

    // Send decoded text to frontend for display
    spindle.sendToFrontend({
      type: 'sp_stream_update',
      chatId: state.activeChatId,
      decodedText: displayText,
      rawIsJson: true,
    })
  }

  // Run the stream guard
  maybeAbortOnStreamLoop(rawText, unwrapped)
})

// ─── Generation Ended Handler ────────────────────────────────────────────────

spindle.on('GENERATION_ENDED', async (payload) => {
  if (!state.active) return

  const { chatId, messageId, content } = payload ?? {}
  if (chatId && state.activeChatId && chatId !== state.activeChatId) return

  // Final unwrap
  const rawText = String(content ?? state.accumulatedStreamText ?? '')
  const unwrapped = tryUnwrapStructuredOutput(rawText)
  const finalText = typeof unwrapped === 'string' ? unwrapped : rawText

  const displayText = stripHidePrefill(finalText)
  state.lastAppliedText = finalText

  // Apply final decoded text to the chat message via chat_mutation
  if (messageId) {
    try {
      const granted = await spindle.permissions.getGranted()
      if (granted.includes('chat_mutation')) {
        // Try different possible API shapes for chat mutation
        const chatMut = (spindle as any).chatMutation ?? (spindle as any).chat_mutation ?? (spindle as any).messages
        if (chatMut?.update) {
          await chatMut.update(messageId, { content: finalText })
          spindle.log.info(`[SP] Applied decoded text to message ${messageId} via chatMutation.update`)
        } else if (chatMut?.updateMessage) {
          await chatMut.updateMessage(messageId, { content: finalText })
          spindle.log.info(`[SP] Applied decoded text to message ${messageId} via chatMutation.updateMessage`)
        } else {
          spindle.log.warn(`[SP] chat_mutation permission granted but no update method found. Available: ${chatMut ? Object.keys(chatMut).join(', ') : 'null'}`)
        }
      } else {
        spindle.log.info(`[SP] chat_mutation not granted — decoded text sent to frontend only`)
      }
    } catch (err) {
      spindle.log.warn(`[SP] Failed to apply decoded text via chat_mutation: ${err}`)
    }
  }

  // Send final text to frontend
  spindle.sendToFrontend({
    type: 'sp_generation_complete',
    chatId: state.activeChatId,
    decodedText: displayText,
    fullText: finalText,
  })

  // Reset state
  state.active = false
  clearHidePrefillState()
  clearContinueState()
  resetStreamGuard()
  state.accumulatedStreamText = ''
})

// ─── Generation Stopped Handler ──────────────────────────────────────────────

spindle.on('GENERATION_STOPPED', async (payload) => {
  if (!state.active) return

  const { chatId, content } = payload ?? {}
  if (chatId && state.activeChatId && chatId !== state.activeChatId) return

  // Best-effort final decode
  const rawText = String(content ?? state.accumulatedStreamText ?? '')
  const unwrapped = tryUnwrapStructuredOutput(rawText)
  const finalText = typeof unwrapped === 'string' ? unwrapped : state.lastAppliedText || rawText

  const displayText = stripHidePrefill(finalText)

  spindle.sendToFrontend({
    type: 'sp_generation_complete',
    chatId: state.activeChatId,
    decodedText: displayText,
    fullText: finalText,
  })

  state.active = false
  clearHidePrefillState()
  clearContinueState()
  resetStreamGuard()
  state.accumulatedStreamText = ''
})

// ─── Frontend Message Handler ────────────────────────────────────────────────

spindle.onFrontendMessage(async (payload: any, userId) => {
  // Capture userId for operator-scoped extensions (lesson from Mode Toggles port)
  if (userId && userId !== currentUserId) {
    spindle.log.info(`[SP] Captured userId from frontend: ${userId}`)
    currentUserId = userId
  }

  switch (payload?.type) {
    case 'get_settings': {
      spindle.sendToFrontend({ type: 'settings', settings })
      break
    }

    case 'update_settings': {
      const updates = payload.settings
      if (updates && typeof updates === 'object') {
        Object.assign(settings, updates)
        await saveSettings()
        spindle.sendToFrontend({ type: 'settings', settings })
      }
      break
    }

    case 'get_connections': {
      try {
        // Pass userId for operator-scoped extensions
        const connections = await spindle.connections.list(currentUserId)
        spindle.sendToFrontend({ type: 'connections', connections })
      } catch (err) {
        spindle.log.warn(`[SP] Failed to list connections: ${err}`)
        spindle.sendToFrontend({ type: 'connections', connections: [] })
      }
      break
    }

    case 'get_state': {
      spindle.sendToFrontend({
        type: 'state',
        active: state.active,
        patternMode: state.patternMode,
      })
      break
    }
  }
})

// ─── Initialization ──────────────────────────────────────────────────────────

;(async () => {
  await loadSettings()
  
  // Log available APIs for debugging
  try {
    const granted = await spindle.permissions.getGranted()
    spindle.log.info(`[SP] Granted permissions: [${granted.join(', ')}]`)
  } catch (e) {
    spindle.log.warn(`[SP] Could not check permissions: ${e}`)
  }

  spindle.log.info(`[SP] ${EXT_NAME} backend loaded (v1.0.0) — settings.enabled=${settings.enabled}`)
})()
