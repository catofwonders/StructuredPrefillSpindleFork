import type { SpindleFrontendContext } from 'lumiverse-spindle-types'

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

interface ConnectionProfile {
  id: string
  name: string
  provider: string
  model: string
  is_default: boolean
}

let ctx: SpindleFrontendContext
let currentSettings: Settings | null = null
let connections: ConnectionProfile[] = []
let drawerTab: any = null

// ─── Settings UI Builder ─────────────────────────────────────────────────────

function buildSettingsHtml(): string {
  const s = currentSettings
  if (!s) return '<p style="color:var(--lumiverse-text-muted)">Loading settings...</p>'

  return `
    <div class="sp-settings" style="display:flex;flex-direction:column;gap:12px;padding:4px 0;">

      <div class="sp-section">
        <div class="sp-section-header" data-section="general">General</div>
        <div class="sp-section-body" data-section-body="general">
          <label class="sp-checkbox">
            <input type="checkbox" data-key="enabled" ${s.enabled ? 'checked' : ''}>
            <span>Enabled</span>
          </label>
          <label class="sp-checkbox">
            <input type="checkbox" data-key="hide_prefill_in_display" ${s.hide_prefill_in_display ? 'checked' : ''}>
            <span>Hide prefill text in final message (use <code>[[keep]]</code> to keep tail visible)</span>
          </label>
        </div>
      </div>

      <div class="sp-section">
        <div class="sp-section-header" data-section="schema">Schema</div>
        <div class="sp-section-body sp-collapsed" data-section-body="schema">
          <label class="sp-field">
            <span>Minimum characters after prefix</span>
            <input type="number" data-key="min_chars_after_prefix" value="${s.min_chars_after_prefix}" min="1" max="10000" step="1">
            <small>Normal generations only. Continue uses its own minimal constraint.</small>
          </label>
          <label class="sp-field">
            <span>Newline token (encoded in schema)</span>
            <input type="text" data-key="newline_token" value="${escHtml(s.newline_token)}" placeholder="\\n">
          </label>
        </div>
      </div>

      <div class="sp-section">
        <div class="sp-section-header" data-section="prefill_gen">Prefill Generator (<code>[[pg]]</code>)</div>
        <div class="sp-section-body sp-collapsed" data-section-body="prefill_gen">
          <label class="sp-checkbox">
            <input type="checkbox" data-key="prefill_gen_enabled" ${s.prefill_gen_enabled ? 'checked' : ''}>
            <span>Replace <code>[[pg]]</code> using a separate LLM call</span>
          </label>
          <label class="sp-field">
            <span>Extra prompt</span>
            <textarea data-key="prefill_gen_extra_prompt" rows="3">${escHtml(s.prefill_gen_extra_prompt)}</textarea>
          </label>
          <label class="sp-field">
            <span>Extra prompt role</span>
            <select data-key="prefill_gen_extra_prompt_role">
              <option value="system" ${s.prefill_gen_extra_prompt_role === 'system' ? 'selected' : ''}>System</option>
              <option value="user" ${s.prefill_gen_extra_prompt_role === 'user' ? 'selected' : ''}>User</option>
              <option value="assistant" ${s.prefill_gen_extra_prompt_role === 'assistant' ? 'selected' : ''}>Assistant</option>
            </select>
          </label>
          <label class="sp-field">
            <span>Connection profile</span>
            <select data-key="prefill_gen_connection_id" id="sp-pg-connection">
              <option value="">&lt;None&gt;</option>
              ${connections.map(c => `<option value="${escHtml(c.id)}" ${c.id === s.prefill_gen_connection_id ? 'selected' : ''}>${escHtml(c.name)} (${escHtml(c.provider)}/${escHtml(c.model)})</option>`).join('')}
            </select>
          </label>
          <label class="sp-field">
            <span>Max tokens</span>
            <input type="number" data-key="prefill_gen_max_tokens" value="${s.prefill_gen_max_tokens}" min="1" max="2048" step="1">
          </label>
          <label class="sp-field">
            <span>Stop strings (one per line)</span>
            <textarea data-key="prefill_gen_stop" rows="3" placeholder="\\n">${escHtml(s.prefill_gen_stop)}</textarea>
          </label>
          <label class="sp-checkbox">
            <input type="checkbox" data-key="prefill_gen_keep_matched_stop_string" ${s.prefill_gen_keep_matched_stop_string ? 'checked' : ''}>
            <span>Append matched stop string to generated prefill</span>
          </label>
          <label class="sp-field">
            <span>Timeout (ms)</span>
            <input type="number" data-key="prefill_gen_timeout_ms" value="${s.prefill_gen_timeout_ms}" min="500" max="120000" step="100">
          </label>
        </div>
      </div>

      <div class="sp-section">
        <div class="sp-section-header" data-section="continue">Continue</div>
        <div class="sp-section-body sp-collapsed" data-section-body="continue">
          <label class="sp-field">
            <span>Overlap characters</span>
            <input type="number" data-key="continue_overlap_chars" value="${s.continue_overlap_chars}" min="0" max="120" step="1">
          </label>
          <small style="color:var(--lumiverse-text-muted)"><code>[[pg]]</code> is not used for Continue.</small>
        </div>
      </div>

      <div class="sp-section">
        <div class="sp-section-header" data-section="antislop">Anti-Slop</div>
        <div class="sp-section-body sp-collapsed" data-section-body="antislop">
          <label class="sp-field">
            <span>Banned words (one per line)</span>
            <textarea data-key="anti_slop_ban_list" rows="5" placeholder="ozone&#10;Elara&#10;&mdash;">${escHtml(s.anti_slop_ban_list)}</textarea>
          </label>
        </div>
      </div>

      <div class="sp-section">
        <div class="sp-section-header" data-section="compat">Connection Compatibility</div>
        <div class="sp-section-body sp-collapsed" data-section-body="compat">
          <small style="color:var(--lumiverse-text-muted);display:block;margin-bottom:8px;">
            The extension auto-detects what structured-output mode each connection supports.
            If a connection was wrongly marked (e.g. the proxy was temporarily down), reset this
            to let the extension re-probe from scratch.
          </small>
          <button type="button" data-action="reset_compat" class="sp-button">
            Reset Connection Compatibility
          </button>
          <div data-compat-status style="margin-top:8px;font-size:12px;color:var(--lumiverse-text-muted);"></div>
        </div>
      </div>

      <div class="sp-section">
        <div class="sp-section-header" data-section="help">Help</div>
        <div class="sp-section-body sp-collapsed" data-section-body="help">
          <small style="color:var(--lumiverse-text-muted)">
            Docs: <a href="https://rentry.org/structuredprefill" target="_blank" rel="noopener noreferrer" style="color:var(--lumiverse-accent)">rentry.org/structuredprefill</a>
            <br><br>
            Ported from SillyTavern to Lumiverse Spindle format.
          </small>
        </div>
      </div>
    </div>
  `
}

function escHtml(s: string): string {
  return String(s ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
}

// ─── Settings Event Wiring ───────────────────────────────────────────────────

function wireSettingsListeners(root: HTMLElement): void {
  // Section toggle (accordion)
  root.querySelectorAll('.sp-section-header').forEach(header => {
    header.addEventListener('click', () => {
      const section = header.getAttribute('data-section')
      if (!section) return
      const body = root.querySelector(`[data-section-body="${section}"]`)
      if (body) body.classList.toggle('sp-collapsed')
    })
  })

  // Checkbox inputs
  root.querySelectorAll<HTMLInputElement>('input[type="checkbox"][data-key]').forEach(input => {
    input.addEventListener('change', () => {
      const key = input.getAttribute('data-key')
      if (!key || !currentSettings) return
      ;(currentSettings as any)[key] = input.checked
      ctx.sendToBackend({ type: 'update_settings', settings: { [key]: input.checked } })
    })
  })

  // Text/number inputs
  root.querySelectorAll<HTMLInputElement>('input[type="text"][data-key], input[type="number"][data-key]').forEach(input => {
    input.addEventListener('change', () => {
      const key = input.getAttribute('data-key')
      if (!key || !currentSettings) return
      const value = input.type === 'number' ? Number(input.value) : input.value
      ;(currentSettings as any)[key] = value
      ctx.sendToBackend({ type: 'update_settings', settings: { [key]: value } })
    })
  })

  // Textareas
  root.querySelectorAll<HTMLTextAreaElement>('textarea[data-key]').forEach(textarea => {
    textarea.addEventListener('input', () => {
      const key = textarea.getAttribute('data-key')
      if (!key || !currentSettings) return
      ;(currentSettings as any)[key] = textarea.value
      ctx.sendToBackend({ type: 'update_settings', settings: { [key]: textarea.value } })
    })
  })

  // Select inputs
  root.querySelectorAll<HTMLSelectElement>('select[data-key]').forEach(select => {
    select.addEventListener('change', () => {
      const key = select.getAttribute('data-key')
      if (!key || !currentSettings) return
      ;(currentSettings as any)[key] = select.value
      ctx.sendToBackend({ type: 'update_settings', settings: { [key]: select.value } })
    })
  })

  // Action buttons
  root.querySelectorAll<HTMLButtonElement>('button[data-action]').forEach(button => {
    button.addEventListener('click', () => {
      const action = button.getAttribute('data-action')
      if (action === 'reset_compat') {
        ctx.sendToBackend({ type: 'clear_blocklist' })
        const status = root.querySelector<HTMLElement>('[data-compat-status]')
        if (status) status.textContent = 'Resetting…'
      }
    })
  })
}

function renderSettingsInTab(): void {
  if (!drawerTab) return
  drawerTab.root.innerHTML = buildSettingsHtml()
  wireSettingsListeners(drawerTab.root)
}

// ─── Styles ──────────────────────────────────────────────────────────────────

function injectStyles(): () => void {
  return ctx.dom.addStyle(`
    .sp-settings {
      font-size: 13px;
      line-height: 1.5;
    }

    .sp-section {
      border: 1px solid var(--lumiverse-border);
      border-radius: var(--lumiverse-radius);
      overflow: hidden;
    }

    .sp-section-header {
      padding: 8px 12px;
      background: var(--lumiverse-fill-subtle);
      border-bottom: 1px solid var(--lumiverse-border);
      font-weight: 600;
      cursor: pointer;
      user-select: none;
      display: flex;
      align-items: center;
      gap: 6px;
      transition: background var(--lumiverse-transition-fast);
    }

    .sp-section-header:hover {
      background: var(--lumiverse-fill);
    }

    .sp-section-header code {
      font-size: 11px;
      padding: 1px 4px;
      background: var(--lumiverse-fill);
      border-radius: 3px;
    }

    .sp-section-body {
      padding: 10px 12px;
      display: flex;
      flex-direction: column;
      gap: 10px;
    }

    .sp-section-body.sp-collapsed {
      display: none;
    }

    .sp-checkbox {
      display: flex;
      align-items: flex-start;
      gap: 8px;
      cursor: pointer;
    }

    .sp-checkbox input[type="checkbox"] {
      margin-top: 3px;
      flex-shrink: 0;
    }

    .sp-checkbox span {
      color: var(--lumiverse-text);
    }

    .sp-checkbox code {
      font-size: 11px;
      padding: 1px 4px;
      background: var(--lumiverse-fill-subtle);
      border-radius: 3px;
    }

    .sp-field {
      display: flex;
      flex-direction: column;
      gap: 4px;
    }

    .sp-field span {
      font-weight: 500;
      color: var(--lumiverse-text);
    }

    .sp-field input,
    .sp-field select,
    .sp-field textarea {
      padding: 6px 8px;
      border: 1px solid var(--lumiverse-border);
      border-radius: var(--lumiverse-radius);
      background: var(--lumiverse-fill);
      color: var(--lumiverse-text);
      font-size: 13px;
      font-family: inherit;
    }

    .sp-field input:focus,
    .sp-field select:focus,
    .sp-field textarea:focus {
      outline: none;
      border-color: var(--lumiverse-accent);
    }

    .sp-field small {
      color: var(--lumiverse-text-dim);
      font-size: 11px;
    }

    .sp-field textarea {
      resize: vertical;
      min-height: 60px;
    }

    .sp-button {
      background: var(--lumiverse-surface-2, #2a2a2a);
      color: var(--lumiverse-text, #fff);
      border: 1px solid var(--lumiverse-border, #444);
      padding: 6px 12px;
      border-radius: 4px;
      cursor: pointer;
      font-size: 12px;
      font-family: inherit;
    }

    .sp-button:hover {
      background: var(--lumiverse-surface-3, #3a3a3a);
      border-color: var(--lumiverse-accent, #888);
    }

    .sp-button:active {
      transform: translateY(1px);
    }
  `)
}

// ─── Setup ───────────────────────────────────────────────────────────────────

const GEAR_ICON = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 20 20" fill="currentColor"><path fill-rule="evenodd" d="M7.84 1.804A1 1 0 018.82 1h2.36a1 1 0 01.98.804l.331 1.652a6.993 6.993 0 011.929 1.115l1.598-.54a1 1 0 011.186.447l1.18 2.044a1 1 0 01-.205 1.251l-1.267 1.113a7.047 7.047 0 010 2.228l1.267 1.113a1 1 0 01.206 1.25l-1.18 2.045a1 1 0 01-1.187.447l-1.598-.54a6.993 6.993 0 01-1.929 1.115l-.33 1.652a1 1 0 01-.98.804H8.82a1 1 0 01-.98-.804l-.331-1.652a6.993 6.993 0 01-1.929-1.115l-1.598.54a1 1 0 01-1.186-.447l-1.18-2.044a1 1 0 01.205-1.251l1.267-1.114a7.05 7.05 0 010-2.227L1.821 7.773a1 1 0 01-.206-1.25l1.18-2.045a1 1 0 011.187-.447l1.598.54A6.993 6.993 0 017.51 3.456l.33-1.652zM10 13a3 3 0 100-6 3 3 0 000 6z" clip-rule="evenodd"/></svg>`

export function setup(context: SpindleFrontendContext) {
  ctx = context

  // Inject styles
  const removeStyles = injectStyles()

  // Register drawer tab
  drawerTab = ctx.ui.registerDrawerTab({
    id: 'settings',
    title: 'StructuredPrefill',
    iconSvg: GEAR_ICON,
  })

  // Request initial data
  ctx.sendToBackend({ type: 'get_settings' })
  ctx.sendToBackend({ type: 'get_connections' })

  // Handle backend messages
  ctx.onBackendMessage((payload: any) => {
    switch (payload?.type) {
      case 'settings': {
        currentSettings = payload.settings
        renderSettingsInTab()
        break
      }

      case 'connections': {
        connections = Array.isArray(payload.connections) ? payload.connections : []
        renderSettingsInTab()
        break
      }

      case 'sp_activated': {
        // Could show an indicator in the UI that structured prefill is active
        break
      }

      case 'sp_stream_update': {
        // The backend decoded the JSON stream — we could apply it to DOM here
        // if Lumiverse's own streaming renderer shows raw JSON.
        // For now, the backend handles this via chat_mutation.
        break
      }

      case 'sp_generation_complete': {
        // Generation finished — final decoded text available
        break
      }

      case 'blocklist_cleared': {
        if (drawerTab) {
          const status = drawerTab.root.querySelector('[data-compat-status]')
          if (status) {
            status.textContent = 'Compatibility cache cleared. Next message will re-probe.'
            setTimeout(() => { if (status) status.textContent = '' }, 5000)
          }
        }
        break
      }
    }
  })

  // Cleanup
  return () => {
    removeStyles()
    if (drawerTab) drawerTab.destroy()
  }
}
