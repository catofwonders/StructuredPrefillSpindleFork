# StructuredPrefill — Lumiverse Spindle Port

Prefill-like behavior using Structured Outputs (`json_schema`) instead of assistant-role prefills.

**Ported from**: [SillyTavern StructuredPrefill](https://github.com/mia13165/StructuredPrefill) by mia13165

## What it does

1. Add an **assistant-role** message at the bottom of your prompt (your prefill template)
2. Send a message like normal
3. StructuredPrefill auto-activates (when the provider supports it) and forces the LLM response to begin with your prefill text via `json_schema` structured output
4. The JSON wrapper is decoded in real-time during streaming — you see clean text, not JSON

## Installation

```bash
# Via the Extensions panel or REST API:
POST /api/v1/spindle/install
{ "github_url": "https://github.com/<your-fork>/structured-prefill-spindle" }
```

### Required Permissions

| Permission | Why |
|---|---|
| `interceptor` | Modify the prompt to inject `json_schema` structured output constraints |
| `generation` | Fire the `[[pg]]` prefill generator (separate LLM call) + inspect connection profiles |
| `chat_mutation` | Apply decoded text back to chat messages after generation |

Grant them via the Extensions panel or:
```bash
curl -X POST http://localhost:7860/api/v1/spindle/structured_prefill/permissions \
  -H 'Content-Type: application/json' \
  -d '{ "grant": ["interceptor", "generation", "chat_mutation"] }'
```

## Building

```bash
# Install types
bun add -d lumiverse-spindle-types

# Build
bun run build
```

Or let Lumiverse auto-build from `src/` on install.

## Template Syntax

| Slot | Description |
|---|---|
| `[[keep]]` | Everything before this marker is hidden from display; everything after stays visible |
| `[[end]]` / `[[stop]]` / `[[eos]]` | Force the model to stop at this exact point |
| `[[pg]]` | Replaced by output from the Prefill Generator (separate LLM call) |
| `[[name]]` | Matches any known character/user name |
| `[[any]]` | Matches any text (non-greedy) |
| `[[re:PATTERN]]` | Raw regex passthrough |

## Architecture Notes (vs SillyTavern version)

| SillyTavern | Lumiverse Spindle |
|---|---|
| `CHAT_COMPLETION_SETTINGS_READY` event | `spindle.registerInterceptor()` |
| `generateData.json_schema = ...` | `context.response_format` / `context.json_schema` mutation |
| `extension_settings[name]` | `spindle.storage.getJson('settings.json')` |
| `eventSource.on('STREAM_TOKEN_RECEIVED')` | `spindle.on('STREAM_TOKEN_RECEIVED')` |
| `eventSource.on('MESSAGE_RECEIVED')` | `spindle.on('GENERATION_ENDED')` |
| DOM manipulation (jQuery) | `spindle.sendToFrontend()` → frontend DOM helper |
| Settings HTML in extensions panel | Drawer tab via `ctx.ui.registerDrawerTab()` |
| `saveSettingsDebounced()` | `spindle.storage.setJson()` |
| `toastr.error()` | `spindle.toast.error()` |
| Connection profiles via `extension_settings.connectionManager` | `spindle.connections.list()` |

## Supported Providers

Works with any OpenAI-compatible provider that supports `response_format: { type: 'json_schema' }`:
- OpenAI (GPT-4o, GPT-4.1, etc.)
- OpenRouter (most models)
- Google Gemini (via Lumiverse's parameter forwarding)
- Any custom OpenAI-compatible endpoint

**Not compatible** with providers that remap `json_schema` to prompt hacks:
- Anthropic/Claude (uses tool-based structured output instead)
- AI21, DeepSeek, Moonshot, ZAI, SiliconFlow

## License

Same as the original SillyTavern extension.
