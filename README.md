# ReBath IT Helper

A Microsoft Teams support bot that triages basic IT issues, walks users through safe
self-help steps (restart the app, reboot, clear cache/cookies, try InPrivate), and—if
the problem persists—collects the details and **creates a real incident in ServiceNow**.

Built on the Microsoft 365 Agents Toolkit and the Teams SDK v2 (`@microsoft/teams.apps`).

## What the bot does

1. **Triage** — recognizes greetings, help requests, and issue descriptions.
2. **Classify, don't improvise** — when a user describes a problem, the AI's only job
   is to classify it into one of a fixed set of categories and extract a few fields
   (affected system, start time, error text). The AI never writes the user-facing
   reply. Every word the user sees during self-help comes from an approved library
   (`responses.ts`) that you control.
3. **Approved self-help** — the bot shows the safe, first-line steps for that category
   (close/reopen, reboot, clear cache/cookies, try InPrivate, check the connection).
   Security and account/sign-in issues skip self-help entirely and go straight to a
   ticket, with evidence-preservation guidance.
4. **Smart intake** — if self-help fails, it opens a ticket pre-filled from the
   classification, so it asks only for what's still missing (usually one or two
   follow-ups, not ten).
5. **ServiceNow incident creation** — on confirmation it calls the ServiceNow Table
   API to create an `incident` and returns the real `INC` number.

## The hard-lock guarantee

The model cannot show troubleshooting advice in its own words, because the code has no
path that sends model-authored prose to the user. `aiHelper.ts` returns structured
JSON only (a category plus extracted fields); `responses.ts` holds every user-facing
string. If OpenAI is unavailable, a deterministic keyword classifier
(`classifyOffline`) picks the category instead — it still only *routes* to an approved
response, never generates text — and security/account issues still escalate. To change
what users are told, edit `responses.ts`.

## Safety & privacy features

- **Secret/PII scrubbing** (`security.ts`): user text is redacted for credentials, MFA
  codes, card/SSN numbers, and long tokens before it is stored, logged, or sent to
  ServiceNow. The OpenAI call additionally strips email addresses.
- **No raw logging**: console logs use a redacted, truncated preview—never the raw
  message.
- **Tenant guard**: set `ALLOWED_TENANT_ID` to restrict the bot to your own Entra
  tenant. If unset, a startup warning is logged.
- **Graceful degradation**: if OpenAI or ServiceNow is not configured, the bot still
  works (keyword triage, locally-saved tickets) and tells the user what's happening.

## Project layout

| File | Purpose |
| - | - |
| `app.ts` | Main message handler: triage, self-help routing, intake state machine, submission |
| `responses.ts` | The approved response library + categories + offline classifier. **Edit this to change what users see.** |
| `aiHelper.ts` | OpenAI call for classification + extraction only (scrubs input first; no user-facing text) |
| `serviceNow.ts` | ServiceNow Table API client (caller lookup + incident creation) |
| `security.ts` | Redaction helpers and the tenant access guard |
| `index.ts` | Starts the app |
| `config.ts` | Bot auth config from environment |
| `appPackage/` | Teams app manifest and icons |
| `env/` | Microsoft 365 Agents Toolkit environment files |
| `infra/` | Azure provisioning templates |

## Configuration

Copy `.env.example` and fill in the values for your hosting environment. Required for
full functionality:

- Teams bot auth: `CLIENT_ID`, `CLIENT_SECRET`, `TENANT_ID`
- ServiceNow: `SERVICENOW_INSTANCE`, `SERVICENOW_USERNAME`, `SERVICENOW_PASSWORD`
- Recommended: `ALLOWED_TENANT_ID`, `OPENAI_API_KEY`

See `.env.example` for the full list and notes.

## Run locally

```bash
npm install
npm run build
npm start
```

Or use the Microsoft 365 Agents Toolkit / Playground as before (`npm run dev`).

## Commands

| Command | Action |
| - | - |
| `/ticket` | Start a new IT support ticket |
| `/status` | Show current ticket progress |
| `/reboot` | Show safe reboot guidance |
| `/cancel` | Cancel the current ticket intake |
| `/reset`  | Clear the conversation state |
| `/help`   | Show the help message |

You can also just describe the problem in plain language.

## Production hardening checklist

- [ ] Set `ALLOWED_TENANT_ID` to your Entra tenant.
- [ ] Use a dedicated, least-privilege ServiceNow integration account.
- [ ] Prefer OAuth over Basic auth for ServiceNow (see the note in `serviceNow.ts`).
- [ ] Confirm your OpenAI account's data-retention settings meet your policy, or run a
      self-hosted model. The bot scrubs input, but configuration is still your call.
- [ ] Replace the in-memory `LocalStorage` with a durable store if you need ticket
      intake to survive restarts (see "Known limitations").

## Known limitations

- **State is in-memory.** `LocalStorage` from the SDK is per-process. A restart or a
  scaled-out second instance loses in-progress intake. For production, back state with
  a database or distributed cache.
- **Caller matching is best-effort.** ServiceNow caller lookup is by email; if the
  Teams account doesn't expose a UPN/email or it doesn't match a ServiceNow user, the
  incident falls back to `SERVICENOW_DEFAULT_CALLER_ID` (or no caller).
- **Redaction is heuristic.** The scrubber catches common secret shapes, not every
  possible one. It is defense-in-depth, not a guarantee.
