# Lumen Onboarding: deploy checklist

Everything you need to update to take the current build live. There are two
surfaces, deployed separately: the Netlify site and the Apps Script Web App.

## 1. Netlify site

Source of truth for the web app. Build is `npm run build` (publish `dist`,
functions `netlify/functions`), per netlify.toml.

Files changed from the original baseline (deploy all of these):
- src/lumen.jsx: the client app. Rebuilt into dist/ by `npm run build`. Do not
  hand-edit dist; edit the source and rebuild.
- public/dashboard.html: Proserv dashboard (static, served as-is).
- public/sales.html: consultant link generator (static).
- netlify/functions/chat.js: Anthropic proxy (system prompt + caching).
- netlify/functions/session.js: session store (completed and in-progress).
- netlify/functions/seed.js: seed store (prepared client profiles).
- netlify/functions/sheet.js: Sheet generation proxy (forwards to Apps Script).
- netlify/functions/stalled-check.js: scheduled function (runs hourly, no route)
  that Slack-alerts on onboardings idle over 24h. Registered by its own schedule
  export, so no netlify.toml change is needed.

Environment variables (Site settings, Environment variables):
Required:
- ANTHROPIC_API_KEY: the chat will not run without it.
- DASHBOARD_TOKEN: gates all dashboard and seed reads. Pick a strong secret; the
  dashboard page prompts for it.
- APPS_SCRIPT_WEBAPP_URL: the deployed Apps Script Web App URL (section 2).
- APPS_SCRIPT_SECRET: shared secret, must equal the Apps Script SHARED_SECRET.
Optional:
- SEED_TTL_DAYS: seed retention in days (default 90; 0 disables expiry).
- SEED_WRITE_TOKEN: if set, the Sales page must supply this token to generate a
  link (it prompts the consultant and caches it per browser tab). Recommended for
  production so the seed store cannot be written by anyone spoofing the Origin
  header. Unset leaves seed writes origin-only.
- URL: set automatically by Netlify, do not set by hand.
- SLACK_BOT_TOKEN: needed only for the stalled-onboarding alert (stalled-check.js).
  Same bot token the Apps Script uses; set it here too so the Netlify function can
  post. Unset means the stalled alert is a no-op (everything else still works).
- SLACK_CHANNEL: channel id for the stalled alert (default C097154H39N, matches
  the completion alert).
- STALLED_HOURS: idle threshold in hours before an in-progress session is flagged
  stalled (default 24).
- GOOGLE_* (six vars): only for the fallback Google-API auth paths. Not needed
  while the Apps Script path is in use.

Function timeout (IMPORTANT — fixes the "didn't go through" failures on long
replies): the chat proxy is non-streaming, so the model's whole reply must be
generated inside the function's execution window. Netlify's default synchronous
timeout is ~10 seconds; a longer reply (big recap turns, large imports) gets the
function KILLED mid-generation and the client sees the retry banner. Raise it:
Site configuration > Functions > Functions timeout (or ask Netlify support on
some plans) to the 26-second maximum. The code side is already matched to that
window: the token ceiling is 2000 and the proxy self-aborts at 24s with a clean
error the client can retry.

Note: /dashboard and /sales carry noindex but are NOT access-gated in code. Put
them behind Netlify password protection or Identity before real-client use.

## 2. Apps Script Web App

A separate project. Runs on your Google account, copies the requirements
template into the Proserv folder, and posts the Slack alert.

File:
- apps-script/onboarding-sheet-webapp.gs: paste the full contents into the Apps
  Script editor, then Deploy, Manage deployments, edit the existing deployment,
  New version. Editing the existing deployment keeps the same Web App URL, so
  APPS_SCRIPT_WEBAPP_URL on Netlify does not change.

Script Properties (Project Settings, Script Properties):
Required:
- SHARED_SECRET: must equal Netlify APPS_SCRIPT_SECRET.
- SLACK_BOT_TOKEN: bot token for the completion alert (no alert if unset).
- DASHBOARD_URL: base dashboard URL; powers the "View full session" deep-link.
- PIPELINE_SHEET_ID: the tracker sheet id
  (1cU_-4GhgpK_YqzWR8pnC16IAEm00NA5Tsr7GyKHjvLY) for IC/TAM @mentions.
- SLACK_IDS_JSON: JSON map of staff name to Slack user id, for the @mentions.
  The roster lives here, not in the repo.
Optional:
- SLACK_CHANNEL: target channel (default C097154H39N).
- SLACK_ESCALATION: comma-separated Slack ids to ping on TW-Core or no-match.

## 3. What changed (per-version log)

This section previously tracked v43-v47 by hand; the build is now at v73. The
authoritative per-version log is versions/CHANGES.md (one row per version).
Highlights since v47 that affect deployment:
- Deploy ALL of section 1's files and redeploy the Apps Script (section 2) — the
  Apps Script gained idempotency (v46), Slack-mention escaping and Sheets
  formula-injection guards (v72), so the redeploy is not optional.
- Set the Netlify function timeout to 26s (see section 1) — pairs with the v73
  token-ceiling change to eliminate the long-reply timeout failures.
- v50-v51 hardened session/seed writes (Origin required; optional
  SEED_WRITE_TOKEN); v72 brought chat.js/sheet.js to the same standard and fixed
  a dashboard XSS — deploy dashboard.html with the functions, not separately.

### Older notes (v43 to v47)

- v43: the in-app demo Slack card was aligned to the real message. Internal demo
  only, no production behaviour.
- v44: dashboard "Link sent" rows (seeds generated but never opened) plus seed
  TTL (SEED_TTL_DAYS). Touches seed.js, session.js, dashboard.html.
- v45: tools/ab-harness.mjs, a dev-only cost-vs-quality test. Not deployed.
- v46: send-path durability. Sheet idempotency (a retry no longer creates a
  duplicate Sheet or re-fires Slack) plus save-before-Sheet, so the Slack
  deep-link never points at an unsaved session. Touches lumen.jsx and the Apps
  Script (needs the redeploy in section 2).
- v47: the A/B harness was extended to measure the history-cache lever (see the
  open item on finding 21). Dev-only, not deployed.

Deploy impact of v43 to v47: rebuild the Netlify site (src/lumen.jsx,
public/dashboard.html, netlify/functions/seed.js, netlify/functions/session.js)
and redeploy the Apps Script (onboarding-sheet-webapp.gs).

## Not deployed
- tools/ab-harness.mjs: local dev tool. Run with ANTHROPIC_API_KEY set; it makes
  real API calls (a few dollars per run). Nothing it does reaches production.
- versions/: per-version snapshots and CHANGES.md, kept for diffing and history.
