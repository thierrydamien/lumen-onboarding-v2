# Lumen Onboarding — live suite

Three apps on one Netlify site, sharing one session store.

- `public/sales.html` — Sales generates a client link (static, no build)
- `src/lumen.jsx` + `src/chat-main.jsx` — the client onboarding chat (built by Vite)
- `public/dashboard.html` — Proserv dashboard (static, no build)
- `netlify/functions/chat.js` — Anthropic proxy (holds the API key)
- `netlify/functions/session.js` — session store (Netlify Blobs)
- `netlify/functions/seed.js` — seed store for pre-filled client profiles
  (Netlify Blobs); keeps consultant notes out of the client link
- `netlify/functions/sheet.js` — converts the sent brief into an editable Google
  Sheet (Drive API). Inert until the Google env vars below are set.

## One-time setup

1. Push this folder to a Git repo and connect it to a new Netlify site
   (or drag-drop after running `npm install && npm run build` locally).
2. In Netlify: Site settings > Environment variables, add:
   - `ANTHROPIC_API_KEY = <your key>` — required for the chat proxy.
   - `DASHBOARD_TOKEN = <a long random string>` — required for the dashboard.
     Session reads and consultant-notes reads are locked until this is set, and
     the dashboard prompts for it in the browser.
   - Google Sheet generation (optional; the brief still sends without it). Set
     `GOOGLE_DRIVE_FOLDER_ID` to the target folder, then ONE auth path:
     - Path A, OAuth as a real user (writes into that user's My Drive folder on
       their quota; no Workspace admin needed). Use this to target a folder owned
       by an account you control: `GOOGLE_OAUTH_CLIENT_ID`,
       `GOOGLE_OAUTH_CLIENT_SECRET`, `GOOGLE_OAUTH_REFRESH_TOKEN`.
     - Path B, service account + domain-wide delegation (impersonate a real user;
       a Workspace admin must authorize the SA for the Drive scope):
       `GOOGLE_SERVICE_ACCOUNT_EMAIL`, `GOOGLE_SERVICE_ACCOUNT_PRIVATE_KEY`,
       `GOOGLE_IMPERSONATE_SUBJECT` (the user to act as).
     - Path C, service account into a Shared Drive (add the SA as Content Manager;
       no impersonation): `GOOGLE_SERVICE_ACCOUNT_EMAIL`,
       `GOOGLE_SERVICE_ACCOUNT_PRIVATE_KEY`. A bare SA has no Drive quota, so
       without a Shared Drive (or Path B impersonation) the upload fails on quota.
3. Deploy. Build command `npm run build`, publish dir `dist`, functions
   dir `netlify/functions` are already set in `netlify.toml`.

> NOT YET SMOKE-TESTED: `sheet.js` is written against the Drive v3 REST API but
> has not been run against real Google credentials. Verify the round trip on a
> deploy before relying on it. GOVERNANCE: it writes client data (possibly PII)
> to Google Drive — confirm folder, sharing scope, and retention with the ISO
> 42001 owner first.

## URLs after deploy

- `/sales` — internal, for the Sales team to generate links
- `/chat?s=...` — the client link Sales hands out
- `/dashboard` — internal, for Proserv

`/sales`, `/chat`, `/dashboard` are redirect aliases to the `.html` files.

## How the pieces connect

- Sales POSTs the client profile (including confidential notes) to
  `/.netlify/functions/seed`, which stores it and returns an opaque id. The
  client link carries only that id (`?s=<id>`). Notes never travel in the URL.
- The chat fetches the CLIENT-SAFE seed fields (no notes) by id and runs seeded.
- The chat calls `/.netlify/functions/chat`. The key never touches the browser.
- On "send", the chat builds the XLSX brief and (if Google is configured) POSTs
  it to `/.netlify/functions/sheet`, which creates an editable Google Sheet,
  shares it with the client's email, and returns the link. The chat then POSTs
  the completed brief (with the Sheet URL) to `/.netlify/functions/session`,
  stored in Netlify Blobs. If Sheet generation is not configured or fails, the
  brief still sends; the client just gets no Sheet link.
- The dashboard GETs `/.netlify/functions/session` for the list and `?id=` for a
  full record, and GETs `/.netlify/functions/seed?id=` for the consultant notes.
  All dashboard reads send the `DASHBOARD_TOKEN`; notes are returned only to a
  caller holding that token.

## Notes / to confirm

- **Access control**: `/sales` and `/dashboard` pages are not gated at the page
  level and should sit behind Netlify password protection or Identity before real
  use (they are marked `noindex`, which is only a crawler hint). Dashboard data
  reads are token-gated via `DASHBOARD_TOKEN`. The `/chat` link is unguessable
  but public by design (the client has no login).
- **Netlify Blobs**: verify `@netlify/blobs` resolves on your Netlify build
  (it is included automatically on recent runtimes). If the dashboard shows a
  store error, check the package version against current Netlify docs.
- **Governance**: completed briefs (client data, possibly PII) live in Netlify
  Blobs on this site. Confirm retention and access with whoever owns the ISO
  42001 work before real client data goes in.
- **Cost**: each ~15 min chat is roughly 15-25 Anthropic calls. The proxy caps
  model and max_tokens server-side so a tampered client cannot inflate spend.
