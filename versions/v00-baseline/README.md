# Lumen Onboarding — live suite

Three apps on one Netlify site, sharing one session store.

- `public/sales.html` — Sales generates a client link (static, no build)
- `src/lumen.jsx` + `src/chat-main.jsx` — the client onboarding chat (built by Vite)
- `public/dashboard.html` — Proserv dashboard (static, no build)
- `netlify/functions/chat.js` — Anthropic proxy (holds the API key)
- `netlify/functions/session.js` — session store (Netlify Blobs)

## One-time setup

1. Push this folder to a Git repo and connect it to a new Netlify site
   (or drag-drop after running `npm install && npm run build` locally).
2. In Netlify: Site settings > Environment variables, add:
   `ANTHROPIC_API_KEY = <your key>`
3. Deploy. Build command `npm run build`, publish dir `dist`, functions
   dir `netlify/functions` are already set in `netlify.toml`.

## URLs after deploy

- `/sales` — internal, for the Sales team to generate links
- `/chat?c=...` — the client link Sales hands out
- `/dashboard` — internal, for Proserv

`/sales`, `/chat`, `/dashboard` are redirect aliases to the `.html` files.

## How the pieces connect

- Sales encodes the client profile into the `?c=` param. The chat decodes it
  and runs seeded.
- The chat calls `/.netlify/functions/chat`. The key never touches the browser.
- On "send", the chat POSTs the completed brief to
  `/.netlify/functions/session`, stored in Netlify Blobs.
- The dashboard GETs `/.netlify/functions/session` for the list and
  `?id=` for a full record.

## Notes / to confirm

- **Access control**: none of these pages are gated yet. `/sales` and
  `/dashboard` are internal and should sit behind Netlify password protection
  or Identity before real use. The `/chat` link is unguessable but public by
  design (the client has no login).
- **Netlify Blobs**: verify `@netlify/blobs` resolves on your Netlify build
  (it is included automatically on recent runtimes). If the dashboard shows a
  store error, check the package version against current Netlify docs.
- **Governance**: completed briefs (client data, possibly PII) live in Netlify
  Blobs on this site. Confirm retention and access with whoever owns the ISO
  42001 work before real client data goes in.
- **Cost**: each ~15 min chat is roughly 15-25 Anthropic calls. The proxy caps
  model and max_tokens server-side so a tampered client cannot inflate spend.
