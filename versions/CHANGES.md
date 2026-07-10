# Lumen onboarding — fix log

Each version below is one fix, applied on top of the previous. Full snapshots of
the files changed in each fix live in `versions/vNN-<slug>/`. `v00-baseline` is
the untouched starting point. Git commits on branch `claude/lumen-onboarding-fixes`
carry the same steps for diffing.

Severity key: B = blocker, C = confusing, M = minor.

| v | Sev | Fix | Files |
|---|-----|-----|-------|
| 01 | M | Remove dead `getSP()` (full duplicate of the server system prompt) and unused `_unusedLegacyClientAPI` stub. Kills prompt drift + shrinks bundle. | src/lumen.jsx |
| 02 | B | Remove false "editable Google Sheets copy / email / consultant sees every update" promise. FinishCard sent-copy + STEP 7 now describe what actually happens (brief sent, consultant follows up in 2 business days). Removed the never-rendering Sheets link button. | src/lumen.jsx, netlify/functions/chat.js |
| 03 | B | Remove pause/resume + auto-save promises (not implemented live) and the "(simulated in demo)" badge shown to real clients. step1Desc (6 langs), stepper badge, prompt START, prompt CHECKPOINT A. NOTE: real resume is a follow-up (needs stable session id + per-turn save + load-on-mount; not verifiable against Blobs here). | src/lumen.jsx, netlify/functions/chat.js |
| 04 | C | Rewrite PACING so it no longer says "then STOP. Wait for the user's next message" after a widget ack (which contradicted FORWARD MOTION and could strand the client). Now: acknowledge + ask the next question in the same turn; only forbids stacking a second widget. | netlify/functions/chat.js |
| 05 | C | Expert flow (STEP 4A) now shows [WIDGET:MARKETS] like the guided flow, so INFERRED SETUP (languages/timezone) and OBJECTIVES have a defined trigger for experts instead of relying on the gap sweep. | netlify/functions/chat.js |
| 06 | C | Add a manual "Review and send your brief" link below the input (shown once progress >= 15% and < 100%). Gives an escape hatch if the model never emits percent:100; ExportModal still gates the actual send on readiness and lets the client fill gaps. | src/lumen.jsx |
