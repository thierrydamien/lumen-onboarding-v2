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
| 07 | C | Merge topic sources by name in ExportModal so urls/hashtags/comments from the %%TOPICS%% marker survive even when the topic-cards widget (name/keywords/rationale only) was used. Previously those fields were silently dropped from the brief and workbook. | src/lumen.jsx |
| 08 | C | (A) mergeCdata now unions channels/reports/alerts by name so a partial marker re-emit can't wipe earlier entries. (B) Reports and alerts are now editable in the review modal (add/edit/remove) and flow from that editable state into the brief, giving a human backstop if the model mis-captured them. | src/lumen.jsx |
| 09 | M | (A) sanitizeIn() collapses injected `%%` runs in client-authored text (typed messages, widget payloads, seeded opener) so a client can't corrupt marker parsing. (B) gw/gwp now pick the latest widget submission of a type (by message-index) instead of the first, avoiding stale reads. | src/lumen.jsx |
| 10 | C | Localize widget chrome (buttons, hints, placeholders, tooltips) into all 6 UI languages via a new WI18N table + WL() helper; thread lang into ChipSelector/RankedSelector/UserForm/TopicCards/QueriesWidget. Option values stay English (product taxonomy, stored in English in the brief). Fixes the all-English form inside a non-English chat. | src/lumen.jsx |
