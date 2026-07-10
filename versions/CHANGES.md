# Lumen onboarding — fix log

Each version below is one fix, applied on top of the previous. Full snapshots of
the files changed in each fix live in `versions/vNN-<slug>/`. `v00-baseline` is
the untouched starting point. Git commits on branch `claude/lumen-onboarding-fixes`
carry the same steps for diffing.

Severity key: B = blocker, C = confusing, M = minor.

| v | Sev | Fix | Files |
|---|-----|-----|-------|
| 01 | M | Remove dead `getSP()` (full duplicate of the server system prompt) and unused `_unusedLegacyClientAPI` stub. Kills prompt drift + shrinks bundle. | src/lumen.jsx |
