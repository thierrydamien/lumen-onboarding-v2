/**
 * Lumen onboarding -> Google Sheet (Apps Script Web App).
 *
 * WHY: runs as YOUR Google account (like the survey script), so it can copy the
 * master template into the Proserv folder on your quota. No service account /
 * OAuth / domain-wide delegation. It COPIES the template (exact tabs/formatting)
 * and fills in the client's responses, then shares + posts to Slack.
 *
 * DEPLOY:
 *   1. script.google.com > New project. Paste this file in.
 *   2. Project Settings > Script Properties:
 *        SHARED_SECRET   = <a long random string>   (same value in Netlify APPS_SCRIPT_SECRET)
 *        SLACK_BOT_TOKEN = <your xoxb- token>        (optional; enables the completion alert)
 *        SLACK_CHANNEL   = <channel id>              (optional; defaults below)
 *        DASHBOARD_URL   = https://<site>/dashboard  (optional; adds a "View full
 *                          session" deep-link to the Slack alert, e.g.
 *                          https://onboarding.hootsuite.com/dashboard)
 *      Optional — IC/TAM @mentions in a threaded reply (reused from the survey
 *      script). Kept in Script Properties, NOT in this file, so the staff roster
 *      never lives in the repo:
 *        PIPELINE_SHEET_ID = <tracker spreadsheet id>   (enables the lookup)
 *        SLACK_IDS_JSON    = {"full name":"U012...", …} (name -> Slack user id)
 *        SLACK_ESCALATION  = U012 U345                  (space/comma-separated ids
 *                            pinged for TW Core clients and no-match cases)
 *   3. Deploy > New deployment > type Web app. Execute as: Me. Who has access:
 *      Anyone. Copy the /exec URL into Netlify env APPS_SCRIPT_WEBAPP_URL.
 *
 * SECURITY: "Anyone" access is why every request must carry the shared secret.
 * Keep the URL + secret + Slack token server-side (Netlify env / Script
 * Properties); never hardcode secrets here.
 *
 * NOTE: the Business Objectives / Users / Topics / Social channels tabs are
 * populated by matching the template's own labels and headers at runtime, so it
 * survives minor template edits. The Reports/Dashboards/Alerts tab is filled
 * best-effort (its layout could not be fully verified) and may need a manual
 * check; adjust REPORT_COLS / ALERT_COLS below once confirmed.
 */

const TEMPLATE_ID = "1VC7nIstJw-H4XMqVPe8stQRIwV88S2bZv6sZGgNPus0"; // "do not modify" master (copied, never edited)
const DEST_FOLDER_ID = "1BacQuILUAGSKcuUzEwY-iVCh37gt72rY";
const DEFAULT_SLACK_CHANNEL = "C097154H39N";

// Pipeline-tracker lookup for the IC/TAM @mention reply (optional; configured via
// Script Properties, see the header). Structural layout of the tracker only —
// no employee data lives here. Adjust if the tracker's columns/tabs differ.
const PIPELINE_TABS = ["ClosedWon", "Pipeline"];                    // searched in order
const PIPELINE_COL = { account: 1, talkwalker: 21, ic: 22, tam: 23 }; // cols B, V, W, X (0-based)
const PIPELINE_MIN_SCORE = 0.4;  // below this: treat as no match
const MENTION_MIN_SCORE = 0.7;   // below this: "low confidence" note / fall back to plain name

function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || "{}");
    const expected = PropertiesService.getScriptProperties().getProperty("SHARED_SECRET");
    if (!expected || body.secret !== expected) return json_({ error: "unauthorized" });

    // Idempotency. A slow copy/populate can outlast the client's 45s timeout; the
    // user then re-sends (or a proxy retries) with the SAME sessionId. Without a
    // guard, every retry copies a NEW Sheet AND re-fires the Slack alert. Keyed by
    // sessionId in the script cache (auto-expiring, so no unbounded growth): once a
    // Sheet exists for a sessionId, return that URL and do nothing else. A short
    // script lock serialises the check-and-claim, so an overlapping retry (the
    // first request still running server-side after the client aborted at 45s)
    // can't slip through; purely sequential retries are covered even if the lock
    // can't be taken. No sessionId (older client) => no dedup, same as before.
    const cache = CacheService.getScriptCache();
    const idemKey = body.sessionId ? "sheet_done_" + body.sessionId : null;
    const lock = LockService.getScriptLock();
    let locked = false;
    try { lock.waitLock(25000); locked = true; } catch (lockErr) { /* proceed unlocked */ }
    try {
      if (idemKey) {
        const prior = cache.get(idemKey);
        if (prior) return json_({ url: prior, idempotent: true });
      }

      const brief = body.brief || {};
      const company = brief.company || {};
      const clientName = company.name || body.company || "Client";
      const title = clientName + " - Lumen Requirements Document";

      // Copy the master template into the destination folder (never touch the master).
      const copy = DriveApp.getFileById(TEMPLATE_ID).makeCopy(title, DriveApp.getFolderById(DEST_FOLDER_ID));
      const ss = SpreadsheetApp.openById(copy.getId());

      // Populate each tab best-effort; one failure must not abort the rest.
      safe_(function () { fillBusinessObjectives_(ss, company); });
      safe_(function () { fillUsers_(ss, brief.users || []); });
      safe_(function () { fillTopics_(ss, brief.topics || []); });
      safe_(function () { fillChannels_(ss, brief.channels || []); });
      safe_(function () { fillReportsAlerts_(ss, brief.reports || [], brief.alerts || []); });
      safe_(function () { fillQueries_(ss, brief.queries || ""); });
      SpreadsheetApp.flush();

      const url = copy.getUrl();

      // Share with the client as editor (sends them a Google notification email).
      const clientEmail = company.email || body.clientEmail || "";
      if (clientEmail && /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(clientEmail)) {
        try { DriveApp.getFileById(copy.getId()).addEditor(clientEmail); } catch (err) { /* non-fatal */ }
      }

      // Claim the sessionId BEFORE the Slack post: if the post is slow, a retry
      // arriving now sees the claim and won't create a second Sheet or re-alert.
      if (idemKey) cache.put(idemKey, url, 21600); // 6h (cache max); real retries land far sooner

      try { postCompletionSlack_(body, company, url); } catch (err) { /* non-fatal */ }

      return json_({ url: url });
    } finally {
      if (locked) { try { lock.releaseLock(); } catch (relErr) { /* ignore */ } }
    }
  } catch (err) {
    return json_({ error: String(err) });
  }
}

// ---- populate helpers -------------------------------------------------------

function norm_(v) { return String(v == null ? "" : v).toLowerCase().replace(/\s+/g, " ").trim(); }
function safe_(fn) { try { fn(); } catch (e) { /* per-tab best-effort */ } }

function sheetLike_(ss, re) {
  const sheets = ss.getSheets();
  for (let i = 0; i < sheets.length; i++) if (re.test(norm_(sheets[i].getName()))) return sheets[i];
  return null;
}

// Business Objectives: key/value layout. Match column A labels, write column B.
function fillBusinessObjectives_(ss, c) {
  const sh = sheetLike_(ss, /business objectives/);
  if (!sh) return;
  const vals = sh.getDataRange().getValues();
  const contact = [c.contact, c.email].filter(Boolean).join(" – ");
  const rules = [
    [/^date\b/, todayStr_()],
    [/requirements completed by/, c.contact || ""],
    [/relevant geographic markets/, c.markets || ""],
    [/key languages/, c.languages || ""],
    [/business objectives/, c.objectives || ""],
    [/objective detail/, c.objectiveDetails || ""],
    [/^industry|industry\/vertical|industry sector/, c.industry || ""],
    [/planned use cases/, c.useCase || ""],
    [/preferred onboarding language/, c.onboardingLanguage || "English"],
    [/preferred time zone/, c.timezone || ""],
    [/teams.*departments|departments.*platform/, c.teams || ""],
    [/main point of contact/, contact],
  ];
  for (let r = 0; r < vals.length; r++) {
    const a = norm_(vals[r][0]);
    if (!a) continue;
    for (let k = 0; k < rules.length; k++) {
      if (rules[k][0].test(a)) {
        if (rules[k][1] !== "") sh.getRange(r + 1, 2).setValue(rules[k][1]);
        break;
      }
    }
  }
}

// Detect a header row (the row matching the most field regexes) and return
// { row: 0-based header index, cols: {field: colIndex} }. null if not confident.
function detectHeader_(sh, fieldMap, maxScan) {
  const vals = sh.getDataRange().getValues();
  let best = null, bestScore = 0;
  const scan = Math.min(vals.length, maxScan || 25);
  for (let r = 0; r < scan; r++) {
    const cols = {}; let score = 0;
    for (let cIdx = 0; cIdx < vals[r].length; cIdx++) {
      const cell = norm_(vals[r][cIdx]);
      if (!cell) continue;
      for (const field in fieldMap) {
        if (!(field in cols) && fieldMap[field].test(cell)) { cols[field] = cIdx; score++; }
      }
    }
    if (score > bestScore) { bestScore = score; best = { row: r, cols: cols }; }
  }
  return bestScore >= 2 ? best : null;
}

function writeRows_(sh, header, items, toRow) {
  const firstDataRow = header.row + 2; // 1-based row just below the header
  for (let i = 0; i < items.length; i++) {
    const rowVals = toRow(items[i]);
    for (const field in header.cols) {
      const v = rowVals[field];
      if (v != null && v !== "") sh.getRange(firstDataRow + i, header.cols[field] + 1).setValue(v);
    }
  }
}

function fillUsers_(ss, users) {
  if (!users.length) return;
  const sh = sheetLike_(ss, /users/);
  if (!sh) return;
  const header = detectHeader_(sh, { firstName: /first name/, lastName: /last name/, role: /role|department/, email: /e-?mail/, access: /access/ });
  if (!header) return;
  writeRows_(sh, header, users, function (u) {
    return { firstName: u.firstName || "", lastName: u.lastName || "", role: u.role || "", email: u.email || "", access: u.access || "" };
  });
}

function fillTopics_(ss, topics) {
  if (!topics.length) return;
  const sh = sheetLike_(ss, /topic/);
  if (!sh) return;
  const header = detectHeader_(sh, { group: /group/, name: /topic.*name|filter name/, keywords: /keyword/, urls: /url/, hashtags: /hashtag/, comments: /comment/ });
  if (!header) return;
  writeRows_(sh, header, topics, function (t) {
    return { group: t.group || "", name: t.name || "", keywords: t.keywords || "", urls: t.urls || "", hashtags: t.hashtags || "", comments: t.comments || t.rationale || "" };
  });
}

function fillChannels_(ss, channels) {
  if (!channels.length) return;
  const sh = sheetLike_(ss, /channel|social/);
  if (!sh) return;
  const header = detectHeader_(sh, { author: /author/, type: /channel type|^type$/, url: /url/, owned: /owned|public/ });
  if (!header) return;
  writeRows_(sh, header, channels, function (ch) {
    return { author: ch.author || "", type: ch.type || "", url: ch.url || "", owned: ch.owned || "" };
  });
}

// Best-effort: the Reports/Dashboards/Alerts layout could not be fully verified.
// Reports detected by name/objective/details/comments; alerts by name/type/details.
function fillReportsAlerts_(ss, reports, alerts) {
  const sh = sheetLike_(ss, /report|dashboard|alert/);
  if (!sh) return;
  if (reports.length) {
    const rh = detectHeader_(sh, { name: /report|dashboard|^name|title/, objective: /objective/, details: /detail|kpi|time/, comments: /comment/ });
    if (rh) writeRows_(sh, rh, reports, function (r) {
      return { name: r.name || "", objective: r.objective || "", details: r.details || "", comments: r.comments || "" };
    });
  }
  if (alerts.length) {
    const ah = detectHeader_(sh, { name: /alert|^name/, type: /type|trigger/, details: /detail|kpi|threshold|time/, comments: /comment/ });
    if (ah) writeRows_(sh, ah, alerts, function (a) {
      return { name: a.name || "", type: a.type || "", details: a.details || "", comments: a.comments || "" };
    });
  }
}

// Migrated queries: the master template has no dedicated tab, so the client's
// pasted/uploaded existing setup was dropped on this path. Add a tab for it when
// present (mirrors the fallback XLSX's "Migrated queries" sheet).
function fillQueries_(ss, queries) {
  const q = String(queries == null ? "" : queries).trim();
  if (!q || q === "__skip__") return;
  const sh = ss.insertSheet("Migrated queries");
  sh.getRange(1, 1).setValue("Migrated queries (client's original content, as submitted)");
  sh.getRange(2, 1).setValue("Reference for rebuilding queries in Lumen. May contain untranslated syntax from the client's previous tool.");
  const lines = q.split("\n");
  for (let i = 0; i < lines.length; i++) sh.getRange(4 + i, 1).setValue(lines[i]);
}

function todayStr_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "GMT", "yyyy-MM-dd");
}

// ---- Slack ------------------------------------------------------------------

function postCompletionSlack_(body, company, url) {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty("SLACK_BOT_TOKEN");
  if (!token) return;
  const channel = props.getProperty("SLACK_CHANNEL") || DEFAULT_SLACK_CHANNEL;

  const name = company.name || body.company || "(unnamed)";
  const contactName = body.contactName || company.contact || "";
  const email = company.email || body.clientEmail || "";
  const contact = [contactName, email].filter(Boolean).join(" · ") || "—";
  const topics = body.topicsCount != null ? String(body.topicsCount) : "—";
  const users = body.usersCount != null ? String(body.usersCount) : "—";

  // Two-column fields (Block Kit renders these 2-up), richest first.
  const fields = [slackField_("Client", name), slackField_("Contact", contact),
    slackField_("Topics", topics), slackField_("Users", users)];
  if (company.industry) fields.push(slackField_("Industry", company.industry));
  if (company.markets)  fields.push(slackField_("Markets", company.markets));
  const topObjective = firstObjective_(company.objectives);
  if (topObjective)     fields.push(slackField_("Top objective", topObjective));

  // Context line: only truthful bits. "Prepared by" appears only if the caller
  // actually supplies it (the sales page does not capture a signed-in user yet);
  // the edit-access note holds whenever the client was shared on the Sheet.
  const context = [];
  if (body.preparedBy) context.push("Link prepared by " + slackEsc_(body.preparedBy));
  if (email) context.push("Client has edit access to the Sheet until the review call");

  // Deep-link to the client's session detail in the Proserv dashboard (full
  // brief + handoff + notes). Needs the DASHBOARD_URL Script Property and the
  // sessionId the client passes through; omitted if either is missing. The
  // dashboard is token-gated, so the link opens but still asks for the token.
  const dashUrl = props.getProperty("DASHBOARD_URL");
  const sessionLink = (dashUrl && body.sessionId)
    ? dashUrl + (dashUrl.indexOf("?") === -1 ? "?" : "&") + "id=" + encodeURIComponent(body.sessionId)
    : "";
  const links = [];
  if (url) links.push("<" + url + "|:page_facing_up: Open the requirements document>");
  if (sessionLink) links.push("<" + sessionLink + "|:mag: View full session>");

  const blocks = [
    { type: "section", text: { type: "mrkdwn", text: ":white_check_mark: *Lumen onboarding completed* — a requirements document was created." } },
    { type: "section", fields: fields.slice(0, 10) },
  ];
  if (links.length) blocks.push({ type: "section", text: { type: "mrkdwn", text: links.join("   ·   ") } });
  if (context.length) blocks.push({ type: "context", elements: [{ type: "mrkdwn", text: context.join("   ·   ") }] });

  // Plain-text fallback for notifications / accessibility (required by Slack).
  const fallback = "Lumen onboarding completed — " + name + " (" + contact + "). Topics: " + topics + " · Users: " + users;

  const ts = slackPost_(token, { channel: channel, text: fallback, blocks: blocks });

  // Threaded reply: match the client against the pipeline tracker and @mention the
  // assigned IC/TAM (reused from the survey script). Best-effort — a lookup failure
  // or missing config never affects the base alert posted above.
  if (ts) safe_(function () { postAssigneeMentions_(token, channel, ts, name); });
}

// Post a Slack message; return its ts (the thread anchor) or null on failure.
function slackPost_(token, payload) {
  const res = UrlFetchApp.fetch("https://slack.com/api/chat.postMessage", {
    method: "post", contentType: "application/json",
    headers: { Authorization: "Bearer " + token },
    payload: JSON.stringify(payload), muteHttpExceptions: true,
  });
  const data = JSON.parse(res.getContentText() || "{}");
  if (!data.ok) Logger.log("Slack post failed: " + res.getContentText());
  return data.ts || null;
}

// A Block Kit two-column field: bold label above an escaped value.
function slackField_(label, value) {
  return { type: "mrkdwn", text: "*" + label + ":*\n" + slackEsc_(value) };
}
// Escape the three characters Slack treats specially in mrkdwn, so a client
// value (company name, markets, etc.) can't break the block or inject a link.
function slackEsc_(s) {
  return String(s == null ? "" : s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}
// "1. Reputation Management, 2. Competitive Intelligence" -> "Reputation Management".
function firstObjective_(objectives) {
  if (!objectives) return "";
  return String(objectives).split(",")[0].replace(/^\s*\d+[\.\)]\s*/, "").trim();
}

// ---- IC/TAM @mention reply (ported from the survey script) ------------------

// Post the threaded reply that @mentions the assigned IC/TAM for this client.
// No-op unless PIPELINE_SHEET_ID is set. Client data is read from the tracker;
// the name->id roster comes from SLACK_IDS_JSON.
function postAssigneeMentions_(token, channel, threadTs, clientName) {
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty("PIPELINE_SHEET_ID") || !clientName) return;

  let match = null;
  for (let i = 0; i < PIPELINE_TABS.length && !match; i++) match = findBestMatch_(PIPELINE_TABS[i], clientName);

  const escalate = escalationMentions_();
  let text;
  if (match) {
    const isCore = match.talkwalker && String(match.talkwalker).trim().toLowerCase() === "core";
    const icMention = match.ic ? getSlackMention_(match.ic) + " (IC)" : null;
    const tamMention = match.tam ? getSlackMention_(match.tam) + " (TAM)" : null;
    const mentions = [icMention, tamMention].filter(Boolean).join("  ");
    const sourceNote = match.source === "ClosedWon" ? " _(existing client)_" : "";
    const confidenceNote = match.score < MENTION_MIN_SCORE ? " _(low confidence — please verify)_" : "";
    const coreNote = isCore ? ("\n:tw-core: *TW Core client*" + (escalate ? " — " + escalate + " please take note." : " — please take note.")) : "";
    text = mentions
      ? "Matched: *" + match.accountName + "*" + sourceNote + confidenceNote + "\n" + mentions + coreNote
      : "Matched: *" + match.accountName + "*" + sourceNote + " — no IC or TAM assigned yet." + coreNote;
  } else {
    text = "No pipeline match for *" + clientName + "* — please assign manually." + (escalate ? "\n" + escalate : "");
  }
  slackPost_(token, { channel: channel, text: text, thread_ts: threadTs });
}

// Best fuzzy match for clientName in one tracker tab; null below PIPELINE_MIN_SCORE.
function findBestMatch_(tabName, clientName) {
  const sheetId = PropertiesService.getScriptProperties().getProperty("PIPELINE_SHEET_ID");
  if (!sheetId) return null;
  let sheet;
  try { sheet = SpreadsheetApp.openById(sheetId).getSheetByName(tabName); } catch (e) { return null; }
  if (!sheet) return null;
  const data = sheet.getDataRange().getValues();
  let best = null, bestScore = 0;
  for (let i = 1; i < data.length; i++) {
    const account = data[i][PIPELINE_COL.account];
    if (!account) continue;
    const score = similarity_(clientName, String(account));
    if (score > bestScore) {
      bestScore = score;
      best = { accountName: account, talkwalker: data[i][PIPELINE_COL.talkwalker], ic: data[i][PIPELINE_COL.ic], tam: data[i][PIPELINE_COL.tam], score: score, source: tabName };
    }
  }
  return best && best.score >= PIPELINE_MIN_SCORE ? best : null;
}

// name -> "<@Uxxxx>" via the roster, with a fuzzy fallback; plain name if no
// confident match (so the reply is still readable).
function getSlackMention_(name) {
  if (!name) return "";
  const ids = slackIds_();
  const normalized = String(name).toLowerCase().trim();
  if (ids[normalized]) return "<@" + ids[normalized] + ">";
  let bestKey = null, bestScore = 0;
  for (const key in ids) {
    const score = similarity_(normalized, key);
    if (score > bestScore) { bestScore = score; bestKey = key; }
  }
  if (bestScore >= MENTION_MIN_SCORE && bestKey) return "<@" + ids[bestKey] + ">";
  return String(name);
}

let _slackIdsCache = null;
function slackIds_() {
  if (_slackIdsCache) return _slackIdsCache;
  try { _slackIdsCache = JSON.parse(PropertiesService.getScriptProperties().getProperty("SLACK_IDS_JSON") || "{}"); }
  catch (e) { _slackIdsCache = {}; }
  return _slackIdsCache;
}

function escalationMentions_() {
  const raw = PropertiesService.getScriptProperties().getProperty("SLACK_ESCALATION") || "";
  return raw.split(/[\s,]+/).filter(function (x) { return x; }).map(function (x) { return "<@" + x + ">"; }).join(" ");
}

// Fuzzy string similarity (token overlap + edit distance + substring bonus).
function fuzzyNorm_(s) {
  if (!s) return "";
  return String(s).toLowerCase()
    .replace(/[,\.'"()\-]/g, "")
    .replace(/\b(inc|llc|ltd|corp|corporation|company|co|the|and|of|for|group|solutions|services|international|global|technologies|technology|association|university|national)\b/g, "")
    .replace(/\s+/g, " ").trim();
}
function levenshtein_(a, b) {
  const m = a.length, n = b.length, dp = [];
  for (let i = 0; i <= m; i++) {
    dp[i] = [i];
    for (let j = 1; j <= n; j++) {
      dp[i][j] = i === 0 ? j : (a[i - 1] === b[j - 1] ? dp[i - 1][j - 1] : 1 + Math.min(dp[i - 1][j], dp[i][j - 1], dp[i - 1][j - 1]));
    }
  }
  return dp[m][n];
}
function similarity_(a, b) {
  const na = fuzzyNorm_(a), nb = fuzzyNorm_(b);
  if (!na || !nb) return 0;
  const uniqA = {}, uniqB = {};
  na.split(" ").forEach(function (x) { if (x) uniqA[x] = 1; });
  nb.split(" ").forEach(function (x) { if (x) uniqB[x] = 1; });
  const aKeys = Object.keys(uniqA), bKeys = Object.keys(uniqB);
  const inter = aKeys.filter(function (x) { return uniqB[x]; }).length;
  const tokenScore = inter / Math.max(aKeys.length, bKeys.length, 1);
  const subBonus = (na.indexOf(nb) !== -1 || nb.indexOf(na) !== -1) ? 0.3 : 0;
  const dist = levenshtein_(na, nb);
  const editScore = 1 - dist / Math.max(na.length, nb.length, 1);
  return Math.min(1, tokenScore * 0.5 + editScore * 0.3 + subBonus);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
