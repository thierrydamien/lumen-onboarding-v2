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

function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || "{}");
    const expected = PropertiesService.getScriptProperties().getProperty("SHARED_SECRET");
    if (!expected || body.secret !== expected) return json_({ error: "unauthorized" });

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
    SpreadsheetApp.flush();

    const url = copy.getUrl();

    // Share with the client as editor (sends them a Google notification email).
    const clientEmail = company.email || body.clientEmail || "";
    if (clientEmail && /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(clientEmail)) {
      try { DriveApp.getFileById(copy.getId()).addEditor(clientEmail); } catch (err) { /* non-fatal */ }
    }

    try { postCompletionSlack_(body, company, url); } catch (err) { /* non-fatal */ }

    return json_({ url: url });
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

  UrlFetchApp.fetch("https://slack.com/api/chat.postMessage", {
    method: "post",
    contentType: "application/json",
    headers: { Authorization: "Bearer " + token },
    payload: JSON.stringify({ channel: channel, text: fallback, blocks: blocks }),
    muteHttpExceptions: true,
  });
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

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
