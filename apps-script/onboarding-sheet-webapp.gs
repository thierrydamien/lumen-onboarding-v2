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
      // Resolve the URL up front, right after the copy exists. Everything the client
      // ultimately needs — the dashboard writeback and the confirmation email — uses
      // it, so computing it here (before the fills / flush) means a hiccup in a later
      // step can never leave us without the link to hand back.
      const url = copy.getUrl();
      // Claim the sessionId the moment the Sheet exists, BEFORE the (slow) fills,
      // share, email and Slack. A retry that arrives while this run is still filling
      // then sees the claim and returns THIS url instead of copying a second Sheet
      // and re-firing the branded email + Slack. Returning a still-filling Sheet's
      // url on a retry beats a duplicate document.
      if (idemKey) cache.put(idemKey, url, 21600); // 6h (cache max); real retries land far sooner
      // Log the actual tab names so a fill that can't find its tab is diagnosable
      // from the execution log (the fill matches tabs by name at runtime).
      console.log("Requirements Sheet tabs: " + ss.getSheets().map(function (s) { return s.getName(); }).join(" | "));

      // Populate each tab best-effort; one failure must not abort the rest.
      safe_(function () { fillBusinessObjectives_(ss, company); });
      safe_(function () { fillUsers_(ss, brief.users || []); });
      safe_(function () { fillTopics_(ss, brief.topics || []); });
      safe_(function () { fillChannels_(ss, brief.channels || []); });
      safe_(function () { fillReportsAlerts_(ss, brief.reports || [], brief.alerts || []); });
      safe_(function () { fillQueries_(ss, brief.queries || ""); });
      // flush() forces pending writes; the setValues above have already applied, so a
      // rare flush error must not abort the run before the writeback + email.
      safe_(function () { SpreadsheetApp.flush(); }, "flush");

      // Push the Sheet link to the dashboard's session store now (best-effort), so
      // the "Open Sheet" link appears even when the client aborted before receiving
      // the URL (a long/heavy session can outlast the client's 30s wait; this script
      // still finishes). Authenticated server-side with the shared secret.
      safe_(function () { updateSessionSheetUrl_(body.sessionId, url, body.dashboardOrigin); }, "session writeback");

      // Share the Sheet with the client AND send our own branded confirmation
      // email (thanks + link + next steps), localised to their onboarding
      // language. shareWithClient_ suppresses Google's generic "shared a document"
      // notification when the Drive advanced service is enabled; otherwise it falls
      // back to a notifying addEditor share so access always works. The email send
      // is best-effort (safe_) so it can never break the Sheet return.
      const clientEmail = company.email || body.clientEmail || "";
      if (clientEmail && /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(clientEmail)) {
        shareWithClient_(copy.getId(), clientEmail);
        safe_(function () {
          const first = String(company.contact || "").trim().split(/\s+/)[0] || "";
          sendClientEmail_(clientEmail, first, url, company.onboardingLanguage || "English");
        });
      }

      try { postCompletionSlack_(body, company, url); } catch (err) { /* non-fatal */ }

      return json_({ url: url });
    } finally {
      if (locked) { try { lock.releaseLock(); } catch (relErr) { /* ignore */ } }
    }
  } catch (err) {
    return json_({ error: String(err) });
  }
}

// ---- client confirmation email ---------------------------------------------
// After the Requirements Sheet is created we send the client OUR OWN branded
// email (thanks + link + next steps), localised to the language they did the
// onboarding in, instead of relying on Google's generic "shared a document"
// notification. Terms mirror the app (DE "Briefing", ES "resumen", AR "ملخص").
// No timeframe is promised for the follow-up, by design.

function htmlEsc_(s) {
  return String(s == null ? "" : s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

// Share the Sheet with the client WITHOUT Google's notification email when the
// Drive advanced service is enabled (Apps Script editor > Services > Drive API).
// Falls back to a notifying addEditor share when it isn't, so access always works
// (the client then gets Google's mail plus ours until the service is enabled).
function shareWithClient_(fileId, email) {
  try {
    if (typeof Drive !== "undefined" && Drive.Permissions) {
      if (Drive.Permissions.create) { Drive.Permissions.create({ role: "writer", type: "user", emailAddress: email }, fileId, { sendNotificationEmail: false }); return; }
      if (Drive.Permissions.insert) { Drive.Permissions.insert({ role: "writer", type: "user", value: email }, fileId, { sendNotificationEmails: false }); return; }
    }
  } catch (err) { /* fall through to notifying share */ }
  try { DriveApp.getFileById(fileId).addEditor(email); } catch (err) { /* non-fatal */ }
}

var CLIENT_EMAIL_I18N = {
  English: { subject: "Your Lumen setup brief", hi: "Hi",
    thanks: "Thanks for taking the time to walk through your Lumen setup with us.",
    intro: "We've turned everything you shared into a setup brief, and shared a copy with you:",
    button: "Open your setup brief",
    next: "What happens next: one of our consultants will reach out to book a short review call, where we'll finalise your setup together. Nothing is live yet; the brief is the starting point for that call.",
    extra: "If you think of anything else in the meantime (a competitor, a campaign, a colleague to add), bring it to the review call or share it with your Lumen contact.",
    signoff: "Thanks,\nYour team at Hootsuite" },
  French: { subject: "Votre brief de configuration Lumen", hi: "Bonjour",
    thanks: "Merci d'avoir pris le temps de préparer votre configuration Lumen avec nous.",
    intro: "Nous avons rassemblé tout ce que vous nous avez indiqué dans un brief de configuration, dont voici une copie :",
    button: "Ouvrir votre brief de configuration",
    next: "Prochaine étape : l'un de nos consultants vous contactera pour planifier un court appel de revue, au cours duquel nous finaliserons votre configuration ensemble. Rien n'est encore actif ; le brief est le point de départ de cet appel.",
    extra: "Si vous pensez à autre chose d'ici là (un concurrent, une campagne, un collègue à ajouter), apportez-le à l'appel de revue ou communiquez-le à votre contact Lumen.",
    signoff: "Merci,\nVotre équipe chez Hootsuite" },
  German: { subject: "Ihr Lumen Setup-Briefing", hi: "Hallo",
    thanks: "vielen Dank, dass Sie sich die Zeit genommen haben, Ihre Lumen-Einrichtung mit uns durchzugehen.",
    intro: "Wir haben alles, was Sie uns mitgeteilt haben, in einem Setup-Briefing zusammengefasst. Hier ist Ihre Kopie:",
    button: "Ihr Setup-Briefing öffnen",
    next: "Wie es weitergeht: Einer unserer Berater meldet sich bei Ihnen, um einen kurzen Review-Termin zu vereinbaren, bei dem wir Ihre Einrichtung gemeinsam finalisieren. Es ist noch nichts aktiv; das Briefing ist der Ausgangspunkt für diesen Termin.",
    extra: "Wenn Ihnen in der Zwischenzeit noch etwas einfällt (ein Wettbewerber, eine Kampagne, ein hinzuzufügender Kollege), bringen Sie es zum Review-Termin mit oder teilen Sie es Ihrem Lumen-Kontakt mit.",
    signoff: "Danke,\nIhr Team bei Hootsuite" },
  Spanish: { subject: "Su resumen de configuración de Lumen", hi: "Hola",
    thanks: "Gracias por dedicar tiempo a preparar su configuración de Lumen con nosotros.",
    intro: "Hemos reunido todo lo que nos indicó en un resumen de configuración, del que aquí tiene una copia:",
    button: "Abrir su resumen de configuración",
    next: "Qué sucede después: uno de nuestros consultores se pondrá en contacto para agendar una breve llamada de revisión, en la que finalizaremos su configuración juntos. Todavía no hay nada activo; el resumen es el punto de partida de esa llamada.",
    extra: "Si se le ocurre algo más mientras tanto (un competidor, una campaña, un colega para añadir), llévelo a la llamada de revisión o compártalo con su contacto de Lumen.",
    signoff: "Gracias,\nSu equipo en Hootsuite" },
  Italian: { subject: "Il tuo brief di configurazione Lumen", hi: "Ciao",
    thanks: "Grazie per aver dedicato del tempo a definire la tua configurazione Lumen con noi.",
    intro: "Abbiamo raccolto tutto ciò che ci hai indicato in un brief di configurazione, di cui trovi qui una copia:",
    button: "Apri il tuo brief di configurazione",
    next: "Cosa succede dopo: uno dei nostri consulenti ti contatterà per fissare una breve call di revisione, durante la quale finalizzeremo insieme la tua configurazione. Non è ancora attivo nulla; il brief è il punto di partenza per quella call.",
    extra: "Se ti viene in mente qualcos'altro nel frattempo (un concorrente, una campagna, un collega da aggiungere), portalo alla call di revisione o comunicalo al tuo contatto Lumen.",
    signoff: "Grazie,\nIl tuo team di Hootsuite" },
  Arabic: { subject: "ملخص إعداد Lumen الخاص بك", hi: "مرحبًا",
    thanks: "شكرًا لك على الوقت الذي خصصته لإعداد Lumen معنا.",
    intro: "لقد جمعنا كل ما شاركته معنا في ملخص إعداد، وهذه نسخة منه:",
    button: "افتح ملخص الإعداد الخاص بك",
    next: "ما الذي يحدث بعد ذلك: سيتواصل معك أحد مستشارينا لتحديد موعد مكالمة مراجعة قصيرة نُكمل خلالها إعدادك معًا. لا شيء مُفعّل بعد؛ الملخص هو نقطة البداية لتلك المكالمة.",
    extra: "إذا خطر لك أي شيء آخر في هذه الأثناء (منافس، حملة، زميل تريد إضافته)، فأحضره إلى مكالمة المراجعة أو شاركه مع جهة اتصالك في Lumen.",
    signoff: "شكرًا،\nفريقك في Hootsuite" }
};

function sendClientEmail_(email, firstName, url, lang) {
  var t = CLIENT_EMAIL_I18N[lang] || CLIENT_EMAIL_I18N.English;
  var rtl = (lang === "Arabic");
  var dir = rtl ? "rtl" : "ltr", align = rtl ? "right" : "left", comma = rtl ? "،" : ",";
  var greet = t.hi + (firstName ? " " + firstName : "");
  var html =
    '<div dir="' + dir + '" style="font-family:Arial,Helvetica,sans-serif;color:#0f172a;font-size:15px;line-height:1.6;max-width:560px;margin:0 auto;text-align:' + align + '">' +
      '<p style="margin:0 0 14px">' + htmlEsc_(greet) + comma + '</p>' +
      '<p style="margin:0 0 14px">' + htmlEsc_(t.thanks) + '</p>' +
      '<p style="margin:0 0 16px">' + htmlEsc_(t.intro) + '</p>' +
      '<p style="margin:0 0 20px"><a href="' + htmlEsc_(url) + '" style="display:inline-block;background:#7C3AED;color:#ffffff;text-decoration:none;font-weight:bold;padding:12px 22px;border-radius:8px">' + htmlEsc_(t.button) + '</a></p>' +
      '<p style="margin:0 0 14px">' + htmlEsc_(t.next) + '</p>' +
      '<p style="margin:0 0 18px">' + htmlEsc_(t.extra) + '</p>' +
      '<p style="margin:0;color:#475569">' + htmlEsc_(t.signoff).replace(/\n/g, "<br>") + '</p>' +
    '</div>';
  var plain = greet + comma + "\n\n" + t.thanks + "\n\n" + t.intro + "\n" + url + "\n\n" + t.next + "\n\n" + t.extra + "\n\n" + t.signoff;
  MailApp.sendEmail({ to: email, subject: t.subject, htmlBody: html, body: plain, name: "Lumen Onboarding" });
}

// Push the Sheet URL back to the Netlify session store so the dashboard shows the
// "Open Sheet" link even when the client timed out before receiving it. The client
// normally saves the URL, but a long/heavy session can outlast its 30s wait while
// this script keeps running to completion. Authenticated with the same SHARED_SECRET
// the chat proxy uses; the endpoint only accepts a sheetUrl update to an existing
// record. Best-effort: any failure is logged and ignored (the client path still
// works when it doesn't time out). The session endpoint origin is derived from the
// DASHBOARD_URL Script Property.
function updateSessionSheetUrl_(sessionId, url, originArg) {
  if (!sessionId || !url) return;
  const props = PropertiesService.getScriptProperties();
  const secret = props.getProperty("SHARED_SECRET");
  // Prefer the origin the caller (sheet.js) passed — it is always the site whose
  // dashboard/session store we must reach. Fall back to the DASHBOARD_URL property
  // for older callers that don't send it.
  const dash = String(originArg || props.getProperty("DASHBOARD_URL") || "");
  const origin = (dash.match(/^https?:\/\/[^\/]+/) || [])[0];
  if (!secret || !origin) { console.log("updateSessionSheetUrl_: no secret/origin (originArg='" + (originArg || "") + "')"); return; }
  const res = UrlFetchApp.fetch(origin + "/.netlify/functions/session", {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify({ secret: secret, id: sessionId, sheetUrl: url }),
    muteHttpExceptions: true,
  });
  const code = res.getResponseCode();
  if (code < 200 || code >= 300) console.log("session sheetUrl writeback HTTP " + code + ": " + res.getContentText());
}

// ---- populate helpers -------------------------------------------------------

function norm_(v) { return String(v == null ? "" : v).toLowerCase().replace(/\s+/g, " ").trim(); }
function safe_(fn, label) { try { fn(); } catch (e) { console.log("safe_ swallowed" + (label ? " [" + label + "]" : "") + ": " + e); } }

function sheetLike_(ss, re) {
  const sheets = ss.getSheets();
  for (let i = 0; i < sheets.length; i++) if (re.test(norm_(sheets[i].getName()))) return sheets[i];
  return null;
}

// Business Objectives: key/value layout. Match column A labels, write column B.
function fillBusinessObjectives_(ss, c) {
  const sh = sheetLike_(ss, /business objectives|launch requirements/) ||
    findTabByLabels_(ss, [/business objectives/, /preferred onboarding language/, /geographic markets/, /main point of contact/, /requirements completed by/]);
  if (!sh) { console.log("fillBusinessObjectives_: no tab matched (name or labels) — see the tab list above"); return; }
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
        // Resilient per-cell write: a reject-input dropdown (e.g. Time Zone) used to
        // throw and abort the WHOLE tab via safe_, dropping every later field —
        // which is why "Main Point of Contact" (the last rule) kept the template
        // placeholder. writeCell_ clears validation and retries, then skips one bad
        // cell without losing the rest.
        writeCell_(sh, r + 1, 2, rules[k][1]);
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

// Fallback tab finder: when a tab's NAME doesn't match (the template can be renamed
// — this is exactly why the Users tab wasn't filling), pick the tab whose CONTENT
// best matches the given header fields. Callers try the name match first; this is
// the safety net so a rename never silently drops a whole section.
function findTabByHeader_(ss, fieldMap) {
  const sheets = ss.getSheets();
  let best = null, bestScore = 0;
  for (let i = 0; i < sheets.length; i++) {
    const h = detectHeader_(sheets[i], fieldMap);
    const sc = h ? Object.keys(h.cols).length : 0;
    if (sc > bestScore) { bestScore = sc; best = sheets[i]; }
  }
  return best;
}

// Content fallback for key/value tabs (Business Objectives) that have no header
// row to detect. Picks the tab whose column A carries the most of the given label
// patterns. Requires >= 3 so it can't latch onto an unrelated tab.
function findTabByLabels_(ss, res) {
  const sheets = ss.getSheets();
  let best = null, bestScore = 0;
  for (let i = 0; i < sheets.length; i++) {
    const vals = sheets[i].getDataRange().getValues();
    const seen = {}; let score = 0;
    for (let r = 0; r < Math.min(vals.length, 40); r++) {
      const a = norm_(vals[r][0]);
      if (!a) continue;
      for (let k = 0; k < res.length; k++) if (!seen[k] && res[k].test(a)) { seen[k] = true; score++; }
    }
    if (score > bestScore) { bestScore = score; best = sheets[i]; }
  }
  return bestScore >= 3 ? best : null;
}

// Neutralise spreadsheet formula injection: a client cell value starting with
// = + - @ (or a leading control char) is coerced to text with a leading apostrophe,
// so "=IMPORTDATA(...)" / "=HYPERLINK(...)" is stored literally and never executes
// when the client-shared Sheet is opened or exported to CSV.
function cellSafe_(v) {
  if (typeof v !== "string") return v;
  return /^[=+\-@\t\r]/.test(v) ? "'" + v : v;
}
// Write one cell resiliently. The template's list tabs use "reject input" data
// validation (dropdowns for channel type, access rights, owned/public, etc.) and
// a value the dropdown doesn't allow makes setValue THROW. A merged non-anchor
// cell throws too. The old code let the first such throw abort the whole tab (via
// safe_), so only one partial row ever filled. Here each cell is independent: try
// to write; on failure clear that cell's validation and retry; still failing, log
// and skip — so every writable cell lands regardless of the others.
// Remove protections that would block writes to a data region: any RANGE
// protection overlapping [top..top+height-1] x [left..left+width-1], plus any
// whole-SHEET protection. Scoped to the fill region so header/instruction/formula
// protections outside it survive. Best-effort and fully guarded so it is a no-op
// where the Sheets service or getProtections isn't available (e.g. the test harness).
function clearRegionProtections_(sh, top, left, height, width) {
  const bottom = top + height - 1, right = left + width - 1;
  try {
    const rp = sh.getProtections(SpreadsheetApp.ProtectionType.RANGE) || [];
    for (let i = 0; i < rp.length; i++) {
      try {
        const rg = rp[i].getRange();
        const pT = rg.getRow(), pB = pT + rg.getNumRows() - 1, pL = rg.getColumn(), pR = pL + rg.getNumColumns() - 1;
        if (pT <= bottom && pB >= top && pL <= right && pR >= left) rp[i].remove();
      } catch (e) { /* one protection we can't inspect/remove — skip it */ }
    }
    const sps = sh.getProtections(SpreadsheetApp.ProtectionType.SHEET) || [];
    for (let j = 0; j < sps.length; j++) { try { sps[j].remove(); } catch (e) {} }
  } catch (e) { console.log("clearRegionProtections_ on '" + sh.getName() + "': " + e); }
}

// Grow the sheet so `lastRow` (1-based) is addressable. No-op when it already fits.
function ensureRows_(sh, lastRow) {
  try {
    const mr = sh.getMaxRows();
    if (lastRow > mr) sh.insertRowsAfter(mr, lastRow - mr);
  } catch (e) { console.log("ensureRows_ (" + lastRow + ") on '" + sh.getName() + "': " + e); }
}
function writeCell_(sh, row, col1, value) {
  if (value == null || value === "") return true;
  const rng = sh.getRange(row, col1);
  try { rng.setValue(cellSafe_(value)); return true; }
  catch (e1) {
    try { rng.setDataValidation(null); rng.setValue(cellSafe_(value)); return true; }
    catch (e2) { console.log("writeCell_ skip r" + row + " c" + col1 + " on '" + sh.getName() + "': " + e2); return false; }
  }
}
function writeRows_(sh, header, items, toRow) {
  if (!items.length) return;
  // BATCHED write. Per-cell setValue + a per-cell merge probe means one backend
  // round-trip PER CELL, which on a full brief is slow enough to push the run past
  // the timeout. Instead: probe merges ONCE, clear validation on the block ONCE
  // (so "reject input" dropdowns can't block the write), then read-modify-write the
  // whole block in a SINGLE setValues. Falls back to resilient per-cell only if the
  // batch is rejected. Flat templates (no merges) behave as one row per item, so
  // the previously-working path is unchanged.
  const cols = [];
  for (const f in header.cols) cols.push(header.cols[f]);
  const minCol = Math.min.apply(null, cols), maxCol = Math.max.apply(null, cols);
  const width = maxCol - minCol + 1;
  const firstDataRow = header.row + 2; // 1-based row just below the header

  // Some tabs give each item a multi-row MERGED slot; writing to consecutive rows
  // would stomp inside slot 1. Detect slots from one merge probe over the region.
  // Clamp the probe height to the sheet's real row count: an unclamped
  // items.length*6+6 overshoots the grid for ~13+ merged-slot items (e.g. Topics),
  // and getRange THROWS out of bounds — which the catch swallowed, dropping ALL
  // merge detection so every item after the first collapsed onto a slot interior.
  let merges = [];
  const probeRows = Math.max(1, Math.min(items.length * 6 + 6, sh.getMaxRows() - firstDataRow + 1));
  try { merges = sh.getRange(firstDataRow, minCol + 1, probeRows, width).getMergedRanges() || []; } catch (e) {}
  function slotAt(r) {
    let anchor = r, bottom = r;
    for (let k = 0; k < merges.length; k++) {
      const top = merges[k].getRow(), bot = top + merges[k].getNumRows() - 1;
      if (r >= top && r <= bot) { if (top < anchor) anchor = top; if (bot > bottom) bottom = bot; }
    }
    return { anchor: anchor, next: bottom + 1 };
  }
  const anchors = [];
  let row = firstDataRow, maxRow = firstDataRow;
  for (let i = 0; i < items.length; i++) {
    const s = slotAt(row);
    anchors.push(s.anchor);
    if (s.anchor > maxRow) maxRow = s.anchor;
    row = s.next;
  }
  // The template can end a section right at its header (e.g. the alert sub-header is
  // the last row), and Sheets does NOT auto-grow on write. Make sure the grid has
  // enough rows for everything we're about to write, or getValues/setValues throws
  // and the whole section silently drops.
  ensureRows_(sh, maxRow);
  const region = sh.getRange(firstDataRow, minCol + 1, maxRow - firstDataRow + 1, width);
  // A protected tab/range blocks every write (setValue/setValues throw) and
  // clearDataValidations does NOT touch protection — the signature is a tab that
  // stays empty with a clean run. The template's Users tab (it governs admin
  // access) is the likely one to carry protection, which a template copy inherits.
  // Clear protection over the DATA region (and any whole-sheet protection) so the
  // fill can write; scoped to the rows we fill, so header/instruction protections
  // are left alone. Best-effort — never breaks the fill.
  clearRegionProtections_(sh, firstDataRow, minCol + 1, maxRow - firstDataRow + 1, width);
  try { region.clearDataValidations(); } catch (e) {}
  let grid = null;
  try { grid = region.getValues(); } catch (e) { grid = null; }
  if (grid) {
    for (let i = 0; i < items.length; i++) {
      const rv = toRow(items[i]);
      const gi = anchors[i] - firstDataRow;
      if (gi < 0 || gi >= grid.length) continue;
      for (const field in header.cols) {
        const v = rv[field];
        if (v != null && v !== "") grid[gi][header.cols[field] - minCol] = cellSafe_(v);
      }
    }
    try {
      region.setValues(grid);
      console.log("writeRows_ '" + sh.getName() + "': wrote " + items.length + " item(s) (batched)");
      return;
    } catch (e) { console.log("writeRows_ '" + sh.getName() + "' batch rejected, per-cell fallback: " + e); }
  }
  let fails = 0;
  for (let i = 0; i < items.length; i++) {
    const rv = toRow(items[i]);
    for (const field in header.cols) if (!writeCell_(sh, anchors[i], header.cols[field] + 1, rv[field])) fails++;
  }
  console.log("writeRows_ '" + sh.getName() + "': wrote " + items.length + " item(s) per-cell, " + fails + " cell(s) skipped");
}

function fillUsers_(ss, users) {
  if (!users.length) return;
  // Find by tab name first, then fall back to CONTENT (in case the tab was renamed
  // — a rename is exactly what silently dropped the Users section before). The
  // content fields (e-mail + access + first/last name) don't occur together on any
  // other tab, so the fallback can't mis-latch.
  const sh = sheetLike_(ss, /user|team/) ||
    findTabByHeader_(ss, { firstName: /first ?name/, lastName: /last ?name|surname/, email: /e-?mail/, access: /access|permission|licen/ });
  if (!sh) { console.log("fillUsers_: no tab matched /user|team/ (and no content match) — see the tab list above"); return; }
  console.log("fillUsers_: tab='" + sh.getName() + "', received " + users.length + " user(s)");
  const header = detectHeader_(sh, { firstName: /first ?name|^name$/, lastName: /last ?name|surname/, role: /role|department|team/, email: /e-?mail/, access: /access|permission|licen/ });
  if (!header) { console.log("fillUsers_: header not detected on '" + sh.getName() + "'"); return; }
  writeRows_(sh, header, users, function (u) {
    return { firstName: u.firstName || "", lastName: u.lastName || "", role: u.role || "", email: u.email || "", access: u.access || "" };
  });
}

function fillTopics_(ss, topics) {
  if (!topics.length) return;
  const sh = sheetLike_(ss, /topic/) ||
    findTabByHeader_(ss, { group: /group/, keywords: /keyword/, urls: /url/, hashtags: /hashtag/ });
  if (!sh) { console.log("fillTopics_: no tab matched /topic/ (and no content match) — see the tab list above"); return; }
  const header = detectHeader_(sh, { group: /group/, name: /topic.*name|filter name/, keywords: /keyword/, urls: /url/, hashtags: /hashtag/, comments: /comment/ });
  if (!header) { console.log("fillTopics_: header not detected on '" + sh.getName() + "'"); return; }
  writeRows_(sh, header, topics, function (t) {
    return { group: t.group || "", name: t.name || "", keywords: t.keywords || "", urls: t.urls || "", hashtags: t.hashtags || "", comments: t.comments || t.rationale || "" };
  });
}

function fillChannels_(ss, channels) {
  if (!channels.length) return;
  const sh = sheetLike_(ss, /channel|social/) ||
    findTabByHeader_(ss, { author: /author/, type: /channel type/, url: /url/, owned: /owned|public/ });
  if (!sh) { console.log("fillChannels_: no tab matched /channel|social/ (and no content match) — see the tab list above"); return; }
  const header = detectHeader_(sh, { author: /author/, type: /channel type|^type$/, url: /url/, owned: /owned|public/ });
  if (!header) { console.log("fillChannels_: header not detected on '" + sh.getName() + "'"); return; }
  writeRows_(sh, header, channels, function (ch) {
    return { author: ch.author || "", type: ch.type || "", url: ch.url || "", owned: ownedLabel_(ch.owned) };
  });
}

// The brief carries owned as a boolean or the strings "true"/"false" (the model's
// CHANNELS marker), which Sheets renders as TRUE/FALSE. Map to the template's
// "Owned"/"Public" wording; pass through any already-worded value; blank stays blank.
function ownedLabel_(v) {
  if (v === true || v === "true" || v === "TRUE") return "Owned";
  if (v === false || v === "false" || v === "FALSE") return "Public";
  return v || "";
}

// Reports/Dashboards/Alerts: ONE tab with TWO stacked sections, each with its own
// sub-header row (and a generic legend row just above the first one):
//   "Dahboard / Report | Group Name | Name | Details (time frame, KPIs, etc.) | Comments"
//   "Alert | Name | Type | Details (time frame, KPIs, etc.) | Comments"
// Detect BOTH sub-headers. The template leaves only a couple of empty rows between
// them, so when there are more reports than that gap we insert rows just above the
// alert sub-header (shifting it, and its data area, down) rather than overwriting
// it. Then each section fills under its own header.
function fillReportsAlerts_(ss, reports, alerts) {
  reports = reports || []; alerts = alerts || [];
  if (!reports.length && !alerts.length) return;
  const sh = sheetLike_(ss, /report|dashboard|alert/) ||
    findTabByHeader_(ss, { name: /^name$/, type: /^type$/, details: /detail|time frame|kpi/, comments: /comment/ });
  if (!sh) { console.log("fillReportsAlerts_: no tab matched (and no content match) — see the tab list above"); return; }

  // Reports/dashboards sub-header: first col "Dahboard / Report" (template typo kept),
  // plus a "Group Name" column that the alert/legend rows don't have.
  const repH = detectHeader_(sh, { type: /dahboard|dashboard/, group: /group/, name: /^name$/, details: /detail|time frame|kpi/, comments: /comment/ });
  // Alerts sub-header: first col is exactly "Alert", and it has a "Type" column.
  const alH = detectHeader_(sh, { alert: /^alert$/, name: /^name$/, type: /^type$/, details: /detail|time frame|kpi/, comments: /comment/ });

  // Protect the alert sub-header from report overflow.
  if (reports.length && repH && alH && alH.row > repH.row) {
    const reportFirstData = repH.row + 2;   // 1-based first writable report row
    const alertHeaderRow = alH.row + 1;      // 1-based alert sub-header row
    const capacity = alertHeaderRow - reportFirstData;
    if (reports.length > capacity) {
      const need = reports.length - capacity;
      try { sh.insertRowsBefore(alertHeaderRow, need); alH.row += need; }
      catch (e) { console.log("fillReportsAlerts_ insertRows failed: " + e); }
    }
  }

  if (reports.length && repH) {
    // objective -> "Group Name" (the sub-header renames the legend's "Main objective"
    // column to Group Name). The "Dahboard / Report" type column is left for the team.
    writeRows_(sh, repH, reports, function (r) {
      return { group: r.objective || "", name: r.name || "", details: r.details || "", comments: r.comments || "" };
    });
  } else if (reports.length) {
    console.log("fillReportsAlerts_: reports sub-header not detected on '" + sh.getName() + "'");
  }

  if (alerts.length && alH) {
    writeRows_(sh, alH, alerts, function (a) {
      return { name: a.name || "", type: a.type || "", details: a.details || "", comments: a.comments || "" };
    });
  } else if (alerts.length) {
    console.log("fillReportsAlerts_: alert sub-header not detected on '" + sh.getName() + "'");
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
  for (let i = 0; i < lines.length; i++) sh.getRange(4 + i, 1).setValue(cellSafe_(lines[i]));
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
  // Escape name/contact: Slack parses broadcast (<!channel>) and link (<url|text>)
  // syntax in the `text` field too, so a company literally named "<!channel>" would
  // fire a channel-wide ping. The Block Kit fields are already slackEsc_'d via
  // slackField_; this closes the same gap in the fallback string.
  const fallback = "Lumen onboarding completed — " + slackEsc_(name) + " (" + slackEsc_(contact) + "). Topics: " + topics + " · Users: " + users;

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
      ? "Matched: *" + slackEsc_(match.accountName) + "*" + sourceNote + confidenceNote + "\n" + mentions + coreNote
      : "Matched: *" + slackEsc_(match.accountName) + "*" + sourceNote + " — no IC or TAM assigned yet." + coreNote;
  } else {
    // Escape the client-entered company name: an unescaped "<!channel>" here would
    // fire a real Slack broadcast (the base alert already escapes via slackField_).
    text = "No pipeline match for *" + slackEsc_(clientName) + "* — please assign manually." + (escalate ? "\n" + escalate : "");
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
