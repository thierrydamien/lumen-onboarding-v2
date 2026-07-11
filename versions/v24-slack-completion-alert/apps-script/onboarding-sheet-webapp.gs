/**
 * Lumen onboarding -> Google Sheet (Apps Script Web App).
 *
 * WHY: this runs as YOUR Google account (like the survey script), so it can create
 * files in the Proserv folder on your quota. No service account / OAuth / domain-
 * wide delegation needed. The Netlify `sheet.js` function POSTs the brief here.
 *
 * DEPLOY:
 *   1. script.google.com > New project. Paste this file in.
 *   2. Editor > Services (+) > add "Drive API" (advanced service; used for the
 *      XLSX -> Google Sheet conversion below).
 *   3. Project Settings > Script Properties > add SHARED_SECRET = <a long random
 *      string>. Put the SAME value in Netlify env APPS_SCRIPT_SECRET.
 *   4. (Optional Slack alert on completion) add Script Property SLACK_BOT_TOKEN =
 *      <your xoxb- token> and, if different from the default below, SLACK_CHANNEL.
 *      The bot must be a member of that channel. Keep the token here in Script
 *      Properties, never hardcoded.
 *   5. Deploy > New deployment > type Web app. Execute as: Me. Who has access:
 *      Anyone. Copy the /exec URL into Netlify env APPS_SCRIPT_WEBAPP_URL.
 *
 * SECURITY: "Anyone" means anyone with the URL can invoke it, so every request is
 * rejected unless it carries the shared secret. Keep the URL + secret server-side
 * (they live in Netlify env; the browser never sees them). Do NOT hardcode secrets
 * here — use Script Properties.
 */

// Destination folder (same one the survey script writes to).
const DEST_FOLDER_ID = "1BacQuILUAGSKcuUzEwY-iVCh37gt72rY";
// Slack channel for the completion alert (override via Script Property SLACK_CHANNEL).
const DEFAULT_SLACK_CHANNEL = "C097154H39N";

function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || "{}");

    const expected = PropertiesService.getScriptProperties().getProperty("SHARED_SECRET");
    if (!expected || body.secret !== expected) return json_({ error: "unauthorized" });
    if (!body.xlsxBase64) return json_({ error: "missing_xlsx" });

    const name = String(body.filename || "Lumen Setup Brief").replace(/\.xlsx$/i, "");
    const blob = Utilities.newBlob(
      Utilities.base64Decode(body.xlsxBase64),
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      name + ".xlsx"
    );

    // Convert the uploaded XLSX into a native Google Sheet inside the folder.
    const file = Drive.Files.insert(
      { title: name, mimeType: "application/vnd.google-apps.spreadsheet", parents: [{ id: DEST_FOLDER_ID }] },
      blob,
      { convert: true, supportsAllDrives: true }
    );

    // Share with the client as editor (this sends them a Google notification email).
    if (body.clientEmail && /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(body.clientEmail)) {
      try { DriveApp.getFileById(file.id).addEditor(body.clientEmail); } catch (err) { /* non-fatal */ }
    }

    const url = "https://docs.google.com/spreadsheets/d/" + file.id + "/edit";

    // Slack alert: a client finished the assistant and a brief was created.
    // Non-fatal — a Slack failure must not fail the Sheet response.
    try { postCompletionSlack_(body, url); } catch (err) { /* non-fatal */ }

    return json_({ url: url });
  } catch (err) {
    return json_({ error: String(err) });
  }
}

// Posts the onboarding-completed alert. No-op unless SLACK_BOT_TOKEN is set.
function postCompletionSlack_(body, url) {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty("SLACK_BOT_TOKEN");
  if (!token) return;
  const channel = props.getProperty("SLACK_CHANNEL") || DEFAULT_SLACK_CHANNEL;

  const contact = [body.contactName, body.clientEmail].filter(Boolean).join(" · ");
  const counts = (body.topicsCount != null || body.usersCount != null)
    ? "Topics: " + (body.topicsCount != null ? body.topicsCount : "?") + " · Users: " + (body.usersCount != null ? body.usersCount : "?") + "\n"
    : "";
  const text =
    ":white_check_mark: *Lumen onboarding completed* — a setup brief was created.\n" +
    "Client: *" + (body.company || "(unnamed)") + "*\n" +
    (contact ? "Contact: " + contact + "\n" : "") +
    counts +
    "<" + url + "|Open the setup brief>";

  UrlFetchApp.fetch("https://slack.com/api/chat.postMessage", {
    method: "post",
    contentType: "application/json",
    headers: { Authorization: "Bearer " + token },
    payload: JSON.stringify({ channel: channel, text: text }),
    muteHttpExceptions: true,
  });
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
