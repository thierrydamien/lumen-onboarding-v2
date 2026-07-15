// Google Sheet generation (Netlify Functions v2).
//   POST /.netlify/functions/sheet  { xlsxBase64, filename, clientEmail?, company? }
//     -> { url }   (a Google Sheet converted from the brief's XLSX)
//
// SETUP (inert until one path below is configured).
//   Path D - Apps Script Web App (PREFERRED here; runs as a real Google account,
//     so it writes into that account's Drive folder with no service account /
//     OAuth / delegation). See apps-script/onboarding-sheet-webapp.gs.
//       APPS_SCRIPT_WEBAPP_URL, APPS_SCRIPT_SECRET
//   The Google-API paths below need GOOGLE_DRIVE_FOLDER_ID for the target folder.
//   Path A - OAuth as a real user (writes into that user's My Drive folder, on
//     their quota; no Workspace admin needed):
//       GOOGLE_OAUTH_CLIENT_ID, GOOGLE_OAUTH_CLIENT_SECRET, GOOGLE_OAUTH_REFRESH_TOKEN
//   Path B - Service account + domain-wide delegation (impersonate a real user;
//     needs a Workspace admin to authorize the SA for the Drive scope):
//       GOOGLE_SERVICE_ACCOUNT_EMAIL, GOOGLE_SERVICE_ACCOUNT_PRIVATE_KEY,
//       GOOGLE_IMPERSONATE_SUBJECT (the user to act as)
//   Path C - Service account into a Shared Drive (SA is a Content Manager; no
//     impersonation). A bare SA has no personal Drive quota, so without a Shared
//     Drive or impersonation the create fails on quota:
//       GOOGLE_SERVICE_ACCOUNT_EMAIL, GOOGLE_SERVICE_ACCOUNT_PRIVATE_KEY
// Selection: if a refresh token is set, Path A is used; else the service account
// (Path B if GOOGLE_IMPERSONATE_SUBJECT is set, else Path C).
// GOVERNANCE: this writes client data (possibly PII) into Google Drive. Confirm
// folder location, sharing scope, and retention with the ISO 42001 owner before
// enabling in production.
//
// UNTESTED IN THIS ENVIRONMENT: written against the Drive v3 REST API but not
// executed here (no credentials / Google runtime). Smoke-test on a real deploy.

import crypto from "node:crypto";

const MAX_BODY_BYTES = 3_000_000; // base64 XLSX, generous
export const config = { path: "/.netlify/functions/sheet" };

const b64url = (buf) =>
  Buffer.from(buf).toString("base64").replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");

const TOKEN_URL = "https://oauth2.googleapis.com/token";
const DRIVE_SCOPE = "https://www.googleapis.com/auth/drive";

// True if any path is configured, so an unconfigured deploy degrades (501).
function sheetsConfigured() {
  return !!(process.env.APPS_SCRIPT_WEBAPP_URL ||
    process.env.GOOGLE_OAUTH_REFRESH_TOKEN ||
    (process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL && process.env.GOOGLE_SERVICE_ACCOUNT_PRIVATE_KEY));
}

async function tokenFrom(params) {
  const res = await fetch(TOKEN_URL, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams(params),
  });
  const data = await res.json().catch(() => ({}));
  if (!res.ok || !data.access_token) throw new Error("token_failed:" + (data.error || res.status));
  return data.access_token;
}

// Path A: OAuth refresh token for a real user.
function getOAuthToken(clientId, clientSecret, refreshToken) {
  return tokenFrom({ grant_type: "refresh_token", client_id: clientId, client_secret: clientSecret, refresh_token: refreshToken });
}

// Paths B/C: service account JWT. `subject` set = domain-wide delegation (act as
// that user); unset = act as the service account itself (needs a Shared Drive).
function getServiceAccountToken(email, privateKey, subject) {
  const now = Math.floor(Date.now() / 1000);
  const header = { alg: "RS256", typ: "JWT" };
  const claim = { iss: email, scope: DRIVE_SCOPE, aud: TOKEN_URL, iat: now, exp: now + 3600 };
  if (subject) claim.sub = subject;
  const unsigned = `${b64url(JSON.stringify(header))}.${b64url(JSON.stringify(claim))}`;
  const signature = crypto.createSign("RSA-SHA256").update(unsigned).sign(privateKey);
  const assertion = `${unsigned}.${b64url(signature)}`;
  return tokenFrom({ grant_type: "urn:ietf:params:oauth:grant-type:jwt-bearer", assertion });
}

// Resolve an access token from whichever path is configured (A > B/C).
function resolveAccessToken() {
  const refreshToken = process.env.GOOGLE_OAUTH_REFRESH_TOKEN;
  if (refreshToken) return getOAuthToken(process.env.GOOGLE_OAUTH_CLIENT_ID, process.env.GOOGLE_OAUTH_CLIENT_SECRET, refreshToken);
  const email = process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL;
  const privateKey = (process.env.GOOGLE_SERVICE_ACCOUNT_PRIVATE_KEY || "").replace(/\\n/g, "\n");
  return getServiceAccountToken(email, privateKey, process.env.GOOGLE_IMPERSONATE_SUBJECT || "");
}

export default async (req) => {
  if (req.method !== "POST") return json(405, { error: "method_not_allowed" });

  // Same-origin friction (chat page is same-origin).
  const origin = req.headers.get("origin");
  const siteURL = process.env.URL;
  if (origin && siteURL && new URL(origin).host !== new URL(siteURL).host) {
    return json(403, { error: "forbidden_origin" });
  }

  const folderId = process.env.GOOGLE_DRIVE_FOLDER_ID || "";
  if (!sheetsConfigured()) {
    // Not configured — let the client degrade gracefully (brief still sends).
    return json(501, { error: "sheets_not_configured" });
  }

  const rawBody = await req.text();
  if (rawBody.length > MAX_BODY_BYTES) return json(413, { error: "payload_too_large" });
  let body;
  try { body = JSON.parse(rawBody); }
  catch { return json(400, { error: "bad_json" }); }

  const { xlsxBase64, brief, filename, clientEmail, company, contactName, topicsCount, usersCount, sessionId } = body || {};
  const name = (typeof filename === "string" && filename) || `Lumen Setup Brief${company ? " - " + company : ""}`;

  // Path D (preferred when set): hand off to an Apps Script Web App that runs as a
  // real Google account. It COPIES the master requirements template and fills in
  // the structured brief, so the output matches the template exactly. URL + secret
  // are server-side env; the browser never sees them.
  const appsUrl = process.env.APPS_SCRIPT_WEBAPP_URL;
  if (appsUrl) {
    if (!brief || typeof brief !== "object") return json(400, { error: "missing_brief" });
    try {
      const r = await fetch(appsUrl, {
        method: "POST", headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ secret: process.env.APPS_SCRIPT_SECRET || "", brief, filename: name, clientEmail: clientEmail || "", company: company || "", contactName: contactName || "", topicsCount, usersCount, sessionId: sessionId || "" }),
      });
      const d = await r.json().catch(() => ({}));
      if (!r.ok || !d.url) {
        console.error("Apps Script sheet failed", r.status, d && d.error);
        return json(502, { error: "sheet_failed" });
      }
      return json(200, { url: d.url });
    } catch (err) {
      console.error("Apps Script sheet unreachable", err);
      return json(502, { error: "sheet_failed" });
    }
  }

  // Google-API fallback paths convert the client-built XLSX instead.
  if (!xlsxBase64 || typeof xlsxBase64 !== "string") return json(400, { error: "missing_xlsx" });
  try {
    const token = await resolveAccessToken();

    // Multipart upload of the XLSX with a Google-Sheets target mimeType, so Drive
    // converts it to a native Sheet on the way in.
    const boundary = "lumen" + crypto.randomUUID();
    const meta = { name: name.replace(/\.xlsx$/i, ""), mimeType: "application/vnd.google-apps.spreadsheet" };
    if (folderId) meta.parents = [folderId];
    const pre =
      `--${boundary}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n` +
      `${JSON.stringify(meta)}\r\n` +
      `--${boundary}\r\nContent-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet\r\n\r\n`;
    const post = `\r\n--${boundary}--`;
    const uploadBody = Buffer.concat([Buffer.from(pre, "utf8"), Buffer.from(xlsxBase64, "base64"), Buffer.from(post, "utf8")]);

    const up = await fetch(
      "https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&supportsAllDrives=true&fields=id,webViewLink",
      { method: "POST", headers: { Authorization: `Bearer ${token}`, "Content-Type": `multipart/related; boundary=${boundary}` }, body: uploadBody }
    );
    const file = await up.json().catch(() => ({}));
    if (!up.ok || !file.id) {
      console.error("Drive upload failed", up.status, JSON.stringify(file && file.error));
      return json(502, { error: "upload_failed" });
    }

    // Share with the client (as editor) if we have their email. sendNotificationEmail
    // makes Google email them the link — this is the "you'll get an email" path.
    if (clientEmail && /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(clientEmail)) {
      await fetch(
        `https://www.googleapis.com/drive/v3/files/${file.id}/permissions?sendNotificationEmail=true&supportsAllDrives=true`,
        { method: "POST", headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
          body: JSON.stringify({ role: "writer", type: "user", emailAddress: clientEmail }) }
      ).catch((e) => console.error("Share failed (non-fatal)", e));
    }

    const url = file.webViewLink || `https://docs.google.com/spreadsheets/d/${file.id}/edit`;
    return json(200, { url });
  } catch (err) {
    console.error("Sheet generation failed", err);
    return json(502, { error: "sheet_failed" });
  }
};

function json(status, obj) {
  return new Response(JSON.stringify(obj), { status, headers: { "Content-Type": "application/json" } });
}
