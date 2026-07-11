// Google Sheet generation (Netlify Functions v2).
//   POST /.netlify/functions/sheet  { xlsxBase64, filename, clientEmail?, company? }
//     -> { url }   (a Google Sheet converted from the brief's XLSX)
//
// SETUP REQUIRED (this function is inert until configured):
//   - GOOGLE_SERVICE_ACCOUNT_EMAIL      service account address
//   - GOOGLE_SERVICE_ACCOUNT_PRIVATE_KEY  its private key (PEM; \n may be escaped)
//   - GOOGLE_DRIVE_FOLDER_ID (recommended) target folder, ideally a Shared Drive
//     the service account is a member of (a bare service account has no personal
//     Drive quota, so uploads generally MUST land in a Shared Drive).
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

async function getAccessToken(email, privateKey) {
  const now = Math.floor(Date.now() / 1000);
  const header = { alg: "RS256", typ: "JWT" };
  const claim = {
    iss: email,
    scope: "https://www.googleapis.com/auth/drive",
    aud: "https://oauth2.googleapis.com/token",
    iat: now,
    exp: now + 3600,
  };
  const unsigned = `${b64url(JSON.stringify(header))}.${b64url(JSON.stringify(claim))}`;
  const signature = crypto.createSign("RSA-SHA256").update(unsigned).sign(privateKey);
  const assertion = `${unsigned}.${b64url(signature)}`;
  const res = await fetch("https://oauth2.googleapis.com/token", {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({ grant_type: "urn:ietf:params:oauth:grant-type:jwt-bearer", assertion }),
  });
  const data = await res.json().catch(() => ({}));
  if (!res.ok || !data.access_token) throw new Error("token_failed:" + (data.error || res.status));
  return data.access_token;
}

export default async (req) => {
  if (req.method !== "POST") return json(405, { error: "method_not_allowed" });

  // Same-origin friction (chat page is same-origin).
  const origin = req.headers.get("origin");
  const siteURL = process.env.URL;
  if (origin && siteURL && new URL(origin).host !== new URL(siteURL).host) {
    return json(403, { error: "forbidden_origin" });
  }

  const email = process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL;
  const privateKey = (process.env.GOOGLE_SERVICE_ACCOUNT_PRIVATE_KEY || "").replace(/\\n/g, "\n");
  const folderId = process.env.GOOGLE_DRIVE_FOLDER_ID || "";
  if (!email || !privateKey) {
    // Not configured — let the client degrade gracefully (brief still sends).
    return json(501, { error: "sheets_not_configured" });
  }

  const rawBody = await req.text();
  if (rawBody.length > MAX_BODY_BYTES) return json(413, { error: "payload_too_large" });
  let body;
  try { body = JSON.parse(rawBody); }
  catch { return json(400, { error: "bad_json" }); }

  const { xlsxBase64, filename, clientEmail, company } = body || {};
  if (!xlsxBase64 || typeof xlsxBase64 !== "string") return json(400, { error: "missing_xlsx" });
  const name = (typeof filename === "string" && filename) || `Lumen Setup Brief${company ? " - " + company : ""}`;

  try {
    const token = await getAccessToken(email, privateKey);

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
