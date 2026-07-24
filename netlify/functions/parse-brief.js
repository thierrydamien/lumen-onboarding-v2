// parse-brief: server-side parser for the Lumen Onboarding Brief template.
//
// Two inputs, one parser:
//   { base64 }   -> an uploaded .xlsx the rep filled from our template
//   { sheetUrl } -> a pasted Google Sheet link. We NEVER fetch the raw URL
//                   (SSRF): we validate host === docs.google.com, extract the
//                   sheet id, and fetch the export endpoint WE build. This only
//                   works when the sheet is shared "anyone with the link can
//                   view"; a private sheet returns Google's HTML login page, which
//                   we detect (not a zip) and report as not_accessible so the rep
//                   downloads and uploads instead.
//
// Sales fills a FIXED template (label in column A, value in column B). We match
// only known labels and ignore everything else, so stray rows can't break it. A
// signature cell gates non-template inputs. We scan ALL tabs for the signature so
// the template still parses when it is one tab in a larger workbook.
//
// Output splits into two channels with different downstream semantics:
//   - brief : client-appropriate facts the chat SURFACES and confirms.
//   - notes : the single "internal notes" row -> the CONFIDENTIAL notes channel
//             the chat must never reveal.

import * as XLSX from "xlsx";

const SIGNATURE = "Lumen Onboarding Brief";
const MAX_BYTES = 2 * 1024 * 1024; // a filled template is tiny; never a data dump
const FETCH_MS = 6000;             // bound the Google fetch inside the function wall-clock

export const FIELD_SPEC = [
  { section: "BASIC INFORMATION" },
  { key: "company",     label: "Company name",                     target: "form:company" },
  { key: "brands",      label: "Key brands or products",           target: "brief:Key brands / products" },
  { key: "industry",    label: "Industry",                         target: "form:industry" },
  { key: "markets",     label: "Key markets or regions",           target: "brief:Key markets / regions" },
  { key: "languages",   label: "Monitoring languages",             target: "brief:Monitoring languages", hint: "languages to monitor content in (not the onboarding UI language)" },
  { key: "objectives",  label: "Business objectives",              target: "brief:Business objectives" },
  { key: "painpoints",  label: "Current tool or pain points",      target: "brief:Current tool / pain points" },

  { section: "COMPETITORS (up to 3)" },
  { key: "competitor1", label: "Competitor 1",                     target: "competitor" },
  { key: "competitor2", label: "Competitor 2",                     target: "competitor" },
  { key: "competitor3", label: "Competitor 3",                     target: "competitor" },

  { section: "TOPICS AND USE CASES" },
  { key: "usecase1",    label: "Primary use case",                 target: "brief:Primary use case", hint: "e.g. Brand Health, Competitive Intelligence, Issue Tracking, Campaign Tracking, Trend Research" },
  { key: "usecase2",    label: "Secondary use case",               target: "brief:Secondary use case" },
  { key: "issues",      label: "Known issues or crisis keywords",  target: "brief:Known issues / crisis keywords" },
  { key: "campaign",    label: "Campaign name and hashtags",       target: "brief:Campaign name / hashtags" },

  { section: "CHANNELS" },
  { key: "channels",    label: "Owned social channels (URLs)",     target: "brief:Owned social channels (URLs)", hint: "full URLs, comma-separated or one per line" },

  { section: "CLIENT CONTACT (optional)" },
  { key: "contactName", label: "Client contact name",              target: "form:contactName" },
  { key: "contactEmail",label: "Client contact email",             target: "form:email" },

  { section: "INTERNAL (not shown to the client)" },
  { key: "notes",       label: "Internal notes for the onboarding team", target: "notes" },
];

const FIELDS = FIELD_SPEC.filter(f => f.key);
const norm = s => String(s == null ? "" : s).replace(/\s+/g, " ").trim().toLowerCase();

function json(status, body) {
  return new Response(JSON.stringify(body), { status, headers: { "Content-Type": "application/json" } });
}

// Match a column-A label to a spec field: exact normalized match, else prefix
// (tolerates a rep appending "(optional)" etc.).
function matchField(labelCell) {
  const n = norm(labelCell);
  if (!n) return null;
  for (const f of FIELDS) {
    const fl = norm(f.label);
    if (n === fl || n.startsWith(fl) || fl.startsWith(n)) return f;
  }
  return null;
}

// Parse a workbook buffer into the two channels. Scans every sheet for the
// signature row so the template parses even as one tab among many.
function extractFromWorkbook(buf) {
  let wb;
  try { wb = XLSX.read(buf, { type: "buffer" }); }
  catch { return { error: "unreadable" }; }

  let rows = null;
  for (const name of wb.SheetNames) {
    const ws = wb.Sheets[name];
    if (!ws) continue;
    const r = XLSX.utils.sheet_to_json(ws, { header: 1, blankrows: false, defval: "" });
    const firstRow = (r[0] || []).map(norm).join(" ");
    if (firstRow.includes(norm(SIGNATURE))) { rows = r; break; }
  }
  if (!rows) return { error: "not_template" };

  const found = {};
  for (const row of rows) {
    if (!Array.isArray(row) || row.length < 2) continue;
    const f = matchField(row[0]);
    if (!f) continue;
    const val = String(row[1] == null ? "" : row[1]).trim();
    if (val) found[f.key] = val;
  }

  const form = {};
  const briefLines = [];
  const competitors = [];
  let notes = "";
  for (const f of FIELDS) {
    const val = found[f.key];
    if (!val) continue;
    if (f.target === "competitor") { competitors.push(val); continue; }
    if (f.target === "notes") { notes = val; continue; }
    if (f.target.startsWith("form:")) { form[f.target.slice(5)] = val; continue; }
    if (f.target.startsWith("brief:")) { briefLines.push(f.target.slice(6) + ": " + val); }
  }
  if (competitors.length) briefLines.push("Competitors: " + competitors.join(", "));

  const filledCount = Object.keys(found).length;
  if (!filledCount) return { error: "template_empty" };
  return { ok: true, form, brief: briefLines.join("\n"), notes, competitors, filledCount };
}

// SSRF guard: accept ONLY a docs.google.com spreadsheet URL, and return just the
// id so the caller builds the export URL itself. Sheet ids are long; require >=20
// chars so we don't mis-match a "/d/e/" published-link segment.
export function sheetIdFrom(u) {
  if (typeof u !== "string" || !u) return null;
  let host;
  try { host = new URL(u).host.toLowerCase(); } catch { return null; }
  if (host !== "docs.google.com") return null;
  const m = u.match(/\/spreadsheets\/d\/([A-Za-z0-9_-]{20,})/);
  return m ? m[1] : null;
}

// Fetch the sheet's xlsx export. Only reachable for link-viewable sheets; a
// private sheet redirects to a login page (HTML, not a zip) which we detect.
async function fetchSheetXlsx(id) {
  const url = "https://docs.google.com/spreadsheets/d/" + id + "/export?format=xlsx";
  const ctrl = new AbortController();
  const timer = setTimeout(() => ctrl.abort(), FETCH_MS);
  try {
    const resp = await fetch(url, { redirect: "follow", signal: ctrl.signal });
    if (!resp.ok) return { error: "not_accessible" };
    const buf = Buffer.from(await resp.arrayBuffer());
    if (buf.length > MAX_BYTES) return { error: "too_large", mb: (buf.length / 1048576).toFixed(1) };
    // .xlsx is a zip: starts with "PK". Anything else (e.g. an HTML login page
    // for a private sheet) means we could not actually read the sheet.
    if (buf.length < 4 || buf[0] !== 0x50 || buf[1] !== 0x4b) return { error: "not_accessible" };
    return { buf };
  } catch (e) {
    return { error: e && e.name === "AbortError" ? "fetch_timeout" : "fetch_failed" };
  } finally {
    clearTimeout(timer);
  }
}

export default async (req) => {
  if (req.method !== "POST") return json(405, { error: "method_not_allowed" });

  let body;
  try { body = await req.json(); } catch { return json(400, { error: "bad_request" }); }

  let buf;
  if (typeof body.sheetUrl === "string" && body.sheetUrl.trim()) {
    const id = sheetIdFrom(body.sheetUrl.trim());
    if (!id) return json(422, { error: "bad_url" });
    const r = await fetchSheetXlsx(id);
    if (r.error) return json(r.error === "too_large" ? 413 : 422, r);
    buf = r.buf;
  } else if (typeof body.base64 === "string" && body.base64) {
    try { buf = Buffer.from(body.base64, "base64"); } catch { return json(400, { error: "unreadable" }); }
    if (!buf.length) return json(400, { error: "empty" });
    if (buf.length > MAX_BYTES) return json(413, { error: "too_large", mb: (buf.length / 1048576).toFixed(1) });
  } else {
    return json(400, { error: "no_file" });
  }

  const out = extractFromWorkbook(buf);
  if (out.error) return json(422, out);
  return json(200, out);
};
