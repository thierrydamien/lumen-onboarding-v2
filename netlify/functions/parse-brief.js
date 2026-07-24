// parse-brief: server-side parser for the Talkwalker/Lumen "Media Brief Form"
// template (the fixed "MAKE A COPY" workbook sales already use).
//
// Two inputs, one parser:
//   { base64 }   -> an uploaded .xlsx (a filled copy of the template)
//   { sheetUrl } -> a pasted Google Sheet link. We NEVER fetch the raw URL
//                   (SSRF): we validate host === docs.google.com, extract the id,
//                   and fetch the export endpoint WE build. Works only when the
//                   sheet is shared "anyone with the link can view"; a private
//                   sheet returns Google's HTML login page, which we detect (not a
//                   zip) and report as not_accessible so the rep uploads instead.
//
// HOW WE PARSE A MESSY, FIXED TEMPLATE: the template's label cells are identical
// in every copy (reps fill values, never edit labels), so we embed the exact label
// set below. In an upload, any cell whose text matches a known label IS a label;
// every other non-empty cell is a rep-entered VALUE. Each value is attached to its
// nearest label (same row to the left, else nearest label above), which matches the
// form's "label left/above, value right/below" layout. This tolerates the merged
// cells and irregular positions without hardcoding fragile addresses. The whole
// brief is client-facing (no confidential-notes field), so it all becomes the
// surfaceable brief the chat confirms; company and industry also pre-fill the form.

import * as XLSX from "xlsx";

const MAX_BYTES = 2 * 1024 * 1024;
const FETCH_MS = 6000;
const SIGNATURE_RE = /media brief form/;

const norm = s => String(s == null ? "" : s).replace(/\s+/g, " ").trim();
const low = s => norm(s).toLowerCase();

// Exact label strings from the template (all tabs). Used only to distinguish
// fixed labels from rep-entered values, so nothing here needs to be pretty.
const LABELS = new Set([
  "media brief form for a custom demo on talkwalker platform",
  "basic information (compulsory)",
  "company name/ holding company",
  "1-2 key brands/product names",
  "key geographical markets/ regions (area to focus for the demo)",
  'key languages (please add english keywords & translation in the "additional info" tab)',
  "industry",
  "1-2 business questions you hope to solve with social listening or objectives to be achieved from the tool",
  "for users of any listening tool: any painpoint/ missing feature/area of difficulty in your current tool?",
  "demo use case 1*",
  "key concerns for the use case / metrics you would like to see (sentiment etc)",
  "demo use case 2*",
  "social media channels active in the past 30 days (links only) facebook, instagram, twitter, youtube, app stores",
  "comments/remarks/ anything to highlight essential for you to see on the demo (if any):",
  "a &b)", "c)", "d)", "e)", "f)", "1)", "2)", "3)",
  "brand health/ pr measurement/ social measurement/ brand influencer discovery",
  "competitor intelligence (max 3 competitors for demo)",
  "key media sources/influencer list (if applicable)",
  "competitor name 1", "competitor name 2", "competitor name 3",
  "geographical markets", "additional remarks (if any)",
  "industry trends/ consumer insights (non-brand specific)",
  "objective or business purpose of understanding trends or consumer insights in the industry/ about the product",
  "specific product category/ industry",
  "1-3 industry areas/ consumer attributes of the product you want to discover (e.g aspects that you want to slice & dice the data further)",
  "issue tracking/ crisis management/ reputation management",
  "current known issues", "known issue keywords",
  "potential issues or areas of concern for your company/industry",
  "social campaign performance tracking (only feasible if you have a campaign hashtag that is active currently)",
  "campaign/ event name & #hashtags", "campaign/ event mechanics (eg retweet & win etc)",
  "period (has to be within the past 30 days)", "known influencers/ media",
  "web sources and urls",
]);

// Structural cells that are labels (so they aren't read as values) but must NOT be
// emitted as fields or own a value: the title, section banners, markers, the
// instruction note. A value near one of these is attributed to the next real label.
const SKIP = new Set([
  "media brief form for a custom demo on talkwalker platform",
  "basic information (compulsory)",
  "a &b)", "c)", "d)", "e)", "f)", "1)", "2)", "3)",
  "brand health/ pr measurement/ social measurement/ brand influencer discovery",
  "competitor intelligence (max 3 competitors for demo)",
  "industry trends/ consumer insights (non-brand specific)",
  "issue tracking/ crisis management/ reputation management",
  "social campaign performance tracking (only feasible if you have a campaign hashtag that is active currently)",
]);
function isSkip(l) { return SKIP.has(l) || l.startsWith("*please fill out"); }
function isLabel(l) { return LABELS.has(l) || l.startsWith("*please fill out"); }

// Prettier display for the verbose labels; others fall through trimmed.
const RELABEL = {
  "company name/ holding company": "Company",
  "1-2 key brands/product names": "Key brands / products",
  "key geographical markets/ regions (area to focus for the demo)": "Key markets / regions",
  'key languages (please add english keywords & translation in the "additional info" tab)': "Key languages",
  "1-2 business questions you hope to solve with social listening or objectives to be achieved from the tool": "Business questions / objectives",
  "for users of any listening tool: any painpoint/ missing feature/area of difficulty in your current tool?": "Current-tool pain points",
  "demo use case 1*": "Demo use case 1",
  "demo use case 2*": "Demo use case 2",
  "key concerns for the use case / metrics you would like to see (sentiment etc)": "Key metrics / concerns",
  "social media channels active in the past 30 days (links only) facebook, instagram, twitter, youtube, app stores": "Social channels",
  "comments/remarks/ anything to highlight essential for you to see on the demo (if any):": "Comments / remarks",
  "objective or business purpose of understanding trends or consumer insights in the industry/ about the product": "Trend/consumer objective",
  "specific product category/ industry": "Product category / industry",
  "1-3 industry areas/ consumer attributes of the product you want to discover (e.g aspects that you want to slice & dice the data further)": "Industry areas / attributes",
  "web sources and urls": "Web sources / URLs",
};
function display(text) {
  const l = low(text);
  if (RELABEL[l]) return RELABEL[l];
  const t = norm(text);
  return t.length > 60 ? t.slice(0, 57) + "…" : t;
}

function json(status, body) {
  return new Response(JSON.stringify(body), { status, headers: { "Content-Type": "application/json" } });
}

// Parse one sheet into ordered { labelCell, value } pairs plus any orphan values.
function pairsForSheet(rows) {
  const labelCells = []; // { r, c, text }
  const valueCells = []; // { r, c, text }
  for (let r = 0; r < rows.length; r++) {
    const row = rows[r] || [];
    for (let c = 0; c < row.length; c++) {
      const text = norm(row[c]);
      if (!text) continue;
      (isLabel(low(text)) ? labelCells : valueCells).push({ r, c, text });
    }
  }
  // Attach each value to its nearest owning (non-skip) label: same row to the left,
  // else nearest label above in the same column.
  const byOwner = new Map(); // key "r,c" -> { label, values: [] }
  const orphans = [];
  const ownable = labelCells.filter(l => !isSkip(low(l.text)));
  for (const v of valueCells) {
    let owner = null;
    for (const l of ownable) if (l.r === v.r && l.c < v.c && (!owner || l.c > owner.c)) owner = l;
    if (!owner) for (const l of ownable) if (l.c === v.c && l.r < v.r && (!owner || l.r > owner.r)) owner = l;
    if (!owner) { orphans.push(v.text); continue; }
    const key = owner.r + "," + owner.c;
    if (!byOwner.has(key)) byOwner.set(key, { label: owner, values: [] });
    byOwner.get(key).values.push(v.text);
  }
  // Emit in reading order (top-to-bottom, left-to-right by owner position).
  const ordered = [...byOwner.values()].sort((a, b) => a.label.r - b.label.r || a.label.c - b.label.c);
  return { pairs: ordered.map(o => ({ label: o.label.text, value: o.values.join("; ") })), orphans };
}

function parseMediaBrief(wb) {
  const dataSheets = wb.SheetNames.filter(n => low(n) !== "dropdown list");
  let signed = false;
  for (const n of wb.SheetNames) {
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, blankrows: false, defval: "" });
    if ((rows[0] || []).some(c => SIGNATURE_RE.test(low(c)))) { signed = true; break; }
  }
  if (!signed) return { error: "not_template" };

  const form = {};
  const sections = [];
  let filledCount = 0;
  for (const name of dataSheets) {
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, blankrows: false, defval: "" });
    const { pairs, orphans } = pairsForSheet(rows);
    const lines = [];
    for (const p of pairs) {
      filledCount++;
      lines.push(display(p.label) + ": " + p.value);
      // Company / industry pre-fill the form (Basic Info tab only, first hit wins).
      if (low(name).includes("basic")) {
        const l = low(p.label);
        if (!form.company && l.includes("company name")) form.company = p.value;
        if (!form.industry && l === "industry") form.industry = p.value;
      }
    }
    for (const o of orphans) { filledCount++; lines.push(o); }
    if (lines.length) sections.push("[" + name + "]\n" + lines.join("\n"));
  }
  if (!filledCount) return { error: "template_empty" };
  return { ok: true, form, brief: sections.join("\n\n").slice(0, 7000), filledCount };
}

// SSRF guard: only a docs.google.com spreadsheet URL; return just the id so the
// caller builds the export URL itself. Ids are long, so require >=20 chars.
export function sheetIdFrom(u) {
  if (typeof u !== "string" || !u) return null;
  let host;
  try { host = new URL(u).host.toLowerCase(); } catch { return null; }
  if (host !== "docs.google.com") return null;
  const m = u.match(/\/spreadsheets\/d\/([A-Za-z0-9_-]{20,})/);
  return m ? m[1] : null;
}

async function fetchSheetXlsx(id) {
  const url = "https://docs.google.com/spreadsheets/d/" + id + "/export?format=xlsx";
  const ctrl = new AbortController();
  const timer = setTimeout(() => ctrl.abort(), FETCH_MS);
  try {
    const resp = await fetch(url, { redirect: "follow", signal: ctrl.signal });
    if (!resp.ok) return { error: "not_accessible" };
    const buf = Buffer.from(await resp.arrayBuffer());
    if (buf.length > MAX_BYTES) return { error: "too_large", mb: (buf.length / 1048576).toFixed(1) };
    // .xlsx is a zip: starts with "PK". A private sheet returns an HTML login page.
    if (buf.length < 4 || buf[0] !== 0x50 || buf[1] !== 0x4b) return { error: "not_accessible" };
    return { buf };
  } catch (e) {
    return { error: e && e.name === "AbortError" ? "fetch_timeout" : "fetch_failed" };
  } finally {
    clearTimeout(timer);
  }
}

function readWorkbook(buf) {
  try { return { wb: XLSX.read(buf, { type: "buffer" }) }; }
  catch { return { error: "unreadable" }; }
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

  const { wb, error } = readWorkbook(buf);
  if (error) return json(422, { error });
  const out = parseMediaBrief(wb);
  if (out.error) return json(422, out);
  return json(200, out);
};
