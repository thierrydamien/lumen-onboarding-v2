// parse-brief: server-side parser for the Lumen Onboarding Brief template (.xlsx).
//
// Sales fills a FIXED template (label in column A, value in column B) and uploads
// it on the sales page BEFORE generating a client link. We parse only the known
// labels and ignore everything else, so a rep adding stray rows or notes can't
// break extraction. A signature cell (A1) gates non-template uploads.
//
// Two OUTPUT channels with different downstream semantics:
//   - brief : client-appropriate facts (brands, markets, competitors, channels,
//             campaign, issues). The chat is told to SURFACE and CONFIRM these.
//   - notes : the single "internal notes" row -> the existing CONFIDENTIAL notes
//             channel the chat must never reveal.
// Keeping them separate avoids leaking internal framing to the client.

import * as XLSX from "xlsx";

const SIGNATURE = "Lumen Onboarding Brief";
const MAX_BYTES = 2 * 1024 * 1024; // reps upload a small filled template, never a data dump

// Canonical template spec. ONE source of truth: the generator script builds the
// downloadable .xlsx from this, and the parser matches uploads against it.
// target: where a filled value goes. "form:*" prefills a visible sales field;
// "brief:*" is a surfaceable fact; "competitor" collects into a list; "notes" is
// the confidential channel. label is matched case-insensitively, exact or prefix.
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

// Match a sheet's column-A label to a spec field: exact normalized match, else
// prefix (tolerates a rep appending "(optional)" etc. to a label cell).
function matchField(labelCell) {
  const n = norm(labelCell);
  if (!n) return null;
  for (const f of FIELDS) {
    const fl = norm(f.label);
    if (n === fl || n.startsWith(fl) || fl.startsWith(n)) return f;
  }
  return null;
}

export default async (req) => {
  if (req.method !== "POST") return json(405, { error: "method_not_allowed" });

  let body;
  try { body = await req.json(); } catch { return json(400, { error: "bad_request" }); }
  const b64 = typeof body.base64 === "string" ? body.base64 : "";
  if (!b64) return json(400, { error: "no_file" });

  let buf;
  try { buf = Buffer.from(b64, "base64"); } catch { return json(400, { error: "unreadable" }); }
  if (!buf.length) return json(400, { error: "empty" });
  if (buf.length > MAX_BYTES) return json(413, { error: "too_large", mb: (buf.length / 1048576).toFixed(1) });

  let rows;
  try {
    const wb = XLSX.read(buf, { type: "buffer" });
    const ws = wb.Sheets[wb.SheetNames[0]];
    if (!ws) return json(422, { error: "unreadable" });
    rows = XLSX.utils.sheet_to_json(ws, { header: 1, blankrows: false, defval: "" });
  } catch {
    return json(422, { error: "unreadable" });
  }

  // Signature gate: A1 (or anywhere in the first row) must carry the template title.
  const firstRow = (rows[0] || []).map(norm).join(" ");
  if (!firstRow.includes(norm(SIGNATURE))) {
    return json(422, { error: "not_template" });
  }

  // Build a label->value map from the sheet, keyed on column A.
  const found = {};
  for (const row of rows) {
    if (!Array.isArray(row) || row.length < 2) continue;
    const f = matchField(row[0]);
    if (!f) continue;
    const val = String(row[1] == null ? "" : row[1]).trim();
    if (val) found[f.key] = val;
  }

  // Assemble outputs.
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
  if (!filledCount) return json(422, { error: "template_empty" });

  const brief = briefLines.join("\n");
  return json(200, { ok: true, form, brief, notes, competitors, filledCount });
};
