// preview-brief: shows a sales rep how the onboarding assistant will read a
// free-form brief BEFORE the link is generated, so they can catch bad formatting.
// Runs the brief past the same model with a focused "read it back" instruction and
// returns a plain-text summary of the facts it extracted plus anything unclear.
//
// Internal tool + a paid model call, so it is guarded like seed.js writes: a
// same-origin Origin header, and (when SEED_WRITE_TOKEN is set) the same
// x-app-write-token the Sales page already caches for generating links.

const MODEL = "claude-sonnet-4-6";
const ANTHROPIC_URL = "https://api.anthropic.com/v1/messages";
const ANTHROPIC_VERSION = "2023-06-01";
const MAX_BRIEF = 8000;
const MAX_TOKENS = 450;
const TIMEOUT_MS = 9000; // stay under the function wall-clock

const SYSTEM = "You are helping a sales rep sanity-check a client brief before it is handed to an onboarding assistant. The rep pasted the brief in the next message. Reply with a short, plain read-back of ONLY the facts the assistant would treat as known about this client, so the rep can confirm it was understood. Use a compact bullet list (lines starting with '- ') grouped simply where it applies: Company, Brands/products, Markets, Languages, Industry, Competitors, Channels, Campaign, Use case, Other. Include only what is actually present; never invent or infer beyond what is written. After the bullets, add one final line starting 'Unclear: ' — list anything ambiguous, malformed, or that you could not interpret; if it all reads cleanly, write 'Unclear: nothing, this reads cleanly.' Keep the whole reply under 150 words. Plain text only, no markdown headings or bold.";

function json(status, body) {
  return new Response(JSON.stringify(body), { status, headers: { "Content-Type": "application/json" } });
}

export default async (req) => {
  if (req.method !== "POST") return json(405, { error: "method_not_allowed" });

  // Origin check (layered, not a strong control on its own).
  const origin = req.headers.get("origin");
  const siteURL = process.env.URL;
  if (siteURL) {
    let ok = false;
    try { ok = !!origin && new URL(origin).host === new URL(siteURL).host; } catch { ok = false; }
    if (!ok) return json(403, { error: "forbidden_origin" });
  }
  // Same write-token posture as seed.js: enforced only when configured.
  const writeToken = process.env.SEED_WRITE_TOKEN;
  if (writeToken && req.headers.get("x-app-write-token") !== writeToken) {
    return json(401, { error: "unauthorized" });
  }

  let body;
  try { body = await req.json(); } catch { return json(400, { error: "bad_request" }); }
  const brief = typeof body.brief === "string" ? body.brief.trim() : "";
  if (!brief) return json(400, { error: "empty" });
  if (brief.length > MAX_BRIEF) return json(413, { error: "too_large" });

  const key = process.env.ANTHROPIC_API_KEY;
  if (!key) { console.error("ANTHROPIC_API_KEY is not set"); return json(500, { error: "not_configured" }); }

  const ac = new AbortController();
  const timer = setTimeout(() => ac.abort(), TIMEOUT_MS);
  try {
    const res = await fetch(ANTHROPIC_URL, {
      method: "POST",
      headers: { "Content-Type": "application/json", "x-api-key": key, "anthropic-version": ANTHROPIC_VERSION },
      body: JSON.stringify({ model: MODEL, max_tokens: MAX_TOKENS, system: SYSTEM, messages: [{ role: "user", content: brief }] }),
      signal: ac.signal,
    });
    if (!res.ok) { console.error("preview upstream error", res.status); return json(502, { error: "upstream_error" }); }
    const data = await res.json();
    const interpretation = Array.isArray(data.content)
      ? data.content.filter(b => b && b.type === "text").map(b => b.text).join("").trim()
      : "";
    if (!interpretation) return json(502, { error: "empty_reply" });
    return json(200, { ok: true, interpretation });
  } catch (e) {
    if (e && e.name === "AbortError") return json(504, { error: "timeout" });
    console.error("preview-brief failed", e && e.message);
    return json(502, { error: "request_failed" });
  } finally {
    clearTimeout(timer);
  }
};
