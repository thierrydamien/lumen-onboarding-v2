// Session store for completed onboarding briefs (Netlify Functions v2 + Blobs).
//   POST /.netlify/functions/session   { session }        -> { id }
//   GET  /.netlify/functions/session                      -> { sessions: [...] }
//   GET  /.netlify/functions/session?id=<id>              -> { session }
//
// In v2, getStore() picks up the site context automatically, so Blobs works
// with no siteID/token wiring. Only dependency: @netlify/blobs.

import { getStore } from "@netlify/blobs";

const STORE = "lumen-sessions";
export const config = { path: "/.netlify/functions/session" };

export default async (req) => {
  let store;
  try { store = getStore(STORE); }
  catch (err) { console.error("Blobs store unavailable", err); return json(500, { error: "store_unavailable" }); }

  const url = new URL(req.url);

  if (req.method === "POST") {
    let body;
    try { body = await req.json(); }
    catch { return json(400, { error: "bad_json" }); }

    const session = body && body.session;
    if (!session || typeof session !== "object") return json(400, { error: "missing_session" });

    const id = session.id || genId();
    const record = { ...session, id, savedAt: new Date().toISOString() };
    try { await store.setJSON(id, record); }
    catch (err) { console.error("Failed to save session", err); return json(502, { error: "save_failed" }); }
    return json(200, { id });
  }

  if (req.method === "GET") {
    const id = url.searchParams.get("id");
    if (id) {
      try {
        const rec = await store.get(id, { type: "json" });
        if (!rec) return json(404, { error: "not_found" });
        return json(200, { session: rec });
      } catch (err) { console.error("Failed to read session", err); return json(502, { error: "read_failed" }); }
    }
    try {
      const { blobs } = await store.list();
      const records = await Promise.all(
        blobs.map((b) => store.get(b.key, { type: "json" }).catch(() => null))
      );
      const sessions = records
        .filter(Boolean)
        .map(summarize)
        .sort((a, b) => (b.savedAt || "").localeCompare(a.savedAt || ""));
      return json(200, { sessions });
    } catch (err) { console.error("Failed to list sessions", err); return json(502, { error: "list_failed" }); }
  }

  return json(405, { error: "method_not_allowed" });
};

function summarize(r) {
  const company = (r.merged && r.merged.company) || {};
  return {
    id: r.id,
    company: company.name || "(unnamed)",
    contact: company.contact || "",
    email: company.email || "",
    objectives: company.objectives || "",
    topicCount: Array.isArray(r.merged && r.merged.topics) ? r.merged.topics.length : 0,
    channelCount: Array.isArray(r.merged && r.merged.channels) ? r.merged.channels.length : 0,
    userCount: Array.isArray(r.users) ? r.users.length : 0,
    status: r.status || "completed",
    durationMs: r.durationMs || null,
    apiCalls: r.apiCalls || null,
    sentAt: r.sentAt || null,
    savedAt: r.savedAt || null,
  };
}

function genId() {
  return "s_" + Date.now().toString(36) + "_" + Math.random().toString(36).slice(2, 8);
}

function json(status, obj) {
  return new Response(JSON.stringify(obj), { status, headers: { "Content-Type": "application/json" } });
}
