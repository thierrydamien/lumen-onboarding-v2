// Seed store for pre-filled client profiles (Netlify Functions v2 + Blobs).
//   POST /.netlify/functions/seed   { seed }            -> { id }
//   GET  /.netlify/functions/seed?id=<id>               -> { seed } CLIENT-SAFE (no notes)
//   GET  /.netlify/functions/seed?id=<id> + x-dashboard-token -> { seed } FULL (incl notes)
//
// WHY THIS EXISTS: the Sales page used to base64-encode the entire seed —
// including confidential consultant notes — into the client's ?c= link. Base64
// is trivially reversible, so any client could read the notes. Now Sales stores
// the seed here and the link carries only an opaque id. The chat page fetches
// only the client-safe fields; the notes are returned solely to the dashboard,
// which authenticates with DASHBOARD_TOKEN.

import { getStore } from "@netlify/blobs";
import crypto from "node:crypto"; // explicit import: globalThis.crypto only exists on Node >= 19; don't rely on it

const STORE = "lumen-seeds";
const MAX_BODY_BYTES = 40_000;
// Retention: stored seeds carry client data (incl. consultant notes), so they
// expire rather than living forever. Lazily enforced on read/list (no cron):
// an expired seed is treated as not-found and deleted when next touched.
// Override the 90-day default with SEED_TTL_DAYS (0 disables expiry).
// Guard a non-numeric SEED_TTL_DAYS: without this, a typo'd value -> NaN -> expiry
// silently disabled (notes-bearing seeds would live forever). Bad value -> 90-day default.
const _ttlDays = process.env.SEED_TTL_DAYS != null ? Number(process.env.SEED_TTL_DAYS) : 90;
// Guard NaN (typo -> 90 default) AND negative (a sign-typo like -1 would make
// SEED_TTL_MS negative and truthy, so isExpired's `age > SEED_TTL_MS` is true for
// EVERY record -> the whole seed store, notes and all, silently deleted on next
// read/list). Clamp to >= 0; 0 disables expiry, which is the intended "keep forever".
const SEED_TTL_MS = Math.max(0, Number.isFinite(_ttlDays) ? _ttlDays : 90) * 86400000;
export const config = { path: "/.netlify/functions/seed" };

function isExpired(rec) {
  if (!SEED_TTL_MS || !rec || !rec.savedAt) return false;
  const t = Date.parse(rec.savedAt);
  return Number.isFinite(t) && (Date.now() - t) > SEED_TTL_MS;
}

// Fields safe to hand back to the (unauthenticated) client chat page. Notes are
// deliberately excluded and only ever returned to an authenticated dashboard.
const CLIENT_SAFE = ["company", "contactName", "email", "industry", "language"];

export default async (req) => {
  let store;
  try { store = getStore(STORE); }
  catch (err) { console.error("Blobs store unavailable", err); return json(500, { error: "store_unavailable" }); }

  const url = new URL(req.url);

  if (req.method === "POST") {
    // Origin check: require a PRESENT, same-origin Origin header (browsers send it
    // on POST). Rejects scripted writes that omit or mismatch it. Layered, not a
    // strong control on its own — Origin is spoofable outside a browser.
    const origin = req.headers.get("origin");
    const siteURL = process.env.URL;
    if (siteURL) {
      let ok = false;
      try { ok = !!origin && new URL(origin).host === new URL(siteURL).host; } catch { ok = false; }
      if (!ok) return json(403, { error: "forbidden_origin" });
    } else {
      console.warn("URL env not set — cannot validate Origin on seed write");
    }
    // Optional write token. Unlike the public chat, the Sales page is an INTERNAL
    // consultant tool, so it can carry a secret the way the dashboard does: when
    // SEED_WRITE_TOKEN is set, Sales sends it as x-app-write-token (prompting the
    // consultant and caching it for the tab, re-prompting on 401). Unset =>
    // origin-only, unchanged, so nothing breaks until you opt in. Recommended for
    // production so the seed store (which holds consultant notes) can't be written
    // by anyone who merely spoofs the Origin header.
    const writeToken = process.env.SEED_WRITE_TOKEN;
    if (writeToken && req.headers.get("x-app-write-token") !== writeToken) {
      return json(401, { error: "unauthorized_write" });
    }

    const rawBody = await req.text();
    if (rawBody.length > MAX_BODY_BYTES) return json(413, { error: "payload_too_large" });

    let body;
    try { body = JSON.parse(rawBody); }
    catch { return json(400, { error: "bad_json" }); }

    const seed = body && body.seed;
    if (!seed || typeof seed !== "object" || Array.isArray(seed)) return json(400, { error: "missing_seed" });
    if (!seed.company) return json(400, { error: "company_required" });

    // Store only known string fields, each length-capped, so a tampered client
    // cannot stuff arbitrary or oversized data into the store.
    // preparedBy (the consultant's name) is stored but is NOT client-safe: it is
    // returned only to the token-gated dashboard, never to the client chat page.
    const clean = {};
    // package (e.g. "core-advanced") scopes how much the assistant gathers. Like
    // notes/preparedBy it is NOT client-safe: chat.js reads it server-side to inject
    // the setup limits, and the dashboard sees it, but the client fetch never returns it.
    for (const k of [...CLIENT_SAFE, "notes", "preparedBy", "package"]) {
      if (seed[k] != null) clean[k] = String(seed[k]).slice(0, 4000);
    }
    const id = "sd_" + crypto.randomUUID();
    const record = { ...clean, id, savedAt: new Date().toISOString() };
    try { await store.setJSON(id, record); }
    catch (err) { console.error("Failed to save seed", err); return json(502, { error: "save_failed" }); }
    // Surface the link's expiry to the Sales page so it can tell the rep how long
    // the link is good for. null when expiry is disabled (SEED_TTL_DAYS=0).
    const expiresAt = SEED_TTL_MS ? new Date(Date.parse(record.savedAt) + SEED_TTL_MS).toISOString() : null;

    // Duplicate-client heads-up: after the write succeeds, look for OTHER
    // non-expired seeds for the same company (case-insensitive, trimmed, exact)
    // so the Sales page can warn the rep a link for this client already exists.
    // Best-effort and non-fatal: the write is already done, so any failure here
    // just yields an empty list rather than failing the request. Reuses the same
    // list()/get pattern as the GET list branch. Bounded to 5 (most recent first).
    let duplicates = [];
    try {
      const target = String(record.company || "").trim().toLowerCase();
      if (target) {
        const { blobs } = await store.list();
        const recs = await Promise.all(blobs.map((b) => store.get(b.key, { type: "json" }).catch(() => null)));
        const matches = [];
        for (const r of recs) {
          if (!r || r.id === id) continue;           // skip the record we just wrote
          if (isExpired(r)) continue;                 // only non-expired seeds count
          if (String(r.company || "").trim().toLowerCase() === target) {
            // preparedBy names an internal consultant. Only include it when the POST
            // was actually authenticated (SEED_WRITE_TOKEN set + matched above); on the
            // Origin-only posture, an attacker who forges Origin could otherwise probe
            // "who made a link for <company>", so return just the timestamp then.
            matches.push(writeToken
              ? { savedAt: r.savedAt || null, preparedBy: r.preparedBy || "" }
              : { savedAt: r.savedAt || null });
          }
        }
        matches.sort((a, b) => (b.savedAt || "").localeCompare(a.savedAt || "")); // most recent first
        duplicates = matches.slice(0, 5);
      }
    } catch (err) { console.error("Duplicate scan failed (non-fatal)", err); }

    return json(200, { id, expiresAt, duplicates });
  }

  if (req.method === "GET") {
    // Notes / the full list are returned only to a caller holding the dashboard
    // token. The client chat page (no token) can only fetch one seed by id and
    // only the client-safe subset.
    const expected = process.env.DASHBOARD_TOKEN;
    const provided = req.headers.get("x-dashboard-token");
    const authed = !!expected && provided === expected;
    const id = url.searchParams.get("id");

    if (id) {
      let rec;
      try { rec = await store.get(id, { type: "json" }); }
      catch (err) { console.error("Failed to read seed", err); return json(502, { error: "read_failed" }); }
      if (!rec) return json(404, { error: "not_found" });
      if (isExpired(rec)) { store.delete(id).catch(() => {}); return json(404, { error: "expired" }); }
      const safe = {};
      for (const k of CLIENT_SAFE) if (rec[k] != null) safe[k] = rec[k];
      const out = authed ? { ...safe, notes: rec.notes || "", preparedBy: rec.preparedBy || "", package: rec.package || "" } : safe;
      return json(200, { seed: out });
    }

    // List: the "link sent" top-of-funnel view for the dashboard. Token-gated;
    // notes are NOT included (they're fetched per-id only). Sweeps expired seeds.
    if (!authed) return json(401, { error: "unauthorized" });
    try {
      const { blobs } = await store.list();
      const recs = await Promise.all(blobs.map((b) => store.get(b.key, { type: "json" }).catch(() => null)));
      const seeds = [];
      for (const r of recs) {
        if (!r) continue;
        if (isExpired(r)) { store.delete(r.id).catch(() => {}); continue; }
        const safe = { id: r.id, savedAt: r.savedAt || null };
        for (const k of CLIENT_SAFE) if (r[k] != null) safe[k] = r[k];
        if (r.preparedBy != null) safe.preparedBy = r.preparedBy; // owner (dashboard only)
        if (r.package != null) safe.package = r.package;           // package code (dashboard only)
        seeds.push(safe);
      }
      seeds.sort((a, b) => (b.savedAt || "").localeCompare(a.savedAt || ""));
      return json(200, { seeds });
    } catch (err) { console.error("Failed to list seeds", err); return json(502, { error: "list_failed" }); }
  }

  return json(405, { error: "method_not_allowed" });
};

function json(status, obj) {
  return new Response(JSON.stringify(obj), { status, headers: { "Content-Type": "application/json" } });
}
