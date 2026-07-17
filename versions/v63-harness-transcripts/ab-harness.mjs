#!/usr/bin/env node
/**
 * Lumen cost-lever A/B harness.
 *
 * Answers "does this cost lever reduce spend WITHOUT reducing quality?" with
 * numbers instead of a guess. It drives a full simulated onboarding against the
 * real Anthropic API for the baseline prompt and for each lever variant, then
 * scores each two ways: a judge model rates overall quality, and a fact-retention
 * probe measures how many specific facts the client stated early the model can
 * still read back at the end. That probe is the sharp signal for finding 21 (the
 * sliding-history lever): a coarse quality score is too noisy to catch "dropped
 * one competitor", but recall catches it directly. Your live prompt is never
 * modified — each lever is just an extra instruction appended in memory here.
 *
 * RUN:  ANTHROPIC_API_KEY=sk-... node tools/ab-harness.mjs
 * ENV (optional): AB_RUNS (default 3), AB_TURNS (16), AB_MAX_HIST (20), AB_MODEL,
 *   AB_JUDGE_MODEL, AB_ONLY (comma-separated config keys to run, e.g.
 *   AB_ONLY=baseline,full-history). Every run also writes ab-transcripts.txt with
 *   the full conversations, so quality can be judged by reading, not the auto-judge.
 * COST: makes real API calls — a few dollars per run. That's the point: spend a
 *       little to know before changing the live prompt. Nothing ships from here.
 *
 * WHAT'S MEASURED: only the ASSISTANT's tokens count toward cost (that's what
 * runs in prod). The simulated client + judge calls are harness overhead. Both
 * prod caches are replicated: the system prompt (v27) and the last-message
 * conversation-history breakpoint (v35). Absolute $/convo still runs a little high
 * (one client per run, no warm cross-session cache), so read the RELATIVE delta
 * between configs — that's what the decision rests on.
 */
import { readFileSync, writeFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const KEY = process.env.ANTHROPIC_API_KEY;
if (!KEY) { console.error("Set ANTHROPIC_API_KEY (e.g. ANTHROPIC_API_KEY=sk-... node tools/ab-harness.mjs)"); process.exit(1); }

const ANTHROPIC_URL = "https://api.anthropic.com/v1/messages";
const VERSION = "2023-06-01";
const ASSISTANT_MODEL = process.env.AB_MODEL || "claude-sonnet-4-6";
const JUDGE_MODEL = process.env.AB_JUDGE_MODEL || "claude-sonnet-4-6";
const RUNS = Number(process.env.AB_RUNS || 3);
const MAX_TURNS = Number(process.env.AB_TURNS || 16);
const RATES = { input: 3, output: 15, cacheWrite: 3.75, cacheRead: 0.30 }; // $/MTok
const MAX_HIST = Number(process.env.AB_MAX_HIST || 20); // prod's MAX_HIST_TURNS (src/lumen.jsx)

// Pull the LIVE system prompt out of chat.js so this never drifts from prod.
const __dir = path.dirname(fileURLToPath(import.meta.url));
const chatSrc = readFileSync(path.join(__dir, "..", "netlify", "functions", "chat.js"), "utf8");
function extract(name) {
  const m = chatSrc.match(new RegExp("const " + name + " = (\"(?:[^\"\\\\]|\\\\.)*\");"));
  if (!m) throw new Error("Could not find " + name + " in netlify/functions/chat.js");
  return JSON.parse(m[1]);
}
const SYSTEM_PROMPT = extract("SYSTEM_PROMPT");

// Each config = a system-prompt lever (`extra`, appended as an extra instruction)
// + a history strategy. `history`:
//   "slide20" replicates prod today (src/lumen.jsx trims to the last MAX_HIST
//             messages) — which slides the prefix and defeats the v35 last-message
//             cache on the back half of every long chat (scan finding 21).
//   "full"    sends the whole growing history so the v35 breakpoint gets a cache
//             hit every turn. Very likely cheaper AND quality-neutral — this
//             harness is how you confirm that before touching prod. Add your own.
const CONFIGS = [
  { key: "baseline", extra: "", history: "slide20" },
  { key: "full-history", extra: "", history: "full" },
  { key: "terser-thought", extra: "\n\nLEVER: Keep the hidden <thought> block to at most TWO short sentences. Plan tersely; do not narrate your reasoning at length.", history: "slide20" },
  { key: "emit-on-change", extra: "\n\nLEVER: Emit the data markers (COMPANY, TOPICS, CHANNELS, REPORTS, ALERTS) ONLY when their values actually changed since your previous message. Always still emit the PROGRESS marker every turn.", history: "slide20" },
];

// Distinctive, INVENTED facts (not real brands) stated by the client early. Using
// invented names is deliberate: a retention check then measures TRUE context recall,
// not the model guessing a real brand from world knowledge. In a long chat these
// early facts fall out of a sliding window unless the model carries them forward —
// which is exactly finding 21's risk, so the probe below quantifies it per config.
const SEED_FACTS = {
  "company":       ["Northwind Athletics"],
  "competitor 1":  ["Velocity Sportswear"],
  "competitor 2":  ["Apex Running Co"],
  "competitor 3":  ["Terra Gear"],
  "campaign":      ["Project Skylark"],
  "market Japan":  ["Japan"],
  "market Brazil": ["Brazil"],
  "market Norway": ["Norway"],
  "contact email": ["dana@northwind.example"],
};
const CLIENT_PERSONA =
  "You are role-playing a CLIENT being onboarded onto Lumen, a social listening tool. You are Dana, marketing lead at Northwind Athletics (athletic footwear & apparel). You ALREADY gave your company, contact email, three competitors, campaign name, and target markets in your first message, so DO NOT re-list those specific names again — if the assistant asks about them, refer to them only generally (e.g. 'the three competitors I mentioned', 'the markets I listed') without repeating the exact names. Your goals: protect brand reputation, track your competitors, and catch customer issues early. Invent any other details consistently. Answer the assistant's latest message the way this client would: natural, cooperative, 1-2 sentences. If it shows options or a widget, just answer in plain language. When the assistant clearly signals the setup is complete and asks for final confirmation, reply exactly 'Yes, looks good.' Output ONLY the client's next message, no quotes, no commentary.";

async function call(model, system, messages, maxTokens) {
  const res = await fetch(ANTHROPIC_URL, {
    method: "POST",
    headers: { "content-type": "application/json", "x-api-key": KEY, "anthropic-version": VERSION },
    body: JSON.stringify({ model, max_tokens: maxTokens, system, messages }),
  });
  const data = await res.json();
  if (!res.ok || data.error) throw new Error("API error: " + JSON.stringify(data.error || res.status));
  return { text: (data.content || []).map((b) => b.text || "").join(""), usage: data.usage || {} };
}

const stripThought = (s) => s.replace(/<thought>[\s\S]*?<\/thought>/gi, "").trim();
const visibleOf = (s) => stripThought(s).replace(/%%[A-Z]+%%[\s\S]*?%%END%%/g, "").replace(/\[(WIDGET|SUGGESTIONS|TOPIC_SUGGESTION)[^\]]*\]/g, "").trim();
const costOf = (u) => (u.input * RATES.input + u.output * RATES.output + u.cacheWrite * RATES.cacheWrite + u.cacheRead * RATES.cacheRead) / 1e6;

// Apply the history strategy and replicate v35: the cache breakpoint goes on the
// last message, so the conversation prefix is billed at the cache-read rate. With
// "slide20" the window slides — message[0] changes each turn once the chat passes
// MAX_HIST, so the prefix stops matching and the read reverts to full price (the
// finding-21 miss); with "full" the prefix is stable and grows, so every turn hits.
function prepMessages(hist, strategy) {
  const win = strategy === "full" ? hist.slice() : hist.slice(-MAX_HIST);
  return win.map((m, i) =>
    i === win.length - 1
      ? { role: m.role, content: [{ type: "text", text: m.content, cache_control: { type: "ephemeral" } }] }
      : m
  );
}

// The client states all retention facts in the OPENING turn, verbatim and
// identical for every config and run. This makes recall a clean measure of the
// history strategy alone: in a long chat this first turn slides out of the
// "slide20" window (so the facts must be recalled from context that no longer
// holds them) but stays in "full". Without this the simulated persona's wording
// varied per run and swamped the signal at low RUNS.
const FACT_INTRO = "Hi! Quick intro before we begin: we're Northwind Athletics (athletic footwear and apparel). Our main contact email is dana@northwind.example. Please make sure we track our three competitors: Velocity Sportswear, Apex Running Co, and Terra Gear. Our current campaign is called Project Skylark, and our priority target markets are Japan, Brazil, and Norway.";

async function runConversation(cfg) {
  const system = [{ type: "text", text: SYSTEM_PROMPT + cfg.extra, cache_control: { type: "ephemeral" } }];
  const hist = [{ role: "user", content: "[BEGIN ONBOARDING] The client just opened their link. " + FACT_INTRO }];
  const transcript = [];
  const usage = { input: 0, output: 0, cacheRead: 0, cacheWrite: 0 };
  for (let turn = 0; turn < MAX_TURNS; turn++) {
    const a = await call(ASSISTANT_MODEL, system, prepMessages(hist, cfg.history), 2000);
    usage.input += a.usage.input_tokens || 0;
    usage.output += a.usage.output_tokens || 0;
    usage.cacheRead += a.usage.cache_read_input_tokens || 0;
    usage.cacheWrite += a.usage.cache_creation_input_tokens || 0;
    hist.push({ role: "assistant", content: stripThought(a.text) });
    const visible = visibleOf(a.text);
    transcript.push({ role: "assistant", text: visible });
    if (/%%PROGRESS%%[\s\S]*?"percent"\s*:\s*100/.test(a.text)) break;
    const c = await call(ASSISTANT_MODEL, [{ type: "text", text: CLIENT_PERSONA }],
      [{ role: "user", content: "Assistant said:\n\"" + visible + "\"\n\nReply as the client." }], 300);
    hist.push({ role: "user", content: c.text.trim() || "Okay." });
    transcript.push({ role: "client", text: c.text.trim() });
  }

  // Fact-retention probe (finding 21). Ask the model to read back the facts the
  // client stated early, using the SAME history strategy, then measure how many
  // survived. This is the signal a coarse quality score misses: under "slide20"
  // the early facts have fallen out of the window on a long chat, so recall drops
  // unless the model carried them forward; under "full" they are always in context.
  // Probe tokens are harness overhead and are NOT counted toward cost (like the
  // client/judge calls) — only the prod assistant turns above are.
  const probeQ = "Before we wrap up, please read back the exact details you have on file: our company name, every competitor you are tracking, our campaign name, all of our target markets, and our main contact email.";
  const probeHist = hist.slice();
  const lastM = probeHist[probeHist.length - 1];
  if (lastM && lastM.role === "user") probeHist[probeHist.length - 1] = { role: "user", content: lastM.content + "\n\n" + probeQ };
  else probeHist.push({ role: "user", content: probeQ });
  // recall === null means the probe FAILED (indistinguishable from a real 0 is a
  // bug we avoid: null renders "n/a", not 0%). A real number (incl. 0) is a measurement.
  let recall = null;
  try {
    const p = await call(ASSISTANT_MODEL, system, prepMessages(probeHist, cfg.history), 600);
    const probeText = visibleOf(p.text).toLowerCase();
    const factList = Object.values(SEED_FACTS).flat();
    const kept = factList.filter((s) => probeText.includes(s.toLowerCase()));
    const missing = factList.filter((s) => !probeText.includes(s.toLowerCase()));
    recall = factList.length ? kept.length / factList.length : null;
    process.stderr.write(`\n[recall ${cfg.key}] ${kept.length}/${factList.length}` + (missing.length ? " missing: " + missing.join(", ") : "") + "\n");
  } catch (e) { process.stderr.write(`\n[recall ${cfg.key}] PROBE FAILED: ${e.message}\n`); }

  return { transcript, usage, recall };
}

const JUDGE_RUBRIC =
  "You are grading a Lumen onboarding transcript for QUALITY (ignore cost). Score each 1-5 (5=excellent): coverage (captured company, goal, competitors, markets, channels, reports/alerts, users), coherence (natural, no stalls/repeats/dead-ends), integrity (stayed consultative, did NOT overstate that the setup is live). Return ONLY minified JSON on ONE line, no prose, no markdown fence, notes under 10 words: {\"coverage\":n,\"coherence\":n,\"integrity\":n,\"overall\":n,\"notes\":\"...\"}";

// Tolerant parse: strip any code fence, grab the JSON object, and if the model
// omitted `overall`, derive it from the sub-scores. Returns null if unrecoverable.
function parseJudge(text) {
  if (!text) return null;
  const m = text.replace(/```(?:json)?/gi, "").match(/\{[\s\S]*\}/);
  if (!m) return null;
  let o; try { o = JSON.parse(m[0]); } catch { return null; }
  if (typeof o.overall === "number") return o;
  const subs = ["coverage", "coherence", "integrity"].map((k) => o[k]).filter((x) => typeof x === "number");
  if (subs.length) { o.overall = subs.reduce((a, b) => a + b, 0) / subs.length; return o; }
  return null;
}
async function judge(transcript) {
  const t = transcript.map((m) => (m.role === "assistant" ? "ASSISTANT: " : "CLIENT: ") + m.text).join("\n\n").slice(0, 20000);
  for (let attempt = 0; attempt < 2; attempt++) {
    const r = await call(JUDGE_MODEL, [{ type: "text", text: JUDGE_RUBRIC }], [{ role: "user", content: t }], 800);
    const parsed = parseJudge(r.text);
    if (parsed) return parsed;
    process.stderr.write("\n[judge parse failed] raw: " + String(r.text || "").slice(0, 200).replace(/\s+/g, " ") + "\n");
  }
  return { overall: null, notes: "judge parse failed" };
}

const avg = (a) => { const v = a.filter((x) => x != null); return v.length ? v.reduce((s, x) => s + x, 0) / v.length : null; };

async function main() {
  // AB_ONLY=baseline,full-history restricts which configs run (cheaper focused runs).
  const only = (process.env.AB_ONLY || "").split(",").map((s) => s.trim()).filter(Boolean);
  const configs = only.length ? CONFIGS.filter((c) => only.includes(c.key)) : CONFIGS;
  const results = {};
  let log = "";
  for (const cfg of configs) {
    const costs = [], scores = [], recalls = [];
    for (let i = 0; i < RUNS; i++) {
      process.stderr.write(`\r${cfg.key}: run ${i + 1}/${RUNS}        `);
      const { transcript, usage, recall } = await runConversation(cfg);
      costs.push(costOf(usage));
      scores.push((await judge(transcript)).overall);
      recalls.push(recall);
      // Save the readable conversation so a HUMAN (or Claude) can judge quality
      // directly, bypassing the noisy auto-judge. This file is the quality gate.
      log += `\n\n===== CONFIG: ${cfg.key} — run ${i + 1}/${RUNS} (recall ${recall == null ? "n/a" : Math.round(recall * 100) + "%"}) =====\n`
        + transcript.map((m) => (m.role === "assistant" ? "ASSISTANT: " : "CLIENT:    ") + m.text).join("\n\n") + "\n";
    }
    results[cfg.key] = { avgCost: avg(costs), avgScore: avg(scores), avgRecall: avg(recalls) };
  }
  try {
    const outPath = path.join(__dir, "..", "ab-transcripts.txt");
    writeFileSync(outPath, `Lumen A/B transcripts — ${RUNS} run(s)/config, ${ASSISTANT_MODEL}\nConfigs: ${configs.map((c) => c.key).join(", ")}\n` + log);
    process.stderr.write("\nFull transcripts written to ab-transcripts.txt (send this file to Claude for the quality read)\n");
  } catch (e) { process.stderr.write("\n(could not write ab-transcripts.txt: " + e.message + ")\n"); }
  process.stderr.write("\r");
  const base = results.baseline;
  console.log("\n=== Lumen cost-lever A/B — " + RUNS + " runs each, " + ASSISTANT_MODEL + " ===\n");
  console.log("config".padEnd(18) + "avg $/convo".padEnd(14) + "quality/5".padEnd(12) + "recall".padEnd(9) + "vs baseline");
  for (const k of Object.keys(results)) {
    const r = results[k];
    const d = base.avgCost ? (r.avgCost - base.avgCost) / base.avgCost * 100 : 0;
    const q = r.avgScore == null ? "n/a" : r.avgScore.toFixed(2);       // null = judge failed, NOT a real 0
    const rc = r.avgRecall == null ? "n/a" : Math.round(r.avgRecall * 100) + "%"; // null = probe failed, NOT a real 0
    console.log(
      k.padEnd(18) +
      ("$" + (r.avgCost ?? 0).toFixed(4)).padEnd(14) +
      q.padEnd(12) +
      rc.padEnd(9) +
      (k === "baseline" ? "—" : (d >= 0 ? "+" : "") + d.toFixed(0) + "% cost")
    );
  }
  console.log("\nColumns: recall = % of the 9 facts the client stated early that the model read");
  console.log("back correctly at the end (the finding-21 signal a coarse quality score misses).");

  // Finding-21 verdict: the decision this harness exists to settle. Never decide
  // off a missing measurement — a null quality/recall means the judge or probe
  // failed, not that the value was zero.
  const f = results["full-history"];
  if (base && f) {
    const costPct = base.avgCost ? (f.avgCost - base.avgCost) / base.avgCost * 100 : 0;
    const qKnown = base.avgScore != null && f.avgScore != null;
    const rKnown = base.avgRecall != null && f.avgRecall != null;
    const qDelta = qKnown ? f.avgScore - base.avgScore : null;
    const rDelta = rKnown ? (f.avgRecall - base.avgRecall) * 100 : null;
    console.log("\n--- Finding 21: full-history vs baseline (slide-" + MAX_HIST + ") ---");
    console.log("  cost " + (costPct >= 0 ? "+" : "") + costPct.toFixed(0) + "%"
      + "   quality " + (qKnown ? (qDelta >= 0 ? "+" : "") + qDelta.toFixed(2) + "/5" : "n/a")
      + "   fact recall " + (rKnown ? (rDelta >= 0 ? "+" : "") + rDelta.toFixed(0) + " pts" : "n/a"));
    if (!qKnown || !rKnown) {
      console.log("  READ: quality and/or recall was not measured this run (judge or probe returned nothing — see the [judge]/[recall] lines on stderr). Do NOT decide; fix/re-run at AB_RUNS>=3.");
    } else {
      const qualityHolds = qDelta >= -0.2 && f.avgRecall >= base.avgRecall - 0.02;
      const notPricier = costPct <= 5;
      if (qualityHolds && notPricier)
        console.log("  READ: quality and recall hold (or improve) and it is not more expensive -> safe to raise/remove MAX_HIST_TURNS. Re-run at higher AB_RUNS to firm up.");
      else if (!qualityHolds)
        console.log("  READ: quality or recall dropped beyond noise -> do NOT adopt; keep the sliding window.");
      else
        console.log("  READ: quality holds but it costs more here -> weigh the trade-off; low AB_RUNS is noisy, re-run higher.");
    }
  }
  console.log("\nRule of thumb: adopt a lever only if cost drops (or holds) AND quality/recall stays within noise of baseline.");
  if (RUNS < 3) console.log("NOTE: AB_RUNS=" + RUNS + " is a smoke run, too noisy to decide. Use AB_RUNS=3+ for the real numbers.");
}
main().catch((e) => { console.error("\nHarness failed:", e.message); process.exit(1); });
