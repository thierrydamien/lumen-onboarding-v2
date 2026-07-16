#!/usr/bin/env node
/**
 * Lumen cost-lever A/B harness.
 *
 * Answers "does this cost lever reduce spend WITHOUT reducing quality?" with
 * numbers instead of a guess. It drives a full simulated onboarding against the
 * real Anthropic API for the baseline prompt and for each lever variant, then a
 * judge model scores the transcripts. Your live prompt is never modified — each
 * lever is just an extra instruction appended in memory here.
 *
 * RUN:  ANTHROPIC_API_KEY=sk-... node tools/ab-harness.mjs
 * ENV (optional): AB_RUNS (default 3), AB_TURNS (16), AB_MODEL, AB_JUDGE_MODEL.
 * COST: makes real API calls — a few dollars per run. That's the point: spend a
 *       little to know before changing the live prompt. Nothing ships from here.
 *
 * WHAT'S MEASURED: only the ASSISTANT's tokens count toward cost (that's what
 * runs in prod). The simulated client + judge calls are harness overhead. The
 * system prompt is cached like prod; the v35 message-prefix cache is not
 * replicated, so absolute $/convo runs a little high — the RELATIVE delta between
 * configs (which is what the decision rests on) is what to read.
 */
import { readFileSync } from "node:fs";
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

// Pull the LIVE system prompt out of chat.js so this never drifts from prod.
const __dir = path.dirname(fileURLToPath(import.meta.url));
const chatSrc = readFileSync(path.join(__dir, "..", "netlify", "functions", "chat.js"), "utf8");
function extract(name) {
  const m = chatSrc.match(new RegExp("const " + name + " = (\"(?:[^\"\\\\]|\\\\.)*\");"));
  if (!m) throw new Error("Could not find " + name + " in netlify/functions/chat.js");
  return JSON.parse(m[1]);
}
const SYSTEM_PROMPT = extract("SYSTEM_PROMPT");

// The levers under test (appended as an extra system instruction). Add your own.
const CONFIGS = [
  { key: "baseline", extra: "" },
  { key: "terser-thought", extra: "\n\nLEVER: Keep the hidden <thought> block to at most TWO short sentences. Plan tersely; do not narrate your reasoning at length." },
  { key: "emit-on-change", extra: "\n\nLEVER: Emit the data markers (COMPANY, TOPICS, CHANNELS, REPORTS, ALERTS) ONLY when their values actually changed since your previous message. Always still emit the PROGRESS marker every turn." },
];

const CLIENT_PERSONA =
  "You are role-playing a CLIENT being onboarded onto Lumen, a social listening tool. You are the marketing lead at Acme Corp, a consumer-goods (footwear & apparel) brand; your email is jane@acmecorp.com. Your goals: protect brand reputation, track competitors (Nike, Adidas, Puma), and catch customer issues early. Invent consistent, realistic details as needed and keep them consistent. Answer the assistant's latest message the way this client would: natural, cooperative, 1-2 sentences. If it shows options or a widget, just answer in plain language. When the assistant clearly signals the setup is complete and asks for final confirmation, reply exactly 'Yes, looks good.' Output ONLY the client's next message — no quotes, no commentary.";

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

async function runConversation(extra) {
  const system = [{ type: "text", text: SYSTEM_PROMPT + extra, cache_control: { type: "ephemeral" } }];
  const hist = [{ role: "user", content: "[BEGIN ONBOARDING] The client just opened their link." }];
  const transcript = [];
  const usage = { input: 0, output: 0, cacheRead: 0, cacheWrite: 0 };
  for (let turn = 0; turn < MAX_TURNS; turn++) {
    const a = await call(ASSISTANT_MODEL, system, hist.slice(-24), 2000);
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
  return { transcript, usage };
}

const JUDGE_RUBRIC =
  "You are grading a Lumen onboarding conversation transcript for QUALITY (ignore cost). Score each 1-5 (5=excellent): coverage (did the assistant capture company, goal, topics/competitors, markets, channels, reports/alerts, users), coherence (natural, no stalls, no repeated questions, no dead ends), integrity (stayed on-task and consultative, no overstating that the setup is live). Return ONLY JSON: {\"coverage\":n,\"coherence\":n,\"integrity\":n,\"overall\":n,\"notes\":\"one sentence\"}.";
async function judge(transcript) {
  const t = transcript.map((m) => (m.role === "assistant" ? "ASSISTANT: " : "CLIENT: ") + m.text).join("\n\n");
  const r = await call(JUDGE_MODEL, [{ type: "text", text: JUDGE_RUBRIC }], [{ role: "user", content: t.slice(0, 20000) }], 400);
  try { return JSON.parse(r.text.match(/\{[\s\S]*\}/)[0]); } catch { return { overall: null, notes: "judge parse failed" }; }
}

const avg = (a) => { const v = a.filter((x) => x != null); return v.length ? v.reduce((s, x) => s + x, 0) / v.length : null; };

async function main() {
  const results = {};
  for (const cfg of CONFIGS) {
    const costs = [], scores = [];
    for (let i = 0; i < RUNS; i++) {
      process.stderr.write(`\r${cfg.key}: run ${i + 1}/${RUNS}        `);
      const { transcript, usage } = await runConversation(cfg.extra);
      costs.push(costOf(usage));
      scores.push((await judge(transcript)).overall);
    }
    results[cfg.key] = { avgCost: avg(costs), avgScore: avg(scores) };
  }
  process.stderr.write("\r");
  const base = results.baseline;
  console.log("\n=== Lumen cost-lever A/B — " + RUNS + " runs each, " + ASSISTANT_MODEL + " ===\n");
  console.log("config".padEnd(18) + "avg $/convo".padEnd(14) + "quality/5".padEnd(12) + "vs baseline");
  for (const k of Object.keys(results)) {
    const r = results[k];
    const d = base.avgCost ? (r.avgCost - base.avgCost) / base.avgCost * 100 : 0;
    console.log(
      k.padEnd(18) +
      ("$" + (r.avgCost ?? 0).toFixed(4)).padEnd(14) +
      ((r.avgScore ?? 0).toFixed(2)).padEnd(12) +
      (k === "baseline" ? "—" : (d >= 0 ? "+" : "") + d.toFixed(0) + "% cost, quality " + (r.avgScore ?? 0).toFixed(2) + " vs " + (base.avgScore ?? 0).toFixed(2))
    );
  }
  console.log("\nRule of thumb: adopt a lever only if cost drops meaningfully AND quality stays within ~0.2 of baseline (noise). Bump AB_RUNS for tighter numbers.");
}
main().catch((e) => { console.error("\nHarness failed:", e.message); process.exit(1); });
