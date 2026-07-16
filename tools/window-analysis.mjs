// API-free analysis of finding 21: does the live slice(-MAX_HIST) window actually
// drop early facts from the model's context in a realistic onboarding? Uses the
// app's own built-in test conversation (loadTestSession in src/lumen.jsx) as a
// representative ~30-message chat, and simulates what the model "sees" each turn.
import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __dir = path.dirname(fileURLToPath(import.meta.url));
const src = readFileSync(path.join(__dir, "..", "src", "lumen.jsx"), "utf8");
const MAX_HIST = Number(process.env.AB_MAX_HIST || 20); // prod's MAX_HIST_TURNS

// Extract the `const tm = [ ... ];` message array (role + content), tolerant of
// escaped quotes. We only need role/content, not the widget metadata.
const tmStart = src.indexOf("const tm = [");
const tmEnd = src.indexOf("];", tmStart);
const tmBlock = src.slice(tmStart, tmEnd);
const re = /\{role:"(assistant|user)",content:"((?:[^"\\]|\\.)*)"/g;
const msgs = [];
let m;
while ((m = re.exec(tmBlock)) !== null) {
  msgs.push({ role: m[1], content: JSON.parse('"' + m[2] + '"') });
}

// Facts a consultant must not lose. Each is one or more surface strings; a fact
// "survives" a window if ANY of its strings appears in ANY message in that window
// (client statement OR an assistant summary that carried it forward).
const FACTS = {
  "company name":            ["Acme Corp"],
  "industry":                ["footwear", "apparel"],
  "competitor Nike":         ["Nike"],
  "competitor Adidas":       ["Adidas"],
  "competitor Puma":         ["Puma"],
  "summer campaign":         ["summer campaign", "Acme Summer", "AcmeSummer"],
  "objective Reputation":    ["Reputation Management", "reputation"],
  "objective Competitive":   ["Competitive Intelligence", "competitor", "competitive"],
  "objective Issue/complaint":["Issue Tracking", "issue", "complaint"],
  "market Canada":           ["Canada"],
  "market Australia":        ["Australia"],
  "market France":           ["France"],
  "language French/German":  ["French", "German"],
  "contact email":           ["jane@acmecorp.com"],
  "team Customer Experience":["Customer Experience"],
};

const has = (text, strings) => strings.some((s) => text.toLowerCase().includes(s.toLowerCase()));

// Assistant turns are the calls that matter (the model generating). For each
// assistant message index T, the model's context was messages[0..T-1] trimmed to
// the last MAX_HIST. Report, per fact, whether it was in-window at each assistant
// turn, and the LAST assistant turn where it was in the RAW window at all.
const assistantTurns = [];
for (let T = 0; T < msgs.length; T++) {
  if (msgs[T].role !== "assistant") continue;
  const ctxFull = msgs.slice(0, T);            // full history the model could have
  const ctxWin = ctxFull.slice(-MAX_HIST);     // what slide20 actually sends
  assistantTurns.push({ T, winText: ctxWin.map((x) => x.content).join("\n"), fullText: ctxFull.map((x) => x.content).join("\n"), winSize: ctxWin.length });
}

// The decisive turns: the two summary turns + the final finish, where the model
// restates the whole brief. If a fact is missing from the window THERE, slide20
// risks a degraded summary; full-history would still have it.
const lastTurns = assistantTurns.slice(-4); // final stretch (summaries + finish)

console.log(`Extracted ${msgs.length} messages (${assistantTurns.length} assistant turns) from loadTestSession.`);
console.log(`MAX_HIST = ${MAX_HIST}. Final assistant turn sees messages ${msgs.length - MAX_HIST}..${msgs.length - 1} (window of ${Math.min(MAX_HIST, msgs.length)}).\n`);

let anyLost = false, lostInWinButInFull = [];
console.log("Fact retention at the FINAL assistant turn (the closing brief/summary):");
const finalTurn = assistantTurns[assistantTurns.length - 1];
for (const [name, strings] of Object.entries(FACTS)) {
  const inWin = has(finalTurn.winText, strings);
  const inFull = has(finalTurn.fullText, strings);
  const mark = inWin ? "kept " : (inFull ? "DROPPED (in full, not in window)" : "n/a  ");
  if (!inWin && inFull) { anyLost = true; lostInWinButInFull.push(name); }
  console.log(`  ${inWin ? "OK " : "!! "} ${name.padEnd(26)} window:${inWin ? "yes" : "no "}  fullHistory:${inFull ? "yes" : "no"}`);
}

console.log("\nAcross the final 4 assistant turns (summaries + finish), facts present only in full history, not in the slide20 window:");
for (const turn of lastTurns) {
  const lost = Object.entries(FACTS).filter(([, s]) => !has(turn.winText, s) && has(turn.fullText, s)).map(([n]) => n);
  console.log(`  turn @msg#${turn.T} (window=${turn.winSize} msgs): ${lost.length ? lost.join(", ") : "none lost"}`);
}

console.log("\n=== VERDICT ===");
if (!anyLost) {
  console.log("At the closing turn, EVERY tracked fact is still inside the slide20 window");
  console.log("(carried forward by the assistant's own running summaries, which stay in the");
  console.log("last 20 messages). So in a representative onboarding, slide20 loses no material");
  console.log("context vs full history -> full-history is essentially a pure COST optimisation");
  console.log("with negligible quality risk. Still confirm empirically with the harness before shipping.");
} else {
  console.log("Facts dropped from the slide20 window at the closing turn (present in full history):");
  console.log("  " + lostInWinButInFull.join(", "));
  console.log("These are the concrete quality risks slide20 carries and full-history would fix.");
  console.log("This RAISES the value of full-history (retention win), but confirm with the harness.");
}
