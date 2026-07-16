import { useState, useRef, useEffect, useCallback, useMemo, memo, Component } from "react";
// xlsx is the bulk of the JS bundle and is only needed when a client uploads a
// file (QUERIES) or exports/sends the brief — never on first paint. Load it
// lazily so it code-splits into its own chunk and doesn't tax the initial load.
let _xlsxMod = null;
async function loadXLSX() {
  if (!_xlsxMod) { const m = await import("xlsx"); _xlsxMod = m.default || m; }
  return _xlsxMod;
}

// ================= LIVE CONFIG =================
// Frontends are served from the same Netlify site as the functions, so these
// are same-origin relative paths (no CORS).
const CHAT_ENDPOINT = "/.netlify/functions/chat";
const SESSION_ENDPOINT = "/.netlify/functions/session";
const SEED_ENDPOINT = "/.netlify/functions/seed";
const SHEET_ENDPOINT = "/.netlify/functions/sheet";
// Demo-only controls (preview / simulate / rewind) are hidden on the live site.
const DEV = false;

// The Sales page stores the profile server-side and puts only an opaque id in
// the client link (?s=<id>). Fetch the CLIENT-SAFE fields (no consultant notes;
// notes are returned only to the token-authenticated dashboard). Returns
// { seed, seedId } or { seed:null, seedId:null }.
// fetch with a hard timeout so a hung request never freezes the UI (a spinner
// that never resolves, a Send button stuck disabled). Aborts after `ms` and
// rejects like any network error; every caller already handles fetch rejection.
async function fetchWithTimeout(url, opts = {}, ms = 30000) {
  const ac = new AbortController();
  const t = setTimeout(() => ac.abort(), ms);
  try { return await fetch(url, { ...opts, signal: ac.signal }); }
  finally { clearTimeout(t); }
}

async function fetchSeedFromURL() {
  const id = new URLSearchParams(location.search).get("s");
  if (!id) return { seed: null, seedId: null };
  // Retry once on a transient failure. If it still fails, KEEP the seedId so the
  // consultant-notes linkage survives (the completed record can still be joined
  // to the seed store) rather than silently downgrading a prepared session to a
  // generic one and orphaning the notes.
  for (let attempt = 0; attempt < 2; attempt++) {
    try {
      const res = await fetchWithTimeout(`${SEED_ENDPOINT}?id=${encodeURIComponent(id)}`, {}, 15000);
      if (res.ok) {
        const data = await res.json();
        const seed = data && data.seed && data.seed.company ? data.seed : null;
        return { seed, seedId: id, seedError: !seed };
      }
    } catch { /* fall through to retry */ }
  }
  return { seed: null, seedId: id, seedError: true };
}

const MIN_MS = 1500;
const P = "#012B3A";
const A = "#7E48EC";   // Lumen purple (sampled from official wordmark)
const NAVY = "#1A3B7B"; // Talkwalker navy (sampled from official wordmark)
const LINK = "#6D28D9"; // interactive purple

const SECTION_KEYS   = ["company","path","topics","channels","reports","users"];
const SECTION_LABELS = { company:"About you", path:"Approach", topics:"What to track", channels:"Where to look", reports:"Reports", users:"Your team" };
const WIDGET_MAX     = { OBJECTIVES:3, TIMEZONE:1 };
const MARKETS_OPT    = ["United States","United Kingdom","France","Germany","Spain","Italy","Netherlands","Canada","Australia","Brazil","Japan","South Korea","India","Middle East","APAC","LATAM","Global"];
const LANG_OPT       = ["English","French","German","Spanish","Italian","Dutch","Portuguese","Japanese","Korean","Mandarin","Arabic","Hindi"];
const OBJ_OPT        = ["Competitive Intelligence","Campaign Optimization","Content Ideation & Recommendation","Reputation Management","Social Measurement","Brand Health Measurement","Issue Tracking","PR Measurement","Influencer Management","Consumer Insights","Trend Research"];
const TEAM_OPT       = ["Marketing","Communications","PR","Brand","Digital","Social Media","Legal","Product","Research","Executive","Customer Experience","Corporate Affairs"];
const TZ_OPT         = ["GMT / UTC","CET (UTC+1)","EET (UTC+2)","GST (UTC+4)","IST (UTC+5:30)","SGT (UTC+8)","JST (UTC+9)","AEST (UTC+10)","EST (UTC-5)","CST (UTC-6)","MST (UTC-7)","PST (UTC-8)"];
const EMAIL_RE       = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
// OBJECTIVES widget data was a plain array before ranking was added; normalize
// both shapes so saved/resumed sessions from older versions keep working.
const normObjectives = d => Array.isArray(d) ? {ranked:d, details:""}
  : (d && typeof d === "object") ? {ranked:d.ranked||[], details:d.details||""}
  : {ranked:[], details:""};
const fmtRanked = d => { const n = normObjectives(d); return n.ranked.length ? n.ranked.map((o,i)=>`${i+1}. ${o}`).join(", ") : ""; };

const SMOKE_PERSONAS = [
  { key:"terse",    label:"Terse expert",       emoji:"🧊" },
  { key:"novice",   label:"Rambling novice",    emoji:"🌀" },
  { key:"offtopic", label:"Off-topic wanderer", emoji:"🦋" },
];


const CLIENT_BASE = "You are role-playing a CLIENT being onboarded onto Lumen, a social listening tool. Your profile: Jane Smith, Marketing Director at Acme Corp, a consumer goods (footwear and apparel) company. Email jane@acmecorp.com. Competitors: Nike, Adidas, Puma. Markets: US and UK.";

// Build the simulated-client profile from the seed a salesperson generated, so a
// simulate run tests the same handoff a real client link produces. Falls back to
// the original Jane Smith / Acme profile when there is no seed.
function buildClientBase(sd) {
  if (!sd || !sd.company) return CLIENT_BASE;
  const contact  = sd.contactName || "the primary contact";
  const company  = sd.company;
  const industry = sd.industry ? ` in ${sd.industry}` : "";
  const email    = sd.email || "the contact's work email";
  return `You are role-playing a CLIENT being onboarded onto Lumen, a social listening tool. Your profile: ${contact}, a contact at ${company}${industry}. Email ${email}. Invent consistent, realistic details for anything not specified (competitors, markets, products, campaigns) and keep them consistent for the entire conversation. Never contradict the profile above.`;
}

// Language directive appended to the simulated client's persona so it answers in
// the same language the session is running in. English needs no directive.
function clientLangLine(uiLang) {
  return uiLang && uiLang !== "English" ? ` Write ALL of your replies in ${uiLang}.` : "";
}

// Seed- and language-aware simulate personas. Same shape as before; the menu only
// surfaces terse/novice/offtopic, rambler/french are kept for completeness.
function buildSmokePrompts(sd, uiLang) {
  const base = buildClientBase(sd), lang = clientLangLine(uiLang);
  return {
    terse:    `${base} Goal: competitive intelligence and brand reputation. STYLE: extremely terse and expert. Maximum 12 words per reply. Use industry jargon (SOV, boolean, sentiment). Slightly impatient. Never volunteer extra information.${lang} Output ONLY the client's next chat message, no quotes, no commentary.`,
    novice:   `${base} Goal: brand reputation, vaguely understood. STYLE: friendly but vague and unsure, new to social listening. 2 to 4 sentences per reply. Start vague ('marketing stuff', 'our competitors I guess') and only give specifics when pushed. Occasionally ask what a term means.${lang} Output ONLY the client's next chat message, no quotes, no commentary.`,
    offtopic: `${base} Goal: brand reputation and competitive intelligence. STYLE: cooperative but easily distracted. Roughly every second reply, FIRST ask one unrelated question (pricing, contract length, an unrelated feature, small talk), THEN briefly answer the assistant's actual question.${lang} Output ONLY the client's next chat message, no quotes, no commentary.`,
    rambler:  `${base} Goal: brand reputation, buried in noise. STYLE: extremely long-winded. Every reply is 4 to 7 sentences containing exactly ONE useful fact wrapped in anecdotes, tangents about colleagues, and hedging. Never structure your answers.${lang} Output ONLY the client's next chat message, no quotes, no commentary.`,
    french:   `${base} Goal: veille concurrentielle et réputation de marque. STYLE: répond UNIQUEMENT en français, poliment, 1 à 3 phrases. Ne passe jamais à l'anglais même si l'assistant écrit en anglais. Output ONLY the client's next chat message, no quotes, no commentary.`,
  };
}

// The exact first message a real client link produces. Shared by the real flow
// (startConvo) and the simulate flow (runSmokeTest) so the two can't drift.
// Carries a language directive when the session isn't in English.
function seededOpener(sd, uiLang) {
  const langDirective = uiLang && uiLang !== "English" ? ` Please conduct the entire conversation in ${uiLang}.` : "";
  if (sd) {
    return `[SEEDED SESSION] Prepared by the Lumen team. Company: ${sd.company}. Contact: ${sd.contactName}${sd.email?` (${sd.email})`:""}.${sd.industry?` Industry: ${sd.industry}.`:""}${sd.notes?` Consultant notes (do not read back to the client): ${sd.notes}.`:""} The client has just opened their link.${langDirective}`;
  }
  return `Hello, I'm ready to get started.${langDirective}`;
}
const SMOKE_WIDGET_DATA = {
  PATH:"guided", QUERIES:"__skip__",
  MARKETS:["United States","United Kingdom"], LANGUAGES:["English"],
  OBJECTIVES:{ranked:["Reputation Management","Competitive Intelligence"],details:""},
  TEAMS:["Marketing","PR"], TIMEZONE:["GMT / UTC"],
  USERS:[{firstName:"Jane",lastName:"Smith",email:"jane@acmecorp.com",role:"Marketing Director",access:"Admin"}],
};
const SMOKE_MAX_TURNS = 22;

// Widget auto-fill for the simulate flow. USERS is derived from the seeded contact
// (instead of hardcoded Jane) so the simulated run matches a real seeded session.
function buildSmokeWidgetData(sd) {
  if (!sd || !sd.company) return SMOKE_WIDGET_DATA;
  const first = (sd.contactName||"").split(" ")[0] || "Alex";
  const last  = (sd.contactName||"").split(" ").slice(1).join(" ") || "";
  const email = sd.email || `${first.toLowerCase()}@${(sd.company||"client").toLowerCase().replace(/[^a-z0-9]/g,"")}.com`;
  return { ...SMOKE_WIDGET_DATA, USERS:[{firstName:first,lastName:last,email,role:"Primary contact",access:"Admin"}] };
}

// ================= CLIENT-SIDE LANGUAGE ==================
// Sales seeds a language; the client can override it on the welcome screen with one
// tap. The choice drives the welcome-screen copy, RTL for Arabic, the seeded opener,
// and the simulated client persona. Only the welcome-screen shell is translated; the
// live conversation follows the LANGUAGE rule in the system prompt.
const UI_LANGS = [
  { code:"English", native:"English" },
  { code:"French",  native:"Français" },
  { code:"German",  native:"Deutsch" },
  { code:"Spanish", native:"Español" },
  { code:"Italian", native:"Italiano" },
  { code:"Arabic",  native:"العربية" },
];

const I18N = {
  English: {
    welcomeTitle:       "Welcome to Lumen Onboarding",
    welcomeTitleSeeded: "Welcome, {name}!",
    welcomeSub:         "We\u2019ll ask about your goals, markets, and team \u2014 then generate your Lumen setup brief.",
    welcomeSubSeeded:   "Your Lumen team prepared this session for {company}. We\u2019ll talk through your goals, markets, and team \u2014 and build your setup brief as we go.",
    step1Title: "About 15 minutes",
    step1Desc:  "About 15 minutes. Pause anytime — reopen this link on the same device and you'll pick up where you left off.",
    step2Title: "A conversation, not a form",
    step2Desc:  "We'll cover your goals, what to track, where your audience talks, reports, and your team.",
    step3Title: "Then we take over",
    step3Desc:  "Your setup brief goes straight to your Lumen team. A consultant contacts you within 2 business days to book your review call.",
    disclaimer: "You'll be chatting with an AI assistant. Everything you share is reviewed by a Lumen consultant before any setup work begins.",
    startBtn:       "Start \u2014 about 15 min",
    startBtnSeeded: "Start {company}'s setup \u2014 about 15 min",
    thinking:       "Assistant is thinking\u2026",
    correctionHint: "Spot something wrong? Just tell me in the chat \u2014 \u201cactually, our main market is Germany\u201d \u2014 and I'll fix it.",
  },
  French: {
    welcomeTitle:       "Bienvenue dans l'intégration Lumen",
    welcomeTitleSeeded: "Bienvenue, {name} !",
    welcomeSub:         "Nous vous poserons des questions sur vos objectifs, vos marchés et votre équipe, puis nous générerons votre brief de configuration Lumen.",
    welcomeSubSeeded:   "Votre équipe Lumen a préparé cette session pour {company}. Nous aborderons vos objectifs, vos marchés et votre équipe, et construirons votre brief au fur et à mesure.",
    step1Title: "Environ 15 minutes",
    step1Desc:  "Environ 15 minutes. Faites une pause quand vous voulez : rouvrez ce lien sur le même appareil et vous reprendrez là où vous vous étiez arrêté.",
    step2Title: "Une conversation, pas un formulaire",
    step2Desc:  "Nous aborderons vos objectifs, ce qu'il faut suivre, où votre audience s'exprime, les rapports et votre équipe.",
    step3Title: "Ensuite, nous prenons le relais",
    step3Desc:  "Votre brief de configuration est transmis directement à votre équipe Lumen. Un consultant vous contacte sous 2 jours ouvrés pour planifier votre appel de révision.",
    disclaimer: "Vous échangez avec un assistant IA. Tout ce que vous partagez est examiné par un consultant Lumen avant tout travail de configuration.",
    startBtn:       "Commencer (environ 15 min)",
    startBtnSeeded: "Démarrer la configuration de {company} (environ 15 min)",
    thinking:       "L'assistant réfléchit\u2026",
    correctionHint: "Vous voyez une erreur ? Dites-le-moi simplement dans le chat \u2014 « en fait, notre marché principal est l'Allemagne » \u2014 et je la corrigerai.",
  },
  German: {
    welcomeTitle:       "Willkommen beim Lumen-Onboarding",
    welcomeTitleSeeded: "Willkommen, {name}!",
    welcomeSub:         "Wir fragen nach Ihren Zielen, Märkten und Ihrem Team und erstellen anschließend Ihr Lumen-Setup-Briefing.",
    welcomeSubSeeded:   "Ihr Lumen-Team hat diese Sitzung für {company} vorbereitet. Wir besprechen Ihre Ziele, Märkte und Ihr Team und erstellen Ihr Setup-Briefing Schritt für Schritt.",
    step1Title: "Etwa 15 Minuten",
    step1Desc:  "Etwa 15 Minuten. Jederzeit pausieren: Öffnen Sie diesen Link auf demselben Gerät erneut und Sie machen dort weiter, wo Sie aufgehört haben.",
    step2Title: "Ein Gespräch, kein Formular",
    step2Desc:  "Wir behandeln Ihre Ziele, was Sie beobachten möchten, wo Ihr Publikum spricht, Berichte und Ihr Team.",
    step3Title: "Dann übernehmen wir",
    step3Desc:  "Ihr Setup-Briefing geht direkt an Ihr Lumen-Team. Ein Berater kontaktiert Sie innerhalb von 2 Werktagen, um Ihren Review-Termin zu vereinbaren.",
    disclaimer: "Sie chatten mit einem KI-Assistenten. Alles, was Sie teilen, wird von einem Lumen-Berater geprüft, bevor die Einrichtung beginnt.",
    startBtn:       "Starten (etwa 15 Min.)",
    startBtnSeeded: "Einrichtung für {company} starten (etwa 15 Min.)",
    thinking:       "Der Assistent denkt nach\u2026",
    correctionHint: "Etwas stimmt nicht? Sagen Sie es mir einfach im Chat \u2014 „eigentlich ist unser Hauptmarkt Deutschland“ \u2014 und ich korrigiere es.",
  },
  Spanish: {
    welcomeTitle:       "Bienvenido a la incorporación de Lumen",
    welcomeTitleSeeded: "¡Bienvenido, {name}!",
    welcomeSub:         "Le preguntaremos por sus objetivos, mercados y equipo, y luego generaremos su resumen de configuración de Lumen.",
    welcomeSubSeeded:   "Su equipo de Lumen preparó esta sesión para {company}. Hablaremos de sus objetivos, mercados y equipo, y crearemos su resumen de configuración sobre la marcha.",
    step1Title: "Unos 15 minutos",
    step1Desc:  "Unos 15 minutos. Haga una pausa cuando quiera: vuelva a abrir este enlace en el mismo dispositivo y continuará donde lo dejó.",
    step2Title: "Una conversación, no un formulario",
    step2Desc:  "Cubriremos sus objetivos, qué monitorizar, dónde habla su audiencia, los informes y su equipo.",
    step3Title: "Después nos encargamos nosotros",
    step3Desc:  "Su resumen de configuración va directamente a su equipo de Lumen. Un consultor le contactará en un plazo de 2 días laborables para agendar su llamada de revisión.",
    disclaimer: "Está chateando con un asistente de IA. Todo lo que comparta es revisado por un consultor de Lumen antes de iniciar cualquier trabajo de configuración.",
    startBtn:       "Comenzar (unos 15 min)",
    startBtnSeeded: "Comenzar la configuración de {company} (unos 15 min)",
    thinking:       "El asistente está pensando\u2026",
    correctionHint: "¿Ve algo incorrecto? Solo dígamelo en el chat \u2014 «en realidad, nuestro mercado principal es Alemania» \u2014 y lo corregiré.",
  },
  Italian: {
    welcomeTitle:       "Benvenuto nell'onboarding di Lumen",
    welcomeTitleSeeded: "Benvenuto, {name}!",
    welcomeSub:         "Ti chiederemo i tuoi obiettivi, i mercati e il team, poi genereremo il tuo brief di configurazione Lumen.",
    welcomeSubSeeded:   "Il tuo team Lumen ha preparato questa sessione per {company}. Parleremo dei tuoi obiettivi, dei mercati e del team, e costruiremo il tuo brief di configurazione strada facendo.",
    step1Title: "Circa 15 minuti",
    step1Desc:  "Circa 15 minuti. Metti in pausa quando vuoi: riapri questo link sullo stesso dispositivo e riprenderai da dove avevi lasciato.",
    step2Title: "Una conversazione, non un modulo",
    step2Desc:  "Copriremo i tuoi obiettivi, cosa monitorare, dove parla il tuo pubblico, i report e il tuo team.",
    step3Title: "Poi ci pensiamo noi",
    step3Desc:  "Il tuo brief di configurazione va direttamente al tuo team Lumen. Un consulente ti contatterà entro 2 giorni lavorativi per fissare la tua call di revisione.",
    disclaimer: "Stai chattando con un assistente IA. Tutto ciò che condividi viene esaminato da un consulente Lumen prima di iniziare qualsiasi attività di configurazione.",
    startBtn:       "Inizia (circa 15 min)",
    startBtnSeeded: "Avvia la configurazione di {company} (circa 15 min)",
    thinking:       "L'assistente sta pensando\u2026",
    correctionHint: "Noti qualcosa di sbagliato? Dimmelo semplicemente in chat \u2014 «in realtà, il nostro mercato principale è la Germania» \u2014 e lo correggerò.",
  },
  Arabic: {
    welcomeTitle:       "مرحبًا بك في إعداد Lumen",
    welcomeTitleSeeded: "مرحبًا، {name}!",
    welcomeSub:         "سنسألك عن أهدافك وأسواقك وفريقك، ثم ننشئ ملخص إعداد Lumen الخاص بك.",
    welcomeSubSeeded:   "أعدّ فريق Lumen هذه الجلسة لـ {company}. سنتحدث عن أهدافك وأسواقك وفريقك، وننشئ ملخص الإعداد الخاص بك خطوة بخطوة.",
    step1Title: "حوالي 15 دقيقة",
    step1Desc:  "حوالي 15 دقيقة. توقف مؤقتًا متى شئت: أعد فتح هذا الرابط على الجهاز نفسه وستتابع من حيث توقفت.",
    step2Title: "محادثة، وليست نموذجًا",
    step2Desc:  "سنغطي أهدافك، وما الذي تريد متابعته، وأين يتحدث جمهورك، والتقارير، وفريقك.",
    step3Title: "ثم نتولى نحن الأمر",
    step3Desc:  "يُرسل ملخص الإعداد الخاص بك مباشرةً إلى فريق Lumen. سيتواصل معك أحد الاستشاريين خلال يومَي عمل لتحديد موعد مكالمة المراجعة.",
    disclaimer: "أنت تتحدث مع مساعد ذكاء اصطناعي. تتم مراجعة كل ما تشاركه من قِبل استشاري Lumen قبل بدء أي عمل إعداد.",
    startBtn:       "ابدأ (حوالي 15 دقيقة)",
    startBtnSeeded: "ابدأ إعداد {company} (حوالي 15 دقيقة)",
    thinking:       "المساعد يفكّر\u2026",
    correctionHint: "لاحظت شيئًا غير صحيح؟ فقط أخبرني في المحادثة \u2014 «في الواقع، سوقنا الرئيسي هو ألمانيا» \u2014 وسأصححه.",
  },
};

function L(key, lang, vars) {
  const dict = I18N[lang] || I18N.English;
  let s = (dict[key] != null ? dict[key] : I18N.English[key]) || "";
  if (vars) for (const k in vars) s = s.replace(new RegExp("\\{" + k + "\\}", "g"), vars[k]);
  return s;
}

// Widget-chrome localization. Option VALUES (markets, objectives, teams,
// timezones) stay in English on purpose — they are Lumen's product taxonomy and
// are stored in English in the brief. Only the chrome (buttons, hints,
// placeholders, tooltips) follows the client's language, so a non-English chat
// no longer renders an all-English form.
const WI18N = {
  English: { "confirm":"Confirm", "skip":"Skip", "add":"+ Add", "customValue":"Type a custom value…", "somethingElse":"Something else? Type it here…", "max":"max", "selected":"selected", "limitReached":"limit reached", "prioritiesHdr":"Your priorities — #1 is where we start", "confirmPriorities":"Confirm priorities", "objDetailsPh":"Anything else about your objectives? (optional)", "firstName":"First name", "lastName":"Last name", "roleDept":"Role / dept", "email":"Email", "invalidEmail":"Invalid email", "addUser":"+ Add user", "confirmUsers":"Confirm users", "topicName":"Topic name", "keywordsPh":"Keywords…", "dragPrioritize":"Drag to prioritize", "kept":"kept", "discarded":"discarded", "pending":"pending", "submitQueries":"Submit queries", "noQueries":"No queries", "importFile":"📎 Or import a file (.txt, .csv, .xlsx)", "pasteQueries":"Paste your existing queries here…", "hintSelectAll":"Select all that apply.", "hintTeams":"Select all teams that will use Lumen.", "hintObjectives":"Pick up to 3, then set their priority — your #1 decides what we build first.", "hintTimezone":"Select your primary timezone.", "phMarket":"Type a market…", "phLanguage":"Type a language…", "phTeam":"Type a team…", "whyMarkets":"So results are scoped to the regions you actually operate in.", "whyTeams":"Helps us tailor dashboards to the people who'll use them.", "whyUsers":"Who should have access — just you for now is fine.", "whyQueries":"If you already track queries elsewhere, we can migrate them.", "whyTopics":"Topics are the subjects Lumen will monitor for you." },
  French: { "confirm":"Confirmer", "skip":"Passer", "add":"+ Ajouter", "customValue":"Saisir une valeur personnalisée…", "somethingElse":"Autre chose ? Saisissez-le ici…", "max":"max", "selected":"sélectionné(s)", "limitReached":"limite atteinte", "prioritiesHdr":"Vos priorités — le n°1 est notre point de départ", "confirmPriorities":"Confirmer les priorités", "objDetailsPh":"Autre chose au sujet de vos objectifs ? (facultatif)", "firstName":"Prénom", "lastName":"Nom", "roleDept":"Rôle / service", "email":"E-mail", "invalidEmail":"E-mail invalide", "addUser":"+ Ajouter un utilisateur", "confirmUsers":"Confirmer les utilisateurs", "topicName":"Nom du sujet", "keywordsPh":"Mots-clés…", "dragPrioritize":"Glissez pour classer par priorité", "kept":"conservés", "discarded":"écartés", "pending":"en attente", "submitQueries":"Envoyer les requêtes", "noQueries":"Aucune requête", "importFile":"📎 Ou importer un fichier (.txt, .csv, .xlsx)", "pasteQueries":"Collez vos requêtes existantes ici…", "hintSelectAll":"Sélectionnez toutes les options applicables.", "hintTeams":"Sélectionnez toutes les équipes qui utiliseront Lumen.", "hintObjectives":"Choisissez-en jusqu'à 3, puis définissez leur priorité : votre n°1 détermine ce que nous configurons en premier.", "hintTimezone":"Sélectionnez votre fuseau horaire principal.", "phMarket":"Saisir un marché…", "phLanguage":"Saisir une langue…", "phTeam":"Saisir une équipe…", "whyMarkets":"Pour que les résultats soient limités aux régions où vous opérez réellement.", "whyTeams":"Nous aide à adapter les tableaux de bord aux personnes qui les utiliseront.", "whyUsers":"Qui doit avoir accès — vous seul pour l'instant, c'est parfait.", "whyQueries":"Si vous suivez déjà des requêtes ailleurs, nous pouvons les migrer.", "whyTopics":"Les sujets sont les thèmes que Lumen surveillera pour vous." },
  German: { "confirm":"Bestätigen", "skip":"Überspringen", "add":"+ Hinzufügen", "customValue":"Eigenen Wert eingeben…", "somethingElse":"Etwas anderes? Hier eingeben…", "max":"max.", "selected":"ausgewählt", "limitReached":"Limit erreicht", "prioritiesHdr":"Ihre Prioritäten — Nr. 1 ist unser Ausgangspunkt", "confirmPriorities":"Prioritäten bestätigen", "objDetailsPh":"Sonst noch etwas zu Ihren Zielen? (optional)", "firstName":"Vorname", "lastName":"Nachname", "roleDept":"Rolle / Abteilung", "email":"E-Mail", "invalidEmail":"Ungültige E-Mail", "addUser":"+ Benutzer hinzufügen", "confirmUsers":"Benutzer bestätigen", "topicName":"Themenname", "keywordsPh":"Schlüsselwörter…", "dragPrioritize":"Zum Priorisieren ziehen", "kept":"behalten", "discarded":"verworfen", "pending":"offen", "submitQueries":"Abfragen senden", "noQueries":"Keine Abfragen", "importFile":"📎 Oder eine Datei importieren (.txt, .csv, .xlsx)", "pasteQueries":"Fügen Sie hier Ihre bestehenden Abfragen ein…", "hintSelectAll":"Wählen Sie alles Zutreffende aus.", "hintTeams":"Wählen Sie alle Teams aus, die Lumen nutzen werden.", "hintObjectives":"Wählen Sie bis zu 3 aus und legen Sie die Priorität fest — Ihre Nr. 1 bestimmt, was wir zuerst einrichten.", "hintTimezone":"Wählen Sie Ihre primäre Zeitzone.", "phMarket":"Markt eingeben…", "phLanguage":"Sprache eingeben…", "phTeam":"Team eingeben…", "whyMarkets":"Damit die Ergebnisse auf die Regionen beschränkt sind, in denen Sie tatsächlich tätig sind.", "whyTeams":"Hilft uns, die Dashboards auf die Personen zuzuschneiden, die sie nutzen.", "whyUsers":"Wer Zugriff haben soll — vorerst reicht es völlig, wenn nur Sie Zugriff haben.", "whyQueries":"Wenn Sie Abfragen bereits anderswo verfolgen, können wir sie migrieren.", "whyTopics":"Themen sind die Bereiche, die Lumen für Sie überwacht." },
  Spanish: { "confirm":"Confirmar", "skip":"Omitir", "add":"+ Añadir", "customValue":"Escriba un valor personalizado…", "somethingElse":"¿Algo más? Escríbalo aquí…", "max":"máx.", "selected":"seleccionado(s)", "limitReached":"límite alcanzado", "prioritiesHdr":"Sus prioridades: el n.º 1 es donde empezamos", "confirmPriorities":"Confirmar prioridades", "objDetailsPh":"¿Algo más sobre sus objetivos? (opcional)", "firstName":"Nombre", "lastName":"Apellidos", "roleDept":"Rol / departamento", "email":"Correo electrónico", "invalidEmail":"Correo no válido", "addUser":"+ Añadir usuario", "confirmUsers":"Confirmar usuarios", "topicName":"Nombre del tema", "keywordsPh":"Palabras clave…", "dragPrioritize":"Arrastre para priorizar", "kept":"conservados", "discarded":"descartados", "pending":"pendientes", "submitQueries":"Enviar consultas", "noQueries":"Sin consultas", "importFile":"📎 O importe un archivo (.txt, .csv, .xlsx)", "pasteQueries":"Pegue aquí sus consultas existentes…", "hintSelectAll":"Seleccione todo lo que corresponda.", "hintTeams":"Seleccione todos los equipos que usarán Lumen.", "hintObjectives":"Elija hasta 3 y ordene su prioridad: su n.º 1 decide qué configuramos primero.", "hintTimezone":"Seleccione su zona horaria principal.", "phMarket":"Escriba un mercado…", "phLanguage":"Escriba un idioma…", "phTeam":"Escriba un equipo…", "whyMarkets":"Para que los resultados se limiten a las regiones donde realmente opera.", "whyTeams":"Nos ayuda a adaptar los paneles a las personas que los usarán.", "whyUsers":"Quién debe tener acceso: por ahora, con usted basta.", "whyQueries":"Si ya sigue consultas en otro sitio, podemos migrarlas.", "whyTopics":"Los temas son los asuntos que Lumen monitorizará para usted." },
  Italian: { "confirm":"Conferma", "skip":"Salta", "add":"+ Aggiungi", "customValue":"Inserisci un valore personalizzato…", "somethingElse":"Qualcos'altro? Scrivilo qui…", "max":"max", "selected":"selezionato/i", "limitReached":"limite raggiunto", "prioritiesHdr":"Le tue priorità — la n.1 è il punto di partenza", "confirmPriorities":"Conferma priorità", "objDetailsPh":"Altro sui tuoi obiettivi? (facoltativo)", "firstName":"Nome", "lastName":"Cognome", "roleDept":"Ruolo / reparto", "email":"E-mail", "invalidEmail":"E-mail non valida", "addUser":"+ Aggiungi utente", "confirmUsers":"Conferma utenti", "topicName":"Nome dell'argomento", "keywordsPh":"Parole chiave…", "dragPrioritize":"Trascina per dare priorità", "kept":"mantenuti", "discarded":"scartati", "pending":"in sospeso", "submitQueries":"Invia query", "noQueries":"Nessuna query", "importFile":"📎 Oppure importa un file (.txt, .csv, .xlsx)", "pasteQueries":"Incolla qui le tue query esistenti…", "hintSelectAll":"Seleziona tutte le opzioni pertinenti.", "hintTeams":"Seleziona tutti i team che useranno Lumen.", "hintObjectives":"Scegline fino a 3, poi imposta la priorità: la n.1 decide cosa configuriamo per primo.", "hintTimezone":"Seleziona il tuo fuso orario principale.", "phMarket":"Inserisci un mercato…", "phLanguage":"Inserisci una lingua…", "phTeam":"Inserisci un team…", "whyMarkets":"Così i risultati sono limitati alle aree in cui operi davvero.", "whyTeams":"Ci aiuta ad adattare le dashboard alle persone che le useranno.", "whyUsers":"Chi deve avere accesso — per ora solo tu va benissimo.", "whyQueries":"Se monitori già delle query altrove, possiamo migrarle.", "whyTopics":"Gli argomenti sono i temi che Lumen monitorerà per te." },
  Arabic: { "confirm":"تأكيد", "skip":"تخطّي", "add":"+ إضافة", "customValue":"أدخل قيمة مخصّصة…", "somethingElse":"شيء آخر؟ اكتبه هنا…", "max":"حد أقصى", "selected":"محدد", "limitReached":"تم بلوغ الحد", "prioritiesHdr":"أولوياتك — رقم 1 هو نقطة البداية", "confirmPriorities":"تأكيد الأولويات", "objDetailsPh":"أي شيء آخر بخصوص أهدافك؟ (اختياري)", "firstName":"الاسم الأول", "lastName":"اسم العائلة", "roleDept":"الدور / القسم", "email":"البريد الإلكتروني", "invalidEmail":"بريد إلكتروني غير صالح", "addUser":"+ إضافة مستخدم", "confirmUsers":"تأكيد المستخدمين", "topicName":"اسم الموضوع", "keywordsPh":"الكلمات المفتاحية…", "dragPrioritize":"اسحب لترتيب الأولوية", "kept":"محتفظ بها", "discarded":"مستبعدة", "pending":"قيد الانتظار", "submitQueries":"إرسال الاستعلامات", "noQueries":"لا توجد استعلامات", "importFile":"📎 أو استورد ملفًا (.txt، .csv، .xlsx)", "pasteQueries":"الصق استعلاماتك الحالية هنا…", "hintSelectAll":"اختر كل ما ينطبق.", "hintTeams":"اختر جميع الفرق التي ستستخدم Lumen.", "hintObjectives":"اختر ما يصل إلى 3، ثم رتّب أولوياتها — رقم 1 يحدد ما نُعدّه أولًا.", "hintTimezone":"اختر منطقتك الزمنية الأساسية.", "phMarket":"أدخل سوقًا…", "phLanguage":"أدخل لغة…", "phTeam":"أدخل فريقًا…", "whyMarkets":"لكي تقتصر النتائج على المناطق التي تعمل فيها فعليًا.", "whyTeams":"يساعدنا على تخصيص لوحات المعلومات للأشخاص الذين سيستخدمونها.", "whyUsers":"من ينبغي أن يملك حق الوصول — الاكتفاء بك وحدك الآن أمر جيد.", "whyQueries":"إذا كنت تتابع استعلامات في مكان آخر، يمكننا نقلها.", "whyTopics":"المواضيع هي ما سيراقبه Lumen نيابةً عنك." },
};
function WL(key, lang) {
  const dict = WI18N[lang] || WI18N.English;
  return (dict[key] != null ? dict[key] : WI18N.English[key]) || "";
}

// QUERIES-widget file-import feedback (shown in the expert flow). Parametrized:
// {name} filename, {n} line cap, {mb} size. QN() substitutes and falls back to English.
const QN18N = {
  English: { "importedTruncated":"Imported the first {n} lines of {name}. Hit Submit and I'll pick out what's relevant — with a file this size, double-check the queries you care about most made it in.", "imported":"Imported {name}. Hit Submit and I'll pick out what's relevant — no need to tidy it up.", "noText":"Couldn't find any text in {name}.", "tooLarge":"That file is {mb} MB — too large to read here. Export just the queries (or paste them directly) and try again.", "unsupported":"That file type isn't supported — use .txt, .csv or .xlsx, or paste the queries directly.", "readError":"Couldn't read that file — try pasting the queries directly instead." },
  French: { "importedTruncated":"Les {n} premières lignes de {name} ont été importées. Cliquez sur Envoyer et je repérerai ce qui est pertinent — avec un fichier de cette taille, vérifiez que les requêtes les plus importantes y figurent.", "imported":"{name} importé. Cliquez sur Envoyer et je repérerai ce qui est pertinent — inutile de faire le tri.", "noText":"Aucun texte trouvé dans {name}.", "tooLarge":"Ce fichier fait {mb} Mo — trop volumineux pour être lu ici. Exportez uniquement les requêtes (ou collez-les directement) et réessayez.", "unsupported":"Ce type de fichier n'est pas pris en charge — utilisez .txt, .csv ou .xlsx, ou collez les requêtes directement.", "readError":"Impossible de lire ce fichier — essayez plutôt de coller les requêtes directement." },
  German: { "importedTruncated":"Die ersten {n} Zeilen von {name} wurden importiert. Klicken Sie auf Senden und ich filtere das Relevante heraus — prüfen Sie bei einer Datei dieser Größe, ob die wichtigsten Abfragen enthalten sind.", "imported":"{name} importiert. Klicken Sie auf Senden und ich filtere das Relevante heraus — Aufräumen ist nicht nötig.", "noText":"In {name} wurde kein Text gefunden.", "tooLarge":"Diese Datei ist {mb} MB groß — zu groß, um sie hier zu lesen. Exportieren Sie nur die Abfragen (oder fügen Sie sie direkt ein) und versuchen Sie es erneut.", "unsupported":"Dieser Dateityp wird nicht unterstützt — verwenden Sie .txt, .csv oder .xlsx, oder fügen Sie die Abfragen direkt ein.", "readError":"Diese Datei konnte nicht gelesen werden — fügen Sie die Abfragen stattdessen direkt ein." },
  Spanish: { "importedTruncated":"Se importaron las primeras {n} líneas de {name}. Pulse Enviar y seleccionaré lo relevante — con un archivo de este tamaño, compruebe que se incluyeron las consultas que más le importan.", "imported":"{name} importado. Pulse Enviar y seleccionaré lo relevante — no hace falta ordenarlo.", "noText":"No se encontró texto en {name}.", "tooLarge":"Este archivo ocupa {mb} MB — demasiado grande para leerlo aquí. Exporte solo las consultas (o péguelas directamente) e inténtelo de nuevo.", "unsupported":"Ese tipo de archivo no es compatible — use .txt, .csv o .xlsx, o pegue las consultas directamente.", "readError":"No se pudo leer ese archivo — pruebe a pegar las consultas directamente." },
  Italian: { "importedTruncated":"Importate le prime {n} righe di {name}. Premi Invia e selezionerò ciò che è pertinente — con un file di queste dimensioni, verifica che le query più importanti siano incluse.", "imported":"{name} importato. Premi Invia e selezionerò ciò che è pertinente — non serve riordinare.", "noText":"Nessun testo trovato in {name}.", "tooLarge":"Questo file è di {mb} MB — troppo grande da leggere qui. Esporta solo le query (o incollale direttamente) e riprova.", "unsupported":"Questo tipo di file non è supportato — usa .txt, .csv o .xlsx, oppure incolla le query direttamente.", "readError":"Impossibile leggere il file — prova a incollare le query direttamente." },
  Arabic: { "importedTruncated":"تم استيراد أول {n} سطرًا من {name}. اضغط إرسال وسأختار ما هو مهم — مع ملف بهذا الحجم، تأكّد من أن أهم الاستعلامات قد أُدرجت.", "imported":"تم استيراد {name}. اضغط إرسال وسأختار ما هو مهم — لا حاجة للترتيب.", "noText":"لم يُعثر على نص في {name}.", "tooLarge":"حجم هذا الملف {mb} ميغابايت — أكبر من أن يُقرأ هنا. صدّر الاستعلامات فقط (أو الصقها مباشرة) وحاول مرة أخرى.", "unsupported":"نوع الملف غير مدعوم — استخدم ‎.txt أو ‎.csv أو ‎.xlsx، أو الصق الاستعلامات مباشرة.", "readError":"تعذّرت قراءة الملف — جرّب لصق الاستعلامات مباشرة بدلاً من ذلك." },
};
function QN(key, lang, vars) {
  const dict = QN18N[lang] || QN18N.English;
  let s = (dict[key] != null ? dict[key] : QN18N.English[key]) || "";
  if (vars) for (const k in vars) s = s.split("{"+k+"}").join(vars[k]);
  return s;
}


const gts   = () => new Date().toLocaleTimeString([], { hour:"2-digit", minute:"2-digit" });
const sleep = ms => new Promise(r => setTimeout(r, ms));

// In-progress autosave to localStorage for same-device pause/resume. Keyed by the
// seed id when present, else a single default slot. Best-effort: any failure
// (private mode, quota, storage disabled) degrades to no-resume without throwing.
const LS_PREFIX = "lumen_onb_v1_";
const lsKey = seedId => LS_PREFIX + (seedId || "default");
function lsLoadDraft(seedId) {
  try {
    const raw = localStorage.getItem(lsKey(seedId));
    if (!raw) return null;
    const o = JSON.parse(raw);
    return (o && Array.isArray(o.messages) && o.messages.length && o.progress && (o.progress.percent || 0) < 100) ? o : null;
  } catch { return null; }
}
function lsSaveDraft(seedId, snap) { try { localStorage.setItem(lsKey(seedId), JSON.stringify(snap)); return true; } catch { return false; } }
// Is localStorage actually writable? (Private mode / quota / disabled storage
// all throw.) Used so the "Saved on this device" badge isn't shown when saving
// silently fails.
function lsProbe() { try { const k = "__lumen_probe__"; localStorage.setItem(k, "1"); localStorage.removeItem(k); return true; } catch { return false; } }
function lsClearDraft(seedId) { try { localStorage.removeItem(lsKey(seedId)); } catch {} }

// Pure fold of one parsed reply's markers onto a cdata object. Used live and to
// rebuild cdata from surviving messages after a rewind.
// Arrays REPLACE wholesale: the system prompt re-emits the FULL array each time,
// so a new non-empty array is the complete current set. (An earlier attempt to
// union by name silently collapsed distinct entries that share a key — e.g. one
// brand's Instagram/X/TikTok channels all keyed on author "Nike" — so it was
// reverted. The rare partial re-emit is recoverable in the editable review modal.)
function mergeCdata(base, pr) {
  const { companyData,topicsData,channelsData,reportsData,alertsData,handoffData } = pr;
  if (!(companyData||topicsData||channelsData||reportsData||alertsData||handoffData)) return base;
  // Wipe guards: a stray re-emit of an EMPTY array (model slip) must not erase
  // data already captured — arrays replace wholesale only when the new one has
  // items. Objects (company/handoff) merge field-by-field with blanks dropped,
  // so a re-emit that forgot a field can't null out a value we already have.
  // The editable review modal remains the human backstop for true removals.
  const keepArr = (next, prev) => (Array.isArray(next) && next.length) ? next : prev;
  const mergeObj = (next, prev) => next ? { ...prev, ...pruneEmpty(next) } : prev;
  return {...base,
    company: mergeObj(companyData, base.company),
    topics: keepArr(topicsData, base.topics),
    channels: keepArr(channelsData, base.channels),
    reports: keepArr(reportsData, base.reports),
    alerts: keepArr(alertsData, base.alerts),
    handoff: mergeObj(handoffData, base.handoff)};
}
// Drop null/undefined/blank values so a marker re-emit with empty fields never
// overwrites a previously captured value.
const pruneEmpty = o => {
  const r = {};
  for (const k in o) { const v = o[k]; if (v != null && v !== "") r[k] = v; }
  return r;
};
const emptyCdata = () => ({company:{},topics:[],channels:[],reports:[],alerts:[]});
function pProg(t) { const m = t.match(/%%PROGRESS%%([\s\S]*?)%%END%%/); try { return m ? JSON.parse(m[1]) : null; } catch { return null; } }
function pMark(t, k) { const m = t.match(new RegExp("%%"+k+"%%(\\[?[\\s\\S]*?\\]?)%%END%%")); try { return m ? JSON.parse(m[1]) : null; } catch { return null; } }
// Neutralize marker delimiters in CLIENT-authored text before it reaches the
// model. Markers are `%%NAME%%...%%END%%`; if a client types "%%END%%" (or the
// model echoes a client value containing it back into a marker), parsing would
// truncate and silently drop that field. Collapsing runs of %% to a single %
// keeps ordinary "50%" intact while removing any delimiter a client could inject.
const sanitizeIn = s => String(s == null ? "" : s).replace(/%%+/g, "%");
// The prompt requires a hidden <thought> block on every reply. It is only needed
// for that turn's planning; re-sending past thoughts back in the history wastes
// input tokens on every call. Strip them before an assistant turn re-enters history
// (markers/widgets stay, so the model keeps the structured context it needs).
const stripThoughtForHistory = t => String(t == null ? "" : t)
  .replace(/<(thought|thoughts|thinking|think)>[\s\S]*?<\/(thought|thoughts|thinking|think)>/g, "")
  .replace(/<(thought|thoughts|thinking|think)>[\s\S]*$/, "")
  .trim();
function stripAll(t) {
  let s = t
    .replace(/%%[A-Z]+%%[\s\S]*?%%END%%/g, "")
    .replace(/\[WIDGET:[A-Z_]+\]/g, "")
    .replace(/<(thought|thoughts|thinking|think)>[\s\S]*?<\/(thought|thoughts|thinking|think)>/g, "")
    .replace(/^\s*<(thought|thoughts|thinking|think)>[\s\S]*$/, "")
    .replace(/\[SUGGESTIONS:[\s\S]*?\]/g, "")
    .replace(/^TOPIC_SUGGESTION\|.*$/gm, "");
  // Safety net: a reply truncated mid-marker leaves an opening %%MARKER%% with no
  // closing %%END%%, which would otherwise render as raw JSON. Markers are always
  // emitted at the start of a reply, so anything from a leftover opener onward is
  // truncated junk — cut it. Same for a half-written [WIDGET:/[SUGGESTIONS: tag.
  const mk = s.search(/%%[A-Z]+%%/);
  if (mk !== -1) s = s.slice(0, mk);
  const tag = s.search(/\[(WIDGET|SUGGESTIONS):[^\]]*$/);
  if (tag !== -1) s = s.slice(0, tag);
  const th = s.search(/<(thought|thoughts|thinking|think)>/);
  if (th !== -1) s = s.slice(0, th);
  return s.trim();
}
// True when a reply contains an opening %%MARKER%% with no matching %%END%% — the
// signature of a response that was cut off mid-emit.
function hasDanglingMarker(t) {
  return /%%[A-Z]+%%/.test(t.replace(/%%[A-Z]+%%[\s\S]*?%%END%%/g, ""));
}
// True when the visible prose implies the setup is already live/running/delivering,
// which it never is until the consultant activates it at the review call. High-
// precision phrase list — the prompt rule handles the long tail; this catches the
// specific overstatements testers flagged and triggers a corrective rewrite.
function overstatesCompletion(t) {
  // Strong present-completion claims (the "now" cases testers flagged) + always-wrong phrases.
  if (/\b((is|are|you'?re|you are) now (set up|live|active|running|configured|enabled|getting|receiving|monitoring|tracking|all set)|will now (get|receive|start|begin)|now getting proactive|delivered on a schedule|you'?re all set|is now live)\b/i.test(t)) return true;
  // Bare "is live/active/running" claims — but NOT when framed conditionally ("once this is live, you'll…"), which is correct.
  if (/\b(this|it|your setup|the setup|everything) is (live|active|running)\b/i.test(t) && !/\b(once|when|after|as soon as|until)\b/i.test(t)) return true;
  return false;
}
function exWid(t) { return [...new Set((t.match(/\[WIDGET:[A-Z_]+\]/g)||[]).map(x => x.replace(/\[WIDGET:|\]/g, "")))]; }
function procTopics(t) {
  const s = [], k = [];
  for (const l of t.split("\n")) {
    if (l.startsWith("TOPIC_SUGGESTION|")) {
      const p = l.split("|");
      if (p.length >= 4) s.push({ name:p[1].trim(), keywords:p[2].trim(), rationale:p[3].trim() });
    } else k.push(l);
  }
  return { suggestions:s, stripped:k.join("\n").trim() };
}
function safeVisible(t) {
  // Progressive display during streaming: hide complete special blocks and
  // truncate at any incomplete marker/widget/thought still being generated.
  let v = stripAll(t);
  const cuts = ["%%", "<thought", "<think", "[WIDGET:", "[SUGGESTIONS:", "TOPIC_SUGGESTION|"];
  let idx = -1;
  for (const c of cuts) { const i = v.indexOf(c); if (i !== -1 && (idx === -1 || i < idx)) idx = i; }
  if (idx !== -1) v = v.slice(0, idx);
  return v.replace(/^\s+/, "");
}
// Chips are pipe-separated per the SP, but models sometimes emit commas anyway.
// Fall back to comma-splitting only at 3+ segments so legitimate single chips
// containing one comma ("Yes, looks good") stay intact.
function splitChips(s) {
  if (s.includes("|")) return s.split("|").map(x=>x.trim()).filter(Boolean);
  const byComma = s.split(",").map(x=>x.trim()).filter(Boolean);
  return byComma.length >= 3 ? byComma : [s.trim()].filter(Boolean);
}
function parseReply(r) {
  const progress = pProg(r), widgets = exWid(r);
  const sm = r.match(/\[SUGGESTIONS:\s*(.+?)\]/);
  const quickReplies = sm && widgets.length === 0 ? splitChips(sm[1]) : [];
  const { suggestions:topicSuggestions, stripped } = procTopics(r);
  const clean = stripAll(stripped);
  if (topicSuggestions.length > 0 && !widgets.includes("TOPICS")) widgets.push("TOPICS");
  return { clean, widgets, topicSuggestions, quickReplies, progress,
    companyData:pMark(r,"COMPANY"), topicsData:pMark(r,"TOPICS"),
    channelsData:pMark(r,"CHANNELS"), reportsData:pMark(r,"REPORTS"), alertsData:pMark(r,"ALERTS"), handoffData:pMark(r,"HANDOFF"), raw:r };
}
function renderText(text) {
  const parts = [], rx = /(\*\*(.+?)\*\*|\[([^\]]+)\]\((https?:\/\/[^\)]+)\))/g;
  let last = 0, m, k = 0;
  while ((m = rx.exec(text)) !== null) {
    if (m.index > last) parts.push(text.slice(last, m.index));
    if (m[0].startsWith("**")) parts.push(<strong key={k++}>{m[2]}</strong>);
    else parts.push(<a key={k++} href={m[4]} target="_blank" rel="noopener noreferrer" style={{color:LINK,textDecoration:"underline"}}>{m[3]}</a>);
    last = m.index + m[0].length;
  }
  if (last < text.length) parts.push(text.slice(last));
  return parts;
}
const MsgText = memo(({ text }) => (
  <span>{text.split("\n").map((l,i,a) => <span key={i}>{renderText(l)}{i < a.length-1 ? <br/> : null}</span>)}</span>
));

function useAudio() {
  const r = useRef(null);
  const init  = useCallback(() => { if (!r.current) r.current = new (window.AudioContext || window.webkitAudioContext)(); }, []);
  const pop   = useCallback(() => { const c=r.current; if (!c||c.state==="suspended") return; const o=c.createOscillator(),g=c.createGain(); o.connect(g); g.connect(c.destination); o.type="sine"; o.frequency.setValueAtTime(450,c.currentTime); o.frequency.exponentialRampToValueAtTime(700,c.currentTime+0.1); g.gain.setValueAtTime(0,c.currentTime); g.gain.linearRampToValueAtTime(0.08,c.currentTime+0.02); g.gain.exponentialRampToValueAtTime(0.001,c.currentTime+0.15); o.start(c.currentTime); o.stop(c.currentTime+0.15); }, []);
  const chime = useCallback(() => { const c=r.current; if (!c||c.state==="suspended") return; [523.25,659.25,783.99,1046.5].forEach((f,i)=>{ const o=c.createOscillator(),g=c.createGain(); o.connect(g); g.connect(c.destination); o.type="sine"; o.frequency.value=f; const t=c.currentTime+i*0.08; g.gain.setValueAtTime(0,t); g.gain.linearRampToValueAtTime(0.05,t+0.02); g.gain.exponentialRampToValueAtTime(0.001,t+0.6); o.start(t); o.stop(t+0.6); }); }, []);
  return { init, pop, chime };
}

function LumenIcon({ size=28, inverse=false }) {
  const sw = size>=48?2.8:2.2, cr = size>=48?5:3.5;
  const tile = inverse?"#ffffff":A, mark = inverse?A:"#ffffff";
  return <svg width={size} height={size} viewBox="0 0 32 32" fill="none"><rect width="32" height="32" rx="8" fill={tile}/><path d="M16 6L16 26M6 16L26 16M9 9L23 23M23 9L9 23" stroke={mark} strokeWidth={sw} strokeLinecap="round"/><circle cx="16" cy="16" r={cr} fill={mark}/></svg>;
}
function Spinner() {
  return <svg width="18" height="18" viewBox="0 0 18 18" fill="none" style={{animation:"spin 0.8s linear infinite"}}><circle cx="9" cy="9" r="7" stroke="rgba(255,255,255,0.3)" strokeWidth="2"/><path d="M9 2a7 7 0 0 1 7 7" stroke="white" strokeWidth="2" strokeLinecap="round"/></svg>;
}
function Ic({ d, size=15 }) {
  return <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.9" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true"><path d={d}/></svg>;
}
const IC = {
  panel:  "M4 5h16M4 12h16M4 19h10",
  sound:  "M11 5 6 9H3v6h3l5 4V5Z M15.5 8.5a5 5 0 0 1 0 7 M18.5 5.5a9 9 0 0 1 0 13",
  mute:   "M11 5 6 9H3v6h3l5 4V5Z M22 9l-6 6 M16 9l6 6",
  moon:   "M21 12.8A9 9 0 1 1 11.2 3 7 7 0 0 0 21 12.8Z",
  sun:    "M12 8a4 4 0 1 0 0 8 4 4 0 0 0 0-8Z M12 2v2 M12 20v2 M4.9 4.9l1.4 1.4 M17.7 17.7l1.4 1.4 M2 12h2 M20 12h2 M4.9 19.1l1.4-1.4 M17.7 6.3l1.4-1.4",
};
function TypingIndicator({ lang }) {
  const [v,setV] = useState(false);
  useEffect(() => { const t = setTimeout(() => setV(true), 300); return () => clearTimeout(t); }, []);
  return <div style={{display:"flex",alignItems:"center",gap:12,minHeight:28}}>{v && <><div style={{display:"flex",gap:4}}>{[0,1,2].map(d => <div key={d} style={{width:6,height:6,borderRadius:"50%",background:P,animation:"bounce 1.4s infinite ease-in-out both",animationDelay:`${d*0.16}s`}}/>)}</div><span style={{fontSize:13,color:"#64748b"}}>{L("thinking", lang)}</span></>}</div>;
}
function Stepper({ progress, dark, compact }) {
  const inactive = dark?"#2d4a6a":"#E7E7EF", muted = dark?"#8aa4c1":"#64748b", F = A, circleBg = dark?"#111f30":"#ffffff";
  return <div style={{display:"flex",alignItems:"flex-start",width:"100%"}}>{SECTION_KEYS.map((key,i) => {
    const done = !!progress.collected?.[key], cur = progress.section === key;
    return <div key={key} style={{flex:1,display:"flex",flexDirection:"column",alignItems:"center",position:"relative"}}>
      {i < SECTION_KEYS.length-1 && <div style={{position:"absolute",top:11,left:"50%",width:"100%",height:2,background:done?F:inactive,zIndex:0,transition:"background 0.4s"}}/>}
      <div style={{width:22,height:22,borderRadius:"50%",border:`2px solid ${done||cur?F:inactive}`,background:done?F:cur?circleBg:circleBg,display:"flex",alignItems:"center",justifyContent:"center",fontSize:10,fontWeight:700,color:done?"white":cur?F:muted,zIndex:1,transition:"all 0.3s"}}>{done?"✓":i+1}</div>
      {(!compact || cur) && <div style={{fontSize:10,marginTop:4,color:done||cur?P:muted,fontWeight:done||cur?600:400,whiteSpace:"nowrap"}}>{SECTION_LABELS[key]}</div>}
    </div>;
  })}</div>;
}

function ChipSelector({ options, max=99, onSubmit, onSkip, placeholder, hint, initialData=[], lang }) {
  const [sel,setSel] = useState(initialData);
  const [custom,setCustom] = useState("");
  const atLim = sel.length >= max;
  const toggle = o => { if (sel.includes(o)) setSel(s=>s.filter(x=>x!==o)); else if (!atLim) setSel(s=>[...s,o]); };
  const addC = () => { const v=custom.trim(); if (v&&!sel.includes(v)&&!atLim) { setSel(s=>[...s,v]); setCustom(""); } };
  return <div style={{marginTop:8}}>
    {hint && <div style={{fontSize:12,color:"#64748b",marginBottom:10}}>{hint}{max<99&&<span style={{marginLeft:6,background:"#ede9fe",color:P,borderRadius:6,padding:"1px 7px",fontSize:11,fontWeight:600}}>{WL("max",lang)} {max}</span>}</div>}
    <div role="group" style={{display:"flex",flexWrap:"wrap",gap:6,marginBottom:10}}>{options.map(o => <button key={o} onClick={()=>toggle(o)} disabled={atLim&&!sel.includes(o)} aria-pressed={sel.includes(o)} style={{padding:"9px 14px",minHeight:38,borderRadius:20,fontSize:12,cursor:atLim&&!sel.includes(o)?"default":"pointer",border:"1px solid",background:sel.includes(o)?P:"transparent",borderColor:sel.includes(o)?P:"#e2e8f0",color:sel.includes(o)?"white":atLim&&!sel.includes(o)?"#cbd5e1":"#64748b",transition:"all 0.15s"}}>{o}</button>)}</div>
    <div style={{display:"flex",gap:6,marginBottom:10}}>
      <input value={custom} onChange={e=>setCustom(e.target.value)} onKeyDown={e=>e.key==="Enter"&&addC()} placeholder={placeholder||WL("customValue",lang)} style={{flex:1,background:"white",border:"1px solid #c4b5fd",borderRadius:8,padding:"7px 11px",fontSize:12,color:"#1e293b",outline:"none"}}/>
      <button onClick={addC} disabled={!custom.trim()||atLim} style={{background:custom.trim()&&!atLim?P:"#e2e8f0",color:"white",border:"none",borderRadius:8,padding:"7px 14px",cursor:custom.trim()&&!atLim?"pointer":"default",fontSize:12,fontWeight:600}}>{WL("add",lang)}</button>
    </div>
    {max<99 && <div style={{fontSize:11,color:atLim?"#dc2626":"#94a3b8",marginBottom:10}}>{sel.length}/{max} {WL("selected",lang)}{atLim?" — "+WL("limitReached",lang):""}</div>}
    <div style={{display:"flex",gap:8}}>
      <button onClick={()=>sel.length>0&&onSubmit(sel)} disabled={sel.length===0} style={{background:sel.length>0?P:"#e2e8f0",color:"white",border:"none",borderRadius:8,padding:"8px 20px",fontSize:13,fontWeight:600,cursor:sel.length>0?"pointer":"default"}}>{WL("confirm",lang)}</button>
      {onSkip && <button onClick={onSkip} style={{background:"transparent",border:"1px solid #e2e8f0",borderRadius:8,padding:"8px 16px",fontSize:13,color:"#64748b",cursor:"pointer"}}>{WL("skip",lang)}</button>}
    </div>
  </div>;
}

function RankedSelector({ options, max=3, onSubmit, onSkip, hint, initialData, lang }) {
  const init = normObjectives(initialData);
  const [sel,setSel]       = useState(init.ranked);
  const [details,setDetails]= useState(init.details);
  const [custom,setCustom] = useState("");
  const atLim = sel.length >= max;
  const toggle = o => { if (sel.includes(o)) setSel(s=>s.filter(x=>x!==o)); else if (!atLim) setSel(s=>[...s,o]); };
  const move   = (i,dir) => setSel(s => { const n=[...s], j=i+dir; if (j<0||j>=n.length) return s; [n[i],n[j]]=[n[j],n[i]]; return n; });
  const addC   = () => { const v=custom.trim(); if (v&&!sel.includes(v)&&!atLim) { setSel(s=>[...s,v]); setCustom(""); } };
  return <div style={{marginTop:8}}>
    {hint && <div style={{fontSize:12,color:"#64748b",marginBottom:10}}>{hint}<span style={{marginLeft:6,background:"#ede9fe",color:P,borderRadius:6,padding:"1px 7px",fontSize:11,fontWeight:600}}>{WL("max",lang)} {max}</span></div>}
    <div style={{display:"flex",flexWrap:"wrap",gap:6,marginBottom:10}}>{options.map(o => <button key={o} onClick={()=>toggle(o)} disabled={atLim&&!sel.includes(o)} style={{padding:"5px 12px",borderRadius:20,fontSize:12,cursor:atLim&&!sel.includes(o)?"default":"pointer",border:"1px solid",background:sel.includes(o)?P:"transparent",borderColor:sel.includes(o)?P:"#e2e8f0",color:sel.includes(o)?"white":atLim&&!sel.includes(o)?"#cbd5e1":"#64748b",transition:"all 0.15s"}}>{o}</button>)}</div>
    <div style={{display:"flex",gap:6,marginBottom:10}}>
      <input value={custom} onChange={e=>setCustom(e.target.value)} onKeyDown={e=>e.key==="Enter"&&addC()} placeholder={WL("somethingElse",lang)} style={{flex:1,background:"white",border:"1px solid #c4b5fd",borderRadius:8,padding:"7px 11px",fontSize:12,color:"#1e293b",outline:"none"}}/>
      <button onClick={addC} disabled={!custom.trim()||atLim} style={{background:custom.trim()&&!atLim?P:"#e2e8f0",color:"white",border:"none",borderRadius:8,padding:"7px 14px",cursor:custom.trim()&&!atLim?"pointer":"default",fontSize:12,fontWeight:600}}>{WL("add",lang)}</button>
    </div>
    {sel.length>0 && <div style={{background:"#f8f9fa",border:"1px solid #e2e8f0",borderRadius:10,padding:"10px 12px",marginBottom:10}}>
      <div style={{fontSize:11,fontWeight:700,color:"#64748b",marginBottom:8,textTransform:"uppercase",letterSpacing:"0.04em"}}>{WL("prioritiesHdr",lang)}</div>
      {sel.map((o,i) => <div key={o} style={{display:"flex",alignItems:"center",gap:8,padding:"6px 0",borderTop:i>0?"1px solid #eef1f5":"none"}}>
        <span style={{width:22,height:22,borderRadius:"50%",background:i===0?A:P,color:"white",fontSize:11,fontWeight:700,display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0}}>{i+1}</span>
        <span style={{flex:1,fontSize:13,color:"#1e293b",fontWeight:i===0?600:400}}>{o}</span>
        <button onClick={()=>move(i,-1)} disabled={i===0} aria-label={`Move ${o} up`} style={{background:"transparent",border:"1px solid #e2e8f0",borderRadius:6,width:24,height:24,cursor:i===0?"default":"pointer",color:i===0?"#cbd5e1":"#64748b",fontSize:11,lineHeight:1}}>▲</button>
        <button onClick={()=>move(i,1)} disabled={i===sel.length-1} aria-label={`Move ${o} down`} style={{background:"transparent",border:"1px solid #e2e8f0",borderRadius:6,width:24,height:24,cursor:i===sel.length-1?"default":"pointer",color:i===sel.length-1?"#cbd5e1":"#64748b",fontSize:11,lineHeight:1}}>▼</button>
        <button onClick={()=>toggle(o)} aria-label={`Remove ${o}`} style={{background:"transparent",border:"none",color:"#ef4444",cursor:"pointer",fontSize:12}}>✕</button>
      </div>)}
    </div>}
    <textarea value={details} onChange={e=>setDetails(e.target.value)} rows={2} placeholder={WL("objDetailsPh",lang)} style={{width:"100%",background:"white",border:"1px solid #e2e8f0",borderRadius:8,padding:"7px 11px",fontSize:12,color:"#1e293b",outline:"none",resize:"vertical",boxSizing:"border-box",marginBottom:10}}/>
    <div style={{fontSize:11,color:atLim?"#dc2626":"#94a3b8",marginBottom:10}}>{sel.length}/{max} {WL("selected",lang)}{atLim?" — "+WL("limitReached",lang):""}</div>
    <div style={{display:"flex",gap:8}}>
      <button onClick={()=>sel.length>0&&onSubmit({ranked:sel,details:details.trim()})} disabled={sel.length===0} style={{background:sel.length>0?P:"#e2e8f0",color:"white",border:"none",borderRadius:8,padding:"8px 20px",fontSize:13,fontWeight:600,cursor:sel.length>0?"pointer":"default"}}>{WL("confirmPriorities",lang)}</button>
      {onSkip && <button onClick={onSkip} style={{background:"transparent",border:"1px solid #e2e8f0",borderRadius:8,padding:"8px 16px",fontSize:13,color:"#64748b",cursor:"pointer"}}>{WL("skip",lang)}</button>}
    </div>
  </div>;
}

function UserForm({ onSubmit, onSkip, initialData=[], lang }) {
  const empty = () => ({ firstName:"", lastName:"", email:"", role:"", access:"Full Tool" });
  const [users,setUsers] = useState(initialData.length>0?initialData:[empty()]);
  const [errors,setErrors] = useState({});
  const upd = (i,k,v) => setUsers(u=>u.map((x,j)=>j===i?{...x,[k]:v}:x));
  const vEmail = (i,v) => setErrors(e=>({...e,[`${i}-email`]:v&&!EMAIL_RE.test(v)?WL("invalidEmail",lang):""}));
  const valid = users.every(u=>u.firstName&&u.email&&EMAIL_RE.test(u.email));
  return <div style={{marginTop:8}}>
    {users.map((u,i) => <div key={i} style={{background:"#f8f9fa",border:"1px solid #e2e8f0",borderRadius:10,padding:"12px 14px",marginBottom:8}}>
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(140px,1fr))",gap:8,marginBottom:8}}>
        {[["firstName",WL("firstName",lang)],["lastName",WL("lastName",lang)],["role",WL("roleDept",lang)]].map(([k,ph]) => <input key={k} value={u[k]} onChange={e=>upd(i,k,e.target.value)} placeholder={ph} style={{background:"white",border:"1px solid #e2e8f0",borderRadius:7,padding:"7px 10px",fontSize:12,color:"#1e293b",outline:"none"}}/>)}
        <div>
          <input value={u.email} onChange={e=>upd(i,"email",e.target.value)} onBlur={e=>vEmail(i,e.target.value)} placeholder={WL("email",lang)} style={{background:"white",border:`1px solid ${errors[`${i}-email`]?"#ef4444":"#e2e8f0"}`,borderRadius:7,padding:"7px 10px",fontSize:12,color:"#1e293b",outline:"none",width:"100%"}}/>
          {errors[`${i}-email`] && <div style={{fontSize:10,color:"#ef4444",marginTop:3}}>{errors[`${i}-email`]}</div>}
        </div>
      </div>
      <div style={{display:"flex",gap:6}}>{["Admin","Full Tool","Read-Only"].map(a => <button key={a} onClick={()=>upd(i,"access",a)} style={{flex:1,padding:"6px 8px",borderRadius:7,fontSize:11,cursor:"pointer",border:"1px solid",background:u.access===a?P:"transparent",borderColor:u.access===a?P:"#e2e8f0",color:u.access===a?"white":"#64748b"}}>{a}</button>)}</div>
    </div>)}
    <div style={{display:"flex",gap:8,marginTop:4}}>
      <button onClick={()=>setUsers(u=>[...u,empty()])} style={{background:"transparent",border:"1px solid #e2e8f0",borderRadius:8,padding:"7px 14px",color:"#64748b",cursor:"pointer",fontSize:12}}>{WL("addUser",lang)}</button>
      <button onClick={()=>valid&&onSubmit(users)} disabled={!valid} style={{background:valid?P:"#e2e8f0",color:"white",border:"none",borderRadius:8,padding:"7px 20px",fontSize:13,fontWeight:600,cursor:valid?"pointer":"default"}}>{WL("confirmUsers",lang)}</button>
      {onSkip && <button onClick={onSkip} style={{background:"transparent",border:"1px solid #e2e8f0",borderRadius:8,padding:"7px 16px",fontSize:13,color:"#64748b",cursor:"pointer"}}>{WL("skip",lang)}</button>}
    </div>
  </div>;
}

function TopicCards({ suggestions, onConfirm, onSkip, lang }) {
  const [cards,setCards] = useState(suggestions.map(s=>({...s,status:"pending",id:Math.random().toString(36).substr(2,9)})));
  const [dragIdx,setDragIdx] = useState(null);
  const upd  = (i,f,v) => setCards(c=>c.map((x,j)=>j===i?{...x,[f]:v}:x));
  const setSt = (i,s) => setCards(c=>c.map((x,j)=>j===i?{...x,status:s}:x));
  // Native HTML5 drag does not fire on touch, so give a tap-friendly reorder too.
  const move = (i,dir) => setCards(c=>{ const n=[...c], j=i+dir; if (j<0||j>=n.length) return c; [n[i],n[j]]=[n[j],n[i]]; return n; });
  const isTouch = typeof window !== "undefined" && ("ontouchstart" in window || (navigator.maxTouchPoints||0) > 0);
  const kept = cards.filter(c=>c.status==="kept");
  return <div style={{marginTop:8}}>
    <div style={{fontSize:11,color:"#64748b",marginBottom:10,display:"flex",justifyContent:"space-between"}}>
      <span>{kept.length} {WL("kept",lang)} · {cards.filter(c=>c.status==="discarded").length} {WL("discarded",lang)} · {cards.filter(c=>c.status==="pending").length} {WL("pending",lang)}</span>
      {!isTouch && <span>☰ {WL("dragPrioritize",lang)}</span>}
    </div>
    {cards.map((c,i) => <div key={c.id} draggable
      onDragStart={e=>{setDragIdx(i);e.dataTransfer.effectAllowed="move";}}
      onDragOver={e=>e.preventDefault()}
      onDrop={e=>{e.preventDefault();if(dragIdx===null||dragIdx===i)return;const nc=[...cards];const[dc]=nc.splice(dragIdx,1);nc.splice(i,0,dc);setCards(nc);setDragIdx(null);}}
      style={{background:c.status==="kept"?"#f0fdf4":c.status==="discarded"?"#fef2f2":"#f8f9fa",border:`1px solid ${c.status==="kept"?"#bbf7d0":c.status==="discarded"?"#fecaca":"#e2e8f0"}`,borderRadius:10,padding:"12px 14px",marginBottom:8,opacity:c.status==="discarded"?0.5:1,display:"flex",alignItems:"center",gap:8}}>
      <div style={{cursor:"grab",padding:"0 8px",color:"#64748b",userSelect:"none"}}>☰</div>
      <div style={{flex:1}}>
        <input value={c.name} onChange={e=>upd(i,"name",e.target.value)} disabled={c.status==="discarded"} placeholder={WL("topicName",lang)} style={{background:"transparent",border:"none",borderBottom:"1px solid #e2e8f0",color:"#1e293b",fontSize:13,fontWeight:600,width:"100%",outline:"none",padding:"2px 0",marginBottom:6}}/>
        <input value={c.keywords} onChange={e=>upd(i,"keywords",e.target.value)} placeholder={WL("keywordsPh",lang)} disabled={c.status==="discarded"} style={{background:"transparent",border:"none",borderBottom:"1px solid #e2e8f0",color:"#1e293b",fontSize:12,width:"100%",outline:"none",padding:"2px 0",marginBottom:6}}/>
        <div style={{fontSize:11,color:"#64748b",fontStyle:"italic"}}>{c.rationale}</div>
      </div>
      <div style={{display:"flex",gap:6,flexShrink:0}}>
        <button onClick={()=>move(i,-1)} disabled={i===0} aria-label="Move topic up" style={{width:32,height:32,borderRadius:8,border:"1px solid #e2e8f0",background:"transparent",color:i===0?"#cbd5e1":"#64748b",cursor:i===0?"default":"pointer",fontSize:12}}>▲</button>
        <button onClick={()=>move(i,1)} disabled={i===cards.length-1} aria-label="Move topic down" style={{width:32,height:32,borderRadius:8,border:"1px solid #e2e8f0",background:"transparent",color:i===cards.length-1?"#cbd5e1":"#64748b",cursor:i===cards.length-1?"default":"pointer",fontSize:12}}>▼</button>
        <button onClick={()=>setSt(i,c.status==="kept"?"pending":"kept")} style={{width:32,height:32,borderRadius:8,border:`1px solid ${c.status==="kept"?"#bbf7d0":"#e2e8f0"}`,background:c.status==="kept"?"#dcfce7":"transparent",color:c.status==="kept"?"#166534":"#64748b",cursor:"pointer",fontSize:16}}>✓</button>
        <button onClick={()=>setSt(i,c.status==="discarded"?"pending":"discarded")} style={{width:32,height:32,borderRadius:8,border:`1px solid ${c.status==="discarded"?"#fecaca":"#e2e8f0"}`,background:c.status==="discarded"?"#fee2e2":"transparent",color:c.status==="discarded"?"#991b1b":"#64748b",cursor:"pointer",fontSize:16}}>✕</button>
      </div>
    </div>)}
    <div style={{display:"flex",gap:8,marginTop:4}}>
      <button onClick={()=>kept.length>0&&onConfirm(kept)} disabled={kept.length===0} style={{background:kept.length>0?P:"#e2e8f0",color:"white",border:"none",borderRadius:8,padding:"8px 20px",fontSize:13,fontWeight:600,cursor:kept.length>0?"pointer":"default"}}>{WL("confirm",lang)} ({kept.length})</button>
      <button onClick={onSkip} style={{background:"transparent",border:"1px solid #e2e8f0",borderRadius:8,padding:"8px 16px",fontSize:13,color:"#64748b",cursor:"pointer"}}>{WL("skip",lang)}</button>
    </div>
  </div>;
}

function PathChoice({ onSubmit }) {
  return <div style={{marginTop:10,display:"flex",gap:10}}>{[{key:"guided",emoji:"🗺️",label:"Guided Setup",desc:"Section by section: markets, languages, objectives and more."},{key:"recommendations",emoji:"⚡",label:"Recommendations",desc:"Share what you're tracking and I'll build smart suggestions."}].map(p => <button key={p.key} onClick={()=>onSubmit(p.key)} style={{flex:1,background:"#f8f9fa",border:"1px solid #e2e8f0",color:"#1e293b",borderRadius:10,padding:"14px 16px",cursor:"pointer",textAlign:"left"}}><div style={{fontSize:18,marginBottom:6}}>{p.emoji}</div><div style={{fontWeight:700,fontSize:13,marginBottom:4}}>{p.label}</div><div style={{fontSize:12,color:"#64748b",lineHeight:1.4}}>{p.desc}</div></button>)}</div>;
}

// Query import limits: extracted file text lands in the API context on submit,
// so cap it before one big agency export blows up the conversation.
const Q_MAX_LINES = 200, Q_MAX_CHARS = 15000, Q_MAX_FILE_BYTES = 2 * 1024 * 1024;
function capQueryText(t) {
  let lines = t.split("\n").map(l=>l.trim()).filter(Boolean);
  let truncated = false;
  if (lines.length > Q_MAX_LINES) { lines = lines.slice(0, Q_MAX_LINES); truncated = true; }
  let out = lines.join("\n");
  if (out.length > Q_MAX_CHARS) { out = out.slice(0, Q_MAX_CHARS); out = out.slice(0, out.lastIndexOf("\n") > 0 ? out.lastIndexOf("\n") : out.length); truncated = true; }
  return { text: out, truncated };
}
function QueriesWidget({ onSubmit, initialData, lang }) {
  const [text,setText] = useState(initialData==="__skip__"||!initialData?"":initialData);
  const [note,setNote] = useState(null);
  const fileRef = useRef(null);
  const ingest = (raw, name) => {
    const { text: capped, truncated } = capQueryText(raw);
    if (!capped) { setNote(QN("noText", lang, { name })); return; }
    setText(t => (t.trim() ? t.trimEnd()+"\n" : "") + capped);
    setNote(truncated
      ? QN("importedTruncated", lang, { n: Q_MAX_LINES, name })
      : QN("imported", lang, { name }));
  };
  const onFile = async e => {
    const f = e.target.files?.[0];
    e.target.value = ""; // allow re-selecting the same file
    if (!f) return;
    setNote(null);
    // Guard before reading: XLSX.read / f.text() load the whole file into memory,
    // so a huge workbook freezes the tab before the line/char cap ever applies.
    if (f.size > Q_MAX_FILE_BYTES) {
      setNote(QN("tooLarge", lang, { mb: (f.size/1048576).toFixed(1) }));
      return;
    }
    try {
      const ext = (f.name.split(".").pop()||"").toLowerCase();
      if (ext === "xlsx" || ext === "xls") {
        const buf = await f.arrayBuffer();
        const XLSX = await loadXLSX();
        const wb = XLSX.read(buf, {type:"array"});
        const rows = [];
        wb.SheetNames.forEach(sn => {
          XLSX.utils.sheet_to_json(wb.Sheets[sn], {header:1, raw:false, defval:""}).forEach(r => {
            const line = r.map(c=>String(c??"").trim()).filter(Boolean).join(" | ");
            if (line) rows.push(line);
          });
        });
        ingest(rows.join("\n"), f.name);
      } else if (ext === "txt" || ext === "csv" || f.type.startsWith("text/")) {
        ingest(await f.text(), f.name);
      } else {
        setNote(QN("unsupported", lang));
      }
    } catch (err) {
      console.error("Query file import failed:", err);
      setNote(QN("readError", lang));
    }
  };
  return <div style={{marginTop:8}}>
    <textarea value={text} onChange={e=>setText(e.target.value)} placeholder={WL("pasteQueries",lang)} rows={4} style={{width:"100%",background:"#f8f9fa",border:"1px solid #e2e8f0",borderRadius:8,padding:"10px 12px",fontSize:12,color:"#1e293b",outline:"none",resize:"vertical",boxSizing:"border-box"}}/>
    <input ref={fileRef} type="file" accept=".txt,.csv,.xlsx,.xls,text/plain,text/csv" onChange={onFile} style={{display:"none"}} aria-hidden="true"/>
    {note && <div style={{fontSize:11,color:"#92400e",background:"#fffbeb",border:"1px solid #fde68a",borderRadius:7,padding:"6px 10px",marginTop:6}}>{note}</div>}
    <div style={{display:"flex",gap:8,marginTop:8,alignItems:"center",flexWrap:"wrap"}}>
      <button onClick={()=>text.trim()&&onSubmit(text.trim())} disabled={!text.trim()} style={{background:text.trim()?P:"#e2e8f0",color:"white",border:"none",borderRadius:8,padding:"8px 20px",fontSize:13,fontWeight:600,cursor:text.trim()?"pointer":"default"}}>{WL("submitQueries",lang)}</button>
      <button onClick={()=>onSubmit("__skip__")} style={{background:"transparent",border:"1px solid #e2e8f0",borderRadius:8,padding:"8px 16px",fontSize:13,color:"#64748b",cursor:"pointer"}}>{WL("noQueries",lang)}</button>
      <button onClick={()=>fileRef.current?.click()} style={{background:"transparent",border:"none",color:LINK,fontSize:12,cursor:"pointer",textDecoration:"underline",padding:"8px 4px"}}>{WL("importFile",lang)}</button>
    </div>
  </div>;
}

function Section({ title, badge, defaultOpen=true, children }) {
  const [open,setOpen] = useState(defaultOpen);
  return <div style={{marginBottom:18}}>
    <button onClick={()=>setOpen(o=>!o)} aria-expanded={open} style={{display:"flex",alignItems:"center",justifyContent:"space-between",width:"100%",background:"transparent",border:"none",cursor:"pointer",padding:"0 0 6px",borderBottom:`2px solid ${P}20`,marginBottom:open?12:0}}>
      <span style={{fontSize:12,fontWeight:700,color:P,textTransform:"uppercase",letterSpacing:"0.06em"}}>{title}{badge!=null && <span style={{marginLeft:8,fontSize:10,fontWeight:600,color:"#64748b",background:"#f1f5f9",borderRadius:8,padding:"1px 7px",textTransform:"none",letterSpacing:0}}>{badge}</span>}</span>
      <span aria-hidden="true" style={{fontSize:11,color:"#94a3b8",transform:open?"rotate(90deg)":"none",transition:"transform 0.15s",display:"inline-block"}}>▶</span>
    </button>
    {open && children}
  </div>;
}

// Bulk import for clients who arrive with a prepared list (e.g. an offline
// template). One item per line; topics support "name | keywords | rationale".
function PasteImport({ label, placeholder, onImport }) {
  const [open,setOpen] = useState(false);
  const [text,setText] = useState("");
  const run = () => {
    const lines = text.split("\n").map(l=>l.trim()).filter(Boolean);
    if (lines.length) onImport(lines);
    setText(""); setOpen(false);
  };
  if (!open) return <button onClick={()=>setOpen(true)} style={{background:"transparent",border:"none",color:LINK,fontSize:12,cursor:"pointer",padding:"6px 0",textDecoration:"underline",marginLeft:10}}>{label}</button>;
  return <div style={{border:`1px solid ${LINK}`,borderRadius:8,padding:"10px 12px",margin:"8px 0",background:"#faf8ff"}}>
    <div style={{fontSize:11,color:"#64748b",marginBottom:6}}>{placeholder}</div>
    <textarea value={text} onChange={e=>setText(e.target.value)} rows={5} style={{width:"100%",border:"1px solid #e2e8f0",borderRadius:7,padding:"7px 10px",fontSize:12,color:"#1e293b",outline:"none",resize:"vertical",boxSizing:"border-box",marginBottom:8}}/>
    <div style={{display:"flex",gap:8}}>
      <button onClick={run} disabled={!text.trim()} style={{background:text.trim()?P:"#e2e8f0",color:"white",border:"none",borderRadius:7,padding:"6px 14px",fontSize:12,fontWeight:600,cursor:text.trim()?"pointer":"default"}}>Import</button>
      <button onClick={()=>{setText("");setOpen(false);}} style={{background:"transparent",border:"1px solid #e2e8f0",borderRadius:7,padding:"6px 12px",fontSize:12,color:"#64748b",cursor:"pointer"}}>Cancel</button>
    </div>
  </div>;
}
const URL_RE = /https?:\/\/[^\s|,]+/;
function guessChanType(url) {
  const u = (url||"").toLowerCase();
  if (u.includes("twitter.")||u.includes("x.com")) return "Twitter/X";
  if (u.includes("linkedin.")) return "LinkedIn";
  if (u.includes("instagram.")) return "Instagram";
  if (u.includes("facebook.")) return "Facebook";
  if (u.includes("youtube.")) return "YouTube";
  if (u.includes("tiktok.")) return "TikTok";
  return "";
}
const GUESS_RE = /suggested by assistant|please verify/i;
const isGuess = tp => GUESS_RE.test(tp?.comments||"") || GUESS_RE.test(tp?.rationale||"");

// Error boundary around the review modal: if anything inside throws, show a
// recoverable message instead of React silently unmounting to a blank screen.
class ModalBoundary extends Component {
  constructor(props){ super(props); this.state={err:null}; }
  static getDerivedStateFromError(err){ return {err}; }
  componentDidCatch(err,info){ console.error("Review modal crashed:",err,info); }
  render(){
    if (!this.state.err) return this.props.children;
    return <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.5)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:1000,padding:16}}>
      <div style={{background:"white",borderRadius:16,maxWidth:440,padding:"28px 28px 22px",boxShadow:"0 16px 48px rgba(0,0,0,0.2)",textAlign:"center"}}>
        <div style={{fontSize:28,marginBottom:10}}>⚠️</div>
        <div style={{fontWeight:700,fontSize:15,color:"#1e293b",marginBottom:8}}>Something went wrong opening your brief</div>
        <div style={{fontSize:12,color:"#64748b",marginBottom:6}}>Your answers are safe. Close this and try again — if it happens twice, let your Lumen contact know.</div>
        <div style={{fontSize:10,color:"#94a3b8",marginBottom:16,fontFamily:"monospace"}}>{String(this.state.err?.message||this.state.err)}</div>
        <button onClick={()=>{this.setState({err:null});this.props.onClose?.();}} style={{background:P,color:"white",border:"none",borderRadius:8,padding:"9px 24px",fontSize:13,fontWeight:600,cursor:"pointer"}}>Close</button>
      </div>
    </div>;
  }
}

function ExportModal({ cdata, wState, messages, onClose, onExport, onSend, sending, sendErr, sent, sheetLink }) {
  // Skipped widgets store the string "__skip__" — returning it caused .join/.map
  // crashes downstream (the "blank screen on Review & send" bug). Treat as null.
  const gw = type => { const es=Object.entries(wState||{}).filter(([k,v])=>k.endsWith(`-${type}`)&&v?.submitted).sort((a,b)=>(parseInt(a[0])||0)-(parseInt(b[0])||0)); const d=es.length?es[es.length-1][1].data:null; return d==="__skip__"?null:d; };
  const historyName = useMemo(() => {
    const m = messages.filter(m=>m.role==="user").map(m=>String(m.content||"")).join(" ")
      .match(/(?:company|we are|we're|I'm from|I work at)[^\w]*([A-Z][A-Za-z0-9& ]{1,40})/);
    return m?.[1]?.trim()||"";
  }, [messages]);
  const objW = normObjectives(gw("OBJECTIVES"));
  const [co,setCo]     = useState({name:"",email:"",industry:"",useCase:"",contact:"",...cdata.company,name:cdata.company?.name||historyName});
  const [mkts,setMkts] = useState((gw("MARKETS")||[]).join(", ")||cdata.company?.markets||"");
  const [langs,setLangs]= useState((gw("LANGUAGES")||[]).join(", ")||cdata.company?.languages||"");
  const [objs,setObjs] = useState(fmtRanked(objW)||cdata.company?.objectives||"");
  const [objDetails,setObjDetails] = useState(objW.details||"");
  const [teams,setTeams]= useState((gw("TEAMS")||[]).join(", ")||cdata.company?.teams||"");
  const [tz,setTz]     = useState(Array.isArray(gw("TIMEZONE"))?gw("TIMEZONE")[0]:(gw("TIMEZONE")||cdata.company?.timezone||""));
  const [users,setUsers]= useState(gw("USERS")||[]);
  // Topics can arrive two ways: the confirmed topic cards (name/keywords/rationale
  // only) and the %%TOPICS%% marker (also urls/hashtags/comments). Merge by name so
  // the marker's urls/hashtags survive into the brief instead of being dropped when
  // the card widget was used.
  const [topics,setTopics]= useState(() => {
    const cards = gw("TOPICS"), markers = cdata.topics || [];
    const byName = {};
    markers.forEach(m => { if (m && m.name) byName[String(m.name).trim().toLowerCase()] = m; });
    const base = (cards && cards.length) ? cards : markers;
    return base.map((tp,i) => {
      const m = byName[String(tp.name||"").trim().toLowerCase()] || {};
      return { name:tp.name||m.name||"", keywords:tp.keywords||m.keywords||"",
        rationale:tp.rationale||m.rationale||"", urls:tp.urls||m.urls||"",
        hashtags:tp.hashtags||m.hashtags||"", comments:tp.comments||m.comments||"",
        group:tp.group||m.group||"",
        id:i, confirmed:!(isGuess(tp)||isGuess(m)) };
    });
  });
  const [chans,setChans]= useState((cdata.channels||[]).map((c,i)=>({...c,id:i})));
  const [reports,setReports]= useState((cdata.reports||[]).map((r,i)=>({...r,id:i})));
  const [alerts,setAlerts]= useState((cdata.alerts||[]).map((a,i)=>({...a,id:i})));
  const emptyUser  = () => ({ firstName:"",lastName:"",email:"",role:"",access:"Full Tool" });
  const emptyChan  = () => ({ author:"",type:"",url:"",owned:"" });
  const emptyReport = () => ({ name:"",objective:"",details:"",comments:"" });
  const emptyAlert  = () => ({ name:"",type:"",details:"",comments:"" });
  const emptyTopic = () => ({ name:"",keywords:"",rationale:"",comments:"",id:Date.now(),confirmed:true });
  const confirmTopic = (i,v) => setTopics(ts=>ts.map((x,j)=>j===i?{...x,confirmed:v,comments:v?(x.comments||"").replace(GUESS_RE,"").replace(/^[\s,-]+|[\s,-]+$/g,"")||"Confirmed by client":x.comments}:x));
  const unconfirmed = topics.filter(t=>!t.confirmed).length;
  // readiness scoring
  const reqChecks = [
    ["Company name", !!co.name],
    ["Contact email", !!co.email && EMAIL_RE.test(co.email)],
    ["Markets", !!mkts.trim()],
    ["Languages", !!langs.trim()],
    ["Objectives", !!objs.trim()],
    ["At least one topic", topics.length>0],
    ["All topics confirmed", topics.length>0 && unconfirmed===0],
    ["At least one user", users.length>0],
  ];
  const passed = reqChecks.filter(c=>c[1]).length;
  const pct = Math.round(passed/reqChecks.length*100);
  const gaps = reqChecks.filter(c=>!c[1]).map(c=>c[0]);
  const ready = gaps.length===0;
  const fld = (label,val,set,multi,req) => <div style={{marginBottom:12}}>
    <div style={{fontSize:11,fontWeight:600,color:"#64748b",marginBottom:4,display:"flex",alignItems:"center",gap:6}}>
      {label}<span style={{fontSize:10,padding:"1px 6px",borderRadius:4,fontWeight:600,background:req?"#fef2f2":"#f1f5f9",color:req?"#dc2626":"#94a3b8"}}>{req?"Required":"Optional"}</span>
    </div>
    {multi
      ? <textarea value={val} onChange={e=>set(e.target.value)} rows={2} style={{width:"100%",border:`1px solid ${req&&!val?"#fca5a5":"#e2e8f0"}`,borderRadius:7,padding:"7px 10px",fontSize:12,color:"#1e293b",outline:"none",resize:"vertical",boxSizing:"border-box"}}/>
      : <input value={val} onChange={e=>set(e.target.value)} style={{width:"100%",border:`1px solid ${req&&!val?"#fca5a5":"#e2e8f0"}`,borderRadius:7,padding:"7px 10px",fontSize:12,color:"#1e293b",outline:"none"}}/>}
  </div>;
  const addBtn = (label,onClick) => <button onClick={onClick} style={{background:"transparent",border:`1px dashed ${LINK}`,color:LINK,borderRadius:8,padding:"6px 14px",fontSize:12,cursor:"pointer",marginTop:6}}>{label}</button>;
  const merged = { company:{...co,markets:mkts,languages:langs,objectives:objs,objectiveDetails:objDetails,teams,timezone:tz}, topics:topics.map(({confirmed,id,...t})=>t), channels:chans.map(({id,...c})=>c), reports:reports.map(({id,...r})=>r), alerts:alerts.map(({id,...a})=>a), queries:gw("QUERIES")||"" };
  const dialogRef = useRef(null);
  useEffect(() => {
    // Dialog a11y: close on Escape and move focus into the dialog on open, so
    // keyboard and screen-reader users aren't stranded behind the chat.
    const onKey = e => { if (e.key === "Escape") onClose(); };
    document.addEventListener("keydown", onKey);
    dialogRef.current && dialogRef.current.focus();
    return () => document.removeEventListener("keydown", onKey);
  }, [onClose]);
  return <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.5)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:1000,padding:16}}>
    <div ref={dialogRef} role="dialog" aria-modal="true" aria-label="Your setup brief" tabIndex={-1} style={{background:"white",borderRadius:16,width:"100%",maxWidth:680,maxHeight:"90vh",display:"flex",flexDirection:"column",boxShadow:"0 16px 48px rgba(0,0,0,0.2)",outline:"none"}}>
      <div style={{padding:"20px 24px 16px",borderBottom:"1px solid #e2e8f0",display:"flex",justifyContent:"space-between",alignItems:"center",flexShrink:0}}>
        <div><div style={{fontWeight:700,fontSize:16,color:"#1e293b"}}>Your setup brief</div><div style={{fontSize:12,color:"#64748b",marginTop:2}}>Everything you've shared, in one place. Open a section to adjust anything.</div></div>
        <button onClick={onClose} aria-label="Close review" style={{background:"transparent",border:"none",fontSize:20,cursor:"pointer",color:"#94a3b8"}}>✕</button>
      </div>
      <div style={{flex:1,overflowY:"auto",padding:"20px 24px"}}>
        <div style={{display:"flex",alignItems:"center",gap:14,marginBottom:20,padding:"12px 14px",borderRadius:10,background:ready?"#f0fdf4":"#fffbeb",border:`1px solid ${ready?"#bbf7d0":"#fde68a"}`}}>
          <div style={{position:"relative",width:52,height:52,flexShrink:0}}>
            <svg width="52" height="52" viewBox="0 0 52 52">
              <circle cx="26" cy="26" r="22" fill="none" stroke="#e2e8f0" strokeWidth="6"/>
              <circle cx="26" cy="26" r="22" fill="none" stroke={ready?"#16a34a":A} strokeWidth="6" strokeLinecap="round" strokeDasharray={`${2*Math.PI*22*pct/100} ${2*Math.PI*22}`} transform="rotate(-90 26 26)"/>
            </svg>
            <div style={{position:"absolute",inset:0,display:"flex",alignItems:"center",justifyContent:"center",fontSize:13,fontWeight:700,color:"#1e293b"}}>{pct}%</div>
          </div>
          <div style={{flex:1,minWidth:0}}>
            <div style={{fontWeight:700,fontSize:13,color:"#1e293b"}}>{ready?"Ready to send":"Almost there"}</div>
            <div style={{fontSize:11,color:"#64748b",margin:"1px 0 2px"}}>{topics.length} topic{topics.length!==1?"s":""} · {chans.length} channel{chans.length!==1?"s":""} · {reports.length+alerts.length} report{(reports.length+alerts.length)!==1?"s":""} · {users.length} user{users.length!==1?"s":""}</div>
            {ready
              ? <div style={{fontSize:12,color:"#166534"}}>All required fields complete and all topics confirmed.</div>
              : <div style={{fontSize:12,color:"#92400e"}}>Still needed: {gaps.join(", ")}</div>}
          </div>
        </div>
        <Section title="About your business" defaultOpen={!co.name||!co.email||!mkts.trim()||!langs.trim()||!objs.trim()}>
          {fld("Company Name",co.name,v=>setCo(c=>({...c,name:v})),false,true)}
          {fld("Contact Email",co.email,v=>setCo(c=>({...c,email:v})),false,true)}
          {fld("Industry",co.industry,v=>setCo(c=>({...c,industry:v})),false,false)}
          {fld("Geographic Markets",mkts,setMkts,false,true)}
          {fld("Key Languages",langs,setLangs,false,true)}
          {fld("Business Objectives",objs,setObjs,false,true)}
          {fld("Objective Details",objDetails,setObjDetails,true,false)}
          {fld("Use Cases",co.useCase,v=>setCo(c=>({...c,useCase:v})),true,false)}
          {fld("Preferred Time Zone",tz,setTz,false,false)}
          {fld("Teams / Departments",teams,setTeams,false,false)}
          {fld("Main Point of Contact",co.contact,v=>setCo(c=>({...c,contact:v})),false,false)}
        </Section>
        <Section title="Your team" badge={users.length} defaultOpen={users.length===0}>
          {users.length===0 && <div style={{fontSize:12,color:"#94a3b8",fontStyle:"italic",marginBottom:8}}>No users captured.</div>}
          {users.map((u,i) => <div key={i} style={{background:"#f8f9fa",borderRadius:8,padding:"10px 12px",marginBottom:8}}>
            <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(200px,1fr))",gap:6,marginBottom:6}}>
              {["firstName","lastName","email","role"].map(k => <input key={k} value={u[k]||""} placeholder={k} onChange={e=>setUsers(us=>us.map((x,j)=>j===i?{...x,[k]:e.target.value}:x))} style={{border:"1px solid #e2e8f0",borderRadius:6,padding:"6px 8px",fontSize:11,outline:"none"}}/>)}
            </div>
            <div style={{display:"flex",alignItems:"center",justifyContent:"space-between"}}>
              <div style={{display:"flex",gap:4}}>{["Admin","Full Tool","Read-Only"].map(a => <button key={a} onClick={()=>setUsers(us=>us.map((x,j)=>j===i?{...x,access:a}:x))} style={{padding:"3px 8px",borderRadius:5,fontSize:10,cursor:"pointer",border:"1px solid",background:u.access===a?P:"transparent",borderColor:u.access===a?P:"#e2e8f0",color:u.access===a?"white":"#64748b"}}>{a}</button>)}</div>
              <button onClick={()=>setUsers(us=>us.filter((_,j)=>j!==i))} style={{background:"transparent",border:"none",color:"#ef4444",cursor:"pointer",fontSize:12}}>✕</button>
            </div>
          </div>)}
          {addBtn("+ Add user", ()=>setUsers(u=>[...u,emptyUser()]))}
        </Section>
        <Section title="What we'll track" badge={topics.length} defaultOpen={topics.length===0||unconfirmed>0}>
          {topics.length===0 && <div style={{fontSize:12,color:"#94a3b8",fontStyle:"italic",marginBottom:8}}>No topics captured.</div>}
          {unconfirmed>0 && <div style={{fontSize:11,color:"#92400e",background:"#fffbeb",border:"1px solid #fde68a",borderRadius:7,padding:"7px 10px",marginBottom:10,display:"flex",gap:6}}><span>⚠</span><span>{unconfirmed} topic{unconfirmed!==1?"s were":" was"} suggested by the assistant. Confirm or drop {unconfirmed!==1?"them":"it"} before handing off.</span></div>}
          {topics.map((tp,i) => { const guess = !tp.confirmed; return <div key={tp.id} style={{background:guess?"#fffbeb":"#f0fdf4",border:`1px solid ${guess?"#fde68a":"#bbf7d0"}`,borderRadius:8,padding:"10px 12px",marginBottom:8}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",gap:8,marginBottom:6}}>
              <div style={{fontSize:10,fontWeight:700,display:"flex",alignItems:"center",gap:5,color:guess?"#d97706":"#16a34a",textTransform:"uppercase",letterSpacing:"0.04em"}}><span>{guess?"⚠":"✓"}</span>{guess?"Assistant guess":"Confirmed"}</div>
              {guess
                ? <div style={{display:"flex",gap:6}}>
                    <button onClick={()=>confirmTopic(i,true)} style={{background:P,color:"#fff",border:"none",borderRadius:6,padding:"4px 10px",fontSize:11,fontWeight:600,cursor:"pointer"}}>Confirm</button>
                    <button onClick={()=>setTopics(ts=>ts.filter((_,j)=>j!==i))} style={{background:"transparent",border:"1px solid #e2e8f0",borderRadius:6,padding:"4px 9px",fontSize:11,color:"#64748b",cursor:"pointer"}}>Drop</button>
                  </div>
                : <button onClick={()=>setTopics(ts=>ts.filter((_,j)=>j!==i))} style={{background:"transparent",border:"none",color:"#ef4444",cursor:"pointer",fontSize:12,flexShrink:0}}>✕</button>}
            </div>
            <div>
              <input value={tp.name||""} placeholder="Topic name" onChange={e=>setTopics(ts=>ts.map((x,j)=>j===i?{...x,name:e.target.value}:x))} style={{width:"100%",border:"none",borderBottom:"1px solid #e2e8f0",fontSize:13,fontWeight:600,outline:"none",background:"transparent",marginBottom:6,padding:"2px 0"}}/>
              <input value={tp.keywords||""} placeholder="Keywords…" onChange={e=>setTopics(ts=>ts.map((x,j)=>j===i?{...x,keywords:e.target.value}:x))} style={{width:"100%",border:"none",borderBottom:"1px solid #e2e8f0",fontSize:12,outline:"none",background:"transparent",padding:"2px 0",marginBottom:6}}/>
              <input value={tp.rationale||tp.comments||""} placeholder="Rationale / comments…" onChange={e=>setTopics(ts=>ts.map((x,j)=>j===i?{...x,rationale:e.target.value,comments:e.target.value}:x))} style={{width:"100%",border:"none",fontSize:11,outline:"none",background:"transparent",padding:"2px 0",color:"#64748b",fontStyle:"italic"}}/>
            </div>
          </div>; })}
          {addBtn("+ Add topic", ()=>setTopics(ts=>[...ts,emptyTopic()]))}
          <PasteImport label="Have a list already? Paste it" placeholder={"One topic per line. Optionally add keywords and a note separated by | — e.g. Nike | \"Nike\" OR @Nike | main competitor"} onImport={lines=>setTopics(ts=>[...ts,...lines.map((l,i)=>{ const p=l.split("|").map(s=>s.trim()); return {name:p[0]||"",keywords:p[1]||"",rationale:p[2]||"",comments:p[2]||"Imported from client list",id:Date.now()+i,confirmed:true}; })])}/>
        </Section>
        <Section title="Where we'll look" badge={chans.length} defaultOpen={false}>
          {chans.length===0 && <div style={{fontSize:12,color:"#94a3b8",fontStyle:"italic",marginBottom:8}}>No channels captured.</div>}
          {chans.map((ch,i) => <div key={ch.id} style={{display:"flex",gap:6,marginBottom:8,alignItems:"center"}}>
            {["author","type","url","owned"].map(k => <input key={k} value={ch[k]||""} placeholder={k} onChange={e=>setChans(cs=>cs.map((x,j)=>j===i?{...x,[k]:e.target.value}:x))} style={{flex:1,border:"1px solid #e2e8f0",borderRadius:6,padding:"6px 8px",fontSize:11,outline:"none"}}/>)}
            <button onClick={()=>setChans(cs=>cs.filter((_,j)=>j!==i))} style={{background:"transparent",border:"none",color:"#ef4444",cursor:"pointer",fontSize:12,flexShrink:0}}>✕</button>
          </div>)}
          {addBtn("+ Add channel", ()=>setChans(cs=>[...cs,{...emptyChan(),id:Date.now()}]))}
          <PasteImport label="Have a list already? Paste it" placeholder={"One channel per line — a URL, a name, or both, e.g. Nike https://twitter.com/nike"} onImport={lines=>setChans(cs=>[...cs,...lines.map((l,i)=>{ const u=l.match(URL_RE)?.[0]||""; const author=l.replace(u,"").replace(/[|,]/g," ").trim(); return {author:author||"",type:guessChanType(u),url:u,owned:"",id:Date.now()+i}; })])}/>
        </Section>
        <Section title="Reports and alerts" badge={reports.length+alerts.length} defaultOpen={false}>
          <div style={{fontSize:11,fontWeight:700,color:"#64748b",textTransform:"uppercase",letterSpacing:"0.04em",marginBottom:6}}>Reports and dashboards</div>
          {reports.length===0 && <div style={{fontSize:12,color:"#94a3b8",fontStyle:"italic",marginBottom:8}}>No reports captured.</div>}
          {reports.map((r,i) => <div key={r.id} style={{display:"flex",gap:6,marginBottom:8,alignItems:"center"}}>
            {["name","objective","details","comments"].map(k => <input key={k} value={r[k]||""} placeholder={k} onChange={e=>setReports(rs=>rs.map((x,j)=>j===i?{...x,[k]:e.target.value}:x))} style={{flex:1,border:"1px solid #e2e8f0",borderRadius:6,padding:"6px 8px",fontSize:11,outline:"none"}}/>)}
            <button onClick={()=>setReports(rs=>rs.filter((_,j)=>j!==i))} style={{background:"transparent",border:"none",color:"#ef4444",cursor:"pointer",fontSize:12,flexShrink:0}}>✕</button>
          </div>)}
          {addBtn("+ Add report", ()=>setReports(rs=>[...rs,{...emptyReport(),id:Date.now()}]))}
          <div style={{fontSize:11,fontWeight:700,color:"#64748b",textTransform:"uppercase",letterSpacing:"0.04em",margin:"14px 0 6px"}}>Alerts</div>
          {alerts.length===0 && <div style={{fontSize:12,color:"#94a3b8",fontStyle:"italic",marginBottom:8}}>No alerts captured.</div>}
          {alerts.map((a,i) => <div key={a.id} style={{display:"flex",gap:6,marginBottom:8,alignItems:"center"}}>
            {["name","type","details","comments"].map(k => <input key={k} value={a[k]||""} placeholder={k} onChange={e=>setAlerts(as=>as.map((x,j)=>j===i?{...x,[k]:e.target.value}:x))} style={{flex:1,border:"1px solid #e2e8f0",borderRadius:6,padding:"6px 8px",fontSize:11,outline:"none"}}/>)}
            <button onClick={()=>setAlerts(as=>as.filter((_,j)=>j!==i))} style={{background:"transparent",border:"none",color:"#ef4444",cursor:"pointer",fontSize:12,flexShrink:0}}>✕</button>
          </div>)}
          {addBtn("+ Add alert", ()=>setAlerts(as=>[...as,{...emptyAlert(),id:Date.now()}]))}
        </Section>
      </div>
      <div style={{padding:"16px 24px",borderTop:"1px solid #e2e8f0",display:"flex",justifyContent:"space-between",alignItems:"center",gap:10,flexShrink:0}}>
        <div style={{fontSize:11,color:ready?"#16a34a":"#92400e"}}>{ready?"✓ Ready to download":`${gaps.length} item${gaps.length!==1?"s":""} to resolve first`}</div>
        <div style={{display:"flex",gap:10,alignItems:"center"}}>
          {sendErr && <div style={{fontSize:11,color:"#dc2626",maxWidth:240,lineHeight:1.4}}>{sendErr==="send-failed"?"We couldn’t send your brief just now. Please check your connection and press Send again.":sendErr}</div>}
          <button onClick={onClose} style={{background:"transparent",border:"1px solid #e2e8f0",borderRadius:8,padding:"9px 20px",fontSize:13,color:"#64748b",cursor:"pointer"}}>Cancel</button>
          {!(sent && sheetLink) && <button onClick={()=>ready&&onExport(merged,users)} disabled={!ready} title={ready?"":"Resolve the readiness gaps first"} style={{background:"transparent",border:`1px solid ${ready?P:"#e2e8f0"}`,color:ready?P:"#94a3b8",borderRadius:8,padding:"9px 18px",fontSize:13,fontWeight:600,cursor:ready?"pointer":"not-allowed"}}>⬇ Download a copy</button>}
          <button onClick={()=>ready&&!sending&&onSend(merged,users)} disabled={!ready||sending} title={ready?"":"Resolve the readiness gaps first"} style={{background:ready?A:"#e2e8f0",color:ready?"white":"#94a3b8",border:"none",borderRadius:8,padding:"9px 24px",fontSize:13,fontWeight:600,cursor:ready&&!sending?"pointer":"not-allowed"}}>{sending?"Sending\u2026":"\ud83d\udce8 Send to my Lumen team"}</button>
        </div>
      </div>
    </div>
  </div>;
}

async function doExport(merged, users, rawMessages) {
  let companyName = merged.company?.name;
  if (!companyName) {
    const m = (rawMessages||[]).filter(m=>m.role==="user").map(m=>m.content||"").join(" ")
      .match(/(?:company|we are|we're|I'm from|I work at)[^\w]*([A-Z][A-Za-z0-9& ]{1,40})/);
    companyName = m?.[1]?.trim()||"Draft";
  }
  const XLSX = await loadXLSX();
  const { wb, filename } = buildWorkbook(XLSX, {...merged, company:{...merged.company, name:companyName}}, users);
  XLSX.writeFile(wb, filename);
}

function FinishCard({ C, cdata, setShowExport, linkCopied, setLinkCopied, sent, sheetLink, onSeeProserv }) {
  return (
    <div style={{display:"flex",justifyContent:"center",marginBottom:24,animation:"slideUpFade 0.5s ease-out forwards"}}>
      <div style={{background:`linear-gradient(135deg,${P}15,${P}08)`,border:`1.5px solid ${P}`,borderRadius:14,padding:"20px 28px",textAlign:"center",maxWidth:460}}>
        <div style={{fontSize:20,marginBottom:8}}>{sent?"\u2705":"\ud83c\udf89"}</div>
        <div style={{fontWeight:700,fontSize:15,color:C.text,marginBottom:6}}>{sent?"Brief sent to your Lumen team":"Setup brief ready"}</div>
        <div style={{fontSize:13,color:C.muted,marginBottom:16,lineHeight:1.5}}>
          {sent
            ? (sheetLink
                ? "Your setup brief has been sent to your Lumen team, and we've shared an editable Google Sheet with you (check your email). Update it anytime before your review call, and your consultant will see the changes. A consultant will be in touch within 2 business days."
                : "Your setup brief has been sent to your Lumen team. A consultant will be in touch within 2 business days to book your review call, where you'll finalise the setup together. You can review or download a copy of your brief below.")
            : "Review your brief, then send it straight to your Lumen team\u2014 nothing to download or email."}
        </div>
        <div style={{display:"flex",gap:8,justifyContent:"center",flexWrap:"wrap"}}>
          {sent && sheetLink && <a href={sheetLink} target="_blank" rel="noopener noreferrer" style={{background:P,color:"white",borderRadius:10,padding:"10px 20px",fontSize:13,fontWeight:600,cursor:"pointer",textDecoration:"none",display:"inline-block"}}>Open your brief (Google Sheet)</a>}
          <button onClick={()=>setShowExport(true)} style={{background:sent&&sheetLink?C.card:A,color:sent&&sheetLink?C.muted:"white",border:sent&&sheetLink?`1px solid ${C.border}`:"none",borderRadius:10,padding:"10px 20px",fontSize:13,fontWeight:600,cursor:"pointer"}}>{sent?(sheetLink?"Review":"Review / download a copy"):"\ud83d\udce8 Review \u0026 send"}</button>
          {sent && onSeeProserv && <button onClick={onSeeProserv} style={{background:"#012B3A",color:"white",border:"none",borderRadius:10,padding:"10px 16px",fontSize:13,fontWeight:600,cursor:"pointer"}}>See what Proserv receives →</button>}
        </div>
      </div>
    </div>
  );
}

function OnboardingApp({ seed, seedId, onBriefSent, onSeeProserv }) {
  const [theme,setTheme]       = useState("light");
  const [sound,setSound]       = useState(false);
  const [persona,setPersona]   = useState("strategist");
  const [uiLang,setUiLang]     = useState(seed?.language || "English");
  const [messages,setMessages] = useState([]);
  const [input,setInput]       = useState("");
  const [loading,setLoading]   = useState(false);
  const [progress,setProgress] = useState({percent:0,collected:{}});
  const [started,setStarted]   = useState(false);
  const [wState,setWState]     = useState({});
  const [saved,setSaved]       = useState(null);
  const [checked,setChecked]   = useState(false);
  const [collapsed,setCollapsed]= useState(true);
  const [rewind,setRewind]     = useState(null);
  const [showExport,setShowExport]= useState(false);
  const [cdata,setCdata]       = useState(emptyCdata());
  const [retryMsg,setRetryMsg] = useState(null);
  const [initErr,setInitErr]   = useState(null); // "start" | "resume" | null — first-turn/resume API failure, offers retry
  const [draftOk,setDraftOk]   = useState(lsProbe); // is on-device draft saving actually working?
  const [showPanel,setShowPanel] = useState(() => typeof window !== "undefined" && window.innerWidth > 1080);
  const [linkCopied,setLinkCopied] = useState(false);
  const [smoke,setSmoke]       = useState(null);
  const [showSmokeMenu,setShowSmokeMenu] = useState(false);
  const [sent,setSent]         = useState(false);
  const [sending,setSending]   = useState(false);
  const [sendErr,setSendErr]   = useState(null);
  const [sheetLink,setSheetLink] = useState(null);
  const [streamTxt,setStreamTxt] = useState("");
  const [ww,setWw] = useState(typeof window !== "undefined" ? window.innerWidth : 1200);
  useEffect(() => { const f = () => setWw(window.innerWidth); window.addEventListener("resize", f); return () => window.removeEventListener("resize", f); }, []);
  const mob = ww < 640;

  const botRef  = useRef(null);
  const histRef = useRef([]);
  const taRef   = useRef(null);
  const msgRef  = useRef(null);
  const prevPct = useRef(0);
  const sndRef  = useRef(sound);
  const personaRef = useRef(persona);
  const prevSecRef = useRef(null);
  const sidRef  = useRef(null);
  const apiCountRef = useRef(0);
  const usageRef = useRef({ input: 0, output: 0, cacheRead: 0, cacheWrite: 0 });
  const sendingRef = useRef(false); // synchronous double-send guard (state lags a fast double-click)
  const startedAtRef = useRef(null);
  const saveT   = useRef(null);
  const wRef    = useRef(wState);
  const smokeStop = useRef(false);
  useEffect(() => { sndRef.current = sound; }, [sound]);
  useEffect(() => { personaRef.current = persona; }, [persona]);
  useEffect(() => { wRef.current = wState; }, [wState]);

  const { init, pop, chime } = useAudio();
  const dark = theme === "dark";
  const C = useMemo(() => dark
    ? {bg:"#0d1b2a",card:"#111f30",border:"#1e3048",muted:"#8aa4c1",text:"#c8d8e8",hi:"#1a2f4a",uBg:"#1e3a5f",uTx:"#d0e8ff",wTx:"#a89af0"}
    : {bg:"#F7F7FA",card:"#ffffff",border:"#E7E7EF",muted:"#64748b",text:"#1e293b",hi:"#F1F0F7",uBg:P,uTx:"#F2F7F8",wTx:LINK}
  , [dark]);

  useEffect(() => { botRef.current?.scrollIntoView({behavior:"smooth"}); }, [messages, loading]);
  useEffect(() => { if (progress.percent===100&&prevPct.current<100&&sndRef.current) chime(); prevPct.current=progress.percent; }, [progress.percent, chime]);

  // On mount, offer to resume an in-progress draft saved on this device.
  useEffect(() => {
    const draft = lsLoadDraft(seedId);
    if (draft) setSaved(draft);
    setChecked(true);
  }, []);

  // Autosave the in-progress draft after each turn, until sent. Two sinks:
  //  - localStorage: the full draft (messages/history) for resume on this device.
  //  - server (from the first real answer on): a TRIMMED snapshot — structured
  //    progress only, no messages/history — so Proserv sees live and stalled
  //    sessions, not just completed ones. Keyed by session id, marked
  //    in_progress; the completed record overwrites it on send (same id).
  useEffect(() => {
    if (!started || sent || messages.length === 0) return;
    if (saveT.current) clearTimeout(saveT.current);
    saveT.current = setTimeout(() => {
      setDraftOk(lsSaveDraft(seedId, { messages, progress, wState, cdata, history: histRef.current, uiLang, sid: sidRef.current, startedAt: startedAtRef.current, savedAt: Date.now() }));
      // Server upsert. Best-effort, never blocks the chat. Skipped while a send is
      // in flight so a late autosave can't overwrite the completed record.
      const pct = (progress && progress.percent) || 0;
      const hasRealAnswer = !!(cdata.company && cdata.company.name) || pct > 0;
      if (hasRealAnswer && !sendingRef.current) {
        const usersW = gwp("USERS");
        const inProgress = {
          id: sidRef.current,
          status: "in_progress",
          percent: pct,
          merged: { company: cdata.company || {}, topics: cdata.topics || [], channels: cdata.channels || [], reports: cdata.reports || [], alerts: cdata.alerts || [], queries: gwp("QUERIES") || "" },
          users: Array.isArray(usersW) ? usersW : [],
          handoff: cdata.handoff || null,
          seedId: seedId || null,
          seed: seed || null,
          durationMs: startedAtRef.current ? (Date.now() - startedAtRef.current) : null,
          apiCalls: apiCountRef.current,
          tokens: { ...usageRef.current },
          lastActiveAt: new Date().toISOString(),
        };
        fetchWithTimeout(SESSION_ENDPOINT, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ session: inProgress }) }, 15000).catch(() => {});
      }
    }, 600);
    return () => { if (saveT.current) clearTimeout(saveT.current); };
  }, [messages, progress, wState, cdata, started, sent, uiLang, seedId, seed]);

  const resetSession = useCallback(() => {
    sidRef.current = crypto.randomUUID();
    setSent(false); setSendErr(null);
    setStarted(false); setMessages([]); setProgress({percent:0,collected:{}});
    setWState({}); setCdata(emptyCdata());
    setSaved(null); setRetryMsg(null); histRef.current = [];
    prevSecRef.current = null; prevPct.current = 0;
    apiCountRef.current = 0; usageRef.current = { input: 0, output: 0, cacheRead: 0, cacheWrite: 0 };
  }, []);

  const loadTestSession = useCallback(() => {
    const tm = [
      {role:"assistant",content:"Hi! I'm your Lumen onboarding assistant. To get started, what's your company name?",timestamp:"09:00",widgets:[],quickReplies:[],raw:""},
      {role:"user",content:"Acme Corp",timestamp:"09:00",raw:"Acme Corp"},
      {role:"assistant",content:"Great to meet you, Acme Corp. What's the best contact email for the setup?",timestamp:"09:01",widgets:[],quickReplies:[],raw:""},
      {role:"user",content:"jane@acmecorp.com",timestamp:"09:01",raw:"jane@acmecorp.com"},
      {role:"assistant",content:"Perfect. From the name I'm guessing you're in consumer goods or retail, is that right? And before we dive in: what are you hoping to get out of Lumen?",timestamp:"09:02",widgets:[],quickReplies:[],raw:""},
      {role:"user",content:"Yes, consumer goods, mainly footwear and apparel. We want to protect our brand reputation, keep ahead of competitors like Nike, Adidas and Puma, and catch any customer issues before they blow up.",timestamp:"09:02",raw:"Yes, consumer goods, mainly footwear and apparel. We want to protect our brand reputation, keep ahead of competitors like Nike, Adidas and Puma, and catch any customer issues before they blow up."},
      {role:"assistant",content:"That's really helpful. Brand reputation, competitive tracking, and early issue detection gives us a clear direction. Would you like me to guide you section by section, or would you rather share what you're tracking and have me build recommendations?",timestamp:"09:03",widgets:[],quickReplies:[],raw:""},
      {role:"user",content:"✓ PATH: Guided path",isWidget:true,timestamp:"09:03"},
      {role:"assistant",content:"Great, we'll go section by section. Do you have any existing queries from another tool you'd like to migrate?",timestamp:"09:04",widgets:[],quickReplies:[],raw:""},
      {role:"user",content:"✓ QUERIES: Skipped",isWidget:true,timestamp:"09:04"},
      {role:"assistant",content:"No problem, we'll build fresh. Before the structured bits, let's line up what to track. Beyond your own brand, what should we monitor?",timestamp:"09:05",widgets:[],quickReplies:[],raw:""},
      {role:"user",content:"Our brand and product lines, our three competitors, our summer campaign, the broader sustainability-in-fashion conversation, and customer service complaints.",timestamp:"09:05",raw:"Our brand and product lines, our three competitors, our summer campaign, the broader sustainability-in-fashion conversation, and customer service complaints."},
      {role:"assistant",content:"Strong list. That gives us eight topics spanning your brand, your product lines, the three competitors, your summer campaign, an industry theme, and a customer-service risk lens. Which geographic markets matter most to you?",timestamp:"09:06",widgets:[],quickReplies:[],raw:""},
      {role:"user",content:"✓ MARKETS: United States, United Kingdom, France, Germany, Canada, Australia",isWidget:true,timestamp:"09:06"},
      {role:"assistant",content:"Six markets across North America, Europe and APAC. Which languages should we monitor?",timestamp:"09:07",widgets:[],quickReplies:[],raw:""},
      {role:"user",content:"✓ LANGUAGES: English, French, German",isWidget:true,timestamp:"09:07"},
      {role:"assistant",content:"Given your goal of protecting brand reputation, tracking competitors, and catching issues early, I'd put Reputation Management first, then Competitive Intelligence, then Issue Tracking. Pick up to 3 and set the order — your top priority decides what we build first.",timestamp:"09:08",widgets:[],quickReplies:[],raw:""},
      {role:"user",content:"✓ OBJECTIVES: 1. Reputation Management, 2. Competitive Intelligence, 3. Issue Tracking",isWidget:true,timestamp:"09:08"},
      {role:"assistant",content:"Priorities locked: Reputation Management first, then Competitive Intelligence, then Issue Tracking. Since issue tracking made your top 3, what kind of issue are you most concerned about catching early, product complaints, executive reputation, or something else?",timestamp:"09:09",widgets:[],quickReplies:[],raw:""},
      {role:"user",content:"Mostly product complaints and anything that could escalate quickly on social.",timestamp:"09:09",raw:"Mostly product complaints and anything that could escalate quickly on social."},
      {role:"assistant",content:"Quick check so far: Acme Corp, consumer goods (footwear and apparel), focused on brand reputation, competitive intelligence and issue tracking (in that priority order), tracking eight topics (your brand, product lines, Nike, Adidas, Puma, your summer campaign, sustainability in fashion, and customer-service issues) across six markets in three languages. Does that look right?",timestamp:"09:10",widgets:[],quickReplies:[],raw:""},
      {role:"user",content:"Yes, that's right.",timestamp:"09:10",raw:"Yes, that's right."},
      {role:"assistant",content:"Great. Which teams or departments will be using Lumen?",timestamp:"09:11",widgets:[],quickReplies:[],raw:""},
      {role:"user",content:"✓ TEAMS: Marketing, Communications, PR, Brand, Social Media, Customer Experience",isWidget:true,timestamp:"09:11"},
      {role:"assistant",content:"And what's your primary timezone for scheduling reports and alerts?",timestamp:"09:12",widgets:[],quickReplies:[],raw:""},
      {role:"user",content:"✓ TIMEZONE: GMT / UTC",isWidget:true,timestamp:"09:12"},
      {role:"assistant",content:"Now let's capture your social channels. Which accounts should we monitor, your own and any competitors?",timestamp:"09:13",widgets:[],quickReplies:[],raw:""},
      {role:"user",content:"Our Twitter, LinkedIn, Instagram and Facebook, plus the main Twitter accounts for Nike and Adidas and Puma's Instagram.",timestamp:"09:13",raw:"Our Twitter, LinkedIn, Instagram and Facebook, plus the main Twitter accounts for Nike and Adidas and Puma's Instagram."},
      {role:"assistant",content:"Updating the picture: six markets in English, French and German, six teams from Marketing to Customer Experience, monitoring your four owned channels plus three competitor accounts. All good before we finish up?",timestamp:"09:14",widgets:[],quickReplies:[],raw:""},
      {role:"user",content:"Looks good.",timestamp:"09:14",raw:"Looks good."},
      {role:"assistant",content:"Would you like any dashboards, reports, or alerts? For example a brand health dashboard, a competitive share-of-voice report, or a crisis alert on negative sentiment spikes?",timestamp:"09:15",widgets:[],quickReplies:[],raw:""},
      {role:"user",content:"Yes please: a brand health dashboard, a competitive share-of-voice report, a weekly executive summary, a crisis alert for negative sentiment, and an alert when a competitor's campaign spikes.",timestamp:"09:15",raw:"Yes please: a brand health dashboard, a competitive share-of-voice report, a weekly executive summary, a crisis alert for negative sentiment, and an alert when a competitor's campaign spikes."},
      {role:"assistant",content:"Almost there. Who needs access to Lumen, and at what level?",timestamp:"09:16",widgets:[],quickReplies:[],raw:""},
      {role:"user",content:"✓ USERS: 5 user(s)",isWidget:true,timestamp:"09:16"},
      {role:"assistant",content:"Here's the full picture for Acme Corp: focused on brand reputation, competitive intelligence and issue tracking, tracking eight topics across your brand, product lines, three competitors, a campaign, an industry theme and customer-service issues, in six markets across three languages, for six teams, with four reports and dashboards, two alerts, and five users. Does that sound right, or is there anything you'd like to adjust?",timestamp:"09:17",widgets:[],quickReplies:["Yes, looks good","I'd like to adjust something"],raw:""},
      {role:"user",content:"Yes, looks good",timestamp:"09:17",raw:"Yes, looks good",isChip:true,chipLabel:"Yes, looks good"},
      {role:"assistant",content:"Wonderful! Thank you for completing your Lumen onboarding. Your setup brief is ready to download below, just click the button to review and download it. A consultant will be in touch shortly to arrange a review session and kick off your implementation. In the meantime, one thing you can get a head start on is generating access tokens for your channels: see [token basics](https://helpcenter.talkwalker.com/s/article/token-basics).",timestamp:"09:18",widgets:[],quickReplies:[],raw:""},
    ];
    const tw = {
      "3-MARKETS":    {submitted:true,data:["United States","United Kingdom","France","Germany","Canada","Australia"]},
      "3-LANGUAGES":  {submitted:true,data:["English","French","German"]},
      "3-OBJECTIVES": {submitted:true,data:{ranked:["Reputation Management","Competitive Intelligence","Issue Tracking"],details:"Most concerned about product complaints escalating quickly on social."}},
      "3-TEAMS":      {submitted:true,data:["Marketing","Communications","PR","Brand","Social Media","Customer Experience"]},
      "3-TIMEZONE":   {submitted:true,data:["GMT / UTC"]},
      "3-USERS":      {submitted:true,data:[
        {firstName:"Jane",lastName:"Smith",email:"jane@acmecorp.com",role:"Marketing Director",access:"Admin"},
        {firstName:"John",lastName:"Doe",email:"john@acmecorp.com",role:"PR Manager",access:"Full Tool"},
        {firstName:"Marie",lastName:"Laurent",email:"marie@acmecorp.com",role:"Comms Lead (FR)",access:"Full Tool"},
        {firstName:"Tom",lastName:"Becker",email:"tom@acmecorp.com",role:"Social Media Manager",access:"Full Tool"},
        {firstName:"Sarah",lastName:"Chen",email:"sarah@acmecorp.com",role:"Insights Analyst",access:"Read-Only"},
      ]},
    };
    const tc = {
      company:  {name:"Acme Corp",email:"jane@acmecorp.com",industry:"Consumer Goods (Footwear & Apparel)",useCase:"Protect brand reputation, track competitors (Nike, Adidas, Puma), and catch customer issues early before they escalate",contact:"Jane Smith"},
      topics: [
        {name:"Acme Corp Brand",keywords:'"Acme Corp" OR "AcmeCorp" OR @AcmeCorp',urls:"https://acmecorp.com",hashtags:"#AcmeCorp",comments:"Primary brand monitoring"},
        {name:"Acme Product Lines",keywords:'"Acme Air" OR "Acme Pro" OR "Acme Classic"',urls:"https://acmecorp.com/products",hashtags:"#AcmeAir",comments:"Product-level tracking"},
        {name:"Nike",keywords:'"Nike" OR @Nike OR "Just Do It"',urls:"https://nike.com",hashtags:"#Nike",comments:"Competitor, client-confirmed"},
        {name:"Adidas",keywords:'"Adidas" OR @adidas',urls:"https://adidas.com",hashtags:"#Adidas",comments:"Competitor, client-confirmed"},
        {name:"Puma",keywords:'"Puma" OR @PUMA',urls:"https://puma.com",hashtags:"#Puma",comments:"Competitor, client-confirmed"},
        {name:"Acme Summer Campaign",keywords:'"Acme Summer" OR #AcmeSummer25',urls:"",hashtags:"#AcmeSummer25",comments:"Campaign tracking, Jun to Aug"},
        {name:"Sustainability in Fashion",keywords:'"sustainable fashion" OR "ethical sourcing" OR "circular fashion"',urls:"",hashtags:"#SustainableFashion",comments:"Industry theme"},
        {name:"Customer Service Issues",keywords:'"Acme" AND (refund OR complaint OR "poor service" OR broken)',urls:"",hashtags:"",comments:"Risk and crisis early warning"},
      ],
      channels: [
        {author:"Acme Corp",type:"Twitter/X",url:"https://twitter.com/acmecorp",owned:"Owned"},
        {author:"Acme Corp",type:"LinkedIn",url:"https://linkedin.com/company/acmecorp",owned:"Owned"},
        {author:"Acme Corp",type:"Instagram",url:"https://instagram.com/acmecorp",owned:"Owned"},
        {author:"Acme Corp",type:"Facebook",url:"https://facebook.com/acmecorp",owned:"Owned"},
        {author:"Nike",type:"Twitter/X",url:"https://twitter.com/nike",owned:"Public"},
        {author:"Adidas",type:"Twitter/X",url:"https://twitter.com/adidas",owned:"Public"},
        {author:"Puma",type:"Instagram",url:"https://instagram.com/puma",owned:"Public"},
      ],
      reports: [
        {name:"Brand Health Dashboard",objective:"Reputation Management",details:"Real-time, all markets, sentiment and volume",comments:"Priority setup"},
        {name:"Competitive Share of Voice",objective:"Competitive Intelligence",details:"Weekly, Acme vs Nike, Adidas, Puma",comments:""},
        {name:"Weekly Executive Summary",objective:"Reputation Management",details:"Weekly email to leadership, top themes and movers",comments:""},
      ],
      alerts:  [
        {name:"Crisis Alert",type:"Sentiment spike",details:"Negative sentiment > 20% in 1 hour",comments:"High priority"},
        {name:"Competitor Campaign Alert",type:"Volume spike",details:"Competitor mention volume up > 50% day over day",comments:""},
      ],
    };
    setStarted(true); setMessages(tm); setWState(tw); setCdata(tc); setSaved(null);
    setProgress({percent:100,collected:{company:true,path:true,topics:true,channels:true,reports:true,users:true}});
    histRef.current = tm.filter(m=>!m.isWidget).map(m=>({role:m.role,content:m.raw||m.content}));
  }, []);

  const MAX_HIST_TURNS = 20;

  const callAPI = useCallback(async (hist, onDelta, sysExtra="") => {
    const trimmed = hist.slice(-MAX_HIST_TURNS);
    apiCountRef.current += 1;
    // The system prompt lives server-side in the chat function; the client only
    // flags whether the OVERSTATE correction pass is needed.
    const res = await fetchWithTimeout(CHAT_ENDPOINT, {
      method:"POST", headers:{"Content-Type":"application/json"},
      body: JSON.stringify({ messages:trimmed, maxTokens:4000, overstateFix: !!sysExtra })
    }, 60000);
    if (!res.ok) throw new Error(`api_${res.status}`);
    const d = await res.json();
    if (d.error) throw new Error("api_error");
    if (d.usage) {
      const u = usageRef.current;
      u.input     += d.usage.input_tokens || 0;
      u.output    += d.usage.output_tokens || 0;
      u.cacheRead += d.usage.cache_read_input_tokens || 0;
      u.cacheWrite+= d.usage.cache_creation_input_tokens || 0;
    }
    return (d.content||[]).map(b=>b.text||"").join("");
  }, []);

  const OVERSTATE_FIX = "\n\nCORRECTION — REWRITE REQUIRED: Your previous reply implied the setup is already live, running, or delivering results. It is NOT — nothing is active until the consultant activates it at the review call. Rewrite your reply keeping all %% markers identical, but change the visible prose to use only future or conditional framing (\"once your consultant activates this, you'll…\", \"this will be set up to…\"). Do not use \"is now set up\", \"you're now getting\", \"will now get\", \"delivered on a schedule\", \"up and running\", or \"you're all set\".";

  const callAPILive = useCallback(async hist => {
    setStreamTxt("");
    let raw = await callAPI(hist, t => setStreamTxt(safeVisible(t)));
    // Fail-safe: every assistant turn must carry a PROGRESS marker, and no marker
    // should be left unterminated. A missing PROGRESS marker or a truncated
    // (dangling) marker is the strongest signal of a malformed or cut-off
    // generation — retry once silently rather than showing the client a derailed
    // reply or dropping the data that was mid-emit when it truncated.
    if (!raw.includes("%%PROGRESS%%") || hasDanglingMarker(raw)) {
      console.warn("malformed reply (missing PROGRESS or truncated marker) — retrying once");
      setStreamTxt("");
      raw = await callAPI(hist, t => setStreamTxt(safeVisible(t)));
    }
    // Expectation guard: never show the client language implying the setup is
    // already live. Unlike a blind retry, this re-runs WITH an explicit corrective
    // so the rewrite actually differs, then accepts the result.
    if (overstatesCompletion(stripAll(raw))) {
      console.warn("overstated completion detected — retrying with corrective");
      setStreamTxt("");
      raw = await callAPI(hist, t => setStreamTxt(safeVisible(t)), OVERSTATE_FIX);
    }
    return raw;
  }, [callAPI]);

  // DEV/simulate only — unreachable in the live build (DEV=false hides the menu).
  // The live proxy rejects client-supplied system prompts by design, so if this
  // is ever re-enabled it must go through a separate dev-only endpoint.
  const callClientAPI = useCallback(async (systemPrompt, hist) => {
    throw new Error("simulate_unavailable_in_live_build");
  }, []);

  const inferPct = useCallback(() => {
    const sub = t => Object.entries(wRef.current).some(([k,v])=>k.endsWith(`-${t}`)&&(v===true||v?.submitted));
    if (sub("USERS")) return 80;
    if (["MARKETS","OBJECTIVES","TEAMS"].every(sub)) return 60;
    if (sub("TOPICS")) return 40;
    if (sub("PATH")) return 15;
    return 0;
  }, []);

  const applyCdata = useCallback(pr => {
    setCdata(p => mergeCdata(p, pr));
  }, []);

  const sendToAPI = useCallback(async (rawTxt, isRetry=false) => {
    // Strip any injected marker delimiters from client input before it reaches
    // the model (see sanitizeIn). Covers typed messages and widget payloads.
    const txt = sanitizeIn(rawTxt);
    // Ensure exactly one trailing user turn for this call. On a fresh send we
    // push it; on retry it's still there from the prior attempt. Either way, if
    // the call fails we pop it back off so a subsequent message can't leave two
    // user turns in a row (which the API rejects, bricking the session).
    const last = histRef.current[histRef.current.length-1];
    const alreadyQueued = isRetry && last?.role==="user" && last?.content===txt;
    if (!alreadyQueued) histRef.current.push({role:"user",content:txt});
    setRetryMsg(null);
    setLoading(true);
    try {
      const t0 = Date.now(), raw = await callAPILive(histRef.current), el = Date.now()-t0;
      if (el < MIN_MS) await sleep(MIN_MS-el);
      const pr = parseReply(raw);
      const { clean,widgets,topicSuggestions,quickReplies,progress:prog } = pr;
      // Dead-reply guard: after callAPILive's retries a reply can still come back
      // malformed (e.g. truncated on a large import), which strips to nothing and
      // carries no widget/chips — an empty bubble that leaves the flow with nothing
      // to click. Treat that as a failure and offer retry rather than hang.
      const actionable = clean.trim() || widgets.length || topicSuggestions.length || quickReplies.length;
      if (!actionable) throw new Error("empty_reply");
      if (prog) setProgress(prog);
      else setProgress(p=>({...p,percent:Math.max(p.percent,inferPct())}));
      applyCdata(pr);
      histRef.current.push({role:"assistant",content:stripThoughtForHistory(raw)});
      if (sndRef.current) pop();
      const dv = maybeDivider(prog);
      setMessages(p=>[...p,...(dv?[dv]:[]),{role:"assistant",content:clean,widgets,topicSuggestions,quickReplies,timestamp:gts(),raw}]);
    } catch(e) {
      if (histRef.current[histRef.current.length-1]?.role==="user") histRef.current.pop();
      setRetryMsg(txt);
    } finally {
      setLoading(false);
    }
  }, [callAPI, pop, inferPct, applyCdata]);

  const handleSend = useCallback(async (merged, users) => {
    // Double-send guard: `sending` is React state and lags a fast double-click,
    // so a ref (updates synchronously) is what actually prevents two records,
    // two Sheets, or two Slack alerts from one impatient double-tap.
    if (sendingRef.current) return;
    sendingRef.current = true;
    if (saveT.current) clearTimeout(saveT.current); // cancel any pending in-progress autosave so it can't land after the completed record
    setSending(true); setSendErr(null);
    try {
      const XLSX = await loadXLSX();
      const { wb, filename } = buildWorkbook(XLSX, merged, users || []);
      const sentAt = new Date();

      // Generate the editable Google Sheet from the brief's workbook. Best-effort:
      // if Sheets isn't configured (501) or the call fails, the brief still sends;
      // the client just doesn't get a Sheet link. Never blocks the confirmation.
      let sheetUrl = null;
      try {
        const xlsxBase64 = XLSX.write(wb, { type: "base64", bookType: "xlsx" });
        const sres = await fetchWithTimeout(SHEET_ENDPOINT, {
          method: "POST", headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ sessionId: sidRef.current, xlsxBase64, brief: { ...merged, company: { ...merged.company, onboardingLanguage: uiLang }, users: users || [] }, filename, clientEmail: merged.company?.email || "", company: merged.company?.name || "", contactName: merged.company?.contact || "", topicsCount: (merged.topics || []).length, usersCount: (users || []).length }),
        }, 45000);
        if (sres.ok) { const sd = await sres.json().catch(() => ({})); sheetUrl = sd.url || null; }
      } catch (e) { console.error("Sheet generation failed (non-fatal)", e); }
      setSheetLink(sheetUrl);

      const record = {
        id: sidRef.current,
        merged, users: users || [],
        handoff: cdata.handoff || {
          maturity: "",
          goalInOwnWords: merged.company?.useCase || "",
          hesitations: "",
          aiSuggestedUnconfirmed: "",
          followUps: "Session sent before the assistant produced a full handoff — review the brief directly and confirm the gaps at the call.",
          consultantTips: "",
        },
        queries: merged.queries || "",
        seed: seed || null,
        seedId: seedId || null,
        sheetUrl,
        durationMs: startedAtRef.current ? (Date.now() - startedAtRef.current) : null,
        apiCalls: apiCountRef.current,
        tokens: { ...usageRef.current },
        status: "completed",
        sentAt: sentAt.toISOString(),
      };

      let saveOk = false;
      try {
        const res = await fetchWithTimeout(SESSION_ENDPOINT, {
          method:"POST", headers:{"Content-Type":"application/json"},
          body: JSON.stringify({ session: record })
        }, 15000);
        if (!res.ok) throw new Error(`save_${res.status}`);
        saveOk = true;
      } catch (e) { console.error("Session save failed", e); }

      // "Delivered" = it reached Proserv by at least one durable channel: the
      // session store (dashboard) OR the Sheet (which also fires the Slack alert
      // and drops the file in the Proserv folder). If BOTH failed, don't show a
      // false "sent" — keep the draft and the modal open so the client can retry
      // instead of walking away thinking it went through.
      if (!saveOk && !sheetUrl) {
        setSendErr("send-failed");
        return; // the `finally` below still re-enables the Send button
      }

      onBriefSent?.({ ...record, filename, sentAt });
      setSent(true); setShowExport(false);
      // Clear the resume draft only once the record is safely stored; if the save
      // failed but the Sheet carried it through, keep the draft so the session can
      // still be re-sent later to populate the dashboard.
      if (saveOk) lsClearDraft(seedId);
      // Bring the "Brief sent" confirmation into view — without this the modal just
      // closes and the client is left looking at empty scroll space (reads as a blank
      // screen / no confirmation).
      requestAnimationFrame(() => { if (msgRef.current) msgRef.current.scrollTop = msgRef.current.scrollHeight; });
      if (sndRef.current) chime();
    } catch (e) {
      // Catastrophic (e.g. buildWorkbook / XLSX threw before anything was sent).
      // Never leave the client on a stuck spinner or a false success.
      console.error("Send failed", e);
      setSendErr("send-failed");
    } finally {
      setSending(false);
      sendingRef.current = false;
    }
  }, [chime, cdata, onBriefSent, seed, seedId, uiLang]);

  const maybeDivider = useCallback(prog => {
    const sec = prog?.section;
    if (!sec) return null;
    const prev = prevSecRef.current;
    prevSecRef.current = sec;
    if (!prev || prev === sec) return null;
    const pi = SECTION_KEYS.indexOf(prev), ni = SECTION_KEYS.indexOf(sec);
    if (pi === -1 || ni === -1 || ni <= pi) return null;
    const remaining = SECTION_KEYS.length - ni;
    return { role:"divider", label:`${SECTION_LABELS[prev]} — done`, sub: remaining>0?`${remaining} section${remaining!==1?"s":""} to go`:"", timestamp:gts() };
  }, []);

  const widgetSum = (type, data) =>
    type==="OBJECTIVES" ? (()=>{ const n=normObjectives(data); return fmtRanked(n)+(n.details?` — ${n.details}`:""); })()
    : ["MARKETS","LANGUAGES","TEAMS","TIMEZONE"].includes(type) ? (Array.isArray(data)?data.join(", "):data)
    : type==="USERS"   ? `${data.length} user(s)`
    : type==="QUERIES" ? (data==="__skip__"?"Skipped":"Submitted")
    : type==="TOPICS"  ? `${data.length} topics confirmed`
    : data==="recommendations" ? "Recommendations path" : "Guided path";

  // What the MODEL receives. The visible chip stays short ("Submitted"), but the
  // model needs the actual content — for QUERIES the pasted/imported text itself,
  // for USERS the real names/roles. Sending only the summary meant the model
  // never saw the queries at all.
  const widgetApiPayload = (type, data) =>
    type==="QUERIES" && data!=="__skip__" ? `Full pasted/imported content below — extract what's relevant per your IMPORTED CONTENT instructions:\n${data}`
    : type==="USERS" && Array.isArray(data) ? data.map(u=>`${u.firstName} ${u.lastName} <${u.email}> — ${u.role||"no role"} — ${u.access}`).join("; ")
    : widgetSum(type, data);

  const runSmokeTest = useCallback(async personaKey => {
    const per = SMOKE_PERSONAS.find(p=>p.key===personaKey);
    if (!per) return;
    init(); smokeStop.current = false;
    resetSession(); setStarted(true); setSmoke({label:per.label,emoji:per.emoji,turn:0});
    // Match a real seeded session: pre-fill the captured company data and build the
    // simulate persona + USERS auto-fill from the seed and the chosen language.
    if (seed) setCdata(p=>({...p, company:{name:seed.company||"", email:seed.email||"", industry:seed.industry||"", useCase:"", contact:seed.contactName||""}}));
    const smokePrompts = buildSmokePrompts(seed, uiLang);
    const smokeWidgetData = buildSmokeWidgetData(seed);
    const hist = [{role:"user",content:seededOpener(seed, uiLang)}];
    const local = [];
    const pushMsg = m => { local.push(m); setMessages([...local]); };
    let pct = 0;
    try {
      for (let turn=0; turn<SMOKE_MAX_TURNS && !smokeStop.current; turn++) {
        setSmoke({label:per.label,emoji:per.emoji,turn:turn+1});
        setLoading(true);
        const raw = await callAPI(hist);
        setLoading(false);
        const pr = parseReply(raw);
        hist.push({role:"assistant",content:raw});
        if (pr.progress) { setProgress(pr.progress); pct = pr.progress.percent; }
        applyCdata(pr);
        const aIdx = local.length;
        pushMsg({role:"assistant",content:pr.clean,widgets:pr.widgets,topicSuggestions:pr.topicSuggestions,quickReplies:pr.quickReplies,timestamp:gts(),raw});
        if (pct >= 100 || smokeStop.current) break;
        await sleep(300);
        if (pr.widgets.length > 0) {
          const w = pr.widgets[0];
          const data = w==="TOPICS" ? pr.topicSuggestions : smokeWidgetData[w];
          if (data===undefined || (w==="TOPICS"&&(!data||!data.length))) {
            setWState(p=>({...p,[`${aIdx}-${w}`]:{submitted:true,data:"__skip__"}}));
            pushMsg({role:"user",content:`Skipped ${w}`,isWidget:true,timestamp:gts()});
            hist.push({role:"user",content:`[Widget skipped — ${w}]`});
          } else if (data==="__skip__") {
            setWState(p=>({...p,[`${aIdx}-${w}`]:{submitted:true,data:"__skip__"}}));
            pushMsg({role:"user",content:`Skipped ${w}`,isWidget:true,timestamp:gts()});
            hist.push({role:"user",content:`[Widget skipped — ${w}]`});
          } else {
            const sum = widgetSum(w, data);
            setWState(p=>({...p,[`${aIdx}-${w}`]:{submitted:true,data}}));
            pushMsg({role:"user",content:`✓ ${w}: ${sum}`,isWidget:true,timestamp:gts()});
            hist.push({role:"user",content:`[Widget submitted — ${w}]: ${widgetApiPayload(w, data)}`});
          }
        } else {
          setLoading(true);
          const reply = await callClientAPI(smokePrompts[per.key], hist);
          setLoading(false);
          if (!reply) break;
          pushMsg({role:"user",content:reply,timestamp:gts(),raw:reply});
          hist.push({role:"user",content:reply});
        }
      }
    } catch(e) {
      pushMsg({role:"assistant",content:"⚠ Smoke test stopped: API error. You can take over manually from here.",widgets:[],quickReplies:[],timestamp:gts(),raw:""});
    } finally {
      histRef.current = hist;
      setLoading(false);
      setSmoke(null);
    }
  }, [callAPI, callClientAPI, init, resetSession, applyCdata, seed, uiLang]);

  const doRewind = useCallback(idx => {
    const nm = messages.slice(0,idx+1); setMessages(nm);
    histRef.current = nm.filter(m=>!m.isWidget).map(m=>({role:m.role,content:m.raw||m.content}));
    let rp = {percent:0,collected:{}};
    for (let i=nm.length-1;i>=0;i--) { if (nm[i].role==="assistant"&&nm[i].raw) { const p=pProg(nm[i].raw); if(p){rp=p;break;} } }
    setProgress(rp);
    // Rebuild cdata from the surviving replies so data captured after the rewind
    // point (topics, channels, company fields) doesn't linger into the brief.
    let cd = emptyCdata();
    for (const m of nm) { if (m.role==="assistant"&&m.raw) cd = mergeCdata(cd, parseReply(m.raw)); }
    setCdata(cd);
    setWState(p=>{const n={...p};Object.keys(n).forEach(k=>{if(parseInt(k.split("-")[0])>idx)delete n[k];});return n;});
    // Reset section-divider and completion-chime trackers, else the next turn
    // fires a spurious divider or re-chimes.
    prevSecRef.current = rp?.section || null;
    prevPct.current = rp?.percent || 0;
    setRewind(null);
  }, [messages]);

  const startConvo = useCallback(async () => {
    init();
    const sd = seed, keepSid = sidRef.current;
    resetSession(); setStarted(true); setLoading(true);
    startedAtRef.current = Date.now(); apiCountRef.current = 0;
    if (sd && keepSid) {
      sidRef.current = keepSid;
      setCdata(p=>({...p, company:{name:sd.company||"", email:sd.email||"", industry:sd.industry||"", useCase:"", contact:sd.contactName||""}}));
    }
    if (msgRef.current) msgRef.current.scrollTop = 0;
    const ini = { role:"user", content: sanitizeIn(seededOpener(sd, uiLang)) };
    histRef.current = [ini];
    setInitErr(null);
    try {
      const raw = await callAPILive([ini]);
      const { clean,widgets,topicSuggestions,quickReplies,progress:prog } = parseReply(raw);
      if (prog) setProgress(prog);
      histRef.current.push({role:"assistant",content:stripThoughtForHistory(raw)});
      prevSecRef.current = prog?.section || "company";
      setMessages([{role:"assistant",content:clean,widgets,topicSuggestions,quickReplies,timestamp:gts(),raw}]);
    } catch (e) {
      // Without this, a failed first turn left a permanent "Assistant is thinking…"
      // spinner with no way out. Clear it and offer a retry instead.
      console.error("startConvo failed", e);
      setInitErr("start");
    } finally {
      setLoading(false);
    }
  }, [callAPI, init, resetSession, seed, uiLang]);

  const resumeConvo = useCallback(async () => {
    init(); if (!saved) return;
    setStarted(true); setLoading(true); setInitErr(null);
    if (saved.uiLang) setUiLang(saved.uiLang);
    if (saved.sid) sidRef.current = saved.sid;
    startedAtRef.current = saved.startedAt || Date.now();
    setMessages(saved.messages); setProgress(saved.progress); setWState(saved.wState||{});
    prevSecRef.current = saved.progress?.section || null;
    if (saved.cdata) setCdata(saved.cdata);
    // Clone the saved history so a failed attempt + retry can't stack two
    // "[RESUMING SESSION]" markers onto the same array reference.
    histRef.current = [...(saved.history||[]), {role:"user",content:"[RESUMING SESSION] The client is returning to continue their onboarding."}];
    try {
      const raw = await callAPILive(histRef.current);
      const { clean,widgets,topicSuggestions,quickReplies,progress:prog } = parseReply(raw);
      if (prog) setProgress(prog);
      histRef.current.push({role:"assistant",content:stripThoughtForHistory(raw)});
      if (sndRef.current) pop();
      const dv = maybeDivider(prog);
      setMessages(p=>[...p,...(dv?[dv]:[]),{role:"assistant",content:clean,widgets,topicSuggestions,quickReplies,timestamp:gts(),raw}]);
      setSaved(null); // only clear the resume draft once we've actually continued
    } catch (e) {
      // Keep `saved` so the retry can re-resume; clear the spinner and surface a retry.
      console.error("resumeConvo failed", e);
      setInitErr("resume");
    } finally {
      setLoading(false);
    }
  }, [saved, callAPI, init, pop]);

  const sendMsg = useCallback(async (ov, chip) => {
    init();
    const txt = ov!==undefined ? ov.trim() : input.trim();
    if (!txt||loading) return;
    setInput(""); if (taRef.current) taRef.current.style.height = "auto";
    setMessages(p=>[...p,{role:"user",content:txt,timestamp:gts(),raw:txt,isChip:!!chip,chipLabel:chip}]);
    await sendToAPI(txt);
  }, [input, loading, sendToAPI, init]);

  const onWSubmit = useCallback((mi, type, data) => {
    const key = `${mi}-${type}`;
    const isUp = !!wRef.current[key];
    const sum = widgetSum(type, data);
    // State updater stays pure; the message + API call happen here, once.
    setWState(prev => ({...prev,[key]:{submitted:true,data}}));
    setMessages(m=>[...m,{role:"user",content:`${isUp?"✎ Updated":"✓"} ${type}: ${sum}`,isWidget:true,timestamp:gts()}]);
    sendToAPI(`[Widget ${isUp?"updated":"submitted"} — ${type}]: ${widgetApiPayload(type, data)}`);
  }, [sendToAPI]);

  const onWSkip = useCallback((mi, type) => {
    const key = `${mi}-${type}`;
    setWState(p=>({...p,[key]:{submitted:true,data:"__skip__"}}));
    setMessages(m=>[...m,{role:"user",content:`Skipped ${type}`,isWidget:true,timestamp:gts()}]);
    sendToAPI(`[Widget skipped — ${type}]`);
  }, [sendToAPI]);

  const renderWidget = useCallback((type, mi, topicSuggestions) => {
    const key = `${mi}-${type}`;
    if (Object.entries(wState).some(([k,v])=>k!==key&&k.endsWith(`-${type}`)&&(v===true||v?.submitted))) return null;
    const ws = wState[key], sub = ws===true||ws?.submitted===true;
    if (sub) return <div style={{padding:"12px 16px",background:C.hi,borderRadius:10,border:`1px solid ${C.border}`,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
      <div style={{fontSize:13,color:C.text,fontWeight:600}}>✓ {ws?.data==="__skip__"?"Skipped":"Submitted"}</div>
      <button onClick={()=>setWState(p=>({...p,[key]:{...p[key],submitted:false}}))} style={{background:"transparent",border:`1px solid ${LINK}`,color:LINK,borderRadius:6,padding:"6px 14px",fontSize:12,cursor:"pointer"}}>Edit</button>
    </div>;
    const pd = ws?.data, os = d=>onWSubmit(mi,type,d), sk = ()=>onWSkip(mi,type);
    const userPrefill = pd || (cdata.company?.email ? [{
      firstName:(cdata.company.contact||"").split(" ")[0]||"",
      lastName:(cdata.company.contact||"").split(" ").slice(1).join(" "),
      email:cdata.company.email, role:"", access:"Admin"
    }] : []);
    const WHY = { MARKETS:WL("whyMarkets",uiLang), TEAMS:WL("whyTeams",uiLang), USERS:WL("whyUsers",uiLang), QUERIES:WL("whyQueries",uiLang), TOPICS:WL("whyTopics",uiLang) };
    return <div>
      {WHY[type] && <div style={{fontSize:11,color:C.muted,margin:"0 0 6px",fontStyle:"italic"}}>{WHY[type]}</div>}
      {type==="PATH"      && <PathChoice onSubmit={d=>onWSubmit(mi,type,d)}/>}
      {type==="QUERIES"   && <QueriesWidget onSubmit={os} initialData={pd} lang={uiLang}/>}
      {type==="TOPICS"    && topicSuggestions?.length>0 && <TopicCards suggestions={topicSuggestions} onConfirm={os} onSkip={sk} lang={uiLang}/>}
      {type==="MARKETS"   && <ChipSelector options={MARKETS_OPT}  onSubmit={os} onSkip={sk} placeholder={WL("phMarket",uiLang)}   hint={WL("hintSelectAll",uiLang)}    initialData={pd||[]} lang={uiLang}/>}
      {type==="LANGUAGES" && <ChipSelector options={LANG_OPT}     onSubmit={os} onSkip={sk} placeholder={WL("phLanguage",uiLang)} hint={WL("hintSelectAll",uiLang)}  initialData={pd||[]} lang={uiLang}/>}
      {type==="OBJECTIVES"&& <RankedSelector options={OBJ_OPT}   onSubmit={os} onSkip={sk} max={WIDGET_MAX.OBJECTIVES}    hint={WL("hintObjectives",uiLang)} initialData={pd} lang={uiLang}/>}
      {type==="TEAMS"     && <ChipSelector options={TEAM_OPT}     onSubmit={os} onSkip={sk} placeholder={WL("phTeam",uiLang)}     hint={WL("hintTeams",uiLang)}      initialData={pd||[]} lang={uiLang}/>}
      {type==="TIMEZONE"  && <ChipSelector options={TZ_OPT}       onSubmit={os} onSkip={sk} max={WIDGET_MAX.TIMEZONE}      hint={WL("hintTimezone",uiLang)}   initialData={pd||[]} lang={uiLang}/>}
      {type==="USERS"     && <UserForm onSubmit={os} onSkip={sk} initialData={userPrefill} lang={uiLang}/>}
    </div>;
  }, [wState, onWSubmit, onWSkip, C, cdata, uiLang]);

  if (!checked) return <div style={{height:"100vh",display:"flex",alignItems:"center",justifyContent:"center"}}>Loading…</div>;

  const SHOW = 6, canCollapse = messages.length>SHOW, vStart = canCollapse&&collapsed ? messages.length-SHOW : 0;
  const last = messages[messages.length-1], showQR = last?.role==="assistant"&&last?.quickReplies?.length>0&&!loading;
  const done = progress.percent === 100;

  const gwp = type => { const es=Object.entries(wState).filter(([k,v])=>k.endsWith(`-${type}`)&&(v===true||v?.submitted)).sort((a,b)=>(parseInt(a[0])||0)-(parseInt(b[0])||0)); return es.length?es[es.length-1][1].data:null; };
  const fmtV = v => { if (v==null||v===""||(Array.isArray(v)&&!v.length)) return null; if (v==="__skip__") return "Skipped"; return Array.isArray(v)?v.join(", "):String(v); };
  const topicsList = (cdata.topics?.length?cdata.topics:Array.isArray(gwp("TOPICS"))?gwp("TOPICS"):[]);
  const usersList  = Array.isArray(gwp("USERS"))?gwp("USERS"):[];
  const sideCol = ww >= 1280;
  const panelRows = [
    ["Company", fmtV(cdata.company?.name)],
    ["Email", fmtV(cdata.company?.email)],
    ["Industry", fmtV(cdata.company?.industry)],
    ["Goal / use case", fmtV(cdata.company?.useCase)],
    ["Markets", fmtV(gwp("MARKETS"))],
    ["Languages", fmtV(gwp("LANGUAGES"))],
    ["Objectives", gwp("OBJECTIVES")==="__skip__" ? "Skipped" : (fmtRanked(gwp("OBJECTIVES")) || fmtV(cdata.company?.objectives))],
    ["Teams", fmtV(gwp("TEAMS"))],
    ["Timezone", fmtV(gwp("TIMEZONE"))],
    ["Topics", topicsList.length?`${topicsList.length}: `+topicsList.map(t=>t.name).filter(Boolean).join(", "):null],
    ["Channels", cdata.channels?.length?`${cdata.channels.length}: `+cdata.channels.map(c=>c.author).filter(Boolean).join(", "):null],
    ["Reports", cdata.reports?.length?cdata.reports.map(r=>r.name).filter(Boolean).join(", "):null],
    ["Alerts", cdata.alerts?.length?cdata.alerts.map(a=>a.name).filter(Boolean).join(", "):null],
    ["Users", usersList.length?usersList.map(u=>`${u.firstName} (${u.access})`).join(", "):null],
  ];

  return (
    <div dir={uiLang==="Arabic"?"rtl":"ltr"} style={{fontFamily:"'Inter', Arial, sans-serif",height:"100%",background:C.bg,display:"flex",flexDirection:"column",color:C.text,overflow:"hidden"}}>
      <style>{`@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700&display=swap');@keyframes slideUpFade{from{opacity:0;transform:translateY(8px)}to{opacity:1;transform:translateY(0)}}@keyframes bounce{0%,80%,100%{transform:translateY(0);opacity:.5}40%{transform:translateY(-4px);opacity:1}}@keyframes spin{to{transform:rotate(360deg)}}*{box-sizing:border-box}button:focus-visible,a:focus-visible,input:focus-visible,textarea:focus-visible{outline:2px solid #6D28D9 !important;outline-offset:2px !important}@media (prefers-reduced-motion: reduce){*{animation:none !important;transition:none !important}}`}</style>

      {smoke && <div style={{position:"fixed",bottom:140,left:"50%",transform:"translateX(-50%)",zIndex:300,background:"#1e293b",color:"white",borderRadius:10,padding:"10px 16px",fontSize:13,display:"flex",gap:12,alignItems:"center",boxShadow:"0 4px 16px rgba(0,0,0,0.25)",whiteSpace:"nowrap"}}>
        <span>🧪 Smoke test: {smoke.emoji} <strong>{smoke.label}</strong> · turn {smoke.turn}/{SMOKE_MAX_TURNS}</span>
        <button onClick={()=>{smokeStop.current=true;}} style={{background:"#ef4444",border:"none",color:"white",borderRadius:6,padding:"4px 10px",fontSize:12,cursor:"pointer"}}>Stop</button>
      </div>}

      {rewind !== null && <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.4)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:1000}}>
        <div style={{background:"white",borderRadius:14,padding:"24px 28px",maxWidth:360,width:"90%",boxShadow:"0 8px 32px rgba(0,0,0,0.18)"}}>
          <div style={{fontSize:20,marginBottom:8}}>↺</div>
          <div style={{fontWeight:700,fontSize:15,marginBottom:8,color:"#1e293b"}}>Rewind conversation?</div>
          <div style={{fontSize:13,color:"#64748b",marginBottom:20,lineHeight:1.6}}>All messages after this point will be <strong>permanently deleted</strong>.</div>
          <div style={{display:"flex",gap:10}}>
            <button onClick={()=>setRewind(null)} style={{flex:1,background:"transparent",border:"1px solid #e2e8f0",borderRadius:8,padding:"9px",fontSize:13,color:"#64748b",cursor:"pointer"}}>Cancel</button>
            <button onClick={()=>doRewind(rewind)} style={{flex:1,background:"#ef4444",border:"none",borderRadius:8,padding:"9px",fontSize:13,color:"white",fontWeight:600,cursor:"pointer"}}>Yes, rewind</button>
          </div>
        </div>
      </div>}

      {showExport && <ModalBoundary onClose={()=>setShowExport(false)}><ExportModal cdata={cdata} wState={wState||{}} messages={messages} onClose={()=>setShowExport(false)} onExport={(merged,users)=>{doExport(merged,users,messages);}} onSend={handleSend} sending={sending} sendErr={sendErr} sent={sent} sheetLink={sheetLink}/></ModalBoundary>}

      {showPanel && started && <div style={{position:"fixed",top:56,right:0,bottom:0,width:mob?"100%":320,background:C.card,borderLeft:`1px solid ${C.border}`,zIndex:500,overflowY:"auto",padding:"16px 18px",boxShadow:sideCol?"none":"-4px 0 16px rgba(0,0,0,0.08)"}}>
        <div style={{fontSize:11,color:C.muted,margin:"0 0 12px",lineHeight:1.5,background:C.hi,borderRadius:8,padding:"8px 10px"}}>{L("correctionHint", uiLang)}</div>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14}}>
          <div style={{fontWeight:700,fontSize:14,color:C.text}}>📋 Captured so far</div>
          <button onClick={()=>setShowPanel(false)} style={{background:"transparent",border:"none",fontSize:16,cursor:"pointer",color:C.muted}}>✕</button>
        </div>
        {panelRows.map(([label,val]) => <div key={label} style={{marginBottom:12}}>
          <div style={{fontSize:10,fontWeight:700,color:C.muted,textTransform:"uppercase",letterSpacing:"0.05em",marginBottom:3}}>{label}</div>
          {val
            ? <div style={{fontSize:12,color:C.text,lineHeight:1.5}}>{val}</div>
            : <div style={{fontSize:12,color:C.muted,fontStyle:"italic",opacity:0.7}}>Not captured yet</div>}
        </div>)}
      </div>}

      {/* Header */}
      <div style={{background:C.card,borderBottom:`1px solid ${C.border}`,padding:mob?"8px 12px":"0 24px",display:"flex",alignItems:"center",justifyContent:"space-between",minHeight:56,height:mob?"auto":56,flexWrap:mob?"wrap":"nowrap",gap:mob?6:0,flexShrink:0}}>
        <div style={{display:"flex",alignItems:"center",gap:10}}>
          <LumenIcon size={28}/>
          <div>
            <div style={{display:"flex",alignItems:"center",gap:8}}>
              <span style={{display:"inline-flex",flexDirection:"column",lineHeight:1.05}}>
                <span style={{fontWeight:800,fontSize:16,color:A,letterSpacing:"-0.01em"}}>Lumen</span>
                <span style={{fontWeight:700,fontSize:8,color:dark?"#8fa8d8":NAVY,letterSpacing:"0.02em"}}>by Talkwalker</span>
              </span>
              <span style={{color:C.muted,fontSize:12,paddingLeft:2}}>Onboarding Assistant</span>
              {DEV && <button onClick={loadTestSession} title="Loads a completed session so you can preview the export" style={{background:"transparent",border:"1px solid #d97706",color:"#92400e",fontSize:10,fontWeight:600,padding:"2px 8px",borderRadius:4,cursor:"pointer",marginLeft:2}}>▶ Preview completed demo</button>}
              {DEV && <div style={{position:"relative",display:"inline-block"}}>
                <button onClick={()=>setShowSmokeMenu(s=>!s)} disabled={!!smoke} title="Auto-runs a full conversation against a simulated client persona to regression-test the flow" style={{background:"transparent",border:"1px solid #d97706",color:"#92400e",fontSize:10,fontWeight:600,padding:"2px 8px",borderRadius:4,cursor:smoke?"default":"pointer",marginLeft:2,opacity:smoke?0.5:1}}>Simulate conversation</button>
                {showSmokeMenu && <div style={{position:"absolute",top:"115%",left:0,background:"white",border:"1px solid #e2e8f0",borderRadius:8,boxShadow:"0 8px 24px rgba(0,0,0,0.12)",zIndex:600,minWidth:190,overflow:"hidden"}}>
                  {SMOKE_PERSONAS.map(p => <button key={p.key} onClick={()=>{setShowSmokeMenu(false);runSmokeTest(p.key);}} style={{display:"block",width:"100%",textAlign:"left",padding:"8px 12px",background:"transparent",border:"none",cursor:"pointer",fontSize:12,color:"#1e293b"}}>{p.emoji} {p.label}</button>)}
                </div>}
              </div>}
            </div>
            <div style={{fontSize:11,color:C.muted,marginTop:1}}>
              <>Your answers go to your Lumen onboarding team</>
            </div>
          </div>
        </div>
        <div style={{display:"flex",gap:8,alignItems:"center"}}>
          {started && <button onClick={()=>setShowPanel(s=>!s)} aria-label={showPanel?"Hide captured answers":"Show captured answers"} aria-pressed={showPanel} title="Show what's been captured so far" style={{background:showPanel?A:C.card,border:`1px solid ${showPanel?A:C.border}`,borderRadius:"50%",width:32,height:32,cursor:"pointer",color:showPanel?"white":C.muted,display:"inline-flex",alignItems:"center",justifyContent:"center"}}><Ic d={IC.panel}/></button>}
          <button onClick={()=>{init();setSound(s=>!s);}} aria-label={sound?"Turn sound off":"Turn sound on"} style={{background:C.card,border:`1px solid ${C.border}`,borderRadius:"50%",width:32,height:32,cursor:"pointer",color:C.muted,display:"inline-flex",alignItems:"center",justifyContent:"center"}}><Ic d={sound?IC.sound:IC.mute}/></button>
          <button onClick={()=>setTheme(th=>th==="dark"?"light":"dark")} aria-label={dark?"Switch to light mode":"Switch to dark mode"} style={{background:C.card,border:`1px solid ${C.border}`,borderRadius:"50%",width:32,height:32,cursor:"pointer",color:C.muted,display:"inline-flex",alignItems:"center",justifyContent:"center"}}><Ic d={dark?IC.sun:IC.moon}/></button>
        </div>
      </div>

      {/* Stepper */}
      {started && <div style={{background:C.card,borderBottom:`1px solid ${C.border}`,padding:"14px 24px",flexShrink:0}}>
        <div style={{maxWidth:640,margin:"0 auto",display:"flex",alignItems:"flex-end",gap:16}}>
          <div style={{flex:1}}><Stepper progress={progress} dark={dark} compact={mob}/></div>
          {!mob && !sent && draftOk && <div style={{fontSize:11,color:C.muted,whiteSpace:"nowrap",paddingBottom:2}}>✓ Saved on this device</div>}
        </div>
      </div>}

      <div aria-live="polite" style={{position:"absolute",width:1,height:1,overflow:"hidden",clip:"rect(0 0 0 0)",whiteSpace:"nowrap"}}>
        {(messages.filter(m=>m.role==="assistant").slice(-1)[0]?.content)||""}
      </div>
      {/* Messages */}
      <div ref={msgRef} style={{flex:1,overflowY:"auto",padding:"24px 16px",maxWidth:760,width:"100%",margin:"0 auto",alignSelf:"center",transform:sideCol&&showPanel&&started?"translateX(-160px)":"none",transition:"transform 0.25s ease"}}>

        {!started && !saved && (
          <div style={{display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",minHeight:"100%",padding:"40px 24px",textAlign:"center",position:"relative",overflow:"hidden"}}>
            <svg aria-hidden="true" viewBox="0 0 900 240" preserveAspectRatio="none" style={{position:"absolute",top:0,left:0,width:"100%",height:220,pointerEvents:"none"}}>
              <defs><linearGradient id="lw" x1="0" y1="0" x2="1" y2="0"><stop offset="0" stopColor="#7C3AED" stopOpacity="0"/><stop offset="0.5" stopColor="#7C3AED" stopOpacity="0.16"/><stop offset="1" stopColor="#7C3AED" stopOpacity="0"/></linearGradient></defs>
              <path d="M0,150 C180,60 320,220 480,130 C640,40 760,180 900,90 L900,0 L0,0 Z" fill="url(#lw)"/>
              <path d="M0,190 C220,110 380,240 560,150 C720,70 820,200 900,140" fill="none" stroke="#7C3AED" strokeOpacity="0.18" strokeWidth="2"/>
            </svg>
            <div style={{width:88,height:88,borderRadius:24,background:A,display:"flex",alignItems:"center",justifyContent:"center",boxShadow:"0 8px 24px rgba(126,72,236,0.25)"}}><LumenIcon size={52} inverse/></div>
            <h2 style={{margin:"22px 0 8px",color:C.text,fontSize:22,fontWeight:700}}>{seed?L("welcomeTitleSeeded",uiLang,{name:seed.contactName?.split(" ")[0]||seed.company}):L("welcomeTitle",uiLang)}</h2>
            <p style={{color:C.muted,fontSize:14,margin:"0 0 18px",maxWidth:420,lineHeight:1.6}}>{seed?L("welcomeSubSeeded",uiLang,{company:seed.company}):L("welcomeSub",uiLang)}</p>
            <div style={{display:"flex",flexWrap:"wrap",gap:6,justifyContent:"center",alignItems:"center",margin:"0 0 28px"}}>
              <span aria-hidden="true" style={{fontSize:14,marginInlineEnd:2}}>🌐</span>
              {UI_LANGS.map(l => { const on = uiLang===l.code; return (
                <button key={l.code} onClick={()=>setUiLang(l.code)} aria-pressed={on} style={{padding:"5px 12px",borderRadius:20,fontSize:12,cursor:"pointer",border:"1px solid",background:on?A:"transparent",borderColor:on?A:C.border,color:on?"white":C.muted,fontWeight:on?700:500,transition:"all 0.15s"}}>{l.native}</button>
              ); })}
            </div>
            <div style={{width:"100%",maxWidth:480,margin:"0 auto 32px",textAlign:uiLang==="Arabic"?"right":"left"}}>
              {[[L("step1Title",uiLang),L("step1Desc",uiLang)],
                [L("step2Title",uiLang),L("step2Desc",uiLang)],
                [L("step3Title",uiLang),L("step3Desc",uiLang)]].map(([t,d],i) => (
                <div key={i} style={{display:"flex",gap:12,padding:"10px 0",borderBottom:i<2?`1px solid ${C.border}`:"none"}}>
                  <div style={{width:22,height:22,borderRadius:"50%",background:`${A}16`,color:A,fontSize:11,fontWeight:700,display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0,marginTop:1}}>{i+1}</div>
                  <div><div style={{fontSize:13,fontWeight:700,color:C.text,marginBottom:2}}>{t}</div><div style={{fontSize:12,color:C.muted,lineHeight:1.5}}>{d}</div></div>
                </div>
              ))}
            </div>
            <p style={{color:C.muted,fontSize:12,margin:"0 0 20px",maxWidth:440,lineHeight:1.6}}>{L("disclaimer",uiLang)}</p>
            <button onClick={startConvo} style={{background:A,color:"white",border:"none",borderRadius:10,padding:"14px 48px",fontSize:15,fontWeight:700,cursor:"pointer",boxShadow:"0 4px 14px rgba(126,72,236,0.30)"}}>{seed?L("startBtnSeeded",uiLang,{company:seed.company}):L("startBtn",uiLang)}</button>
          </div>
        )}

        {!started && saved && (
          <div style={{display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",minHeight:380,textAlign:"center"}}>
            <LumenIcon size={64}/>
            <h2 style={{margin:"20px 0 8px",color:C.text}}>Welcome back!</h2>
            <p style={{color:C.muted,fontSize:14,margin:"0 0 8px"}}>You have an onboarding session in progress.</p>
            <p style={{color:P,fontSize:13,fontWeight:600,margin:"0 0 24px"}}>{saved?.progress?.percent||0}% complete</p>
            <div style={{display:"flex",gap:12}}>
              <button onClick={resumeConvo} style={{background:P,color:"white",border:"none",borderRadius:10,padding:"13px 28px",cursor:"pointer",fontWeight:600}}>Resume session</button>
              <button onClick={()=>{const keep=sidRef.current;lsClearDraft(seedId);resetSession();sidRef.current=keep;}} style={{background:"transparent",border:`1px solid ${C.border}`,color:C.muted,borderRadius:10,padding:"13px 28px",cursor:"pointer"}}>Start fresh</button>
            </div>
          </div>
        )}

        {canCollapse && collapsed && <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:20,paddingInlineStart:38}}>
          <div style={{height:1,width:20,background:C.border}}/>
          <button onClick={()=>setCollapsed(false)} style={{background:"transparent",border:"none",padding:0,fontSize:12,color:C.muted,cursor:"pointer",textDecoration:"underline"}}>Show {messages.length-SHOW} earlier messages</button>
        </div>}

        {messages.slice(vStart).map((m,ri) => {
          const i = vStart+ri, canRW = m.role==="assistant"&&i>0;
          const canEdit = m.role==="user"&&!m.isWidget&&!loading&&!smoke;
          if (m.role==="divider") return <div key={i} style={{display:"flex",alignItems:"center",gap:10,margin:"6px 0 22px"}} role="separator" aria-label={`${m.label}${m.sub?`, ${m.sub}`:""}`}>
            <div style={{flex:1,height:1,background:C.border}}/>
            <div style={{fontSize:11,fontWeight:600,color:C.muted,whiteSpace:"nowrap"}}>✓ {m.label}{m.sub?<span style={{fontWeight:400}}> · {m.sub}</span>:null}</div>
            <div style={{flex:1,height:1,background:C.border}}/>
          </div>;
          return <div key={i} style={{display:"flex",justifyContent:m.role==="user"?"flex-end":"flex-start",marginBottom:18,animation:m.role==="assistant"?"slideUpFade 0.4s ease-out forwards":"none"}}>
            {m.role==="assistant" && <div style={{flexShrink:0,marginInlineEnd:10,marginTop:2}}><LumenIcon size={28}/></div>}
            <div style={{maxWidth:m.role==="assistant"?"min(88%, 580px)":"78%"}}>
              {m.content && <div>
                <div style={{background:m.role==="user"?(m.isWidget?C.hi:C.uBg):(dark?C.card:"#F5F3FB"),border:`1px solid ${m.role==="user"?(m.isWidget?P:C.border):(dark?C.border:"#E5E0F3")}`,color:m.role==="user"?(m.isWidget?C.wTx:C.uTx):C.text,borderRadius:14,padding:"11px 15px",fontSize:14,lineHeight:1.7,boxShadow:m.role==="assistant"?"0 1px 3px rgba(1,43,58,0.06)":"none"}}>
                  <MsgText text={m.content}/>
                </div>
                <div style={{display:"flex",gap:8,alignItems:"center",justifyContent:m.role==="user"?"flex-end":"flex-start",marginTop:4}}>
                  {DEV && canRW && <button onClick={()=>setRewind(i)} title="Deletes all messages after this point" style={{background:"transparent",border:"none",color:"#94a3b8",cursor:"pointer",fontSize:11,padding:"2px 6px",borderRadius:4,opacity:0.5}}>↺ Rewind to here</button>}
                  {canEdit && <button onClick={()=>{setInput(`Correction, earlier I said "${m.content}". What I actually meant: `);setTimeout(()=>taRef.current?.focus(),50);}} title="Send a correction without deleting any messages" style={{background:"transparent",border:"none",color:"#94a3b8",cursor:"pointer",fontSize:11,padding:"2px 6px",borderRadius:4,opacity:0.6}}>✎ Edit</button>}
                  {m.timestamp && <div style={{fontSize:10,color:C.muted,opacity:0.6}}>{m.timestamp}</div>}
                </div>
              </div>}
              {m.role==="assistant" && m.quickReplies?.length>0 && (()=>{
                const next = messages[i+1];
                const chosen = next?.isChip ? next.chipLabel||next.content : null;
                if (chosen) return <div style={{fontSize:11,color:C.muted,marginTop:6,fontStyle:"italic"}}>You chose: <strong style={{color:C.text}}>{chosen}</strong></div>;
                return null;
              })()}
              {m.role==="assistant" && m.widgets?.map(w => <div key={w} style={{background:C.card,border:`1px solid ${C.border}`,borderLeft:`3px solid ${A}`,borderRadius:12,padding:"12px 14px",marginTop:8,boxShadow:"0 2px 10px rgba(1,43,58,0.08)"}}>
                {renderWidget(w,i,w==="TOPICS"?m.topicSuggestions:null)}
              </div>)}
            </div>
          </div>;
        })}

        {showQR && !loading && <div style={{display:"flex",flexWrap:"wrap",gap:8,marginTop:-8,marginBottom:18,marginInlineStart:38,marginInlineEnd:0}}>
          {last.quickReplies.map((qr,idx) => <button key={idx} onClick={()=>sendMsg(qr,qr)} style={{background:"transparent",border:`1px solid ${LINK}`,color:LINK,borderRadius:16,padding:"6px 14px",fontSize:13,cursor:"pointer",fontWeight:600}}>{qr}</button>)}
        </div>}
        {loading && <div role="status" aria-live="polite" aria-label={WL("thinking",uiLang)} style={{display:"flex",justifyContent:"flex-start",marginBottom:18,animation:"slideUpFade 0.3s ease-out forwards"}}>
          <div style={{flexShrink:0,marginInlineEnd:10,marginTop:2}}><LumenIcon size={28}/></div>
          <div style={{background:dark?C.card:"#F5F3FB",border:`1px solid ${dark?C.border:"#E5E0F3"}`,borderRadius:14,padding:"14px 18px",maxWidth:"88%",boxShadow:"0 1px 3px rgba(1,43,58,0.06)"}}>
            {streamTxt ? <div style={{fontSize:14,lineHeight:1.65,color:C.text,whiteSpace:"pre-wrap"}}>{renderText(streamTxt)}<span style={{opacity:0.5}}>▍</span></div> : <TypingIndicator lang={uiLang}/>}
          </div>
        </div>}

        {retryMsg && !loading && <div style={{display:"flex",justifyContent:"flex-start",marginBottom:18,animation:"slideUpFade 0.3s ease-out forwards"}}>
          <div style={{flexShrink:0,marginInlineEnd:10,marginTop:2}}><LumenIcon size={28}/></div>
          <div style={{background:C.card,border:`1px solid ${C.border}`,borderRadius:12,padding:"14px 18px",display:"flex",alignItems:"center",gap:12}}>
            <span style={{fontSize:13,color:C.muted}}>That didn't go through. Tap Try again to resend.</span>
            <button onClick={()=>sendToAPI(retryMsg,true)} style={{background:P,color:"white",border:"none",borderRadius:7,padding:"6px 14px",fontSize:12,fontWeight:600,cursor:"pointer",whiteSpace:"nowrap"}}>Try again</button>
          </div>
        </div>}

        {initErr && !loading && <div style={{display:"flex",justifyContent:"flex-start",marginBottom:18,animation:"slideUpFade 0.3s ease-out forwards"}}>
          <div style={{flexShrink:0,marginInlineEnd:10,marginTop:2}}><LumenIcon size={28}/></div>
          <div style={{background:C.card,border:`1px solid ${C.border}`,borderRadius:12,padding:"14px 18px",display:"flex",alignItems:"center",gap:12,flexWrap:"wrap"}}>
            <span style={{fontSize:13,color:C.muted}}>We couldn't reach the assistant. Please check your connection and try again.</span>
            <button onClick={()=>{ const t=initErr; setInitErr(null); t==="resume"?resumeConvo():startConvo(); }} style={{background:P,color:"white",border:"none",borderRadius:7,padding:"6px 14px",fontSize:12,fontWeight:600,cursor:"pointer",whiteSpace:"nowrap"}}>Try again</button>
          </div>
        </div>}

        {done && !loading && <FinishCard C={C} cdata={cdata} setShowExport={setShowExport} linkCopied={linkCopied} setLinkCopied={setLinkCopied} sent={sent} sheetLink={sheetLink} onSeeProserv={onSeeProserv}/>}

        <div ref={botRef}/>
      </div>

      {/* Input */}
      {started && <div style={{background:C.card,borderTop:`1px solid ${C.border}`,padding:"12px 16px",flexShrink:0}}>
        <div style={{maxWidth:760,margin:"0 auto",transform:sideCol&&showPanel&&started?"translateX(-160px)":"none",transition:"transform 0.25s ease"}}>
          <div style={{display:"flex",gap:10,alignItems:"flex-end"}}>
            <textarea ref={taRef} value={input}
              onChange={e=>{setInput(e.target.value);if(taRef.current){taRef.current.style.height="auto";taRef.current.style.height=taRef.current.scrollHeight+"px";}}}
              onKeyDown={e=>{if(e.key==="Enter"&&!e.shiftKey){e.preventDefault();sendMsg();}}}
              placeholder={last?.role==="assistant"&&(last.widgets||[]).some(w=>!wState[w]?.submitted)?"Answer above — or just type it here":"Type your reply…"} rows={1}
              style={{flex:1,background:C.bg,border:`1.5px solid ${C.border}`,borderRadius:10,padding:"11px 14px",fontSize:mob?16:14,resize:"none",outline:"none",color:C.text}}/>
            <button onClick={()=>sendMsg()} aria-label="Send message" disabled={!input.trim()&&!loading}
              style={{background:input.trim()&&!loading?A:loading?A:"#e2e8f0",color:"white",border:"none",borderRadius:10,width:44,height:44,cursor:input.trim()||loading?"pointer":"default",flexShrink:0,display:"flex",alignItems:"center",justifyContent:"center"}}>
              {loading?<Spinner/>:<span style={{fontSize:18}}>↑</span>}
            </button>
          </div>
          {!done && progress.percent >= 15 && <div style={{textAlign:"center",marginTop:8}}>
            <button onClick={()=>setShowExport(true)} style={{background:"transparent",border:"none",color:C.muted,fontSize:11,cursor:"pointer",textDecoration:"underline",opacity:0.85}}>Finished early, or stuck? Review and send your brief</button>
          </div>}
        </div>
      </div>}
    </div>
  );
}


// ================= DEMO SHELL =================

function buildWorkbook(XL, merged, users) {
  const co = merged.company || {}, topics = merged.topics || [], channels = merged.channels || [],
        rpts = merged.reports || [], alts = merged.alerts || [];
  const companyName = co.name || "Draft";
  const wb = XL.utils.book_new();

  const boSheet = XL.utils.aoa_to_sheet([
    ["Welcome to your Lumen onboarding setup form!"],
    ["During onboarding, our team will configure your initial setup to help you get started quickly. Please review and complete each section of this sheet including: Business Objectives, Users, Social Channels, Topics/Filters, Reports/Dashboards/Alerts."],
    [],
    ["Field","Instructions / Example","Comments"],
    ["Company Name","",co.name||""],
    ["Date","",new Date().toLocaleDateString()],
    ["Contact Email","",co.email||""],
    ["Industry","",co.industry||""],
    ["Relevant Geographic Markets","Example: US, Germany, UK",co.markets||""],
    ["Key Languages","Example: English, German",co.languages||""],
    ["Business Objectives (Top 3 in priority order)","Example: 1. Reputation Management, 2. Competitive Intelligence, 3. Issue Tracking",co.objectives||""],
    ["Objective Details","Anything else about your objectives",co.objectiveDetails||""],
    ["Planned Use Cases Description","What insights or outcomes are you most interested in?",co.useCase||""],
    ["Preferred Onboarding Language","English",co.onboardingLanguage||"English"],
    ["Preferred Time Zone","CET",co.timezone||""],
    ["Teams/Departments Using Platform","Marketing, Comms, PR",co.teams||""],
    ["Main Point of Contact (Name + Email)","Jane Smith - jane@company.com",co.contact||co.email||""],
    ["Additional Comments or Questions","",""],
  ]);
  boSheet["!merges"] = [{s:{r:0,c:0},e:{r:0,c:2}},{s:{r:1,c:0},e:{r:1,c:2}}];
  boSheet["!cols"]   = [{wch:40},{wch:50},{wch:40}];
  XL.utils.book_append_sheet(wb, boSheet, "Business Objectives");

  const usersAoa = [
    ["Lumen Scoping Project Structure"],
    ["List of Users Requiring Access to the Tool\nAdmin: Full access to Analytics, Dashboards, Reports, IQ Apps, and Settings.\nFull Tool: Full access excluding user management.\nRead-Only: View-only access."],
    [],
    ["First Name","Last Name","Role/Department","E-mail","Access Rights"],
    ...((users && users.length) ? users.map(u=>[u.firstName||"",u.lastName||"",u.role||"",u.email||"",u.access||""]) : [["","","","",""]]),
  ];
  const usersSheet = XL.utils.aoa_to_sheet(usersAoa);
  usersSheet["!merges"] = [{s:{r:0,c:0},e:{r:0,c:4}},{s:{r:1,c:0},e:{r:1,c:4}}];
  usersSheet["!cols"]   = [{wch:15},{wch:15},{wch:20},{wch:30},{wch:15}];
  XL.utils.book_append_sheet(wb, usersSheet, "Users list");

  if (merged.queries && merged.queries !== "__skip__") {
    const qAoa = [
      ["Migrated queries (client's original content, as submitted)"],
      ["Reference for rebuilding queries in Lumen. May contain untranslated syntax from the client's previous tool."],
      [],
      ...String(merged.queries).split("\n").map(l=>[l]),
    ];
    const qSheet = XL.utils.aoa_to_sheet(qAoa);
    qSheet["!merges"] = [{s:{r:0,c:0},e:{r:0,c:2}},{s:{r:1,c:0},e:{r:1,c:2}}];
    qSheet["!cols"]   = [{wch:100}];
    XL.utils.book_append_sheet(wb, qSheet, "Migrated queries");
  }

  const tRows = [
    ["Lumen Project Plan - Topics/Filters"],
    ["Please list the brands, competitors, industry topics, campaigns, or categories you would like to monitor."],
    [],
    ["What is a Topic?","","A Topic is the broad subject you are tracking - the bucket that catches all the data."],
    ["Examples:","","\"Coca-Cola\", \"Apple\", \"Sustainability in Fashion\""],
    ["What is a Filter?","","A Filter is a specific lens you use to sort through the data in a topic."],
    ["Examples:","","Product categories, regions, campaigns, events, holidays, or crisis management topics"],
    ["Full Example:","","Topic: The Walt Disney Company | Filters: Movies, Retail, Disney Parks, Disney+"],
    ["Do you have existing queries to migrate?","","Please export any existing queries and include them below."],
    [],
    ["#","Topics/Filters","Group Name\nExamples: Competitors, Industry, Campaigns","Topic/Filter name","Keywords","URLs","Hashtags","Comments"],
  ];
  topics.forEach((tp,i) => tRows.push([i+1,"",tp.group||"",tp.name||"",tp.keywords||"",tp.urls||"",tp.hashtags||"",tp.rationale||tp.comments||""]));
  while (tRows.length < 31) tRows.push([tRows.length-10,"","","","","","",""]);
  const topicsSheet = XL.utils.aoa_to_sheet(tRows);
  topicsSheet["!merges"] = [{s:{r:0,c:0},e:{r:0,c:7}},{s:{r:1,c:0},e:{r:1,c:7}}];
  topicsSheet["!cols"]   = [{wch:4},{wch:12},{wch:20},{wch:25},{wch:30},{wch:30},{wch:15},{wch:40}];
  XL.utils.book_append_sheet(wb, topicsSheet, "Topics-Filters-Hashtags");

  const chRows = [
    ["Lumen Project Plan - Social channels"],
    ["Please list the social media profiles you wish to track (Brands, Competitors, or Influencers)."],
    [],
    ["What is a Channel?","A Channel in Lumen is a specific social media account or online source that you monitor."],
    ["Why Add Channels?","Adding channels allows you to monitor and analyse any public social media account."],
    ["Which channels should I add?","Include your own brand accounts, competitor channels, influencers, or thought leaders."],
    [],
    ["#","Author name","Channel type","Channel URL","Owned/Public"],
    ...channels.map((c,i) => [i+1,c.author||"",c.type||"",c.url||"",c.owned||""]),
  ];
  while (chRows.length < 33) chRows.push([chRows.length-7,"","","",""]);
  const chSheet = XL.utils.aoa_to_sheet(chRows);
  chSheet["!merges"] = [{s:{r:0,c:0},e:{r:0,c:4}},{s:{r:1,c:0},e:{r:1,c:4}}];
  chSheet["!cols"]   = [{wch:4},{wch:25},{wch:15},{wch:45},{wch:15}];
  XL.utils.book_append_sheet(wb, chSheet, "Social Channels");

  const rdRows = [
    ["Lumen Project Plan - Dashboards/Reports"],
    ["You may request dashboards, reports, or alerts during onboarding. Please ensure your Topics are created first."],
    [],
    ["What is a dashboard?","","A fully customisable, interactive overview of your data, shared via URL and updates live."],
    ["What is a report?","","A customisable snapshot for a specific time period, ideal for scheduled email delivery."],
    ["What is an alert?","","An automated email notification triggered by specific events like spikes in mentions or negative sentiment."],
    [],
    ["Dashboard / report / alert name","Main objective","","Details (time frame, KPIs, etc.)","Comments"],
    ...((rpts.length) ? rpts.map(r=>[r.name||"",r.objective||"","",r.details||"",r.comments||""]) : [["","","","",""]]),
    [],
    ["Alert","Name","Type","Details (time frame, KPIs, etc.)","Comments"],
    ...((alts.length) ? alts.map(a=>["Alert",a.name||"",a.type||"",a.details||"",a.comments||""]) : [["Alert","","","",""]]),
  ];
  const rdSheet = XL.utils.aoa_to_sheet(rdRows);
  rdSheet["!merges"] = [{s:{r:0,c:0},e:{r:0,c:4}},{s:{r:1,c:0},e:{r:1,c:4}}];
  rdSheet["!cols"]   = [{wch:30},{wch:30},{wch:5},{wch:30},{wch:30}];
  XL.utils.book_append_sheet(wb, rdSheet, "Reports-Dashboards-Alerts");

  const filename = "Lumen_Setup_Brief_" + companyName.replace(/\s+/g,"_") + "_" + new Date().toISOString().slice(0,10) + ".xlsx";
  return { wb, filename };
}


const TEAL = "#012B3A", CHERRY = "#FF4C46", MINT = "#DFFFDE";

const EXAMPLE_BRIEF = {
  merged: {
    company:{name:"Acme Corp",email:"jane@acmecorp.com",industry:"Consumer Goods (Footwear & Apparel)",useCase:"Protect brand reputation, track competitors, catch customer issues early",contact:"Jane Smith",markets:"United States, United Kingdom",languages:"English",objectives:"Reputation Management, Competitive Intelligence, Crisis Management",teams:"Marketing, PR",timezone:"GMT / UTC"},
    topics:[
      {name:"Acme Corp Brand",keywords:'"Acme Corp" OR @AcmeCorp',urls:"https://acmecorp.com",hashtags:"#AcmeCorp",comments:"Primary brand monitoring"},
      {name:"Nike",keywords:'"Nike" OR @Nike',urls:"",hashtags:"#Nike",comments:"Competitor, client-confirmed"},
      {name:"Customer Service Issues",keywords:'"Acme" AND (refund OR complaint)',urls:"",hashtags:"",comments:"Crisis early warning"},
    ],
    channels:[{author:"Acme Corp",type:"Instagram",url:"https://instagram.com/acmecorp",owned:"Owned"}],
    reports:[{name:"Brand Health Dashboard",objective:"Reputation Management",details:"Real-time, all markets",comments:""}],
    alerts:[{name:"Crisis Alert",type:"Sentiment spike",details:"Negative sentiment > 20% in 1 hour",comments:""}],
  },
  users:[{firstName:"Jane",lastName:"Smith",email:"jane@acmecorp.com",role:"Marketing Director",access:"Admin"}],
  handoff:{maturity:"Early — knows the pain, not the tooling",goalInOwnWords:"“I want to know when people complain about us before my CEO does”",hesitations:"Unsure about UK market priority; hesitated on competitor list beyond Nike",aiSuggestedUnconfirmed:"Adidas and Puma as competitors; GMT timezone",followUps:"Add 2-3 colleagues as users; confirm owned TikTok handle",consultantTips:"Lead the review call with the crisis-alert setup — that's the outcome she cares about most."},
  filename:"Lumen_Setup_Brief_Acme_Corp_"+new Date().toISOString().slice(0,10)+".xlsx",
  sentAt:new Date(),
};

function DemoBar({ stage, setStage, brief }) {
  const steps = [["sales","1. Sales generates link"],["client","2. Client onboarding chat"],["proserv","3. What Proserv receives"],["dash","4. Proserv dashboard"]];
  return (
    <div style={{background:TEAL,color:"white",display:"flex",alignItems:"center",gap:14,padding:"0 16px",height:44,flexShrink:0,fontSize:12}}>
      <span style={{fontWeight:700,whiteSpace:"nowrap"}}>Lumen Onboarding — Demo</span>
      <div style={{display:"flex",gap:6}}>
        {steps.map(([k,label]) => (
          <button key={k} onClick={()=>setStage(k)}
            style={{background:stage===k?"white":"rgba(255,255,255,0.12)",color:stage===k?TEAL:"white",border:"none",borderRadius:14,padding:"5px 12px",fontSize:11,fontWeight:stage===k?700:400,cursor:"pointer",whiteSpace:"nowrap"}}>
            {label}{k==="proserv"&&brief?" ●":""}
          </button>
        ))}
      </div>
      <span style={{marginLeft:"auto",opacity:0.75,whiteSpace:"nowrap"}}>Chat is live · hosting, sign-in, Drive &amp; Slack are simulated</span>
    </div>
  );
}

function Field({ label, opt, value, onChange, placeholder, area }) {
  const st = {width:"100%",border:"1px solid #e2e8f0",borderRadius:8,padding:"10px 12px",fontSize:13,fontFamily:"inherit",color:TEAL,outline:"none",boxSizing:"border-box",resize:"vertical"};
  return (
    <div style={{marginBottom:14}}>
      <div style={{fontSize:12,fontWeight:700,marginBottom:5}}>{label} {opt && <span style={{fontWeight:400,color:"#5b6b76"}}>(optional)</span>}</div>
      {area ? <textarea value={value} onChange={e=>onChange(e.target.value)} placeholder={placeholder} rows={3} style={st}/>
            : <input value={value} onChange={e=>onChange(e.target.value)} placeholder={placeholder} style={st}/>}
    </div>
  );
}

function SalesStage({ onGenerated }) {
  const [company,setCompany] = useState("");
  const [contactName,setContactName] = useState("");
  const [email,setEmail] = useState("");
  const [industry,setIndustry] = useState("");
  const [language,setLanguage] = useState("English");
  const [notes,setNotes] = useState("");
  const [link,setLink] = useState(null);
  const [copied,setCopied] = useState(false);
  const fillExample = () => { setCompany("Acme Corp"); setContactName("Jane Smith"); setEmail("jane@acmecorp.com"); setIndustry("Consumer goods — footwear and apparel"); setNotes("Enterprise tier. Main interest is competitive intelligence; key competitor is Nike."); };
  const generate = () => {
    if (!company.trim() || !contactName.trim()) return;
    setLink({ url:`https://onboarding.hootsuite.com/?s=${crypto.randomUUID()}`, seed:{company:company.trim(),contactName:contactName.trim(),email:email.trim(),industry:industry.trim(),notes:notes.trim(),language} });
  };
  return (
    <div style={{flex:1,overflowY:"auto",background:"white",color:TEAL,fontFamily:"Arial, sans-serif"}}>
      <div style={{maxWidth:520,margin:"0 auto",padding:"40px 20px"}}>
        <h1 style={{fontSize:20,margin:"0 0 6px"}}>Generate a client onboarding link</h1>
        <p style={{fontSize:13,color:"#5b6b76",margin:"0 0 20px",lineHeight:1.5}}>Internal page for the sales team. Fill in what we already know — the client is greeted by name and never re-asked the basics.</p>

        <div style={{display:"flex",alignItems:"center",gap:8,background:"#f0fdf4",border:"1px solid #bbf7d0",borderRadius:8,padding:"9px 12px",fontSize:12,marginBottom:22}}>
          <span style={{width:8,height:8,borderRadius:"50%",background:"#16a34a",flexShrink:0}}/>
          <span>Signed in as <strong>damien@hootsuite.com</strong></span>
          <span style={{color:"#5b6b76",marginLeft:"auto",fontSize:11}}>Google sign-in, @hootsuite.com only · simulated</span>
        </div>

        <Field label="Company" value={company} onChange={setCompany} placeholder="Acme Corp"/>
        <Field label="Contact name" value={contactName} onChange={setContactName} placeholder="Jane Smith"/>
        <Field label="Contact email" opt value={email} onChange={setEmail} placeholder="jane@acmecorp.com"/>
        <Field label="Industry" opt value={industry} onChange={setIndustry} placeholder="Consumer goods — footwear and apparel"/>
        <div style={{marginBottom:16}}>
          <label style={{display:"block",fontSize:13,fontWeight:700,marginBottom:5}}>Onboarding language <span style={{fontWeight:400,color:"#5b6b76"}}>· the client can change this on their welcome screen</span></label>
          <select value={language} onChange={e=>setLanguage(e.target.value)} style={{width:"100%",border:"1px solid #e2e8f0",borderRadius:8,padding:"11px 12px",fontSize:14,color:TEAL,background:"white",cursor:"pointer"}}>
            {UI_LANGS.map(l=><option key={l.code} value={l.code}>{l.native}</option>)}
          </select>
        </div>
        <Field label="What do you already know about this client?" opt area value={notes} onChange={setNotes} placeholder="Why they're buying, competitors they named, report audience, tier sold, language, anything sensitive — never shown to the client, quietly shapes the assistant's suggestions"/>

        <div style={{display:"flex",gap:10,alignItems:"center"}}>
          <button onClick={generate} disabled={!company.trim()||!contactName.trim()} style={{background:company.trim()&&contactName.trim()?CHERRY:"#f0b5b3",color:"white",border:"none",borderRadius:8,padding:"12px 24px",fontSize:14,fontWeight:700,cursor:company.trim()&&contactName.trim()?"pointer":"default"}}>Generate link</button>
          <button onClick={fillExample} style={{background:"transparent",border:"none",color:"#5b6b76",fontSize:12,cursor:"pointer",textDecoration:"underline"}}>Use example client</button>
        </div>

        {link && (
          <div style={{marginTop:20,background:MINT,border:"1px solid #b9e8b8",borderRadius:10,padding:14}}>
            <strong style={{fontSize:13}}>✓ Link ready</strong>
            <div style={{fontSize:12,wordBreak:"break-all",background:"white",border:"1px solid #e2e8f0",borderRadius:6,padding:"8px 10px",margin:"8px 0 10px"}}>{link.url}</div>
            <div style={{display:"flex",gap:8,alignItems:"center"}}>
              <button onClick={()=>onGenerated(link.seed)} style={{background:TEAL,color:"white",border:"none",borderRadius:6,padding:"9px 16px",fontSize:12,fontWeight:700,cursor:"pointer"}}>Open the link as the client →</button>
              <button onClick={()=>{try{navigator.clipboard?.writeText(link.url);}catch(e){} setCopied(true);setTimeout(()=>setCopied(false),1500);}} style={{background:"white",color:TEAL,border:"1px solid #e2e8f0",borderRadius:6,padding:"9px 14px",fontSize:12,cursor:"pointer"}}>{copied?"Copied ✓":"Copy link"}</button>
            </div>
            <div style={{fontSize:11,color:"#5b6b76",marginTop:8,lineHeight:1.4}}>In production the salesperson pastes this into their email. Here, click through to experience it as the client.</div>
          </div>
        )}
      </div>
    </div>
  );
}

function SlackCard({ variant, brief }) {
  const co = brief?.merged?.company || {};
  const stalled = variant === "stalled";
  const fields = stalled
    ? [["Company", co.name||"Acme Corp"],["Contact email", co.email||"jane@acmecorp.com"],["Progress","60%"],["Last active","2 days ago"]]
    : [["Company", co.name||"—"],["Contact", `${co.contact||"—"} (${co.email||"no email"})`],["Topics", String((brief?.merged?.topics||[]).length)],["Users", String((brief?.users||[]).length)]];
  return (
    <div style={{background:"white",border:"1px solid #ddd",borderRadius:10,overflow:"hidden",fontFamily:"Arial, sans-serif"}}>
      <div style={{background:"#3f0e40",color:"white",padding:"8px 14px",fontSize:12,fontWeight:700}}># proserv-lumen-onboarding</div>
      <div style={{display:"flex",gap:10,padding:"12px 14px"}}>
        <div style={{width:36,height:36,borderRadius:6,background:"#7c6fe0",color:"white",display:"flex",alignItems:"center",justifyContent:"center",fontSize:16,flexShrink:0}}>🦉</div>
        <div style={{flex:1,minWidth:0}}>
          <div style={{fontSize:13,marginBottom:4}}><strong>Lumen Onboarding</strong> <span style={{background:"#e8e8e8",borderRadius:3,padding:"0 4px",fontSize:9,fontWeight:700,color:"#616061"}}>APP</span> <span style={{color:"#616061",fontSize:11}}>{new Date().toLocaleTimeString([],{hour:"2-digit",minute:"2-digit"})}</span></div>
          <div style={{fontSize:14,fontWeight:700,marginBottom:8}}>{stalled?"🟡 Lumen onboarding stalled":"🟢 Lumen setup brief completed"}</div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:"6px 16px",fontSize:12,marginBottom:8}}>
            {fields.map(([k,v]) => <div key={k}><div style={{fontWeight:700}}>{k}:</div><div style={{color:"#1d1c1d"}}>{v}</div></div>)}
          </div>
          {!stalled && <div style={{fontSize:11,color:"#616061",marginBottom:6}}>Link prepared by damien@hootsuite.com · Client has edit access to the Sheet until the review call</div>}
          <div style={{fontSize:12,color:"#1264a3"}}>
            {stalled
              ? <>Partial brief available — a consultant can pick this up or nudge the client. <u>💬 Open session</u></>
              : <><u>📄 Setup brief in Drive</u> &nbsp;·&nbsp; <u>💬 Conversation transcript</u></>}
          </div>
        </div>
      </div>
    </div>
  );
}

function DriveCard({ brief }) {
  const fname = (brief?.filename || EXAMPLE_BRIEF.filename).replace(/\.xlsx$/, "");
  const download = async () => {
    const b = brief || EXAMPLE_BRIEF;
    const XLSX = await loadXLSX();
    const { wb, filename } = buildWorkbook(XLSX, b.merged, b.users);
    XLSX.writeFile(wb, filename);
  };
  return (
    <div style={{background:"white",border:"1px solid #ddd",borderRadius:10,overflow:"hidden",fontFamily:"Arial, sans-serif"}}>
      <div style={{padding:"10px 14px",borderBottom:"1px solid #eee",fontSize:12,color:"#5f6368",display:"flex",alignItems:"center",gap:6}}>
        <span style={{fontSize:14}}>▸</span> Proserv Shared Drive <span>›</span> <strong style={{color:"#202124"}}>Lumen Onboarding Briefs</strong>
      </div>
      <div style={{display:"flex",alignItems:"center",gap:12,padding:"12px 14px"}}>
        <div style={{width:34,height:34,borderRadius:6,background:"#e6f4ea",display:"flex",alignItems:"center",justifyContent:"center",fontSize:16,flexShrink:0}}>📊</div>
        <div style={{flex:1,minWidth:0}}>
          <div style={{fontSize:13,fontWeight:700,color:"#202124",wordBreak:"break-all"}}>{fname}</div>
          <div style={{fontSize:11,color:"#5f6368"}}>just now · uploaded by lumen-onboarding@…iam.gserviceaccount.com</div>
        </div>
        <button onClick={download} style={{background:TEAL,color:"white",border:"none",borderRadius:6,padding:"8px 14px",fontSize:12,fontWeight:700,cursor:"pointer",flexShrink:0}}>⬇ Download (real file)</button>
      </div>
    </div>
  );
}

function ProservStage({ brief, useExample }) {
  return (
    <div style={{flex:1,overflowY:"auto",background:"#f8f9fa",color:TEAL,fontFamily:"Arial, sans-serif"}}>
      <div style={{maxWidth:620,margin:"0 auto",padding:"36px 20px 60px"}}>
        <h1 style={{fontSize:20,margin:"0 0 6px"}}>What Proserv receives</h1>
        <p style={{fontSize:13,color:"#5b6b76",margin:"0 0 24px",lineHeight:1.5}}>The moment the client clicks “Send to my Lumen team”, three things happen automatically — no client download, no email attachment, nothing to chase.</p>

        {!brief && (
          <div style={{background:"#fffbeb",border:"1px solid #fde68a",borderRadius:10,padding:"12px 14px",fontSize:12,marginBottom:20,display:"flex",alignItems:"center",gap:10,flexWrap:"wrap"}}>
            <span>No brief sent yet in this demo — finish the chat in step 2, or</span>
            <button onClick={useExample} style={{background:CHERRY,color:"white",border:"none",borderRadius:6,padding:"6px 12px",fontSize:12,fontWeight:700,cursor:"pointer"}}>show with example data</button>
          </div>
        )}

        <div style={{fontSize:12,fontWeight:700,margin:"0 0 8px",textTransform:"uppercase",letterSpacing:"0.05em",color:"#5b6b76"}}>1 · Slack notification (your existing webhook pattern)</div>
        <SlackCard brief={brief}/>

        <div style={{fontSize:12,fontWeight:700,margin:"24px 0 8px",textTransform:"uppercase",letterSpacing:"0.05em",color:"#5b6b76"}}>2 · Brief lands in the Google Drive folder</div>
        <DriveCard brief={brief}/>
        <div style={{fontSize:11,color:"#5b6b76",margin:"6px 2px 0"}}>The download button generates the actual XLSX — same tabs as the current requirements spreadsheet.</div>

        {(brief?.handoff) && <>
          <div style={{fontSize:12,fontWeight:700,margin:"24px 0 8px",textTransform:"uppercase",letterSpacing:"0.05em",color:"#5b6b76"}}>2b · Consultant handoff — a Google Doc the client never sees</div>
          <div style={{background:"white",border:"1px solid #ddd",borderRadius:10,padding:"14px 16px",fontSize:12,lineHeight:1.7,fontFamily:"Arial, sans-serif"}}>
            {[["Maturity",brief.handoff.maturity],["Goal in their own words",brief.handoff.goalInOwnWords],["Hesitations",brief.handoff.hesitations],["AI-suggested, unconfirmed",brief.handoff.aiSuggestedUnconfirmed],["Follow-ups for the review call",brief.handoff.followUps],["Tips for the consultant",brief.handoff.consultantTips]].map(([k,v])=>v?<div key={k}><strong>{k}:</strong> {v}</div>:null)}
          </div>
          <div style={{fontSize:11,color:"#5b6b76",margin:"6px 2px 0"}}>Generated by the assistant during the conversation — the difference between “reading a spreadsheet” and “being briefed”.</div>
        </>}

        <div style={{fontSize:12,fontWeight:700,margin:"24px 0 8px",textTransform:"uppercase",letterSpacing:"0.05em",color:"#5b6b76"}}>3 · And if a client stalls, nothing is lost</div>
        <SlackCard variant="stalled" brief={brief}/>
        <div style={{fontSize:11,color:"#5b6b76",margin:"6px 2px 0"}}>Sent automatically after 48h of inactivity at 40%+ progress. A half-finished brief becomes a warm follow-up instead of a silent drop-off.</div>

        <div style={{marginTop:28,background:"white",border:"1px solid #e2e8f0",borderRadius:10,padding:"14px 16px",fontSize:12,lineHeight:1.7}}>
          <strong>In this demo:</strong> the conversation, widgets, brief-building and XLSX file are fully real (live AI). The hosting, resumable session links, Google sign-in, Drive upload and Slack posts are visual simulations of the built, ready-to-deploy Netlify app.
        </div>
      </div>
    </div>
  );
}

function DashboardStage({ brief, onOpenHandoff }) {
  const downloadBrief = async () => {
    const b = brief || EXAMPLE_BRIEF;
    const XLSX = await loadXLSX();
    const { wb, filename } = buildWorkbook(XLSX, b.merged, b.users);
    XLSX.writeFile(wb, filename);
  };
  const P = "#012B3A";
  const rows = [
    { co:"Acme Corp", em:"jane@acmecorp.com", by:"tom.reid", st:"completed", pct:100, start:"Jul 1, 09:12", last:"Jul 1, 09:26", min:"14 min", tok:"31k", cost:"$1.24", calls:19, brief:true, handoff:true },
    { co:"Northwind Foods", em:"m.alvarez@northwind.com", by:"sara.kim", st:"in_progress", pct:60, start:"Jul 2, 15:40", last:"Jul 3, 08:55", min:"11 min", tok:"18k", cost:"$0.71", calls:12, brief:false, handoff:false, sofar:true },
    { co:"Helios Bank", em:"p.dubois@heliosbank.eu", by:"tom.reid", st:"stalled", pct:40, start:"Jun 29, 10:02", last:"Jul 1, 10:15", min:"9 min", tok:"14k", cost:"$0.55", calls:9, brief:false, handoff:false, sofar:true },
    { co:"Verde Cosmetics", em:"l.ricci@verdecos.it", by:"sara.kim", st:"seeded", pct:0, start:"—", last:"Jul 3, 07:30", min:"—", tok:"0", cost:"$0.00", calls:0, brief:false, handoff:false },
  ];
  const pill = st => ({
    completed:{bg:"#DFFFDE",c:"#166534",l:"Completed"},
    in_progress:{bg:"#E8F1FE",c:"#1d4ed8",l:"In progress"},
    stalled:{bg:"#FEF3E2",c:"#92400e",l:"Stalled"},
    seeded:{bg:"#f1f5f9",c:"#475569",l:"Link sent"},
  })[st];
  const kpis = [["12","sessions total"],["7/9","completed (78%)"],["14 min","median completion time"],["8.9 h","est. time saved (vs 90 min manual)"],["$9.12","est. AI spend · $1.30/brief"]];
  return (
    <div style={{flex:1,overflowY:"auto",background:"#f7f8fa",fontFamily:"Arial, sans-serif",color:P}}>
      <div style={{background:"white",borderBottom:"1px solid #e2e8f0",padding:"12px 24px",display:"flex",alignItems:"center",gap:12}}>
        <div style={{fontSize:15,fontWeight:700}}>Lumen Onboarding — Dashboard</div>
        <div style={{marginLeft:"auto",fontSize:11,color:"#5b6b76"}}>signed in as damien.thierry@hootsuite.com</div>
      </div>
      <div style={{maxWidth:1080,margin:"0 auto",padding:"20px 16px"}}>
        <div style={{fontSize:11,color:"#5b6b76",marginBottom:12}}>This is the live view at <b>/dashboard.html</b> — Google sign-in, @hootsuite.com only. Example data below; the Acme row is your demo session.</div>
        <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(170px,1fr))",gap:10,marginBottom:16}}>
          {kpis.map(([v,l],i)=>(<div key={i} style={{background:"white",border:"1px solid #e2e8f0",borderRadius:10,padding:"12px 14px"}}>
            <div style={{fontSize:20,fontWeight:700}}>{v}</div><div style={{fontSize:10.5,color:"#5b6b76",marginTop:3,lineHeight:1.4}}>{l}</div></div>))}
        </div>
        <div style={{background:"white",border:"1px solid #e2e8f0",borderRadius:12,overflowX:"auto"}}>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:12.5,minWidth:820}}>
            <thead><tr>{["Client","Status","Progress","Started","Last active","Time","Tokens / cost","Report"].map(h=>(
              <th key={h} style={{textAlign:h==="Time"||h==="Tokens / cost"?"right":"left",fontSize:10,textTransform:"uppercase",letterSpacing:"0.04em",color:"#5b6b76",padding:"9px 12px",borderBottom:"1px solid #e2e8f0",whiteSpace:"nowrap"}}>{h}</th>))}</tr></thead>
            <tbody>{rows.map((r,i)=>{ const pl = pill(r.st); return (
              <tr key={i}>
                <td style={{padding:"10px 12px",borderBottom:"1px solid #eef1f4"}}><div style={{fontWeight:700}}>{r.co}</div><div style={{fontSize:10.5,color:"#5b6b76"}}>{r.em} · seeded by {r.by}</div></td>
                <td style={{padding:"10px 12px",borderBottom:"1px solid #eef1f4"}}><span style={{background:pl.bg,color:pl.c,fontSize:10.5,fontWeight:700,borderRadius:10,padding:"2px 9px",whiteSpace:"nowrap"}}>{pl.l}</span></td>
                <td style={{padding:"10px 12px",borderBottom:"1px solid #eef1f4"}}><div style={{width:80,height:6,background:"#eef1f4",borderRadius:3}}><div style={{width:`${r.pct}%`,height:"100%",background:P,borderRadius:3}}/></div><div style={{fontSize:10.5,color:"#5b6b76",marginTop:2}}>{r.pct}%</div></td>
                <td style={{padding:"10px 12px",borderBottom:"1px solid #eef1f4",whiteSpace:"nowrap"}}>{r.start}</td>
                <td style={{padding:"10px 12px",borderBottom:"1px solid #eef1f4",whiteSpace:"nowrap"}}>{r.last}</td>
                <td style={{padding:"10px 12px",borderBottom:"1px solid #eef1f4",textAlign:"right",whiteSpace:"nowrap"}}>{r.min}{r.sofar&&<div style={{fontSize:10,color:"#5b6b76"}}>so far</div>}</td>
                <td style={{padding:"10px 12px",borderBottom:"1px solid #eef1f4",textAlign:"right",whiteSpace:"nowrap"}}>{r.tok}<div style={{fontSize:10,color:"#5b6b76"}}>{r.cost} · {r.calls} calls</div></td>
                <td style={{padding:"10px 12px",borderBottom:"1px solid #eef1f4",whiteSpace:"nowrap"}}>
                  {r.brief ? <>
                    <button onClick={downloadBrief} style={{background:"none",border:"none",color:"#0b6b3a",fontWeight:700,fontSize:11.5,cursor:"pointer",padding:0,textDecoration:"underline"}}>Open brief ⤓</button><br/>
                    <button onClick={onOpenHandoff} style={{background:"none",border:"none",color:"#5b6b76",fontSize:11,cursor:"pointer",padding:0,textDecoration:"underline"}}>Consultant handoff ↗</button>
                  </> : "—"}
                </td>
              </tr>); })}</tbody>
          </table>
        </div>
        <div style={{fontSize:10.5,color:"#5b6b76",marginTop:12,lineHeight:1.5}}>Estimates: time saved = completed sessions × (90 min manual baseline − median completion time). AI cost from exact per-call token metering at $3/$15 per MTok in/out. In production the “Open brief” link goes to the delivered Google Sheet in Drive.</div>
      </div>
    </div>
  );
}

export default function Demo() {
  const [stage,setStage] = useState("sales");
  const [seed,setSeed] = useState(null);
  const [brief,setBrief] = useState(null);
  const [chatKey,setChatKey] = useState(0);
  return (
    <div style={{height:"100vh",display:"flex",flexDirection:"column",overflow:"hidden"}}>
      <DemoBar stage={stage} setStage={setStage} brief={brief}/>
      {stage==="sales" && <SalesStage onGenerated={sd=>{setSeed(sd);setBrief(null);setChatKey(k=>k+1);setStage("client");}}/>}
      {stage==="client" && (
        <div style={{flex:1,minHeight:0}}>
          <OnboardingApp key={chatKey} seed={seed} onBriefSent={setBrief} onSeeProserv={()=>setStage("proserv")}/>
        </div>
      )}
      {stage==="proserv" && <ProservStage brief={brief} useExample={()=>setBrief(EXAMPLE_BRIEF)}/>}
      {stage==="dash" && <DashboardStage brief={brief} onOpenHandoff={()=>setStage("proserv")}/>}
    </div>
  );
}

// ================= LIVE CLIENT CHAT ENTRY =================
// Standalone client-facing page: no demo tab shell. Fetches the client-safe seed
// the Sales page stored under ?s=<id> (consultant notes never reach the browser),
// then runs the onboarding chat full-bleed. onBriefSent writes to the session
// store (handled inside OnboardingApp.handleSend); here we just need a no-op sink
// and no "see Proserv" navigation.
export function LiveChat() {
  const [state, setState] = useState({ loading: true, seed: null, seedId: null });
  useEffect(() => {
    let alive = true;
    fetchSeedFromURL().then(r => { if (alive) setState({ loading: false, seed: r.seed, seedId: r.seedId }); });
    return () => { alive = false; };
  }, []);
  if (state.loading) return <div style={{height:"100vh",display:"flex",alignItems:"center",justifyContent:"center",fontFamily:"'Inter', Arial, sans-serif",color:"#64748b"}}>Loading…</div>;
  return (
    <div style={{height:"100vh",display:"flex",flexDirection:"column",overflow:"hidden"}}>
      <div style={{flex:1,minHeight:0}}>
        <OnboardingApp seed={state.seed} seedId={state.seedId} onBriefSent={()=>{}} onSeeProserv={null}/>
      </div>
    </div>
  );
}
