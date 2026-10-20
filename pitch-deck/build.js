// Cedarworks Apps investor pitch deck (v3: readability pass + full AppForge section)
// Usage: node build.js [output.pptx]   (images are read from this script's folder)
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const I = require("react-icons/lu");
const fs = require("fs"), path = require("path");

const OUT = process.argv[2] || "Cedarworks-Apps-Pitch-Deck.pptx";

// Brand palette (brand/README.md)
const GREEN = "143D28", GREEN2 = "1E5238", CREAM = "F4EBD5", ORANGE = "D98B3C",
  TAN = "ECE1C7", PAPER = "FAF6EC", INK = "1B2A21", MUTED = "4E5D53", SAGE = "B7CDB6", WHITE = "FFFFFF";
const HEAD = "Cambria", BODY = "Calibri";

const CEDAR = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="#d98b3c" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"><path d="M12 22V8"/><path d="M12 8c0-2.5 1.6-4.2 4-4.5C16 6 14.4 7.8 12 8Z"/><path d="M12 8c0-2.5-1.6-4.2-4-4.5C8 6 9.6 7.8 12 8Z"/><path d="M12 13c0-2.2 1.5-3.7 3.6-4C15.6 11.3 14.1 12.8 12 13Z"/><path d="M12 13c0-2.2-1.5-3.7-3.6-4C8.4 11.3 9.9 12.8 12 13Z"/><path d="M12 18c0-2 1.4-3.4 3.3-3.7C15.3 16.4 13.9 17.8 12 18Z"/><path d="M12 18c0-2-1.4-3.4-3.3-3.7C8.7 16.4 10.1 17.8 12 18Z"/></svg>`;

async function svgPng(svg) {
  const buf = await sharp(Buffer.from(svg)).resize(256, 256).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}
async function icon(name, color) {
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(I[name], { color: "#" + color, size: 256 }));
  return svgPng(svg);
}
const shot = f => "image/png;base64," + fs.readFileSync(path.join(__dirname, f)).toString("base64");

(async () => {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9"; // 10 x 5.625
  pres.author = "James Chambers";
  pres.company = "Cedarworks Apps LLC";
  pres.title = "Cedarworks Apps — Investor Pitch Deck";
  pres.subject = "Pre-seed pitch deck: executive summary, problem statement, solution & product, AppForge, traction & metrics, market opportunity, business model, competitive landscape, use of funds, financial projections, team, the ask";

  const cedar = await svgPng(CEDAR);
  const ic = {};
  for (const n of ["LuChurch", "LuTarget", "LuHammer", "LuRocket", "LuSearch", "LuUsers", "LuSmartphone",
    "LuMegaphone", "LuLayers", "LuTrendingUp", "LuClock", "LuBrain", "LuRepeat", "LuUser", "LuWrench", "LuHandshake",
    "LuZap", "LuBriefcase", "LuMail", "LuPhone", "LuGlobe", "LuBot", "LuCode", "LuLayoutDashboard", "LuSparkles",
    "LuPresentation", "LuShare2", "LuDatabase", "LuGauge", "LuCompass", "LuFactory", "LuPenTool", "LuStore",
    "LuMessagesSquare", "LuChartLine", "LuPackage", "LuSend", "LuDollarSign"]) {
    ic[n] = { o: await icon(n, ORANGE), g: await icon(n, GREEN) };
  }

  const slides = [];
  // ---- helpers ----
  function badge(s, name, x, y, d, dark) {
    s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: dark ? GREEN2 : WHITE }, line: { color: ORANGE, width: 1.5 } });
    const p = d * 0.26;
    s.addImage({ data: ic[name][dark ? "o" : "g"], x: x + p, y: y + p, w: d - 2 * p, h: d - 2 * p });
  }
  function logo(s, x, y, d) {
    s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: GREEN }, line: { color: ORANGE, width: 1.5 } });
    const p = d * 0.22;
    s.addImage({ data: cedar, x: x + p, y: y + p, w: d - 2 * p, h: d - 2 * p });
  }
  function eyebrow(s, text, y = 0.4) {
    s.addText(text, { x: 0.5, y, w: 9, h: 0.28, fontFace: BODY, fontSize: 12, bold: true, color: ORANGE, margin: 0, isTextBox: true });
  }
  function title(s, text, dark = false, y = 0.68, size = 26) {
    s.addText(text, { x: 0.5, y, w: 9, h: 0.85, fontFace: HEAD, fontSize: size, bold: true, color: dark ? CREAM : GREEN, margin: 0, valign: "top", isTextBox: true });
  }
  function light() { const s = pres.addSlide(); s.background = { color: PAPER }; slides.push([s, false]); return s; }
  function dark() { const s = pres.addSlide(); s.background = { color: GREEN }; slides.push([s, true]); return s; }
  function card(s, x, y, w, h, fill = WHITE) {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.08, fill: { color: fill },
      line: { color: TAN, width: 0.75 }, shadow: { type: "outer", color: "000000", opacity: 0.08, blur: 6, offset: 2, angle: 90 } });
  }
  function darkCard(s, x, y, w, h) {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.08, fill: { color: GREEN2 }, line: { color: GREEN2 } });
  }
  function strip(s, x, y, w, h) {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.08, fill: { color: TAN }, line: { color: TAN } });
  }
  const T = (s, text, o) => s.addText(text, Object.assign({ fontFace: BODY, fontSize: 13, color: INK, margin: 0, valign: "top", isTextBox: true }, o));
  const bullets = (arr) => arr.map((f, i) => ({ text: f, options: { bullet: { indent: 15 }, breakLine: i < arr.length - 1 } }));
  // bold-lead bullets: [["Lead: ", "rest"], ...]
  const leadBullets = (arr) => {
    const out = [];
    arr.forEach(([b, t], i) => {
      out.push({ text: b, options: { bold: true, bullet: { indent: 15 } } });
      out.push({ text: t, options: { breakLine: i < arr.length - 1 } });
    });
    return out;
  };
  function statCard(s, x, y, w, h, big, small, hi, bigSize = 28) {
    card(s, x, y, w, h, hi ? GREEN : WHITE);
    T(s, big, { x: x + 0.2, y: y + 0.1, w: w - 0.36, h: 0.52, fontFace: HEAD, fontSize: bigSize, bold: true, color: hi ? CREAM : GREEN, valign: "middle" });
    T(s, small, { x: x + 0.2, y: y + 0.64, w: w - 0.36, h: h - 0.68, fontSize: 12, color: hi ? SAGE : MUTED });
  }
  function appShot(s, file, x, y, w, h) {
    s.addShape(pres.shapes.RECTANGLE, { x: x - 0.03, y: y - 0.03, w: w + 0.06, h: h + 0.06, fill: { color: TAN }, line: { color: TAN },
      shadow: { type: "outer", color: "000000", opacity: 0.18, blur: 8, offset: 3, angle: 90 } });
    s.addImage({ data: shot(file), x, y, w, h });
  }
  function numDot(s, n, x, y, d = 0.36) {
    s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: ORANGE }, line: { color: ORANGE } });
    T(s, String(n), { x, y, w: d, h: d, fontSize: 13, bold: true, color: WHITE, align: "center", valign: "middle" });
  }
  // numbered callouts beside a screenshot
  function callouts(s, items, x, y, w, gap) {
    items.forEach(([h, b], i) => {
      const yy = y + i * gap;
      numDot(s, i + 1, x, yy);
      T(s, h, { x: x + 0.5, y: yy - 0.01, w: w - 0.5, h: 0.3, fontSize: 14, bold: true, color: GREEN });
      T(s, b, { x: x + 0.5, y: yy + 0.3, w: w - 0.5, h: gap - 0.35, fontSize: 12.5, color: MUTED });
    });
  }

  // ===================== 1. Title =====================
  {
    const s = dark();
    logo(s, 0.5, 0.55, 0.8);
    T(s, "CEDARWORKS APPS", { x: 1.5, y: 0.72, w: 5, h: 0.45, fontSize: 16, bold: true, color: CREAM, valign: "middle" });
    T(s, "Practical AI apps,\nbuilt to do real work.", { x: 0.5, y: 1.6, w: 6.6, h: 1.7, fontFace: HEAD, fontSize: 40, bold: true, color: CREAM });
    T(s, "An AI app studio with two apps live in the app stores today, and AppForge, our in-house app factory, to build the next eight.", {
      x: 0.5, y: 3.4, w: 6.3, h: 0.9, fontFace: HEAD, italic: true, fontSize: 16, color: SAGE });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 7.2, y: 1.7, w: 2.3, h: 1.6, rectRadius: 0.1, fill: { color: GREEN2 }, line: { color: ORANGE, width: 1.25 } });
    s.addText([
      { text: "Pre-seed round", options: { fontSize: 12, bold: true, color: ORANGE, breakLine: true } },
      { text: "$450K", options: { fontSize: 38, bold: true, color: CREAM, fontFace: HEAD, breakLine: true } },
      { text: "Springfield, MO", options: { fontSize: 12, color: SAGE } },
    ], { x: 7.2, y: 1.78, w: 2.3, h: 1.45, fontFace: BODY, align: "center", valign: "middle", margin: 0, isTextBox: true });
    T(s, "James Chambers, Founder  ·  james@cedarworksapps.com  ·  417-370-2084  ·  cedarworksapps.com", { x: 0.5, y: 4.6, w: 9, h: 0.3, fontSize: 12, color: CREAM });
    s.addNotes("Cedarworks Apps is a Springfield, Missouri AI app studio. We build focused, subscription AI apps for specific professions. Two are live in the Apple App Store and Google Play today. We also built AppForge, our in-house app factory, which finds proven app ideas, writes the code and builds the marketing. We're raising a $450K pre-seed round so I can build full-time, market our two apps and launch eight more.");
  }

  // ===================== 2. Executive summary =====================
  {
    const s = light();
    eyebrow(s, "Executive Summary");
    title(s, "Cedarworks Apps in one minute");
    const items = [
      ["LuLayers", "What we do", "We build focused AI apps for specific professions and sell them as monthly subscriptions."],
      ["LuSmartphone", "Live today", "PastorAI ($47/mo) for pastors and Ready Room ($29/mo) for sales reps, on iOS, Android and web."],
      ["LuFactory", "Our engine: AppForge", "Our in-house app factory finds proven app ideas, writes the code and builds the marketing for each app."],
      ["LuRocket", "The ask", "$450,000 pre-seed to fund full-time building, market our two apps and launch eight more in 18 months."],
    ];
    items.forEach(([n, h, b], i) => {
      const x = 0.5 + (i % 2) * 4.6, y = 1.55 + Math.floor(i / 2) * 1.75;
      card(s, x, y, 4.4, 1.55);
      badge(s, n, x + 0.22, y + 0.25, 0.6);
      T(s, h, { x: x + 1.0, y: y + 0.22, w: 3.2, h: 0.32, fontFace: HEAD, fontSize: 16, bold: true, color: GREEN });
      T(s, b, { x: x + 1.0, y: y + 0.58, w: 3.25, h: 0.9, fontSize: 13, color: INK });
    });
    s.addNotes("The whole story in one slide. We build focused AI apps for specific professions. Two are live. AppForge, our in-house app factory, lets a small team launch many more. We're raising $450,000 to fund the founder full-time, grow the two apps and launch eight more.");
  }

  // ===================== 3. Problem =====================
  {
    const s = light();
    eyebrow(s, "Problem Statement  ·  Customer pain points");
    title(s, "Generic AI doesn't fit how pastors and reps work");
    const rows = [
      ["LuClock", "Pastors are overloaded", "Sermons, Bible studies, devotionals, kids' lessons and social posts fill every week, often with little or no staff. 42% of U.S. pastors considered quitting full-time ministry in the past year (Barna Group, 2022)."],
      ["LuTarget", "Sales reps practice on real prospects", "Reps learn to handle objections on live calls. One fumbled objection can cost a deal worth thousands in commission, and practice tools are mostly built for big companies."],
      ["LuBrain", "General chatbots don't close the gap", "Tools like ChatGPT don't know these jobs. People spend more time writing prompts than getting their work done."],
    ];
    rows.forEach(([n, h, b], i) => {
      const y = 1.55 + i * 1.08;
      badge(s, n, 0.5, y, 0.58);
      T(s, h, { x: 1.3, y: y - 0.02, w: 8.2, h: 0.32, fontFace: HEAD, fontSize: 16, bold: true, color: GREEN });
      T(s, b, { x: 1.3, y: y + 0.32, w: 8.2, h: 0.7, fontSize: 13, color: MUTED });
    });
    T(s, [{ text: "Why we know: ", options: { bold: true } }, { text: "founder James Chambers watched his bishop's weekly workload, and spent 40+ years in sales." }],
      { x: 0.5, y: 4.85, w: 9, h: 0.3, fontSize: 12.5, color: GREEN });
    s.addNotes("The problem: skilled professionals spend hours every week on repeatable work that AI could speed up, but generic AI makes them do the translating. Pastors carry a week of writing and teaching prep, often alone, and burnout is high: Barna found 42% of U.S. pastors considered quitting full-time ministry in the past year. Sales reps have nowhere affordable to practice before a real call. I've seen both first-hand: my bishop's workload, and 40 years of selling.");
  }

  // ===================== 4. Solution overview =====================
  {
    const s = dark();
    eyebrow(s, "Solution & Product  ·  What Cedarworks offers");
    title(s, "One studio, many focused AI apps", true);
    T(s, "Each app does one job for one profession, with no prompting needed. Every app shares one codebase, one AI stack and one operating system, so each new app costs less to launch and run.", {
      x: 0.5, y: 1.4, w: 9, h: 0.7, fontSize: 14, color: CREAM });
    // AppForge box
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 2.35, w: 2.35, h: 2.35, rectRadius: 0.1, fill: { color: GREEN2 }, line: { color: ORANGE, width: 1.5 } });
    badge(s, "LuFactory", 1.35, 2.55, 0.65, true);
    T(s, "AppForge", { x: 0.6, y: 3.3, w: 2.15, h: 0.35, fontFace: HEAD, fontSize: 18, bold: true, color: CREAM, align: "center" });
    T(s, "our in-house app factory", { x: 0.6, y: 3.68, w: 2.15, h: 0.6, fontSize: 12.5, color: SAGE, align: "center" });
    T(s, "→", { x: 2.9, y: 3.2, w: 0.45, h: 0.5, fontSize: 28, color: ORANGE, align: "center", valign: "middle" });
    // 10 app tiles
    const names = ["PastorAI", "Ready Room", "App 3", "App 4", "App 5", "App 6", "App 7", "App 8", "App 9", "App 10"];
    names.forEach((n, i) => {
      const x = 3.45 + (i % 5) * 1.22, y = 2.35 + Math.floor(i / 5) * 1.22;
      const live = i < 2;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: 1.1, h: 1.1, rectRadius: 0.08,
        fill: { color: live ? ORANGE : GREEN }, line: { color: live ? ORANGE : SAGE, width: 1, dashType: live ? "solid" : "dash" } });
      T(s, n, { x: x + 0.05, y: y + 0.22, w: 1.0, h: 0.4, fontSize: live ? 13 : 13, bold: true, color: live ? WHITE : SAGE, align: "center", valign: "middle" });
      T(s, live ? "live now" : "planned", { x: x + 0.05, y: y + 0.62, w: 1.0, h: 0.3, fontSize: 11, color: live ? WHITE : SAGE, align: "center" });
    });
    T(s, "Apps 3–10 will be chosen from AppForge's market research, not guesswork.", { x: 3.45, y: 4.85, w: 6.05, h: 0.3, fontSize: 12.5, italic: true, color: SAGE });
    s.addNotes("Our solution is a studio, not a single app. Each app does one job for one profession. All of them share our code, AI stack and operating system, so each new app is cheaper to launch and run. Two are live. AppForge, our in-house app factory, will produce the next eight.");
  }

  // ===================== 5. PastorAI =====================
  {
    const s = light();
    eyebrow(s, "Solution & Product  ·  PastorAI  ·  pastorai.io");
    title(s, "PastorAI: an AI ministry assistant for the whole week");
    card(s, 0.5, 1.55, 5.6, 3.5);
    T(s, "What a pastor can create", { x: 0.8, y: 1.72, w: 5.1, h: 0.3, fontSize: 13, bold: true, color: ORANGE });
    const feats = [["Sermon Builder: ", "full outlines with scripture and application"], ["Series Planner: ", "multi-week sermon series"],
      ["Bible Study: ", "small-group guides with questions"], ["Social Media: ", "7 days of posts from one sermon"],
      ["Devotionals and children's lessons", ""], ["Translation ", "into 10 languages"], ["Rehearsal: ", "practice delivering the sermon"]];
    T(s, leadBullets(feats), { x: 0.8, y: 2.08, w: 5.15, h: 2.3, fontSize: 13, paraSpaceAfter: 3 });
    T(s, "“This will save me so much time. It's a great tool.”  — Pastor Dennis, trial user", {
      x: 0.8, y: 4.45, w: 5.15, h: 0.5, fontFace: HEAD, italic: true, fontSize: 12.5, color: GREEN });
    const stats = [["$47/mo", "after a 14-day free trial"], ["~15", "trial users so far"], ["3 platforms", "iOS, Android and web"]];
    stats.forEach(([big, small], i) => statCard(s, 6.4, 1.55 + i * 1.2, 3.1, 1.1, big, small, i === 0, i === 2 ? 22 : 26));
    s.addNotes("PastorAI puts the tools that fill a pastor's week in one place: sermons, series planning, Bible studies, social media, devotionals, children's lessons, translation into 10 languages and sermon rehearsal. It's live on iOS, Android and the web at $47 a month after a 14-day free trial, with about 15 trial users. One of them, Pastor Dennis, told us: 'This will save me so much time. It's a great tool.'");
  }

  // ===================== 6. PastorAI screen =====================
  {
    const s = light();
    eyebrow(s, "Solution & Product  ·  PastorAI  ·  inside the app");
    title(s, "One dashboard for the whole ministry week");
    appShot(s, "pastorai-wide.png", 0.5, 1.55, 6.0, 3.49);
    callouts(s, [["Pick a tool", "Eight ministry tools, one click away in the sidebar."],
      ["Create in minutes", "For example, a full sermon outline with scripture, illustrations and application."],
      ["Try free, then subscribe", "A 14-day free trial, with the $47/month upgrade built into the app."]], 6.8, 1.6, 2.7, 1.15);
    s.addNotes("This is the PastorAI home screen. Every tool is one click away. A pastor picks a tool, answers a few questions, and gets a finished draft in their own voice. The free trial and upgrade are built in, so trials turn into subscriptions inside the app.");
  }

  // ===================== 7. Ready Room =====================
  {
    const s = light();
    eyebrow(s, "Solution & Product  ·  Ready Room  ·  tryreadyroom.com");
    title(s, "Ready Room: a flight simulator for sales calls");
    const stats = [["$29/mo", "or $249/year"], ["~15", "trial users so far"], ["3 platforms", "iOS, Android and web"]];
    stats.forEach(([big, small], i) => statCard(s, 0.5, 1.55 + i * 1.2, 3.1, 1.1, big, small, i === 0, i === 2 ? 22 : 26));
    card(s, 3.9, 1.55, 5.6, 3.5);
    T(s, "How reps practice", { x: 4.2, y: 1.72, w: 5.1, h: 0.3, fontSize: 13, bold: true, color: ORANGE });
    const feats = [["Talk out loud ", "to a voice AI prospect that reacts to how you sell"], ["Two drills: ", "discovery questions and handling objections"],
      ["Three levels: ", "easy, medium and hard"], ["18+ prospects ", "across industries, from roofing to software"],
      ["Score real calls: ", "record live, upload a recording or paste a transcript"], ["AI coach ", "grades every session and says what to fix next"]];
    T(s, leadBullets(feats), { x: 4.2, y: 2.08, w: 5.15, h: 2.4, fontSize: 13, paraSpaceAfter: 3 });
    T(s, "Built on the founder's 40+ years in sales.", { x: 4.2, y: 4.55, w: 5.15, h: 0.4, fontFace: HEAD, italic: true, fontSize: 12.5, color: GREEN });
    s.addNotes("Ready Room lets sales reps rehearse on an AI prospect that talks back out loud. Reps choose a drill, a difficulty and a prospect from 18+ industries, then get graded by an AI coach. They can also upload their real calls for the same scorecard. It's $29 a month or $249 a year, live on iOS, Android and the web, with about 15 trial users. It's built on what I learned in 40 years of selling.");
  }

  // ===================== 8. Ready Room screen =====================
  {
    const s = light();
    eyebrow(s, "Solution & Product  ·  Ready Room  ·  inside the app");
    title(s, "Pick a prospect, pick a drill, start selling");
    appShot(s, "readyroom-wide.png", 0.5, 1.55, 6.0, 3.49);
    callouts(s, [["Choose a drill", "Discovery or objections."], ["Set the difficulty", "Easy, medium or hard."],
      ["Pick a prospect", "18+ industries, each with a realistic AI buyer."], ["Get scored", "An AI coach grades every call, including real calls you upload."]], 6.8, 1.6, 2.7, 0.86);
    s.addNotes("This is Ready Room's practice screen. A rep picks a drill and a difficulty, then picks a prospect, like a homeowner considering a new roof. The AI prospect talks back, and an AI coach scores the call afterward.");
  }

  // ===================== 9. AppForge overview =====================
  {
    const s = dark();
    eyebrow(s, "Solution & Product  ·  AppForge");
    title(s, "AppForge: our in-house app factory", true);
    T(s, "One system that takes us from market research to a finished, marketed app. It's internal only: our competitive edge, not a product we sell.", {
      x: 0.5, y: 1.3, w: 9, h: 0.6, fontSize: 14, color: CREAM });
    const stages = [["LuBrain", "1  Know", "Learns the business, audience and goals", "Brain · Chief of Staff"],
      ["LuSearch", "2  Find", "Finds proven app ideas with real demand", "Top 100 Apps · App Types · Opportunities · Market Analysis · Viability Scorer"],
      ["LuCode", "3  Build", "Plans the app and writes the code", "Projects · Tech Stack · Build · Store Listing"],
      ["LuMegaphone", "4  Market", "Builds a marketing funnel for each app", "Attractive Character · Framework Extractor · One-to-Many Emails, Social, Ads & Presentations"]];
    stages.forEach(([n, h, d, tools], i) => {
      const x = 0.5 + i * 2.3;
      darkCard(s, x, 2.05, 2.1, 3.0);
      badge(s, n, x + 0.2, 2.22, 0.6, true);
      T(s, h, { x: x + 0.2, y: 2.92, w: 1.75, h: 0.32, fontSize: 16, bold: true, color: CREAM });
      T(s, d, { x: x + 0.2, y: 3.27, w: 1.75, h: 0.65, fontSize: 12.5, color: CREAM });
      T(s, tools, { x: x + 0.2, y: 3.95, w: 1.75, h: 1.05, fontSize: 11, color: SAGE });
      if (i < 3) T(s, "›", { x: x + 2.08, y: 3.3, w: 0.24, h: 0.4, fontSize: 22, color: ORANGE, align: "center" });
    });
    s.addNotes("AppForge is the engine behind the plan to go from 2 apps to 10. It works in four steps. Know: it learns about the business and the owner. Find: it finds app ideas with proven demand. Build: it plans the app and writes the code. Market: it builds the marketing funnel. The next four slides walk through each step. AppForge is internal only: it's our competitive advantage, not something we sell.");
  }

  // ===================== 10. AppForge: Know =====================
  {
    const s = light();
    eyebrow(s, "Solution & Product  ·  AppForge step 1: Know");
    title(s, "Brain and Chief of Staff: AppForge learns the business");
    appShot(s, "appforge-brain.png", 0.5, 1.55, 4.8, 3.0);
    T(s, "AppForge's Brain screen", { x: 0.5, y: 4.7, w: 4.8, h: 0.3, fontSize: 11, italic: true, color: MUTED });
    T(s, "Brain", { x: 5.65, y: 1.55, w: 3.85, h: 0.3, fontFace: HEAD, fontSize: 16, bold: true, color: GREEN });
    T(s, "Stores what the AI should know about the owner and business, in 7 areas: business, offers and products, audience, brand voice, story, goals and competitors. It's used to pick apps that fit the owner and to write marketing in their voice.", {
      x: 5.65, y: 1.88, w: 3.85, h: 1.5, fontSize: 12.5, color: INK });
    T(s, "Chief of Staff", { x: 5.65, y: 3.4, w: 3.85, h: 0.3, fontFace: HEAD, fontSize: 16, bold: true, color: GREEN });
    T(s, "An AI business partner grounded in the Brain. Ask it a question, such as “I want to develop an app for…”, and it researches the answer and helps decide what to do next.", {
      x: 5.65, y: 3.73, w: 3.85, h: 1.3, fontSize: 12.5, color: INK });
    s.addNotes("Step one is Know. The Brain stores what the AI should know about the owner and the business, in seven areas: business, offers and products, audience, brand voice, story, goals and competitors. Everything else in AppForge starts from it, so app ideas fit the owner's strengths and marketing sounds like them. The Chief of Staff is an AI business partner grounded in the Brain: you ask it questions, it researches the answers and helps decide what to do next.");
  }

  // ===================== 11. AppForge: Find =====================
  {
    const s = light();
    eyebrow(s, "Solution & Product  ·  AppForge step 2: Find");
    title(s, "Start from apps people already pay for");
    const pts = [["Top 100 Apps Explorer: ", "ranks the top-grossing apps by category and how they make money (subscription, freemium, ads)."],
      ["Build a better version: ", "one click starts a new app from a proven idea, focused on a niche the big apps underserve."],
      ["App Types & Opportunities: ", "suggest specific, scored app ideas."],
      ["Market Analysis & Viability Scorer: ", "check demand, competition and fit before we build."]];
    T(s, leadBullets(pts), { x: 0.5, y: 1.55, w: 3.75, h: 3.5, fontSize: 12.5, paraSpaceAfter: 7 });
    appShot(s, "appforge-top100.png", 4.55, 1.55, 4.95, 3.09);
    T(s, "Top 100 Apps Explorer. Revenue ranges are AppForge's AI estimates.", { x: 4.55, y: 4.75, w: 4.95, h: 0.3, fontSize: 11, italic: true, color: MUTED });
    s.addNotes("Step two is Find. Instead of guessing, AppForge starts from apps people already pay for. The Top 100 Apps Explorer ranks top-grossing apps by category and business model. One click on 'Build a better version' starts a new project aimed at a niche the big apps don't serve well, the way PastorAI focuses on pastors. App Types and Opportunities suggest scored ideas, and Market Analysis and the Viability Scorer check demand, competition and fit before we build anything.");
  }

  // ===================== 12. AppForge: Build =====================
  {
    const s = light();
    eyebrow(s, "Solution & Product  ·  AppForge step 3: Build");
    title(s, "AppForge plans the app and writes the code");
    const cards = [["LuLayoutDashboard", "Projects", "Each app gets its own workspace, plan and milestones."],
      ["LuLayers", "Tech Stack", "Picks the technology, reusing our proven AI, voice and billing setup."],
      ["LuCode", "Build", "Generates the working app code."],
      ["LuStore", "Store Listing", "Writes the App Store and Google Play listing."]];
    cards.forEach(([n, h, b], i) => {
      const x = 0.5 + (i % 2) * 4.6, y = 1.55 + Math.floor(i / 2) * 1.32;
      card(s, x, y, 4.4, 1.17);
      badge(s, n, x + 0.22, y + 0.28, 0.6);
      T(s, h, { x: x + 1.0, y: y + 0.2, w: 3.25, h: 0.32, fontFace: HEAD, fontSize: 16, bold: true, color: GREEN });
      T(s, b, { x: x + 1.0, y: y + 0.55, w: 3.25, h: 0.55, fontSize: 13, color: INK });
    });
    strip(s, 0.5, 4.27, 9, 0.78);
    T(s, [{ text: "Business HQ: ", options: { bold: true } }, { text: "every new app plugs into our shared operating system (dashboards, goals, key metrics, procedures and milestones), so one small team can run ten apps." }],
      { x: 0.75, y: 4.32, w: 8.5, h: 0.68, fontSize: 12.5, color: GREEN, valign: "middle" });
    s.addNotes("Step three is Build. Each app gets a project workspace. AppForge picks the tech stack, reusing the AI, voice and billing setup we already proved in PastorAI and Ready Room, then generates the working app code and writes the app store listing. Every new app also plugs into Business HQ, our shared operating system of dashboards, goals, key metrics and procedures, which is how one small team can run ten apps.");
  }

  // ===================== 13. AppForge: Market =====================
  {
    const s = light();
    eyebrow(s, "Solution & Product  ·  AppForge step 4: Market");
    title(s, "A ready-made marketing funnel for every app");
    const stages = [["Attract", "ads and social posts"], ["Engage", "cold emails and presentations"], ["Convert", "free trial"], ["Subscribe", "monthly or annual plan"]];
    const colors = [GREEN, GREEN2, "4F7F63", ORANGE];
    stages.forEach(([h, b], i) => {
      const w = 4.2 - i * 0.6, x = 0.5 + (4.2 - w) / 2, y = 1.6 + i * 0.85;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h: 0.72, rectRadius: 0.08, fill: { color: colors[i] }, line: { color: colors[i] } });
      T(s, [{ text: h, options: { bold: true, fontSize: 14, breakLine: true } }, { text: b, options: { fontSize: 11.5 } }],
        { x, y: y + 0.06, w, h: 0.62, color: WHITE, align: "center", valign: "middle" });
    });
    const tools = [["Attractive Character: ", "the brand persona and founder story behind the marketing"],
      ["Framework Extractor: ", "turns the owner's know-how into frameworks the content is built on"],
      ["One-to-Many Emails: ", "cold-email campaigns that bring in trial users"],
      ["One-to-Many Social & Ads: ", "social posts and ad copy"],
      ["One-to-Many Presentations: ", "presentations that sell to many people at once"]];
    T(s, leadBullets(tools), { x: 5.0, y: 1.55, w: 4.5, h: 3.0, fontSize: 12.5, paraSpaceAfter: 5 });
    T(s, "All of it is written from the Brain, in the owner's voice.", { x: 5.0, y: 4.7, w: 4.5, h: 0.32, fontSize: 12.5, italic: true, color: GREEN });
    s.addNotes("Step four is Market. AppForge builds a funnel for each app: ads and social posts attract people, cold emails and presentations engage them, the free trial converts them, and they subscribe. Attractive Character builds the brand persona and story, Framework Extractor turns know-how into frameworks the content is built on, and the One-to-Many tools write the emails, social posts, ads and presentations. All of it comes from the Brain, so it sounds like the owner.");
  }

  // ===================== 14. Why AppForge matters =====================
  {
    const s = dark();
    eyebrow(s, "Solution & Product  ·  Why AppForge matters");
    title(s, "Why AppForge matters to investors", true);
    const pts = [["LuSearch", "Less guesswork", "Every new app starts from proven demand and a viability score."],
      ["LuZap", "Faster launches", "AppForge writes the code, so a new app takes far less time to build."],
      ["LuDollarSign", "Lower cost per app", "Shared code, AI stack and Business HQ keep each app's cost low."],
      ["LuMegaphone", "Marketing from day one", "Every app launches with its funnel, emails and ads ready."]];
    pts.forEach(([n, h, b], i) => {
      const x = 0.5 + (i % 2) * 4.6, y = 1.5 + Math.floor(i / 2) * 1.3;
      darkCard(s, x, y, 4.4, 1.15);
      badge(s, n, x + 0.22, y + 0.27, 0.6, true);
      T(s, h, { x: x + 1.0, y: y + 0.18, w: 3.25, h: 0.32, fontSize: 16, bold: true, color: CREAM });
      T(s, b, { x: x + 1.0, y: y + 0.52, w: 3.25, h: 0.6, fontSize: 13, color: SAGE });
    });
    T(s, "PastorAI and Ready Room were built before AppForge, and the lessons from building them went into it. Apps 3–10 will be the first built with AppForge.", {
      x: 0.5, y: 4.2, w: 9, h: 0.75, fontFace: HEAD, italic: true, fontSize: 13.5, color: CREAM });
    s.addNotes("Why this matters: AppForge turns building apps into a repeatable process. New apps start from proven demand, take less time because AppForge writes the code, cost less because everything is shared, and launch with marketing ready. To be clear, PastorAI and Ready Room were built before AppForge; everything we learned building them went into it. Apps 3 through 10 will be the first built with it.");
  }

  // ===================== 15. Traction =====================
  {
    const s = light();
    eyebrow(s, "Traction & Metrics");
    title(s, "Built, shipped and live, before raising a dollar");
    const stats = [["2", "apps live in the app stores"], ["3", "platforms: iOS, Android, web"], ["~30", "trial users across both apps"], ["$0 MRR", "pre-revenue · team of 1"]];
    stats.forEach(([big, small], i) => statCard(s, 0.5 + i * 2.3, 1.5, 2.1, 1.15, big, small, i === 3, i === 3 ? 24 : 30));
    card(s, 0.5, 2.85, 4.4, 2.2);
    T(s, "Milestones reached", { x: 0.75, y: 2.98, w: 3.9, h: 0.28, fontSize: 13, bold: true, color: ORANGE });
    T(s, bullets(["Cedarworks Apps LLC formed (2026)", "PastorAI live on App Store, Google Play and web", "Ready Room live on App Store, Google Play and web",
      "Subscriptions and free trials working", "AppForge built for apps 3–10", "First testimonial from a trial user"]), { x: 0.75, y: 3.3, w: 4.0, h: 1.7, fontSize: 12, paraSpaceAfter: 1 });
    card(s, 5.1, 2.85, 4.4, 2.2);
    T(s, "Goals for the next 90 days", { x: 5.35, y: 2.98, w: 3.9, h: 0.28, fontSize: 13, bold: true, color: ORANGE });
    const kpi = [["Paying subscribers", "75"], ["Monthly revenue (MRR)", "~$2,850"], ["Trials that convert to paid", "20%+"], ["Subscribers who cancel each month", "under 6%"]];
    kpi.forEach(([k, v], i) => {
      const y = 3.35 + i * 0.4;
      T(s, k, { x: 5.35, y, w: 2.9, h: 0.32, fontSize: 12.5, color: INK });
      T(s, v, { x: 8.2, y, w: 1.1, h: 0.32, fontSize: 12.5, bold: true, color: GREEN, align: "right" });
    });
    s.addNotes("We're at the very start. Both apps are complete and live on three platforms, with subscriptions and free trials working, and about 30 people on free trials. We don't have paying customers yet. In the next 90 days we're aiming for 75 paying subscribers, about $2,850 a month in revenue, at least 20% of trials converting to paid, and fewer than 6% of subscribers cancelling each month. We'll report these to investors monthly.");
  }

  // ===================== 16. Market =====================
  {
    const s = light();
    eyebrow(s, "Market Opportunity");
    title(s, "A ~$690M U.S. market for our first two apps");
    s.addShape(pres.shapes.OVAL, { x: 0.6, y: 1.5, w: 3.0, h: 3.0, fill: { color: TAN }, line: { color: TAN } });
    s.addShape(pres.shapes.OVAL, { x: 1.125, y: 2.45, w: 1.95, h: 1.95, fill: { color: SAGE }, line: { color: SAGE } });
    s.addShape(pres.shapes.OVAL, { x: 1.625, y: 3.35, w: 0.95, h: 0.95, fill: { color: GREEN }, line: { color: GREEN } });
    T(s, "TAM ~$690M", { x: 0.6, y: 1.78, w: 3.0, h: 0.3, fontSize: 14, bold: true, color: GREEN, align: "center" });
    T(s, "SAM ~$207M", { x: 1.125, y: 2.72, w: 1.95, h: 0.3, fontSize: 13, bold: true, color: GREEN, align: "center" });
    T(s, "SOM\n~$2M", { x: 1.625, y: 3.52, w: 0.95, h: 0.6, fontSize: 11.5, bold: true, color: CREAM, align: "center" });
    const rows = [
      ["Total market (TAM): ~$690M", "Everyone who could buy: 300K+ U.S. churches × $564/yr (≈$169M) plus ~1.5M U.S. B2B sales reps × $348/yr (≈$522M)."],
      ["Reachable market (SAM): ~$207M", "The ~30% we can reach with online, self-serve sales."],
      ["Our 3-year target (SOM): ~$2M a year", "1% of the reachable market: about 4,500 subscribers."],
    ];
    rows.forEach(([h, b], i) => {
      const y = 1.5 + i * 1.0;
      T(s, h, { x: 4.0, y, w: 5.5, h: 0.3, fontSize: 14, bold: true, color: GREEN });
      T(s, b, { x: 4.0, y: y + 0.32, w: 5.5, h: 0.62, fontSize: 12.5, color: MUTED });
    });
    T(s, "Each of the 8 AppForge apps adds a new market on top of these two.", { x: 4.0, y: 4.5, w: 5.5, h: 0.3, fontSize: 12.5, italic: true, color: GREEN });
    T(s, "Sources: Hartford Institute for Religion Research and U.S. Religion Census (congregations); U.S. Bureau of Labor Statistics (wholesale and manufacturing sales reps). Prices: current PastorAI and Ready Room plans.", {
      x: 0.5, y: 4.83, w: 9, h: 0.34, fontSize: 9.5, italic: true, color: MUTED });
    s.addNotes("Bottom-up sizing at our current prices. The total market counts everyone who could buy: roughly 300,000+ U.S. congregations at $564 a year, plus about 1.5 million U.S. B2B sales reps at $348 a year, around $690 million. About 30% of that, roughly $207 million, is reachable with online self-serve sales. Our three-year target is 1% of that, about $2 million a year from roughly 4,500 subscribers. Every new AppForge app adds a market on top.");
  }

  // ===================== 17. Why now =====================
  {
    const s = light();
    eyebrow(s, "Market Opportunity  ·  Why now & customer segments");
    title(s, "Why the timing is right");
    const cards = [["LuZap", "AI got affordable", "The cost of running language and voice AI has fallen sharply since 2023, so $29–$47/month apps can earn healthy margins."],
      ["LuBrain", "People expect AI help", "Most professionals have tried AI chat and found it doesn't fit their daily work. Focused apps fill that gap."],
      ["LuRocket", "Small teams can ship fast", "App stores handle distribution and billing, and AppForge lets one team launch app after app."]];
    cards.forEach(([n, h, b], i) => {
      const x = 0.5 + i * 3.07;
      card(s, x, 1.5, 2.85, 2.4);
      badge(s, n, x + 0.25, 1.68, 0.55);
      T(s, h, { x: x + 0.25, y: 2.35, w: 2.4, h: 0.32, fontFace: HEAD, fontSize: 15, bold: true, color: GREEN });
      T(s, b, { x: x + 0.25, y: 2.7, w: 2.4, h: 1.15, fontSize: 12.5, color: MUTED });
    });
    strip(s, 0.5, 4.08, 9, 0.97);
    T(s, [{ text: "Who we sell to: ", options: { bold: true } },
      { text: "pastors and staff at small and mid-size churches; individual sales reps and small sales teams; sales managers and church networks buying group plans; and new niches chosen by AppForge." }],
      { x: 0.75, y: 4.14, w: 8.5, h: 0.85, fontSize: 12.5, color: GREEN, valign: "middle" });
    s.addNotes("Why now: AI costs have dropped enough to support affordable subscription apps; professionals have tried generic AI and want tools built for their work; and app stores plus AppForge let a small studio launch quickly. Our customers start with pastors, church staff and individual sales reps, expand to group plans, then to new niches AppForge finds.");
  }

  // ===================== 18. Business model =====================
  {
    const s = light();
    eyebrow(s, "Business Model  ·  How we make money");
    title(s, "Recurring subscriptions across a growing set of apps");
    const streams = [["LuRepeat", "App subscriptions", "Monthly or annual plans with a free trial first, sold through the App Store, Google Play and the web."],
      ["LuUsers", "Group plans", "Multi-seat plans for sales teams and church staffs: one buyer, several users."],
      ["LuWrench", "Custom AI apps", "Custom builds for local businesses: extra revenue and new app ideas."]];
    streams.forEach(([n, h, b], i) => {
      const x = 0.5 + i * 3.07;
      card(s, x, 1.5, 2.85, 2.3);
      badge(s, n, x + 0.25, 1.68, 0.55);
      T(s, h, { x: x + 0.25, y: 2.35, w: 2.4, h: 0.32, fontFace: HEAD, fontSize: 15, bold: true, color: GREEN });
      T(s, b, { x: x + 0.25, y: 2.7, w: 2.4, h: 1.05, fontSize: 12.5, color: MUTED });
    });
    strip(s, 0.5, 4.0, 9, 1.05);
    T(s, [{ text: "Pricing today: ", options: { bold: true } }, { text: "PastorAI $47/month after a 14-day free trial · Ready Room $29/month or $249/year.", options: { breakLine: true } },
      { text: "Why it scales: ", options: { bold: true } }, { text: "every app shares one codebase, AI stack and operating system, so each new app adds revenue with little extra cost." }],
      { x: 0.75, y: 4.07, w: 8.5, h: 0.92, fontSize: 12.5, color: GREEN, valign: "middle" });
    s.addNotes("We make money from subscriptions: monthly or annual plans after a free trial. Group plans let one buyer, like a sales manager or a church, pay for several users. Custom app work for local businesses is a smaller extra revenue line. Because every app shares the same foundation, each new app adds revenue without adding much cost.");
  }

  // ===================== 19. Unit economics =====================
  {
    const s = light();
    eyebrow(s, "Business Model  ·  Unit economics (what each subscriber is worth)");
    title(s, "Each subscriber is worth ~3× what it costs to win");
    const cell = (t, o = {}) => ({ text: t, options: Object.assign({ fontFace: BODY, fontSize: 12.5, color: INK, valign: "middle" }, o) });
    const hdrO = (i) => ({ bold: true, color: i ? CREAM : GREEN, fill: { color: i ? GREEN : TAN }, align: i ? "center" : "left" });
    const rows = [["Monthly price", "$47", "$29"], ["What we keep after the 15% app-store fee", "$39.95", "$24.65"],
      ["Estimated AI and hosting cost per user", "~$4", "~$5"], ["Gross margin (share of revenue left after those costs)", "~77%", "~68%"],
      ["Lifetime value: profit over an average ~20-month subscription", "~$720", "~$390"],
      ["Most we'll spend to win one customer (CAC)", "$240", "$130"], ["Months to earn that cost back", "~7", "~7"]];
    const data = [["Per subscriber", "PastorAI", "Ready Room"].map((h, i) => cell(h, hdrO(i)))]
      .concat(rows.map(r => r.map((c, i) => cell(c, { align: i ? "center" : "left", bold: i > 0, color: i ? GREEN : INK, fill: { color: WHITE } }))));
    s.addTable(data, { x: 0.5, y: 1.5, w: 9, colW: [5.6, 1.7, 1.7], rowH: 0.42, border: { type: "solid", pt: 0.75, color: TAN }, margin: [0, 0.12, 0, 0.12] });
    T(s, "Rows after the price are planning assumptions (5% of subscribers cancel each month), not results yet.", { x: 0.5, y: 4.95 - 0.1, w: 9, h: 0.3, fontSize: 11, italic: true, color: MUTED });
    s.addNotes("Unit economics in plain terms. After the app-store fee and AI costs, we keep roughly 77% of PastorAI revenue and 68% of Ready Room revenue. If the average subscriber stays about 20 months, a PastorAI subscriber is worth about $720 in profit and a Ready Room subscriber about $390. We'll spend no more than a third of that to win each customer, so we earn the cost back in about seven months. These are assumptions until we have real data.");
  }

  // ===================== 20. Go-to-market =====================
  {
    const s = light();
    eyebrow(s, "Go-to-Market Strategy");
    title(s, "How we'll turn trials into paying customers");
    card(s, 0.5, 1.5, 4.4, 3.0);
    badge(s, "LuChurch", 0.75, 1.68, 0.5);
    T(s, "PastorAI", { x: 1.4, y: 1.76, w: 3.2, h: 0.35, fontFace: HEAD, fontSize: 16, bold: true, color: GREEN });
    T(s, bullets(["Facebook and YouTube ads aimed at pastors and church staff", "LinkedIn outreach to ministry leaders",
      "Partnerships with denominations, church networks and seminaries", "Conferences, webinars and free sermon-prep resources"]),
      { x: 0.75, y: 2.35, w: 3.95, h: 2.1, fontSize: 12.5, paraSpaceAfter: 5 });
    card(s, 5.1, 1.5, 4.4, 3.0);
    badge(s, "LuBriefcase", 5.35, 1.68, 0.5);
    T(s, "Ready Room", { x: 6.0, y: 1.76, w: 3.2, h: 0.35, fontFace: HEAD, fontSize: 16, bold: true, color: GREEN });
    T(s, bullets(["LinkedIn and short videos showing live AI role-plays", "Direct outreach to sales managers for team plans",
      "Partnerships with sales trainers, coaches and bootcamps", "App store search and referral rewards"]),
      { x: 5.35, y: 2.35, w: 3.95, h: 2.1, fontSize: 12.5, paraSpaceAfter: 5 });
    T(s, "Powered by AppForge's marketing tools: cold emails, social posts, ads and presentations for every app.", { x: 0.5, y: 4.68, w: 9, h: 0.35, fontSize: 12.5, italic: true, color: GREEN });
    s.addNotes("The biggest use of funds is winning customers. We'll run small, measured tests on each channel, then put more money into whichever channels win paying subscribers most cheaply. AppForge writes the emails, posts, ads and presentations, so marketing each app takes far less time.");
  }

  // ===================== 21. Competition =====================
  {
    const s = light();
    eyebrow(s, "Competitive Landscape  ·  Direct & indirect competitors");
    title(s, "Why Cedarworks wins");
    const hdr = ["", "Indirect: generic AI chat (e.g. ChatGPT)", "Direct: ministry software (e.g. Logos, Pulpit AI)", "Direct: sales practice (e.g. Second Nature, Hyperbound)", "Cedarworks apps"];
    const rows = [["Built for one profession", "No", "Yes", "Yes", "Yes"], ["Ready to use with no prompting", "No", "Partly", "Yes", "Yes"],
      ["Made for small churches and individual reps", "Yes", "Partly", "No (big companies)", "Yes"], ["Self-serve monthly price", "Low", "Varies", "Sales call", "$29–$47"],
      ["Launches new niche apps quickly", "N/A", "No", "No", "AppForge"]];
    const cell = (t, o = {}) => ({ text: t, options: Object.assign({ fontFace: BODY, fontSize: 11.5, color: INK, valign: "middle" }, o) });
    const data = [hdr.map((h, i) => cell(h, { bold: true, fontSize: 11, color: i === 4 ? CREAM : GREEN, fill: { color: i === 4 ? GREEN : TAN }, align: i ? "center" : "left" }))]
      .concat(rows.map(r => r.map((c, i) => cell(c, { align: i ? "center" : "left", bold: i === 4, color: i === 4 ? GREEN : INK, fill: { color: i === 4 ? "E3EDE3" : WHITE } }))));
    s.addTable(data, { x: 0.5, y: 1.45, w: 9, colW: [2.5, 1.6, 1.7, 1.85, 1.35], rowH: [0.75, 0.42, 0.42, 0.42, 0.42, 0.42], border: { type: "solid", pt: 0.75, color: TAN }, margin: [0, 0.08, 0, 0.08] });
    T(s, [{ text: "Our advantages: ", options: { bold: true } }, { text: "we focus on niches big platforms ignore, at prices individuals can afford, with AppForge to reach new markets fast and a founder who can both build and sell." }], {
      x: 0.5, y: 4.4, w: 9, h: 0.65, fontSize: 12.5, color: GREEN });
    s.addNotes("General chatbots make people do the prompting and give generic results. Ministry software like Logos focuses on study and research, and Pulpit AI on repurposing sermon content. Sales-practice platforms like Second Nature and Hyperbound are sold to company sales teams. We sit in between: purpose-built, affordable and self-serve, and AppForge lets us keep adding niches.");
  }

  // ===================== 22. Use of funds =====================
  {
    const s = light();
    eyebrow(s, "Use of Funds");
    title(s, "$450K to grow two apps and launch eight more");
    const funds = [["Marketing and outreach for all apps", 160], ["Founder salary: 18 months full-time", 120], ["Building 8 new apps with AppForge", 90],
      ["AI, hosting and infrastructure", 30], ["Legal, accounting and operations", 25], ["Reserve", 25]];
    const colors = [GREEN, ORANGE, "4F7F63", "8FA98E", "B9A67A", "D9CBA6"];
    s.addChart(pres.charts.DOUGHNUT, [{ name: "Use of funds", labels: funds.map(f => f[0]), values: funds.map(f => f[1]) }], {
      x: 0.4, y: 1.4, w: 3.7, h: 3.6, holeSize: 58, chartColors: colors, showLegend: false, showPercent: true, showValue: false,
      dataLabelColor: WHITE, dataLabelFontSize: 12, dataLabelFontBold: true, showTitle: false });
    funds.forEach(([label, v], i) => {
      const y = 1.45 + i * 0.54;
      s.addShape(pres.shapes.OVAL, { x: 4.5, y: y + 0.1, w: 0.22, h: 0.22, fill: { color: colors[i] }, line: { color: colors[i] } });
      T(s, `$${v}K`, { x: 4.85, y, w: 0.95, h: 0.42, fontFace: HEAD, fontSize: 18, bold: true, color: GREEN, valign: "middle" });
      T(s, label, { x: 5.85, y, w: 3.65, h: 0.42, fontSize: 13.5, color: INK, valign: "middle" });
    });
    T(s, "About 18 months of runway on a lean, founder-led budget.", { x: 4.5, y: 4.8, w: 5, h: 0.3, fontSize: 12.5, italic: true, color: MUTED });
    s.addNotes("$160,000 (about a third) goes to marketing and outreach, our fastest path to revenue. $120,000 is my salary for 18 months, about $80,000 a year, so I can work on Cedarworks full-time and build the next apps. $90,000 funds building eight new apps with AppForge, including contract development, design and testing. The rest covers AI and hosting, legal and operations, and a reserve. That's about 18 months of runway.");
  }

  // ===================== 23. Roadmap =====================
  {
    const s = light();
    eyebrow(s, "Roadmap & Milestones");
    title(s, "From 2 apps to 10 in 18 months");
    const ms = [["Months 0–3", "Turn trials into paying users; test ads; add group plans"],
      ["Months 4–6", "Grow the best ad channels; launch apps 3 and 4"],
      ["Months 7–12", "Launch apps 5 to 7; promote apps to each other's users"],
      ["Months 13–18", "Launch apps 8 to 10; raise a seed round to market all 10 apps at scale"]];
    s.addShape(pres.shapes.LINE, { x: 0.68, y: 1.75, w: 0, h: 2.95, line: { color: TAN, width: 2 } });
    ms.forEach(([h, b], i) => {
      const y = 1.5 + i * 0.86;
      s.addShape(pres.shapes.OVAL, { x: 0.5, y: y + 0.05, w: 0.36, h: 0.36, fill: { color: i === 0 ? ORANGE : WHITE }, line: { color: ORANGE, width: 1.5 } });
      T(s, h, { x: 1.05, y, w: 4.5, h: 0.3, fontSize: 14, bold: true, color: GREEN });
      T(s, b, { x: 1.05, y: y + 0.32, w: 4.5, h: 0.5, fontSize: 12.5, color: MUTED });
    });
    s.addChart(pres.charts.BAR, [{ name: "Paying subscribers (goal)", labels: ["Month 3", "6", "9", "12", "18"], values: [75, 200, 400, 650, 1000] }], {
      x: 5.75, y: 1.4, w: 3.8, h: 3.3, barDir: "col", chartColors: [GREEN], showTitle: true, title: "Paying subscribers (goal)",
      titleFontSize: 12, titleColor: GREEN, titleFontFace: BODY, showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 11,
      dataLabelColor: INK, catAxisLabelColor: MUTED, valAxisHidden: true, catAxisLabelFontSize: 11,
      valGridLine: { style: "none" }, catGridLine: { style: "none" }, showLegend: false, barGapWidthPct: 55 });
    T(s, "Goals, not guarantees: 1,000 subscribers at ~$38 average ≈ $38K a month in revenue.", { x: 5.75, y: 4.72, w: 3.8, h: 0.4, fontSize: 11, italic: true, color: MUTED });
    s.addNotes("Our 18-month goal is ten live apps and about 1,000 paying subscribers. At an average of about $38 a month, that's around $38,000 a month in revenue, which sets us up to raise a seed round on real traction. That next round pays to market apps 3 through 10 at scale. These are goals, not guarantees.");
  }

  // ===================== 24. Financial projections =====================
  {
    const s = light();
    eyebrow(s, "Financial Projections  ·  5-year plan");
    title(s, "Profitable in year 3, ~$4.8M revenue by year 5");
    const yrs = ["Year 1", "Year 2", "Year 3", "Year 4", "Year 5"];
    const subs = [650, 2000, 4500, 8000, 13000];
    const rev = [128, 604, 1482, 2850, 4788];
    const cogs = rev.map(r => Math.round(r * 0.27));
    const opex = [340, 600, 900, 1400, 2000];
    const net = rev.map((r, i) => r - cogs[i] - opex[i]);
    const k = v => (v < 0 ? "−$" : "$") + (Math.abs(v) >= 1000 ? (Math.abs(v) / 1000).toFixed(2) + "M" : Math.abs(v) + "K");
    const cell = (t, o = {}) => ({ text: t, options: Object.assign({ fontFace: BODY, fontSize: 12, color: INK, valign: "middle", align: "center" }, o) });
    const data = [[cell("", { fill: { color: TAN } })].concat(yrs.map(y => cell(y, { bold: true, color: CREAM, fill: { color: GREEN } })))];
    const row = (label, vals, o = {}) => [cell(label, { align: "left", fill: { color: WHITE }, bold: !!o.bold })].concat(vals.map(v => cell(v, Object.assign({ fill: { color: WHITE } }, o))));
    data.push(row("Paying subscribers (end of year)", subs.map(v => v.toLocaleString("en-US"))));
    data.push(row("Revenue", rev.map(k)));
    data.push(row("App-store fees, AI and hosting (~27%)", cogs.map(k)));
    data.push(row("Operating costs (marketing, salaries, other)", opex.map(k)));
    data.push(row("Profit (loss)", net.map(k), { bold: true, color: GREEN }));
    s.addTable(data, { x: 0.5, y: 1.45, w: 9, colW: [3.5, 1.1, 1.1, 1.1, 1.1, 1.1], rowH: 0.42, border: { type: "solid", pt: 0.75, color: TAN }, margin: [0, 0.1, 0, 0.1] });
    T(s, "Key assumptions", { x: 0.5, y: 4.1, w: 9, h: 0.28, fontSize: 13, bold: true, color: ORANGE });
    T(s, "Average price ~$38/month · ~5% of subscribers cancel each month · new apps launch on the roadmap above · this $450K round covers about 18 months, including the founder's salary, and a seed round around month 18 funds marketing apps 3–10 at scale.", {
      x: 0.5, y: 4.4, w: 9, h: 0.6, fontSize: 12, color: MUTED });
    s.addNotes("Five-year projections, built up from subscriber numbers. Paying subscribers grow from 650 to 13,000 as AppForge adds apps. Revenue uses an average price of about $38 a month. About 27% of revenue goes to app-store fees, AI and hosting. Operating costs cover marketing, the founder's salary, contractors and later hires. We expect losses in years 1 and 2, funded by this round and a seed round, and to become profitable in year 3. These are projections, not guarantees.");
  }

  // ===================== 25. Team =====================
  {
    const s = light();
    eyebrow(s, "Team & Founders");
    title(s, "A founder who can both build and sell");
    card(s, 0.5, 1.5, 4.6, 3.55);
    badge(s, "LuUser", 0.78, 1.72, 0.8);
    T(s, "James Chambers", { x: 1.78, y: 1.8, w: 3.2, h: 0.38, fontFace: HEAD, fontSize: 19, bold: true, color: GREEN });
    T(s, "Founder & CEO", { x: 1.78, y: 2.18, w: 3.2, h: 0.3, fontSize: 13, color: ORANGE, bold: true });
    T(s, bullets(["Degree in computer science", "40+ years in sales across many industries", "Started Cedarworks after seeing the workload his bishop carried",
      "Built PastorAI, Ready Room and AppForge himself", "Will work on Cedarworks full-time with this round"]), { x: 0.78, y: 2.75, w: 4.15, h: 2.2, fontSize: 13, paraSpaceAfter: 5 });
    T(s, "Hiring with this round", { x: 5.4, y: 1.5, w: 4.1, h: 0.3, fontSize: 13, bold: true, color: ORANGE });
    const hires = [["LuWrench", "Contract developer", "Speeds up launching the next eight apps"],
      ["LuMegaphone", "Part-time growth marketer", "Runs ads, outreach and partnerships"],
      ["LuHandshake", "Advisors", "Ministry, sales-training and software mentors"]];
    hires.forEach(([n, h, b], i) => {
      const y = 1.95 + i * 1.0;
      badge(s, n, 5.4, y, 0.58);
      T(s, h, { x: 6.15, y, w: 3.35, h: 0.3, fontSize: 15, bold: true, color: GREEN });
      T(s, b, { x: 6.15, y: y + 0.32, w: 3.35, h: 0.5, fontSize: 12.5, color: MUTED });
    });
    s.addNotes("I have a degree in computer science and more than 40 years in sales across different industries, so I can build the products and sell them. Ready Room comes straight from my sales career. PastorAI started when I saw how much my bishop carries every week. I built both apps and AppForge myself. This round lets me work on Cedarworks full-time and adds a contract developer and a part-time marketer, without a heavy payroll.");
  }

  // ===================== 26. The Ask =====================
  {
    const s = dark();
    logo(s, 0.5, 0.45, 0.65);
    eyebrow(s, "The Ask", 1.3);
    T(s, "Raising $450,000 to take Cedarworks from 2 apps to 10", { x: 0.5, y: 1.6, w: 5.7, h: 1.25, fontFace: HEAD, fontSize: 27, bold: true, color: CREAM });
    T(s, bullets(["Fund the founder to build full-time", "Market PastorAI and Ready Room to paying users", "Launch 8 new apps with AppForge, reaching ~1,000 paying subscribers in 18 months"]),
      { x: 0.5, y: 2.85, w: 5.7, h: 1.35, fontSize: 14, color: CREAM, paraSpaceAfter: 4 });
    T(s, "Use of funds: $160K marketing · $120K founder salary (18 months) · $90K new apps · $30K AI and hosting · $25K operations · $25K reserve", { x: 0.5, y: 4.15, w: 5.7, h: 0.5, fontSize: 12.5, color: SAGE });
    T(s, "Pre-seed · terms open to discussion (SAFE or priced round)", { x: 0.5, y: 4.7, w: 5.7, h: 0.3, fontSize: 12.5, italic: true, color: SAGE });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 6.5, y: 1.45, w: 3.0, h: 3.3, rectRadius: 0.1, fill: { color: GREEN2 }, line: { color: ORANGE, width: 1.25 } });
    T(s, "Let's talk", { x: 6.75, y: 1.62, w: 2.5, h: 0.3, fontSize: 13, bold: true, color: ORANGE });
    T(s, "James Chambers", { x: 6.75, y: 1.98, w: 2.5, h: 0.4, fontFace: HEAD, fontSize: 19, bold: true, color: CREAM });
    const contact = [["LuMail", "james@cedarworksapps.com"], ["LuPhone", "417-370-2084"], ["LuGlobe", "cedarworksapps.com"], ["LuSmartphone", "pastorai.io · tryreadyroom.com"]];
    contact.forEach(([n, t], i) => {
      const y = 2.6 + i * 0.5;
      s.addImage({ data: ic[n].o, x: 6.75, y: y + 0.05, w: 0.26, h: 0.26 });
      T(s, t, { x: 7.12, y, w: 2.3, h: 0.36, fontSize: 12, color: CREAM, valign: "middle" });
    });
    s.addNotes("We're raising $450,000 pre-seed to fund me full-time, market our two live apps and launch eight more with AppForge, reaching about 1,000 paying subscribers in 18 months. Thank you. I'd love to show you a live demo of PastorAI, Ready Room and AppForge.");
  }

  // ===================== 27. Key terms =====================
  {
    const s = light();
    eyebrow(s, "Appendix");
    title(s, "Key terms used in this deck");
    const terms = [["TAM / SAM / SOM", "Total market / the part we can reach / our target share"],
      ["MRR / ARR", "Monthly / annual recurring revenue from subscriptions"],
      ["Churn", "Share of subscribers who cancel each month"],
      ["Trial-to-paid conversion", "Share of free-trial users who become paying subscribers"],
      ["Gross margin", "Share of revenue left after app-store fees, AI and hosting"],
      ["LTV (lifetime value)", "Profit from one subscriber over their whole subscription"],
      ["CAC", "Customer acquisition cost: spend to win one subscriber"],
      ["Pre-seed / SAFE", "First outside funding round / an agreement that turns into shares later"]];
    terms.forEach(([h, b], i) => {
      const x = 0.5 + (i % 2) * 4.6, y = 1.4 + Math.floor(i / 2) * 0.94;
      card(s, x, y, 4.4, 0.86);
      T(s, h, { x: x + 0.2, y: y + 0.08, w: 4.0, h: 0.28, fontSize: 13, bold: true, color: GREEN });
      T(s, b, { x: x + 0.2, y: y + 0.37, w: 4.05, h: 0.46, fontSize: 11.5, color: MUTED });
    });
    s.addNotes("A quick reference for the business terms used in this deck.");
  }

  // ---- footers with page numbers ----
  slides.forEach(([s, isDark], i) => {
    s.addText("Cedarworks Apps  ·  Confidential", { x: 0.5, y: 5.24, w: 5, h: 0.25, fontFace: BODY, fontSize: 9, color: isDark ? SAGE : MUTED, margin: 0, isTextBox: true });
    s.addText(`${i + 1} / ${slides.length}`, { x: 8.5, y: 5.24, w: 1, h: 0.25, fontFace: BODY, fontSize: 9, align: "right", color: isDark ? SAGE : MUTED, margin: 0, isTextBox: true });
  });

  await pres.writeFile({ fileName: OUT });
  console.log("wrote", OUT, slides.length, "slides");
})();
