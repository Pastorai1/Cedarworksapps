// Cedarworks Apps investor pitch deck (v4: 18-slide main deck + appendix)
// Usage: node build.js [output.pptx]   (images are read from this script's folder)
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const I = require("react-icons/lu");
const fs = require("fs"), path = require("path");

const OUT = process.argv[2] || "Cedarworks-Apps-Pitch-Deck.pptx";
// Optional personal cell number, kept out of the public repo: CELL_PHONE=xxx-xxx-xxxx node build.js
const CELL = process.env.CELL_PHONE || "";

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
    "LuMessagesSquare", "LuChartLine", "LuPackage", "LuSend", "LuDollarSign", "LuFlag", "LuLockOpen", "LuBuilding2"]) {
    ic[n] = { o: await icon(n, ORANGE), g: await icon(n, GREEN), w: await icon(n, WHITE) };
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
    T(s, "James Chambers, Founder  ·  James@cedarworksapps.com  ·  cedarworksapps.com\nOffice 417-370-2084" + (CELL ? "  ·  Cell " + CELL : ""), { x: 0.5, y: 4.45, w: 9, h: 0.55, fontSize: 12, color: CREAM, lineSpacingMultiple: 1.15 });
    s.addNotes("Cedarworks Apps is a Springfield, Missouri AI app studio. We build focused, subscription AI apps for specific professions. Two are live in the Apple App Store and Google Play today. We also built AppForge, our in-house app factory, which finds proven app ideas, writes the code and builds the marketing. We're raising a $450K pre-seed round so I can build full-time, market our two apps and launch eight more.");
  }

  // ===================== Executive summary =====================
  {
    const s = light();
    eyebrow(s, "Executive Summary");
    title(s, "Cedarworks Apps in one minute");
    const items = [
      ["LuLayers", "What we do", "We build focused AI apps for specific professions and sell them as subscriptions."],
      ["LuSmartphone", "Live today", "PastorAI ($47/mo) for pastors and Ready Room ($29/mo) for sales reps, on iOS, Android and web."],
      ["LuFlag", "Focus first", "The next 6 months prove these two apps. New apps unlock only after we hit 150 paying subscribers."],
      ["LuRocket", "The ask", "$450,000 pre-seed: founder full-time, cold email and marketing, then AppForge apps 3–10."],
    ];
    items.forEach(([n, h, b], i) => {
      const x = 0.5 + (i % 2) * 4.6, y = 1.55 + Math.floor(i / 2) * 1.75;
      card(s, x, y, 4.4, 1.55);
      badge(s, n, x + 0.22, y + 0.25, 0.6);
      T(s, h, { x: x + 1.0, y: y + 0.22, w: 3.2, h: 0.32, fontFace: HEAD, fontSize: 16, bold: true, color: GREEN });
      T(s, b, { x: x + 1.0, y: y + 0.58, w: 3.25, h: 0.9, fontSize: 13, color: INK });
    });
    s.addNotes("The whole story in one slide. We build focused AI apps for specific professions. Two are live. For the next six months we focus on proving those two apps can win paying customers; new apps built with AppForge, our in-house app factory, unlock only after we reach 150 paying subscribers. We're raising $450,000 to make that happen.");
  }

  // ===================== 3. Problem =====================
  {
    const s = light();
    eyebrow(s, "Problem Statement  ·  Customer pain points");
    title(s, "Generic AI doesn't fit how pastors and reps work");
    const rows = [
      ["LuClock", "Pastors are overloaded", "Sermons, Bible studies, devotionals, kids' lessons and social posts fill every week, often with little or no staff. Most of it starts from a blank page, and it leaves little time for people."],
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
    s.addNotes("The problem: skilled professionals spend hours every week on repeatable work that AI could speed up, but generic AI makes them do the translating. Pastors carry a week of writing and teaching prep, often alone, which leaves less time for their people. Sales reps have nowhere affordable to practice before a real call. I've seen both first-hand: my bishop's workload, and 40 years of selling.");
  }

  // ===================== Solution overview =====================
  {
    const s = dark();
    eyebrow(s, "Solution & Product  ·  What Cedarworks offers");
    title(s, "One platform, focused AI apps", true);
    T(s, "Each app does one job for one profession, with no prompting needed. Every app runs on the same code, AI stack and operating system, and is sold the same way: a self-serve subscription with a free trial.", {
      x: 0.5, y: 1.4, w: 9, h: 0.75, fontSize: 14, color: CREAM });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 2.35, w: 2.35, h: 2.35, rectRadius: 0.1, fill: { color: GREEN2 }, line: { color: ORANGE, width: 1.5 } });
    badge(s, "LuFactory", 1.35, 2.55, 0.65, true);
    T(s, "AppForge", { x: 0.6, y: 3.3, w: 2.15, h: 0.35, fontFace: HEAD, fontSize: 18, bold: true, color: CREAM, align: "center" });
    T(s, "our in-house app factory", { x: 0.6, y: 3.68, w: 2.15, h: 0.6, fontSize: 12.5, color: SAGE, align: "center" });
    T(s, "→", { x: 2.9, y: 3.2, w: 0.45, h: 0.5, fontSize: 28, color: ORANGE, align: "center", valign: "middle" });
    const names = ["PastorAI", "Ready Room", "App 3", "App 4", "App 5", "App 6", "App 7", "App 8", "App 9", "App 10"];
    names.forEach((n, i) => {
      const x = 3.45 + (i % 5) * 1.22, y = 2.35 + Math.floor(i / 5) * 1.22;
      const live = i < 2;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: 1.1, h: 1.1, rectRadius: 0.08,
        fill: { color: live ? ORANGE : GREEN }, line: { color: live ? ORANGE : SAGE, width: 1, dashType: live ? "solid" : "dash" } });
      T(s, n, { x: x + 0.05, y: y + 0.22, w: 1.0, h: 0.4, fontSize: 13, bold: true, color: live ? WHITE : SAGE, align: "center", valign: "middle" });
      T(s, live ? "live now" : "after milestone", { x: x + 0.05, y: y + 0.62, w: 1.0, h: 0.3, fontSize: 10.5, color: live ? WHITE : SAGE, align: "center" });
    });
    T(s, "Apps 3–10 unlock only after PastorAI and Ready Room reach 150 paying subscribers.", { x: 3.45, y: 4.85, w: 6.05, h: 0.3, fontSize: 12.5, italic: true, color: SAGE });
    s.addNotes("Our solution is a platform for focused apps. Each app does one job for one profession, and they all share our code, AI stack, operating system and sales model, so each new app is cheaper to launch and run. Two are live. Apps 3 to 10 are planned, but they only unlock after the first two reach 150 paying subscribers.");
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
    const stats = [["$47/mo", "after a 14-day free trial"], ["31", "trial users, plus 1 test user in the UK"], ["3 platforms", "iOS, Android and web"]];
    stats.forEach(([big, small], i) => statCard(s, 6.4, 1.55 + i * 1.2, 3.1, 1.1, big, small, i === 0, i === 2 ? 22 : 26));
    s.addNotes("PastorAI puts the tools that fill a pastor's week in one place: sermons, series planning, Bible studies, social media, devotionals, children's lessons, translation into 10 languages and sermon rehearsal. It's live on iOS, Android and the web at $47 a month after a 14-day free trial, with 31 trial users and a test user in the UK. One of them, Pastor Dennis, told us: 'This will save me so much time. It's a great tool.'");
  }

  // ===================== 7. Ready Room =====================
  {
    const s = light();
    eyebrow(s, "Solution & Product  ·  Ready Room  ·  tryreadyroom.com");
    title(s, "Ready Room: a flight simulator for sales calls");
    const stats = [["$29/mo", "or $249/year"], ["36", "trial users, plus 1 test user in the U.S."], ["3 platforms", "iOS, Android and web"]];
    stats.forEach(([big, small], i) => statCard(s, 0.5, 1.55 + i * 1.2, 3.1, 1.1, big, small, i === 0, i === 2 ? 22 : 26));
    card(s, 3.9, 1.55, 5.6, 3.5);
    T(s, "How reps practice", { x: 4.2, y: 1.72, w: 5.1, h: 0.3, fontSize: 13, bold: true, color: ORANGE });
    const feats = [["Talk out loud ", "to a voice AI prospect that reacts to how you sell"], ["Two drills: ", "discovery questions and handling objections"],
      ["Three levels: ", "easy, medium and hard"], ["18+ prospects ", "across industries, from roofing to software"],
      ["Score real calls: ", "record live, upload a recording or paste a transcript"], ["AI coach ", "grades every session and says what to fix next"]];
    T(s, leadBullets(feats), { x: 4.2, y: 2.08, w: 5.15, h: 2.4, fontSize: 13, paraSpaceAfter: 3 });
    T(s, "Built on the founder's 40+ years in sales.", { x: 4.2, y: 4.55, w: 5.15, h: 0.4, fontFace: HEAD, italic: true, fontSize: 12.5, color: GREEN });
    s.addNotes("Ready Room lets sales reps rehearse on an AI prospect that talks back out loud. Reps choose a drill, a difficulty and a prospect from 18+ industries, then get graded by an AI coach. They can also upload their real calls for the same scorecard. It's $29 a month or $249 a year, live on iOS, Android and the web, with 36 trial users and a test user in the U.S. It's built on what I learned in 40 years of selling.");
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

  // ===================== Traction =====================
  {
    const s = light();
    eyebrow(s, "Traction & Metrics");
    title(s, "Built, shipped and live, before raising a dollar");
    const stats = [["2", "AI apps live in the app stores"], ["3", "platforms: iOS, Android, web"], ["67", "trial users, plus 2 test users (UK and U.S.)"], ["$0 MRR", "pre-revenue · team of 1"]];
    stats.forEach(([big, small], i) => statCard(s, 0.5 + i * 2.3, 1.5, 2.1, 1.15, big, small, i === 3, i === 3 ? 24 : 30));
    card(s, 0.5, 2.85, 4.4, 2.2);
    T(s, "What's already proven", { x: 0.75, y: 2.98, w: 3.9, h: 0.28, fontSize: 13, bold: true, color: ORANGE });
    T(s, bullets(["Two complete AI apps (one with voice AI) shipped solo", "Live on App Store, Google Play and web", "Subscriptions and free trials working",
      "AppForge built and ready for apps 3–10", "First testimonial from a trial user"]), { x: 0.75, y: 3.3, w: 4.0, h: 1.7, fontSize: 12.5, paraSpaceAfter: 2 });
    card(s, 5.1, 2.85, 4.4, 2.2);
    T(s, "What we'll prove next (90 days)", { x: 5.35, y: 2.98, w: 3.9, h: 0.28, fontSize: 13, bold: true, color: ORANGE });
    const kpi = [["Paying subscribers", "75"], ["Monthly revenue (MRR)", "~$2,850"], ["Trials that convert to paid", "20%+"], ["Subscribers who cancel each month", "under 6%"]];
    kpi.forEach(([k, v], i) => {
      const y = 3.35 + i * 0.4;
      T(s, k, { x: 5.35, y, w: 2.9, h: 0.32, fontSize: 12.5, color: INK });
      T(s, v, { x: 8.2, y, w: 1.1, h: 0.32, fontSize: 12.5, bold: true, color: GREEN, align: "right" });
    });
    s.addNotes("We're honest about where we are: the products are built, live on three platforms, with subscriptions and free trials working, and 67 people on free trials (31 on PastorAI and 36 on Ready Room), plus test users in the UK and the U.S. We don't have paying customers yet. Proving people will pay is the first job of this round: in the next 90 days we're aiming for 75 paying subscribers, about $2,850 a month in revenue, at least 20% of trials converting and fewer than 6% of subscribers cancelling each month. We'll report these to investors monthly.");
  }

  // ===================== Focus plan =====================
  {
    const s = light();
    eyebrow(s, "Roadmap & Milestones  ·  Our focus plan");
    title(s, "Prove two apps first, then scale with AppForge");
    // Phase 1
    card(s, 0.5, 1.5, 4.1, 2.75, GREEN);
    T(s, "Phase 1 · Months 0–6", { x: 0.75, y: 1.65, w: 3.6, h: 0.3, fontSize: 13, bold: true, color: ORANGE });
    T(s, "Prove PastorAI and Ready Room", { x: 0.75, y: 1.97, w: 3.6, h: 0.38, fontFace: HEAD, fontSize: 17, bold: true, color: CREAM });
    T(s, bullets(["Cold email and marketing to win paying customers", "Learn which channels win customers most cheaply", "Add church-staff and sales-team plans"]),
      { x: 0.75, y: 2.45, w: 3.65, h: 1.7, fontSize: 12.5, color: CREAM, paraSpaceAfter: 4 });
    // gate
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 4.75, y: 2.1, w: 0.9, h: 1.55, rectRadius: 0.08, fill: { color: ORANGE }, line: { color: ORANGE } });
    s.addImage({ data: ic.LuLockOpen.w, x: 5.0, y: 2.25, w: 0.4, h: 0.4 });
    T(s, "Gate", { x: 4.75, y: 2.72, w: 0.9, h: 0.3, fontSize: 13, bold: true, color: WHITE, align: "center" });
    T(s, "150 paying users", { x: 4.78, y: 3.0, w: 0.84, h: 0.55, fontSize: 10.5, color: WHITE, align: "center" });
    // Phase 2
    card(s, 5.8, 1.5, 3.7, 2.75);
    T(s, "Phase 2 · Months 7–18", { x: 6.05, y: 1.65, w: 3.2, h: 0.3, fontSize: 13, bold: true, color: ORANGE });
    T(s, "Launch apps 3–10 with AppForge", { x: 6.05, y: 1.97, w: 3.3, h: 0.38, fontFace: HEAD, fontSize: 16, bold: true, color: GREEN });
    T(s, bullets(["About one new app every 6 weeks", "Each uses the proven sales playbook from Phase 1", "Cross-promote apps to each other's users"]),
      { x: 6.05, y: 2.45, w: 3.3, h: 1.7, fontSize: 12.5, color: INK, paraSpaceAfter: 4 });
    strip(s, 0.5, 4.42, 9, 0.63);
    T(s, [{ text: "Why this is one company, not two: ", options: { bold: true } }, { text: "every app shares one platform, one AppForge pipeline and one sales playbook: self-serve subscriptions for busy professionals." }],
      { x: 0.75, y: 4.45, w: 8.5, h: 0.57, fontSize: 12.5, color: GREEN, valign: "middle" });
    s.addNotes("This is how we stay focused. Phase one, the first six months, is only about proving PastorAI and Ready Room: cold email and marketing to win paying customers, learning which channels work, and adding team plans. Only when the two apps reach 150 paying subscribers do we unlock phase two, launching apps 3 to 10 with AppForge, about one every six weeks, each using the sales playbook we proved in phase one. The money for new apps is held until that milestone. It's one company because every app shares the same platform, pipeline and playbook.");
  }

  // ===================== Market =====================
  {
    const s = light();
    eyebrow(s, "Market Opportunity");
    title(s, "Two large markets, each with room to grow");
    const cols = [
      ["LuChurch", "PastorAI: churches", "~$169M", "300,000+ U.S. congregations × $564/yr",
        [["Today: ", "individual pastors"], ["Next: ", "church-staff plans"], ["Then: ", "denominations and church networks buying for hundreds of churches"]]],
      ["LuBriefcase", "Ready Room: sales teams", "~$522M", "~1.5M U.S. B2B sales reps × $348/yr",
        [["Today: ", "individual reps"], ["Next: ", "sales-team plans"], ["Then: ", "company-wide sales training and enablement"]]],
    ];
    cols.forEach(([n, h, big, calc, path], i) => {
      const x = 0.5 + i * 4.6;
      card(s, x, 1.45, 4.4, 3.0);
      badge(s, n, x + 0.22, 1.6, 0.5);
      T(s, h, { x: x + 0.85, y: 1.66, w: 3.4, h: 0.35, fontFace: HEAD, fontSize: 15, bold: true, color: GREEN });
      T(s, big, { x: x + 0.25, y: 2.15, w: 1.9, h: 0.5, fontFace: HEAD, fontSize: 26, bold: true, color: ORANGE });
      T(s, "U.S. market at today's prices: " + calc, { x: x + 2.05, y: 2.15, w: 2.2, h: 0.6, fontSize: 11, color: MUTED });
      T(s, leadBullets(path), { x: x + 0.25, y: 2.95, w: 3.95, h: 1.45, fontSize: 12.5, paraSpaceAfter: 3 });
    });
    T(s, [{ text: "Targets: ", options: { bold: true } }, { text: "reachable market ~$207M (the ~30% we can reach online); 3-year goal ~$2M a year in revenue. Organization plans and apps 3–10 are upside on top." }],
      { x: 0.5, y: 4.55, w: 9, h: 0.42, fontSize: 12, color: GREEN });
    T(s, "Sources: Hartford Institute for Religion Research and U.S. Religion Census (congregations); U.S. Bureau of Labor Statistics (wholesale and manufacturing sales reps).", {
      x: 0.5, y: 4.98, w: 9, h: 0.22, fontSize: 9, italic: true, color: MUTED });
    s.addNotes("Each market stands on its own. PastorAI: roughly 300,000+ U.S. congregations at $564 a year is about $169 million, before larger deals. The path grows from individual pastors to church-staff plans to denominations and church networks buying for hundreds of churches at once. Ready Room: about 1.5 million U.S. B2B sales reps at $348 a year is about $522 million, growing from individual reps to team plans to company-wide sales training. About 30% of these markets is reachable online, and our three-year goal is about $2 million a year in revenue, with organization plans and new apps as upside.");
  }

  // ===================== Competition: PastorAI =====================
  {
    const s = light();
    eyebrow(s, "Competitive Landscape  ·  PastorAI vs. faith-tech");
    title(s, "PastorAI: the all-in-one weekly assistant for pastors");
    const hdr = ["Competitor", "What they focus on", "Where PastorAI wins"];
    const rows = [["ChatGPT and general AI (indirect)", "General-purpose chat", "Built for ministry, no prompting, writes in the pastor's voice"],
      ["Logos (Faithlife)", "Bible study and research library", "Creates the week's finished content, not just research"],
      ["Pulpit AI", "Turns a recorded sermon into clips and posts", "Helps before the sermon too: outlines, series, studies, kids' lessons"],
      ["Gloo", "Church engagement and growth platform", "Simple self-serve tool a single pastor can start today"]];
    const cell = (t, o = {}) => ({ text: t, options: Object.assign({ fontFace: BODY, fontSize: 12, color: INK, valign: "middle" }, o) });
    const data = [hdr.map((h, i) => cell(h, { bold: true, color: i === 2 ? CREAM : GREEN, fill: { color: i === 2 ? GREEN : TAN } }))]
      .concat(rows.map(r => r.map((c, i) => cell(c, { bold: i !== 1, color: i === 2 ? GREEN : INK, fill: { color: i === 2 ? "E3EDE3" : WHITE } }))));
    s.addTable(data, { x: 0.5, y: 1.45, w: 9, colW: [2.5, 2.9, 3.6], rowH: [0.42, 0.62, 0.62, 0.62, 0.62], border: { type: "solid", pt: 0.75, color: TAN }, margin: [0, 0.12, 0, 0.12] });
    T(s, [{ text: "Our edge: ", options: { bold: true } }, { text: "one dashboard for the whole ministry week (sermons, series, studies, devotionals, kids' lessons, social posts, translation, rehearsal) at $47/month." }],
      { x: 0.5, y: 4.45, w: 9, h: 0.6, fontSize: 12.5, color: GREEN });
    s.addNotes("PastorAI's competitors. General AI like ChatGPT is the main indirect competitor, but pastors have to do all the prompting. Logos is a Bible study and research library. Pulpit AI turns a recorded sermon into clips and posts after Sunday. Gloo is a church engagement and growth platform. PastorAI's edge is one simple dashboard that creates the whole week's content, before and after the sermon, for $47 a month.");
  }

  // ===================== Competition: Ready Room =====================
  {
    const s = light();
    eyebrow(s, "Competitive Landscape  ·  Ready Room vs. AI sales coaching");
    title(s, "Ready Room: sales practice any rep can afford");
    const hdr = ["Competitor", "What they focus on", "Where Ready Room wins"];
    const rows = [["Second Nature", "AI role-play training for company sales teams", "Self-serve at $29/month; no company contract needed"],
      ["Hyperbound", "AI cold-call role-play for sales teams", "Covers discovery and objections across 18+ industries"],
      ["Gong", "Analyzes recorded sales calls for large teams", "Practice before the call, plus scoring of real calls"],
      ["Practicing with a manager or peer", "Role-play when someone has time", "Available any time, at three difficulty levels"]];
    const cell = (t, o = {}) => ({ text: t, options: Object.assign({ fontFace: BODY, fontSize: 12, color: INK, valign: "middle" }, o) });
    const data = [hdr.map((h, i) => cell(h, { bold: true, color: i === 2 ? CREAM : GREEN, fill: { color: i === 2 ? GREEN : TAN } }))]
      .concat(rows.map(r => r.map((c, i) => cell(c, { bold: i !== 1, color: i === 2 ? GREEN : INK, fill: { color: i === 2 ? "E3EDE3" : WHITE } }))));
    s.addTable(data, { x: 0.5, y: 1.45, w: 9, colW: [2.5, 2.9, 3.6], rowH: [0.42, 0.62, 0.62, 0.62, 0.62], border: { type: "solid", pt: 0.75, color: TAN }, margin: [0, 0.12, 0, 0.12] });
    T(s, [{ text: "Our edge: ", options: { bold: true } }, { text: "the well-funded players sell to big companies. Ready Room serves the individual reps and small teams they skip, then grows into team plans." }],
      { x: 0.5, y: 4.45, w: 9, h: 0.6, fontSize: 12.5, color: GREEN });
    s.addNotes("Ready Room's competitors. Second Nature and Hyperbound offer AI role-play, sold mainly to company sales teams. Gong analyzes recorded calls for large sales organizations. And many reps simply practice with a manager when someone has time. Ready Room serves the individual reps and small teams the big players skip, at $29 a month, then grows into team plans.");
  }

  // ===================== Business model =====================
  {
    const s = light();
    eyebrow(s, "Business Model  ·  How we make money");
    title(s, "Subscriptions today, organization plans next");
    const streams = [["LuRepeat", "App subscriptions", "Monthly or annual plans with a free trial, sold through the app stores and the web."],
      ["LuUsers", "Team plans", "Per-seat plans for church staffs and sales teams: one buyer, several users."],
      ["LuBuilding2", "Organization plans", "Denominations, church networks and companies buying for many churches or reps at once."]];
    streams.forEach(([n, h, b], i) => {
      const x = 0.5 + i * 3.07;
      card(s, x, 1.5, 2.85, 2.25);
      badge(s, n, x + 0.25, 1.66, 0.55);
      T(s, h, { x: x + 0.25, y: 2.33, w: 2.4, h: 0.32, fontFace: HEAD, fontSize: 15, bold: true, color: GREEN });
      T(s, b, { x: x + 0.25, y: 2.68, w: 2.4, h: 1.0, fontSize: 12.5, color: MUTED });
    });
    strip(s, 0.5, 3.92, 9, 1.13);
    T(s, [{ text: "Pricing today: ", options: { bold: true } }, { text: "PastorAI $47/month after a 14-day free trial · Ready Room $29/month or $249/year.", options: { breakLine: true } },
      { text: "Unit economics (planned): ", options: { bold: true } }, { text: "~68–77% gross margin; each subscriber worth ~3× what it costs to win, paid back in ~7 months." }],
      { x: 0.75, y: 3.98, w: 8.5, h: 1.0, fontSize: 12.5, color: GREEN, valign: "middle" });
    s.addNotes("We make money from subscriptions: monthly or annual plans after a free trial. Next come per-seat team plans for church staffs and sales teams, and then organization plans, where a denomination, church network or company buys for many users at once. After app-store fees and AI costs we expect to keep 68 to 77% of revenue, and each subscriber should be worth about three times what it costs to win them. The appendix has the full unit economics.");
  }

  // ===================== 20. Go-to-market =====================
  {
    const s = light();
    eyebrow(s, "Go-to-Market Strategy");
    title(s, "How we'll turn trials into paying customers");
    card(s, 0.5, 1.5, 4.4, 3.0);
    badge(s, "LuChurch", 0.75, 1.68, 0.5);
    T(s, "PastorAI", { x: 1.4, y: 1.76, w: 3.2, h: 0.35, fontFace: HEAD, fontSize: 16, bold: true, color: GREEN });
    T(s, bullets(["Cold email to pastors and church leaders, starting now", "Facebook, YouTube and LinkedIn aimed at pastors and church staff",
      "Partnerships with denominations, church networks and seminaries", "Conferences, webinars and free sermon-prep resources"]),
      { x: 0.75, y: 2.35, w: 3.95, h: 2.1, fontSize: 12.5, paraSpaceAfter: 5 });
    card(s, 5.1, 1.5, 4.4, 3.0);
    badge(s, "LuBriefcase", 5.35, 1.68, 0.5);
    T(s, "Ready Room", { x: 6.0, y: 1.76, w: 3.2, h: 0.35, fontFace: HEAD, fontSize: 16, bold: true, color: GREEN });
    T(s, bullets(["Cold email to sales reps and sales managers, starting now", "LinkedIn and short videos showing live AI role-plays",
      "Partnerships with sales trainers, coaches and bootcamps", "Team plans, app store search and referral rewards"]),
      { x: 5.35, y: 2.35, w: 3.95, h: 2.1, fontSize: 12.5, paraSpaceAfter: 5 });
    T(s, "Cold email first, to win paying customers fast: AppForge writes the campaigns, plus social posts, ads and presentations.", { x: 0.5, y: 4.68, w: 9, h: 0.35, fontSize: 12.5, italic: true, color: GREEN });
    s.addNotes("Our first channel is cold email, starting immediately: targeted campaigns to pastors and church leaders for PastorAI, and to sales reps and sales managers for Ready Room. It's low-cost and fast, and AppForge's One-to-Many Emails tool writes the campaigns. The biggest use of funds is winning customers. Alongside cold email we'll run small, measured tests on each channel, then put more money into whichever channels win paying subscribers most cheaply. AppForge writes the emails, posts, ads and presentations, so marketing each app takes far less time.");
  }

  // ===================== Team =====================
  {
    const s = light();
    eyebrow(s, "Team & Founders");
    title(s, "A founder who can both build and sell");
    card(s, 0.5, 1.45, 5.0, 3.6);
    badge(s, "LuUser", 0.75, 1.62, 0.75);
    T(s, "James Chambers", { x: 1.7, y: 1.68, w: 3.6, h: 0.38, fontFace: HEAD, fontSize: 19, bold: true, color: GREEN });
    T(s, "Founder & CEO · Springfield, MO", { x: 1.7, y: 2.06, w: 3.6, h: 0.3, fontSize: 12.5, color: ORANGE, bold: true });
    T(s, leadBullets([["Technical: ", "degree in computer science"], ["Builder: ", "designed and shipped PastorAI and Ready Room (with voice AI) on iOS, Android and web, solo"],
      ["Systems: ", "built AppForge and Business HQ to run many apps"], ["Seller: ", "40+ years in sales across many industries"],
      ["Why him: ", "lived both problems: his bishop's workload, and decades of selling"]]),
      { x: 0.75, y: 2.5, w: 4.6, h: 2.5, fontSize: 12.5, paraSpaceAfter: 4 });
    T(s, "Hiring with this round", { x: 5.8, y: 1.45, w: 3.7, h: 0.3, fontSize: 13, bold: true, color: ORANGE });
    const hires = [["LuWrench", "Contract developer", "Builds apps 3–10 with AppForge (Phase 2)"],
      ["LuMegaphone", "Part-time growth marketer", "Runs cold email, ads and partnerships"],
      ["LuHandshake", "Advisors (recruiting now)", "Ministry, sales training and software"]];
    hires.forEach(([n, h, b], i) => {
      const y = 1.9 + i * 1.02;
      badge(s, n, 5.8, y, 0.55);
      T(s, h, { x: 6.5, y, w: 3.0, h: 0.3, fontSize: 14, bold: true, color: GREEN });
      T(s, b, { x: 6.5, y: y + 0.32, w: 3.0, h: 0.55, fontSize: 12.5, color: MUTED });
    });
    s.addNotes("I'm the founder. I have a degree in computer science and more than 40 years in sales across different industries, so I can build the products and sell them. I designed and shipped PastorAI and Ready Room, including Ready Room's voice AI, on iOS, Android and the web by myself, and built AppForge and Business HQ so one small team can run many apps. I've lived both problems: I saw my bishop's workload, and I spent decades selling. With this round I go full-time, add a part-time growth marketer for cold email and ads, and in phase two a contract developer. I'm recruiting advisors in ministry, sales training and software now.");
  }

  // ===================== Financial projections =====================
  {
    const s = light();
    eyebrow(s, "Financial Projections  ·  3-year model");
    title(s, "Break-even in year 3, ~$2M annual revenue run-rate");
    const yrs = ["Year 1", "Year 2", "Year 3"];
    const subs = [650, 2000, 4500];
    const rev = [128, 604, 1482];
    const arr = subs.map(n => Math.round(n * 38 * 12 / 1000));
    const cogs = rev.map(r => Math.round(r * 0.27));
    const opex = [340, 600, 900];
    const net = rev.map((r, i) => r - cogs[i] - opex[i]);
    const burn = net.map(n => n < 0 ? Math.round(-n / 12) : 0);
    const k = v => (v < 0 ? "−$" : "$") + (Math.abs(v) >= 1000 ? (Math.abs(v) / 1000).toFixed(2) + "M" : Math.abs(v) + "K");
    const cell = (t, o = {}) => ({ text: t, options: Object.assign({ fontFace: BODY, fontSize: 12.5, color: INK, valign: "middle", align: "center" }, o) });
    const data = [[cell("", { fill: { color: TAN } })].concat(yrs.map(y => cell(y, { bold: true, color: CREAM, fill: { color: GREEN } })))];
    const row = (label, vals, o = {}) => [cell(label, { align: "left", fill: { color: WHITE }, bold: !!o.bold })].concat(vals.map(v => cell(v, Object.assign({ fill: { color: WHITE } }, o))));
    data.push(row("Paying subscribers (end of year)", subs.map(v => v.toLocaleString("en-US"))));
    data.push(row("Annual revenue run-rate at year end (ARR)", arr.map(k)));
    data.push(row("Revenue during the year", rev.map(k)));
    data.push(row("Gross margin (after app-store fees, AI, hosting)", ["~73%", "~73%", "~73%"]));
    data.push(row("Operating costs (marketing, salaries, other)", opex.map(k)));
    data.push(row("Profit (loss)", net.map(k), { bold: true, color: GREEN }));
    data.push(row("Average monthly burn", burn.map(b => b ? "~" + k(b) : "none")));
    s.addTable(data, { x: 0.5, y: 1.4, w: 9, colW: [4.5, 1.5, 1.5, 1.5], rowH: 0.37, border: { type: "solid", pt: 0.75, color: TAN }, margin: [0, 0.1, 0, 0.1] });
    T(s, [{ text: "Runway: ", options: { bold: true } }, { text: "the $450K covers about 18 months, with roughly $120K left at month 18 when we raise a seed round to scale apps 3–10. ", options: {} },
      { text: "Assumptions: ", options: { bold: true } }, { text: "~$38 average monthly price, ~5% monthly cancellations, new apps only after the Phase 1 milestone. Projections, not guarantees." }],
      { x: 0.5, y: 4.48, w: 9, h: 0.65, fontSize: 11.5, color: MUTED });
    s.addNotes("A three-year model, built up from subscriber numbers. Paying subscribers grow from 650 at the end of year 1 to 4,500 at the end of year 3, which is an annual revenue run-rate of about $2 million. After app-store fees, AI and hosting we keep about 73% of revenue. We lose money in years 1 and 2, burning about $21,000 a month in year 1 and $13,000 in year 2, and break even in year 3. The $450,000 covers about 18 months with roughly $120,000 left when we raise a seed round. These are projections, not guarantees.");
  }

  // ===================== Use of funds & milestones =====================
  {
    const s = light();
    eyebrow(s, "Use of Funds  ·  Milestones each dollar buys");
    title(s, "$450K over 18 months, released in two phases");
    const funds = [["Marketing: cold email, ads and outreach", 160], ["Founder salary: 18 months full-time", 120], ["Apps 3–10 with AppForge (Phase 2 only)", 90],
      ["AI, hosting and infrastructure", 30], ["Legal, accounting and operations", 25], ["Reserve", 25]];
    const colors = [GREEN, ORANGE, "4F7F63", "8FA98E", "B9A67A", "D9CBA6"];
    s.addChart(pres.charts.DOUGHNUT, [{ name: "Use of funds", labels: funds.map(f => f[0]), values: funds.map(f => f[1]) }], {
      x: 0.3, y: 1.4, w: 2.75, h: 2.75, holeSize: 55, chartColors: colors, showLegend: false, showPercent: true, showValue: false,
      dataLabelColor: WHITE, dataLabelFontSize: 11, dataLabelFontBold: true, showTitle: false });
    funds.forEach(([label, v], i) => {
      const y = 1.42 + i * 0.5;
      s.addShape(pres.shapes.OVAL, { x: 3.15, y: y + 0.11, w: 0.2, h: 0.2, fill: { color: colors[i] }, line: { color: colors[i] } });
      T(s, `$${v}K`, { x: 3.42, y, w: 0.78, h: 0.42, fontFace: HEAD, fontSize: 15, bold: true, color: GREEN, valign: "middle" });
      T(s, label, { x: 4.2, y, w: 2.65, h: 0.42, fontSize: 11.5, color: INK, valign: "middle" });
    });
    card(s, 7.0, 1.4, 2.5, 3.0);
    T(s, "Milestones", { x: 7.2, y: 1.52, w: 2.1, h: 0.28, fontSize: 13, bold: true, color: ORANGE });
    T(s, leadBullets([["Month 3: ", "75 paying subscribers"], ["Month 6: ", "150 paying (unlocks Phase 2)"], ["Month 12: ", "650 paying; apps 3–6 live"], ["Month 18: ", "~1,000 paying; 10 apps; seed round"]]),
      { x: 7.2, y: 1.85, w: 2.15, h: 2.5, fontSize: 11.5, paraSpaceAfter: 6 });
    T(s, "The $90K for new apps is spent only after the Phase 1 milestone. 1,000 subscribers at ~$38 ≈ $38K a month in revenue.", {
      x: 0.5, y: 4.55, w: 9, h: 0.5, fontSize: 12, italic: true, color: MUTED });
    s.addNotes("Here's where the $450,000 goes and what it buys. $160,000 goes to marketing, led by cold email. $120,000 is my salary for 18 months so I can work full-time. $90,000 builds apps 3 to 10 with AppForge, but it's only spent after phase one hits 150 paying subscribers. The rest covers AI and hosting, legal and operations, and a reserve. Milestones: 75 paying subscribers by month 3, 150 by month 6, 650 and four new apps by month 12, and about 1,000 paying subscribers across 10 apps by month 18, when we raise a seed round.");
  }

  // ===================== The Ask =====================
  {
    const s = dark();
    logo(s, 0.5, 0.45, 0.65);
    eyebrow(s, "The Ask", 1.3);
    T(s, "Raising $450,000 to prove our apps, then scale to 10", { x: 0.5, y: 1.6, w: 5.7, h: 1.25, fontFace: HEAD, fontSize: 26, bold: true, color: CREAM });
    T(s, bullets(["Founder full-time for 18 months", "Cold email and marketing to win paying users fast", "150 paying subscribers unlocks apps 3–10 with AppForge", "~1,000 paying subscribers by month 18, then a seed round"]),
      { x: 0.5, y: 2.8, w: 5.7, h: 1.4, fontSize: 13.5, color: CREAM, paraSpaceAfter: 3 });
    T(s, "Use of funds: $160K marketing · $120K founder salary · $90K apps 3–10 (after the milestone) · $80K AI, operations and reserve", { x: 0.5, y: 4.2, w: 5.7, h: 0.48, fontSize: 12, color: SAGE });
    T(s, "Pre-seed · terms open to discussion (SAFE or priced round)", { x: 0.5, y: 4.75, w: 5.7, h: 0.3, fontSize: 12.5, italic: true, color: SAGE });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 6.5, y: 1.45, w: 3.0, h: 3.3, rectRadius: 0.1, fill: { color: GREEN2 }, line: { color: ORANGE, width: 1.25 } });
    T(s, "Let's talk", { x: 6.75, y: 1.62, w: 2.5, h: 0.3, fontSize: 13, bold: true, color: ORANGE });
    T(s, "James Chambers", { x: 6.75, y: 1.98, w: 2.5, h: 0.4, fontFace: HEAD, fontSize: 19, bold: true, color: CREAM });
    const contact = [["LuMail", "James@cedarworksapps.com"], ["LuPhone", "Office 417-370-2084"]].concat(CELL ? [["LuSmartphone", "Cell " + CELL]] : []).concat([["LuGlobe", "cedarworksapps.com"], ["LuStore", "pastorai.io · tryreadyroom.com"]]);
    contact.forEach(([n, t], i) => {
      const y = 2.5 + i * 0.43;
      s.addImage({ data: ic[n].o, x: 6.75, y: y + 0.05, w: 0.26, h: 0.26 });
      T(s, t, { x: 7.12, y, w: 2.3, h: 0.36, fontSize: 12, color: CREAM, valign: "middle" });
    });
    s.addNotes("We're raising $450,000 pre-seed. It puts me on Cedarworks full-time, funds cold email and marketing to win paying users, and, once we hit 150 paying subscribers, funds AppForge apps 3 to 10. The goal is about 1,000 paying subscribers by month 18, then a seed round. The appendix that follows has the app screens, AppForge in detail, unit economics and key terms. Thank you. I'd love to show you a live demo.");
  }

  // ===================== Appendix divider =====================
  {
    const s = dark();
    eyebrow(s, "Appendix", 2.0);
    T(s, "Product screens, AppForge in detail, unit economics and key terms", { x: 0.5, y: 2.35, w: 8.5, h: 1.2, fontFace: HEAD, fontSize: 26, bold: true, color: CREAM });
    s.addNotes("The appendix has more detail for anyone who wants to go deeper.");
  }
  // ===================== 6. PastorAI screen =====================
  {
    const s = light();
    eyebrow(s, "Appendix  ·  Solution & Product  ·  PastorAI  ·  inside the app");
    title(s, "One dashboard for the whole ministry week");
    appShot(s, "pastorai-wide.png", 0.5, 1.55, 6.0, 3.49);
    callouts(s, [["Pick a tool", "Eight ministry tools, one click away in the sidebar."],
      ["Create in minutes", "For example, a full sermon outline with scripture, illustrations and application."],
      ["Try free, then subscribe", "A 14-day free trial, with the $47/month upgrade built into the app."]], 6.8, 1.6, 2.7, 1.15);
    s.addNotes("This is the PastorAI home screen. Every tool is one click away. A pastor picks a tool, answers a few questions, and gets a finished draft in their own voice. The free trial and upgrade are built in, so trials turn into subscriptions inside the app.");
  }

  // ===================== 8. Ready Room screen =====================
  {
    const s = light();
    eyebrow(s, "Appendix  ·  Solution & Product  ·  Ready Room  ·  inside the app");
    title(s, "Pick a prospect, pick a drill, start selling");
    appShot(s, "readyroom-wide.png", 0.5, 1.55, 6.0, 3.49);
    callouts(s, [["Choose a drill", "Discovery or objections."], ["Set the difficulty", "Easy, medium or hard."],
      ["Pick a prospect", "18+ industries, each with a realistic AI buyer."], ["Get scored", "An AI coach grades every call, including real calls you upload."]], 6.8, 1.6, 2.7, 0.86);
    s.addNotes("This is Ready Room's practice screen. A rep picks a drill and a difficulty, then picks a prospect, like a homeowner considering a new roof. The AI prospect talks back, and an AI coach scores the call afterward.");
  }

  // ===================== 10. AppForge: Know =====================
  {
    const s = light();
    eyebrow(s, "Appendix  ·  Solution & Product  ·  AppForge step 1: Know");
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
    eyebrow(s, "Appendix  ·  Solution & Product  ·  AppForge step 2: Find");
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
    eyebrow(s, "Appendix  ·  Solution & Product  ·  AppForge step 3: Build");
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
    eyebrow(s, "Appendix  ·  Solution & Product  ·  AppForge step 4: Market");
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
    eyebrow(s, "Appendix  ·  Solution & Product  ·  Why AppForge matters");
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

  // ===================== 17. Why now =====================
  {
    const s = light();
    eyebrow(s, "Appendix  ·  Market Opportunity  ·  Why now & customer segments");
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

  // ===================== 19. Unit economics =====================
  {
    const s = light();
    eyebrow(s, "Appendix  ·  Business Model  ·  Unit economics (what each subscriber is worth)");
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
