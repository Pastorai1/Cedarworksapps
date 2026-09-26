const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const I = require("react-icons/lu");

const OUT = process.argv[2] || "Cedarworks-Apps-Pitch-Deck.pptx";

// Brand palette (brand/README.md)
const GREEN = "143D28", GREEN2 = "1E5238", CREAM = "F4EBD5", ORANGE = "D98B3C",
  TAN = "ECE1C7", PAPER = "FAF6EC", INK = "1B2A21", MUTED = "5C6B61", SAGE = "A9C2A8", WHITE = "FFFFFF";
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

(async () => {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9"; // 10 x 5.625
  pres.author = "James Chambers";
  pres.company = "Cedarworks Apps LLC";
  pres.title = "Cedarworks Apps — Investor Pitch Deck";
  pres.subject = "Pre-seed pitch deck: problem statement, solution & product, traction & metrics, market opportunity, business model, competitive landscape, team, use of funds";

  const cedar = await svgPng(CEDAR);
  const ic = {};
  for (const n of ["LuChurch", "LuMic", "LuTarget", "LuHammer", "LuRocket", "LuSearch", "LuUsers", "LuSmartphone",
    "LuDollarSign", "LuMegaphone", "LuLayers", "LuTrendingUp", "LuClock", "LuFileText", "LuBookOpen",
    "LuShieldCheck", "LuMail", "LuPhone", "LuGlobe", "LuBrain", "LuRepeat", "LuChartBar", "LuUser",
    "LuWrench", "LuHandshake", "LuZap", "LuStore", "LuBriefcase"]) {
    ic[n] = { o: await icon(n, ORANGE), g: await icon(n, GREEN), c: await icon(n, CREAM) };
  }

  const TOTAL = 18;
  let num = 0;

  // Brand motif: icon inside an orange-outlined circle (mirrors the logo mark)
  function badge(s, name, x, y, d, dark) {
    s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: dark ? GREEN2 : WHITE }, line: { color: ORANGE, width: 1.5 } });
    const p = d * 0.26;
    s.addImage({ data: ic[name][dark ? "o" : "g"], x: x + p, y: y + p, w: d - 2 * p, h: d - 2 * p });
  }
  function logo(s, x, y, d, dark = true) {
    s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: dark ? GREEN : WHITE }, line: { color: ORANGE, width: 1.5 } });
    const p = d * 0.22;
    s.addImage({ data: cedar, x: x + p, y: y + p, w: d - 2 * p, h: d - 2 * p });
  }
  function footer(s, dark) {
    num++;
    s.addText("Cedarworks Apps  ·  Confidential", { x: 0.5, y: 5.22, w: 5, h: 0.25, fontFace: BODY, fontSize: 8,
      color: dark ? SAGE : MUTED, margin: 0, isTextBox: true });
    s.addText(`${num} / ${TOTAL}`, { x: 8.5, y: 5.22, w: 1, h: 0.25, fontFace: BODY, fontSize: 8, align: "right",
      color: dark ? SAGE : MUTED, margin: 0, isTextBox: true });
  }
  function eyebrow(s, text, y = 0.42, dark = false) {
    s.addText(text, { x: 0.5, y, w: 9, h: 0.26, fontFace: BODY, fontSize: 11, bold: true,
      color: ORANGE, margin: 0, isTextBox: true });
  }
  function title(s, text, dark = false, y = 0.68, size = 26) {
    s.addText(text, { x: 0.5, y, w: 9, h: 0.88, fontFace: HEAD, fontSize: size, bold: true,
      color: dark ? CREAM : GREEN, margin: 0, valign: "top", isTextBox: true });
  }
  function light() { const s = pres.addSlide(); s.background = { color: PAPER }; return s; }
  function dark() { const s = pres.addSlide(); s.background = { color: GREEN }; return s; }
  function card(s, x, y, w, h, fill = WHITE) {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.08, fill: { color: fill },
      line: { color: TAN, width: 0.75 }, shadow: { type: "outer", color: "000000", opacity: 0.08, blur: 6, offset: 2, angle: 90 } });
  }
  const T = (s, text, o) => s.addText(text, Object.assign({ fontFace: BODY, fontSize: 13, color: INK, margin: 0, valign: "top", isTextBox: true }, o));

  const bullets = (arr) => arr.map((f, i) => ({ text: f, options: { bullet: { indent: 14 }, breakLine: i < arr.length - 1 } }));
  function statCard(s, x, y, w, h, big, small, hi, bigSize = 28) {
    card(s, x, y, w, h, hi ? GREEN : WHITE);
    T(s, big, { x: x + 0.22, y: y + 0.1, w: w - 0.4, h: 0.5, fontFace: HEAD, fontSize: bigSize, bold: true, color: hi ? CREAM : GREEN, valign: "middle" });
    T(s, small, { x: x + 0.22, y: y + 0.6, w: w - 0.4, h: h - 0.65, fontSize: 11, color: hi ? SAGE : MUTED });
  }

  // ---------- 1. Title ----------
  {
    const s = dark();
    logo(s, 0.5, 0.55, 0.8);
    s.addText("CEDARWORKS APPS", { x: 1.5, y: 0.72, w: 5, h: 0.45, fontFace: BODY, fontSize: 15, bold: true, color: CREAM, margin: 0, isTextBox: true });
    s.addText("Practical AI apps,\nbuilt to do real work.", { x: 0.5, y: 1.65, w: 6.6, h: 1.7, fontFace: HEAD, fontSize: 40, bold: true, color: CREAM, margin: 0, valign: "top", isTextBox: true });
    s.addText("A small studio shipping focused, subscription AI apps: two live in the app stores today, and AppForge to build the next eight.", {
      x: 0.5, y: 3.45, w: 6.2, h: 0.75, fontFace: HEAD, italic: true, fontSize: 15, color: SAGE, margin: 0, valign: "top", isTextBox: true });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 7.2, y: 1.7, w: 2.3, h: 1.55, rectRadius: 0.1, fill: { color: GREEN2 }, line: { color: ORANGE, width: 1.25 } });
    s.addText([
      { text: "Pre-seed round", options: { fontSize: 10, bold: true, color: ORANGE, breakLine: true } },
      { text: "$250K", options: { fontSize: 36, bold: true, color: CREAM, fontFace: HEAD, breakLine: true } },
      { text: "Springfield, MO", options: { fontSize: 11, color: SAGE } },
    ], { x: 7.2, y: 1.8, w: 2.3, h: 1.35, fontFace: BODY, align: "center", valign: "middle", margin: 0, isTextBox: true });
    s.addText("James Chambers, Founder  ·  james@cedarworksapps.com  ·  417-370-2084  ·  cedarworksapps.com", {
      x: 0.5, y: 4.6, w: 9, h: 0.3, fontFace: BODY, fontSize: 11, color: CREAM, margin: 0, isTextBox: true });
    footer(s, true);
    s.addNotes("Cedarworks Apps is a Springfield, Missouri AI app studio. We build focused, subscription AI apps for specific professions. Two are live in the Apple App Store and Google Play today. We're raising a $250K pre-seed round to market them and use our in-house AppForge engine to launch eight more.");
  }

  // ---------- 2. Executive summary ----------
  {
    const s = light();
    eyebrow(s, "Executive Summary");
    title(s, "Cedarworks Apps at a glance");
    const items = [
      ["Problem Statement", "Pastors and salespeople lose hours to repetitive prep and unpracticed skills. Generic AI doesn't fit their work."],
      ["Solution & Product", "Focused AI apps: PastorAI ($47/mo) for ministry and Ready Room ($29/mo) for sales-call practice."],
      ["Traction & Metrics", "Both apps live on iOS, Android and web; ~30 trial users; $0 MRR (pre-revenue); team of 1."],
      ["Market Opportunity", "~$690M U.S. TAM across 300K+ churches and ~1.5M B2B sales reps; 8 more apps widen it."],
      ["Business Model", "Monthly subscriptions, ~68–77% estimated gross margin, target LTV:CAC of 3:1 or better."],
      ["Competitive Landscape", "Direct: ministry software, sales role-play platforms. Indirect: ChatGPT. We win on focus and price."],
      ["Team", "Founder James Chambers: computer science degree, 40+ years in sales, built both apps."],
      ["The Ask", "$250K pre-seed for about 18 months: grow the two apps and launch 8 more with AppForge."],
    ];
    items.forEach(([h, b], i) => {
      const col = i % 2, row = Math.floor(i / 2);
      const x = 0.5 + col * 4.6, y = 1.55 + row * 0.9;
      card(s, x, y, 4.4, 0.78);
      T(s, h, { x: x + 0.2, y: y + 0.09, w: 4.0, h: 0.24, fontSize: 11, bold: true, color: ORANGE });
      T(s, b, { x: x + 0.2, y: y + 0.33, w: 4.0, h: 0.42, fontSize: 10.5, color: INK });
    });
    footer(s);
    s.addNotes("A one-slide summary of the whole deck, in the same order investors score it: problem, solution, traction, market, business model, competition, team and the ask.");
  }

  // ---------- 3. Problem ----------
  {
    const s = light();
    eyebrow(s, "Problem Statement  ·  Customer pain points");
    title(s, "Generic AI doesn't fit how pastors and reps work");
    const rows = [
      ["LuClock", "Pastors are overloaded", "Sermons, studies, devotionals, kids' lessons, social posts and translation fill the week, often with little or no staff. 42% of U.S. pastors considered quitting full-time ministry in the past year (Barna Group, 2022)."],
      ["LuTarget", "Sales reps practice on real prospects", "Discovery and objection-handling are learned live, on the calls that matter. One fumbled objection can cost a deal worth thousands in commission."],
      ["LuBrain", "General chatbots don't close the gap", "Existing options fall short: chatbots don't know the workflow or vocabulary, and sales-training platforms are priced and sold for enterprise teams."],
    ];
    rows.forEach(([n, h, b], i) => {
      const y = 1.65 + i * 0.84;
      badge(s, n, 0.5, y, 0.56);
      T(s, h, { x: 1.3, y: y - 0.02, w: 8.2, h: 0.3, fontFace: HEAD, fontSize: 15.5, bold: true, color: GREEN });
      T(s, b, { x: 1.3, y: y + 0.3, w: 8.2, h: 0.48, fontSize: 11.5, color: MUTED });
    });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 4.25, w: 9, h: 0.78, rectRadius: 0.08, fill: { color: TAN }, line: { color: TAN } });
    T(s, [
      { text: "Why we know this problem: ", options: { bold: true } },
      { text: "Founder James Chambers watched the weekly workload his bishop carried, and spent 40+ years in sales watching reps learn objection-handling on live deals." },
    ], { x: 0.75, y: 4.33, w: 8.5, h: 0.62, fontSize: 12, color: GREEN, valign: "middle" });
    footer(s);
    s.addNotes("The problem: skilled professionals spend hours every week on repeatable work that AI could speed up, but generic AI makes them do the translating. Pastors carry a week of writing and teaching prep, often alone. Sales reps have nowhere safe to practice before a real call. I've seen both first-hand: my bishop's workload, and 40 years of selling.");
  }

  // ---------- 4. Solution ----------
  {
    const s = dark();
    eyebrow(s, "Solution & Product  ·  What Cedarworks offers");
    title(s, "Focused AI apps that each do one job really well", true);
    T(s, "Cedarworks Apps builds niche, subscription AI apps: one profession, one clear job, a running start in seconds. Each app has its own workflow, prompts and guardrails, so users get results instead of a blank chat box.", {
      x: 0.5, y: 1.7, w: 4.3, h: 1.6, fontSize: 14, color: CREAM });
    T(s, "One studio, many apps. Shared code, shared infrastructure, one founder-led team.", {
      x: 0.5, y: 3.45, w: 4.3, h: 0.8, fontFace: HEAD, italic: true, fontSize: 14, color: SAGE });
    const pts = [["LuTarget", "Built for one job", "Purpose-built workflows, not generic chat"],
      ["LuSmartphone", "Everywhere users are", "iOS, Android and web"],
      ["LuRepeat", "Recurring revenue", "Monthly subscriptions with free trials"],
      ["LuLayers", "A portfolio, not a bet", "Many apps spread the risk and share costs"]];
    pts.forEach(([n, h, b], i) => {
      const y = 1.6 + i * 0.86;
      badge(s, n, 5.3, y, 0.55, true);
      T(s, h, { x: 6.05, y, w: 3.45, h: 0.28, fontSize: 14, bold: true, color: CREAM });
      T(s, b, { x: 6.05, y: y + 0.28, w: 3.45, h: 0.3, fontSize: 11.5, color: SAGE });
    });
    footer(s, true);
    s.addNotes("We don't build one app for everyone. We build many small apps that each serve one audience extremely well. That keeps each product simple and lets us reuse our technology across the portfolio.");
  }

  // App screenshots (recreated from the live apps; sources in pitch-deck/*.html)
  const fs = require("fs"), path = require("path");
  const shot = f => "image/png;base64," + fs.readFileSync(path.join(__dirname, f)).toString("base64");
  function appShot(s, file, x, y, w, h) {
    s.addShape(pres.shapes.RECTANGLE, { x: x - 0.03, y: y - 0.03, w: w + 0.06, h: h + 0.06, fill: { color: TAN }, line: { color: TAN },
      shadow: { type: "outer", color: "000000", opacity: 0.15, blur: 8, offset: 3, angle: 90 } });
    s.addImage({ data: shot(file), x, y, w, h });
  }

  // ---------- 5. PastorAI ----------
  {
    const s = light();
    eyebrow(s, "Solution & Product  ·  PastorAI  ·  pastorai.io");
    title(s, "PastorAI: an AI ministry assistant for the whole week");
    appShot(s, "pastorai-dashboard.png", 0.5, 1.62, 3.6, 3.37);
    card(s, 4.35, 1.6, 3.0, 3.4);
    T(s, "Features", { x: 4.58, y: 1.75, w: 2.6, h: 0.26, fontSize: 11, bold: true, color: ORANGE });
    const feats = ["Sermon Builder: full outlines with scripture", "Series Planner: multi-week series", "Bible Study: small-group guides",
      "Social media: 7 days of posts per sermon", "Devotionals and children's lessons", "Translation into 10 languages", "Sermon rehearsal"];
    T(s, bullets(feats), { x: 4.58, y: 2.04, w: 2.62, h: 2.2, fontSize: 11, paraSpaceAfter: 3 });
    T(s, "\u201CThis will save me so much time. It's a great tool.\u201D \u2014 Pastor Dennis, trial user", {
      x: 4.58, y: 4.3, w: 2.62, h: 0.6, fontFace: HEAD, italic: true, fontSize: 10.5, color: GREEN });
    const stats = [["$47", "per month, free trial first"], ["~15", "trial members"], ["3", "platforms: iOS · Android · web"]];
    stats.forEach(([big, small], i) => statCard(s, 7.55, 1.6 + i * 1.2, 1.95, 1.05, big, small, i === 0, 26));
    footer(s);
    s.addNotes("This is the PastorAI dashboard. Every tool a pastor needs for the week is one click away: Sermon Builder, Series Planner, Bible Study, Social Media, Devotional, Children's Lesson, Translation into 10 languages, and a Rehearsal tool. It's live on iOS, Android and the web at $47 a month after a free trial, with about 15 trial members. One of them, Pastor Dennis, told us: 'This will save me so much time. It's a great tool.'");
  }

  // ---------- 6. Ready Room ----------
  {
    const s = light();
    eyebrow(s, "Solution & Product  ·  Ready Room  ·  tryreadyroom.com");
    title(s, "Ready Room: a flight simulator for sales calls");
    const stats = [["$29", "per month, free trial first"], ["~15", "trial members"], ["3", "platforms: iOS · Android · web"]];
    stats.forEach(([big, small], i) => statCard(s, 0.5, 1.6 + i * 1.2, 1.95, 1.05, big, small, i === 0, 26));
    card(s, 2.65, 1.6, 3.0, 3.4);
    T(s, "Features", { x: 2.88, y: 1.75, w: 2.6, h: 0.26, fontSize: 11, bold: true, color: ORANGE });
    const feats = ["Voice AI prospect that talks back and reacts to how you sell", "Discovery and objection drills at easy, medium and hard",
      "18+ ready-made prospects, from roofing to SaaS", "Score real calls: record live, upload a recording or paste a transcript",
      "AI coach scorecard after every session"];
    T(s, bullets(feats), { x: 2.88, y: 2.04, w: 2.62, h: 2.4, fontSize: 11, paraSpaceAfter: 3 });
    T(s, "Built on 40+ years of the founder's sales experience.", {
      x: 2.88, y: 4.45, w: 2.62, h: 0.45, fontFace: HEAD, italic: true, fontSize: 10.5, color: GREEN });
    appShot(s, "readyroom-app.png", 5.9, 1.62, 3.6, 3.37);
    footer(s);
    s.addNotes("This is Ready Room. A rep picks a prospect from 18+ industry scenarios, chooses a discovery or objection drill and a difficulty, then sells out loud to a voice AI prospect that reacts to how they sell. Reps can also score their real calls by recording live, uploading a recording, or pasting a transcript. An AI coach grades every session. It's live on iOS, Android and the web at $29 a month after a free trial, with about 15 trial members. It's built on what I learned in 40 years of selling.");
  }

  // ---------- 7. AppForge ----------
  {
    const s = dark();
    eyebrow(s, "Solution & Product  ·  AppForge");
    title(s, "AppForge: our engine for finding and building apps", true);
    appShot(s, "appforge-app.png", 0.5, 1.62, 3.6, 3.37);
    const steps = [["LuSearch", "1  Analyze", "Top 100 Apps, Market Analysis and Opportunities find underserved niches with proven demand"],
      ["LuTarget", "2  Score & select", "App Types and the Viability Scorer rank ideas on demand, competition and fit"],
      ["LuHammer", "3  Build", "Tech Stack, Projects and Build reuse the Cedarworks stack of AI, voice and billing"],
      ["LuRocket", "4  Launch & market", "Store Listing, plus one-to-many emails, social posts, ads and presentations"]];
    steps.forEach(([n, h, b], i) => {
      const y = 1.62 + i * 0.86;
      badge(s, n, 4.45, y + 0.04, 0.52, true);
      T(s, h, { x: 5.15, y, w: 4.35, h: 0.28, fontSize: 14, bold: true, color: CREAM });
      T(s, b, { x: 5.15, y: y + 0.29, w: 4.35, h: 0.5, fontSize: 11, color: SAGE });
    });
    footer(s, true);
    s.addNotes("AppForge is our proprietary engine, shown here. It covers the whole path from idea to launch. It analyzes the market (Top 100 Apps, Market Analysis, Opportunities), scores ideas by app type with a Viability Scorer, sets up the tech stack and build, then produces the store listing and one-to-many marketing: emails, social posts, ads and presentations. PastorAI and Ready Room prove the build side works. The next eight apps will be chosen from AppForge's market analysis, not guesswork, and each launches faster and cheaper than the last.");
  }

  // ---------- 8. Traction & Metrics ----------
  {
    const s = light();
    eyebrow(s, "Traction & Metrics");
    title(s, "Built, shipped and live, before raising a dollar");
    const stats = [["2", "AI apps live and selling"], ["3", "platforms: iOS · Android · web"], ["~30", "trial members across both apps"], ["Pre-revenue", "$0 MRR today · team of 1 · launched 2026"]];
    stats.forEach(([big, small], i) => statCard(s, 0.5 + i * 2.3, 1.6, 2.05, 1.12, big, small, i === 3, i === 3 ? 19 : 30));
    // Milestones achieved
    card(s, 0.5, 2.95, 4.35, 2.05);
    T(s, "Milestones achieved", { x: 0.75, y: 3.08, w: 3.9, h: 0.26, fontSize: 11, bold: true, color: ORANGE });
    const done = ["Cedarworks Apps LLC formed (2026)", "PastorAI launched on App Store, Google Play & web",
      "Ready Room launched on App Store, Google Play & web", "Subscription billing and free trials live", "First trial users and first testimonial"];
    T(s, bullets(done), { x: 0.75, y: 3.38, w: 3.95, h: 1.5, fontSize: 11, paraSpaceAfter: 2 });
    // 90-day KPI targets
    card(s, 5.15, 2.95, 4.35, 2.05);
    T(s, "90-day KPI targets", { x: 5.4, y: 3.08, w: 3.9, h: 0.26, fontSize: 11, bold: true, color: ORANGE });
    const kpi = [["Paying subscribers", "75"], ["Monthly recurring revenue", "~$2,850"], ["Trial-to-paid conversion", "20%+"], ["Monthly churn", "under 6%"], ["CAC tracked per channel", "weekly"]];
    kpi.forEach(([k, v], i) => {
      const y = 3.4 + i * 0.3;
      T(s, k, { x: 5.4, y, w: 2.6, h: 0.28, fontSize: 11, color: INK });
      T(s, v, { x: 8.0, y, w: 1.3, h: 0.28, fontSize: 11, bold: true, color: GREEN, align: "right" });
    });
    footer(s);
    s.addNotes("We're at the very start. Both apps are complete and live on three platforms with subscription billing, and about 30 people are on free trials. We don't have paying customers yet. Our first 90-day goals are 75 paying subscribers, about $2,850 a month in recurring revenue, 20%+ trial-to-paid conversion and monthly churn under 6%. We'll report these metrics to investors every month.");
  }

  // ---------- 9. Market ----------
  {
    const s = light();
    eyebrow(s, "Market Opportunity");
    title(s, "A ~$690M U.S. market for our first two apps");
    // Nested circles: TAM > SAM > SOM
    s.addShape(pres.shapes.OVAL, { x: 0.6, y: 1.55, w: 3.1, h: 3.1, fill: { color: TAN }, line: { color: TAN } });
    s.addShape(pres.shapes.OVAL, { x: 1.175, y: 2.55, w: 1.95, h: 1.95, fill: { color: SAGE }, line: { color: SAGE } });
    s.addShape(pres.shapes.OVAL, { x: 1.675, y: 3.45, w: 0.95, h: 0.95, fill: { color: GREEN }, line: { color: GREEN } });
    T(s, "TAM ~$690M", { x: 0.6, y: 1.8, w: 3.1, h: 0.3, fontSize: 13, bold: true, color: GREEN, align: "center" });
    T(s, "SAM ~$207M", { x: 1.175, y: 2.8, w: 1.95, h: 0.3, fontSize: 12, bold: true, color: GREEN, align: "center" });
    T(s, "SOM\n~$2M", { x: 1.675, y: 3.63, w: 0.95, h: 0.6, fontSize: 10.5, bold: true, color: CREAM, align: "center" });
    const rows = [
      ["TAM: total addressable", "300K+ U.S. congregations × $564/yr (≈ $169M), plus ~1.5M U.S. B2B sales reps × $348/yr (≈ $522M)."],
      ["SAM: serviceable", "The ~30% we can reach with digital, self-serve sales: small and mid-size churches, and individual reps and small teams."],
      ["SOM: 3-year target", "1% of SAM, or about 4,500 subscribers and ~$2M in annual recurring revenue."],
      ["Expansion", "Each of the 8 AppForge apps adds a new market on top of these two."],
    ];
    rows.forEach(([h, b], i) => {
      const y = 1.55 + i * 0.8;
      T(s, h, { x: 4.3, y, w: 5.2, h: 0.27, fontSize: 13, bold: true, color: GREEN });
      T(s, b, { x: 4.3, y: y + 0.28, w: 5.2, h: 0.52, fontSize: 11, color: MUTED });
    });
    T(s, "Sources: Hartford Institute for Religion Research & U.S. Religion Census (congregations); U.S. Bureau of Labor Statistics, Occupational Employment & Wage Statistics (sales representatives, wholesale & manufacturing). Prices: current PastorAI and Ready Room plans.", {
      x: 0.5, y: 4.8, w: 9, h: 0.36, fontSize: 8.5, italic: true, color: MUTED });
    footer(s);
    s.addNotes("Bottom-up sizing at our current prices. Congregations: roughly 300,000 to 380,000 in the U.S. (Hartford Institute for Religion Research; U.S. Religion Census). Sales reps: about 1.5 million U.S. wholesale and manufacturing sales representatives (Bureau of Labor Statistics), before counting insurance, real estate and other sales roles. SAM assumes about 30% are reachable through self-serve digital channels. SOM is 1% of SAM over three years.");
  }

  // ---------- 9b. Why now ----------
  {
    const s = light();
    eyebrow(s, "Market Opportunity  ·  Why now & customer segments");
    title(s, "Why the timing is right");
    const cards = [["LuZap", "AI got affordable", "The cost of running language and voice AI has fallen sharply since 2023, so $29–$47/month apps can earn healthy margins."],
      ["LuBrain", "People expect AI help", "Most professionals have now tried AI chat, and found it doesn't fit their daily work. That gap is where focused apps win."],
      ["LuRocket", "Small studios can ship fast", "App stores give global distribution and billing on day one, and AppForge lets one team launch app after app."]];
    cards.forEach(([n, h, b], i) => {
      const x = 0.5 + i * 3.07;
      card(s, x, 1.6, 2.85, 2.2);
      badge(s, n, x + 0.25, 1.8, 0.55);
      T(s, h, { x: x + 0.25, y: 2.5, w: 2.4, h: 0.3, fontFace: HEAD, fontSize: 14.5, bold: true, color: GREEN });
      T(s, b, { x: x + 0.25, y: 2.85, w: 2.4, h: 0.9, fontSize: 11, color: MUTED });
    });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 4.0, w: 9, h: 1.0, rectRadius: 0.08, fill: { color: TAN }, line: { color: TAN } });
    T(s, [
      { text: "Customer segments: ", options: { bold: true } },
      { text: "(1) pastors and staff at small and mid-size churches; (2) individual B2B sales reps and small sales teams; (3) sales managers and church networks buying group plans; (4) new niches chosen by AppForge." },
    ], { x: 0.75, y: 4.08, w: 8.5, h: 0.84, fontSize: 12, color: GREEN, valign: "middle" });
    footer(s);
    s.addNotes("Why now: AI costs have dropped enough to support affordable subscription apps; professionals have tried generic AI and want tools built for their work; and app stores plus AppForge let a small studio launch quickly. Our customer segments start with pastors and church staff and individual sales reps, expand to group plans, and then to new niches chosen by AppForge.");
  }

  // ---------- 10. Business model ----------
  {
    const s = light();
    eyebrow(s, "Business Model  ·  Revenue streams & unit economics");
    title(s, "Recurring subscriptions with healthy unit economics");
    const cell = (t, o = {}) => ({ text: t, options: Object.assign({ fontFace: BODY, fontSize: 11, color: INK, valign: "middle" }, o) });
    const hdrO = (i) => ({ bold: true, color: i ? CREAM : GREEN, fill: { color: i ? GREEN : TAN }, align: i ? "center" : "left" });
    const rows = [["Monthly price", "$47", "$29"], ["Net after app-store fee (15%)", "$39.95", "$24.65"], ["Est. AI + hosting cost / user", "~$4", "~$5"],
      ["Gross margin", "~77%", "~68%"], ["Lifetime value (5% churn, 20 mo)", "~$720", "~$390"], ["Target CAC (LTV:CAC ≥ 3:1)", "≤ $240", "≤ $130"], ["CAC payback", "~7 months", "~7 months"]];
    const data = [["Unit economics (per subscriber)", "PastorAI", "Ready Room"].map((h, i) => cell(h, hdrO(i)))]
      .concat(rows.map(r => r.map((c, i) => cell(c, { align: i ? "center" : "left", bold: i > 0, color: i ? GREEN : INK, fill: { color: WHITE } }))));
    s.addTable(data, { x: 0.5, y: 1.6, w: 5.4, colW: [2.9, 1.25, 1.25], rowH: 0.4, border: { type: "solid", pt: 0.75, color: TAN }, margin: [0, 0.1, 0, 0.1] });
    const streams = [["LuRepeat", "App subscriptions", "Monthly plans with free trials on App Store, Google Play and web"],
      ["LuUsers", "Group plans", "Multi-seat plans for sales teams and church staffs"],
      ["LuWrench", "Custom AI apps", "Custom builds for local businesses: service revenue and new app ideas"]];
    T(s, "Revenue streams", { x: 6.2, y: 1.6, w: 3.3, h: 0.26, fontSize: 11, bold: true, color: ORANGE });
    streams.forEach(([n, h, b], i) => {
      const y = 1.95 + i * 0.95;
      badge(s, n, 6.2, y, 0.5);
      T(s, h, { x: 6.85, y, w: 2.65, h: 0.28, fontSize: 13, bold: true, color: GREEN });
      T(s, b, { x: 6.85, y: y + 0.28, w: 2.65, h: 0.6, fontSize: 10.5, color: MUTED });
    });
    T(s, "Figures after the price rows are planning assumptions, not results.", { x: 0.5, y: 4.85, w: 5.4, h: 0.25, fontSize: 9, italic: true, color: MUTED });
    footer(s);
    s.addNotes("Revenue comes from monthly subscriptions. Apple's Small Business Program and Google Play take 15% on subscriptions. AI and hosting costs are estimates. Lifetime value assumes 5% monthly churn, an average 20-month life, measured as gross profit. We'll hold customer acquisition cost to a third of lifetime value, which pays back in about seven months. Group plans and custom app work add revenue on top.");
  }

  // ---------- 11. Go-to-market ----------
  {
    const s = light();
    eyebrow(s, "Go-to-Market Strategy");
    title(s, "How we'll turn trials into paying customers");
    card(s, 0.5, 1.6, 4.35, 3.4);
    badge(s, "LuChurch", 0.75, 1.8, 0.5);
    T(s, "PastorAI", { x: 1.4, y: 1.88, w: 3.2, h: 0.35, fontFace: HEAD, fontSize: 16, bold: true, color: GREEN });
    const pa = ["Targeted Facebook & YouTube ads to pastors and church staff", "LinkedIn outreach to ministry leaders",
      "Partnerships with denominations, church networks & seminaries", "Pastor conferences, webinars and free sermon-prep resources"];
    T(s, bullets(pa), { x: 0.75, y: 2.5, w: 3.9, h: 2.4, fontSize: 12.5, paraSpaceAfter: 6 });
    card(s, 5.15, 1.6, 4.35, 3.4);
    badge(s, "LuMegaphone", 5.4, 1.8, 0.5);
    T(s, "Ready Room", { x: 6.05, y: 1.88, w: 3.2, h: 0.35, fontFace: HEAD, fontSize: 16, bold: true, color: GREEN });
    const rr = ["LinkedIn and short-form video ads showing live AI role-plays", "Direct outreach to sales managers for team plans, drawing on 40 years of sales contacts",
      "Sales-trainer, coach and bootcamp affiliate partners", "App Store search optimization and referral rewards"];
    T(s, bullets(rr), { x: 5.4, y: 2.5, w: 3.9, h: 2.4, fontSize: 12.5, paraSpaceAfter: 6 });
    footer(s);
    s.addNotes("The biggest use of funds is customer acquisition. We'll run small, measured ad tests on each channel, then double down on the ones with the lowest cost per paying subscriber. Free trials are already built into both apps.");
  }

  // ---------- 12. Competition ----------
  {
    const s = light();
    eyebrow(s, "Competitive Landscape  ·  Direct & indirect competitors");
    title(s, "Why Cedarworks wins");
    const hdr = ["", "Indirect: generic AI chat (e.g. ChatGPT)", "Direct: ministry software (e.g. Logos, Pulpit AI)", "Direct: sales role-play (e.g. Second Nature, Hyperbound)", "Cedarworks apps"];
    const rows = [["Built for one profession", "No", "Yes", "Yes", "Yes"], ["Ready to use with no prompting", "No", "Partly", "Yes", "Yes"],
      ["Built for small churches & individual reps", "Yes", "Partly", "No (enterprise)", "Yes"], ["Self-serve monthly price", "Low", "Varies", "Sales-led", "$29–$47"],
      ["Launches new niche products quickly", "N/A", "No", "No", "AppForge"]];
    const cell = (t, o = {}) => ({ text: t, options: Object.assign({ fontFace: BODY, fontSize: 10.5, color: INK, valign: "middle" }, o) });
    const data = [hdr.map((h, i) => cell(h, { bold: true, fontSize: 10, color: i === 4 ? CREAM : GREEN, fill: { color: i === 4 ? GREEN : TAN }, align: i ? "center" : "left" }))]
      .concat(rows.map(r => r.map((c, i) => cell(c, { align: i ? "center" : "left", bold: i === 4, color: i === 4 ? GREEN : INK, fill: { color: i === 4 ? "E3EDE3" : WHITE } }))));
    s.addTable(data, { x: 0.5, y: 1.55, w: 9, colW: [2.5, 1.55, 1.75, 1.85, 1.35], rowH: [0.72, 0.4, 0.4, 0.4, 0.4, 0.4], border: { type: "solid", pt: 0.75, color: TAN }, margin: [0, 0.08, 0, 0.08] });
    T(s, [{ text: "Our advantages: ", options: { bold: true } }, { text: "focus on niches big platforms ignore · self-serve pricing individuals can afford · AppForge speed to new markets · a shared platform that lowers each app's cost · a founder with both the technical and sales background." }], {
      x: 0.5, y: 4.45, w: 9, h: 0.6, fontSize: 11.5, color: GREEN });
    footer(s);
    s.addNotes("People can use general chatbots, but they have to do the prompting and still get generic output. Ministry software such as Logos focuses on study and research, and tools like Pulpit AI focus on repurposing sermon content. Sales role-play platforms such as Second Nature and Hyperbound are sold to company sales teams. We sit in between: purpose-built, affordable, self-serve, and quick to launch in new niches.");
  }

  // ---------- 13. Use of funds ----------
  {
    const s = light();
    eyebrow(s, "Use of Funds");
    title(s, "$250K to scale two apps and launch eight more");
    const funds = [["Customer acquisition: PastorAI & Ready Room ads", 100], ["Build 8 new apps with AppForge", 75],
      ["AI compute, hosting & infrastructure", 25], ["Legal, accounting & app-store operations", 25], ["Founder runway & reserve", 25]];
    const colors = [GREEN, ORANGE, "4F7F63", "B9A67A", SAGE];
    s.addChart(pres.charts.DOUGHNUT, [{ name: "Use of funds", labels: funds.map(f => f[0]), values: funds.map(f => f[1]) }], {
      x: 0.4, y: 1.45, w: 3.7, h: 3.6, holeSize: 58, chartColors: colors, showLegend: false, showPercent: true, showValue: false,
      dataLabelColor: WHITE, dataLabelFontSize: 10, dataLabelFontBold: true, showTitle: false });
    funds.forEach(([label, v], i) => {
      const y = 1.6 + i * 0.66;
      s.addShape(pres.shapes.OVAL, { x: 4.5, y: y + 0.09, w: 0.2, h: 0.2, fill: { color: colors[i] }, line: { color: colors[i] } });
      T(s, `$${v}K`, { x: 4.85, y, w: 0.9, h: 0.38, fontFace: HEAD, fontSize: 17, bold: true, color: GREEN, valign: "middle" });
      T(s, label, { x: 5.8, y, w: 3.7, h: 0.38, fontSize: 12.5, color: INK, valign: "middle" });
    });
    T(s, "About 18 months of runway on a lean, founder-led cost base.", { x: 4.5, y: 4.7, w: 5, h: 0.3, fontSize: 11, italic: true, color: MUTED });
    footer(s);
    s.addNotes("40% goes to customer acquisition for PastorAI and Ready Room, the fastest path to revenue. 30% funds building eight new apps with AppForge, including contract development, design and QA. The rest covers AI compute and hosting, legal and operations, and a founder runway reserve.");
  }

  // ---------- 14. Roadmap ----------
  {
    const s = light();
    eyebrow(s, "Roadmap & Milestones");
    title(s, "From 2 apps to 10: milestones this round unlocks");
    const ms = [["Months 0–3", "Convert trials to paid; launch measured ad tests; add team and church plans"],
      ["Months 4–6", "Scale winning ad channels; AppForge picks and launches apps #3 and #4"],
      ["Months 7–12", "Launch apps #5 to #7; cross-promote across the portfolio"],
      ["Months 13–18", "Launch apps #8 to #10; position for a seed round on proven revenue"]];
    s.addShape(pres.shapes.LINE, { x: 0.8, y: 1.85, w: 0, h: 3.0, line: { color: TAN, width: 2 } });
    ms.forEach(([h, b], i) => {
      const y = 1.6 + i * 0.83;
      s.addShape(pres.shapes.OVAL, { x: 0.62, y: y + 0.07, w: 0.36, h: 0.36, fill: { color: i === 0 ? ORANGE : WHITE }, line: { color: ORANGE, width: 1.5 } });
      T(s, h, { x: 1.2, y: y + 0.02, w: 1.6, h: 0.4, fontSize: 13, bold: true, color: GREEN, valign: "middle" });
      T(s, b, { x: 2.8, y: y + 0.02, w: 2.75, h: 0.75, fontSize: 11.5, color: MUTED });
    });
    s.addChart(pres.charts.BAR, [{ name: "Paying subscribers (target)", labels: ["M3", "M6", "M9", "M12", "M18"], values: [75, 200, 400, 650, 1000] }], {
      x: 5.8, y: 1.5, w: 3.7, h: 3.35, barDir: "col", chartColors: [GREEN], showTitle: true, title: "Paying subscribers (target)",
      titleFontSize: 11, titleColor: GREEN, titleFontFace: BODY, showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 10,
      dataLabelColor: INK, catAxisLabelColor: MUTED, valAxisLabelColor: MUTED, valAxisLabelFontSize: 9, catAxisLabelFontSize: 10,
      valGridLine: { color: "E5DDC8", size: 0.5 }, catGridLine: { style: "none" }, showLegend: false, barGapWidthPct: 60 });
    T(s, "Targets, not forecasts: 1,000 subscribers at a ~$38 blended price ≈ $38K MRR (~$450K ARR).", { x: 5.8, y: 4.85, w: 3.7, h: 0.35, fontSize: 9.5, italic: true, color: MUTED });
    footer(s);
    s.addNotes("Our 18-month goal is ten live apps and about 1,000 paying subscribers across the portfolio. At a blended price near $38 that's roughly $38K a month in recurring revenue, which positions us for a seed round on proven traction. These are targets, not guarantees.");
  }

  // ---------- 14b. Financial projections ----------
  {
    const s = light();
    eyebrow(s, "Financial Projections  ·  5-year plan");
    title(s, "Profitable in year 3, ~$4.8M revenue by year 5");
    const yrs = ["Year 1", "Year 2", "Year 3", "Year 4", "Year 5"];
    const subs = [650, 2000, 4500, 8000, 13000];
    const rev = [128, 604, 1482, 2850, 4788];
    const cogs = rev.map(r => Math.round(r * 0.27));
    const opex = [200, 560, 900, 1400, 2000];
    const net = rev.map((r, i) => r - cogs[i] - opex[i]);
    const k = v => (v < 0 ? "−$" : "$") + (Math.abs(v) >= 1000 ? (Math.abs(v) / 1000).toFixed(2) + "M" : Math.abs(v) + "K");
    const cell = (t, o = {}) => ({ text: t, options: Object.assign({ fontFace: BODY, fontSize: 10.5, color: INK, valign: "middle", align: "center" }, o) });
    const data = [[cell("($, per year)", { bold: true, color: GREEN, fill: { color: TAN }, align: "left" })].concat(yrs.map(y => cell(y, { bold: true, color: CREAM, fill: { color: GREEN } })))];
    const row = (label, vals, o = {}) => [cell(label, { align: "left", fill: { color: WHITE }, bold: !!o.bold })].concat(vals.map(v => cell(v, Object.assign({ fill: { color: WHITE } }, o))));
    data.push(row("Paying subscribers (year-end)", subs.map(v => v.toLocaleString("en-US"))));
    data.push(row("Revenue", rev.map(k)));
    data.push(row("Cost of revenue (~27%)", cogs.map(k)));
    data.push(row("Operating expenses", opex.map(k)));
    data.push(row("Net income", net.map(k), { bold: true, color: GREEN }));
    s.addTable(data, { x: 0.5, y: 1.55, w: 5.6, colW: [1.95, 0.73, 0.73, 0.73, 0.73, 0.73], rowH: 0.38, border: { type: "solid", pt: 0.75, color: TAN }, margin: [0, 0.06, 0, 0.06] });
    s.addChart(pres.charts.BAR, [
      { name: "Revenue", labels: yrs.map(y => y.replace("Year ", "Y")), values: rev },
      { name: "Total expenses", labels: yrs.map(y => y.replace("Year ", "Y")), values: rev.map((r, i) => cogs[i] + opex[i]) },
    ], { x: 6.3, y: 1.45, w: 3.25, h: 2.55, barDir: "col", chartColors: [GREEN, ORANGE], showTitle: true, title: "Revenue vs. expenses ($K)",
      titleFontSize: 10, titleColor: GREEN, titleFontFace: BODY, showLegend: true, legendPos: "b", legendFontSize: 8, legendColor: MUTED,
      catAxisLabelColor: MUTED, valAxisLabelColor: MUTED, valAxisLabelFontSize: 8, catAxisLabelFontSize: 9,
      valGridLine: { color: "E5DDC8", size: 0.5 }, catGridLine: { style: "none" }, barGapWidthPct: 50 });
    T(s, "Key assumptions", { x: 0.5, y: 4.0, w: 5.6, h: 0.25, fontSize: 11, bold: true, color: ORANGE });
    T(s, "Blended price ~$38/month · 15% app-store fee + ~12% AI & hosting · ~5% monthly churn · new apps launch per the roadmap · $250K pre-seed now, seed round around month 18 to fund year-2 growth · break-even during year 3.", {
      x: 0.5, y: 4.27, w: 9, h: 0.6, fontSize: 10.5, color: MUTED });
    T(s, "Projections, not guarantees.", { x: 6.3, y: 4.02, w: 3.25, h: 0.22, fontSize: 9, italic: true, color: MUTED, align: "right" });
    footer(s);
    s.addNotes("Five-year projections, built bottom-up from subscribers. Year-end paying subscribers grow from 650 to 13,000 as AppForge adds apps. Revenue uses a blended price of about $38 a month. Cost of revenue is about 27%: the 15% app-store fee plus about 12% for AI and hosting. Operating expenses cover marketing, contractors and later hires. We expect losses in years 1 and 2, funded by this round and a seed round, and to break even during year 3. These are projections, not guarantees.");
  }

  // ---------- 15. Team ----------
  {
    const s = light();
    eyebrow(s, "Team & Founders");
    title(s, "A founder who knows these customers first-hand");
    card(s, 0.5, 1.6, 4.4, 3.4);
    badge(s, "LuUser", 0.8, 1.82, 0.8);
    T(s, "James Chambers", { x: 1.8, y: 1.9, w: 3.0, h: 0.38, fontFace: HEAD, fontSize: 19, bold: true, color: GREEN });
    T(s, "Founder & CEO", { x: 1.8, y: 2.28, w: 3.0, h: 0.3, fontSize: 12, color: ORANGE, bold: true });
    const jb = ["Degree in computer science: builds the products himself",
      "40+ years in sales across multiple industries: the customer Ready Room is built for",
      "Started Cedarworks after seeing the workload his bishop carried, which led to PastorAI",
      "Built and launched both apps, plus AppForge, the studio's market-analysis and build engine"];
    T(s, bullets(jb), { x: 0.8, y: 2.75, w: 3.95, h: 2.15, fontSize: 11.5, paraSpaceAfter: 4 });
    T(s, "Hiring with this round", { x: 5.25, y: 1.65, w: 4.2, h: 0.3, fontSize: 11, bold: true, color: ORANGE });
    const hires = [["LuWrench", "Contract developer", "Speeds up launches of the next eight apps"],
      ["LuMegaphone", "Part-time growth marketer", "Runs paid acquisition and partnerships"],
      ["LuHandshake", "Advisors", "Ministry, sales-training and SaaS mentors"]];
    hires.forEach(([n, h, b], i) => {
      const y = 2.1 + i * 0.95;
      badge(s, n, 5.25, y, 0.55);
      T(s, h, { x: 6.0, y, w: 3.5, h: 0.28, fontSize: 14, bold: true, color: GREEN });
      T(s, b, { x: 6.0, y: y + 0.29, w: 3.5, h: 0.3, fontSize: 11.5, color: MUTED });
    });
    footer(s);
    s.addNotes("I've spent more than 40 years in sales across different industries, so Ready Room comes from my own experience. I also have a degree in computer science, which is how I built both apps and AppForge myself. PastorAI started when I saw how much my bishop carries every week and wanted to see what I could do. This round adds contract development and part-time marketing help so I can move faster without a heavy payroll.");
  }

  // ---------- 16. The Ask ----------
  {
    const s = dark();
    logo(s, 0.5, 0.5, 0.7);
    eyebrow(s, "The Ask", 1.45);
    s.addText("Raising $250K to take Cedarworks from 2 apps to 10", { x: 0.5, y: 1.75, w: 5.6, h: 1.3, fontFace: HEAD, fontSize: 28, bold: true, color: CREAM, margin: 0, valign: "top", isTextBox: true });
    const pts = ["Scale PastorAI and Ready Room with paid acquisition", "Launch 8 new market-picked apps with AppForge", "Reach ~1,000 paying subscribers in 18 months"];
    T(s, bullets(pts), { x: 0.5, y: 3.15, w: 5.6, h: 1.2, fontSize: 14, color: CREAM, paraSpaceAfter: 6 });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 6.5, y: 1.5, w: 3.0, h: 3.2, rectRadius: 0.1, fill: { color: GREEN2 }, line: { color: ORANGE, width: 1.25 } });
    T(s, "Let's talk", { x: 6.75, y: 1.7, w: 2.5, h: 0.3, fontSize: 11, bold: true, color: ORANGE });
    T(s, "James Chambers", { x: 6.75, y: 2.05, w: 2.5, h: 0.4, fontFace: HEAD, fontSize: 18, bold: true, color: CREAM });
    const contact = [["LuMail", "james@cedarworksapps.com"], ["LuPhone", "417-370-2084"], ["LuGlobe", "cedarworksapps.com"], ["LuSmartphone", "pastorai.io · tryreadyroom.com"]];
    contact.forEach(([n, t], i) => {
      const y = 2.65 + i * 0.47;
      s.addImage({ data: ic[n].o, x: 6.75, y: y + 0.04, w: 0.24, h: 0.24 });
      T(s, t, { x: 7.1, y, w: 2.3, h: 0.32, fontSize: 11, color: CREAM, valign: "middle" });
    });
    T(s, "Use of funds: 40% marketing · 30% building 8 new apps · 10% AI infrastructure · 10% operations · 10% runway", { x: 0.5, y: 4.2, w: 5.6, h: 0.45, fontSize: 11, color: CREAM });
    T(s, "Pre-seed · $250,000 · Terms open to discussion (SAFE or priced)", { x: 0.5, y: 4.7, w: 5.6, h: 0.3, fontSize: 11, italic: true, color: SAGE });
    footer(s, true);
    s.addNotes("We're raising $250K pre-seed. It lets us scale the two apps already in market and launch eight more that AppForge identifies. Thank you. I'd love to show you a live demo of PastorAI and Ready Room.");
  }


  await pres.writeFile({ fileName: OUT });
  console.log("wrote", OUT);
})();
