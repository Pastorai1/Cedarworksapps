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

  const cedar = await svgPng(CEDAR);
  const ic = {};
  for (const n of ["LuChurch", "LuMic", "LuTarget", "LuHammer", "LuRocket", "LuSearch", "LuUsers", "LuSmartphone",
    "LuDollarSign", "LuMegaphone", "LuLayers", "LuTrendingUp", "LuClock", "LuFileText", "LuBookOpen",
    "LuShieldCheck", "LuMail", "LuPhone", "LuGlobe", "LuBrain", "LuRepeat", "LuChartBar", "LuUser",
    "LuWrench", "LuHandshake", "LuZap", "LuStore", "LuBriefcase"]) {
    ic[n] = { o: await icon(n, ORANGE), g: await icon(n, GREEN), c: await icon(n, CREAM) };
  }

  const TOTAL = 15;
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
    s.addText("CEDARWORKS APPS  ·  CONFIDENTIAL", { x: 0.5, y: 5.22, w: 5, h: 0.25, fontFace: BODY, fontSize: 8,
      color: dark ? SAGE : MUTED, charSpacing: 2, margin: 0, isTextBox: true });
    s.addText(`${num} / ${TOTAL}`, { x: 8.5, y: 5.22, w: 1, h: 0.25, fontFace: BODY, fontSize: 8, align: "right",
      color: dark ? SAGE : MUTED, margin: 0, isTextBox: true });
  }
  function eyebrow(s, text, y = 0.42, dark = false) {
    s.addText(text.toUpperCase(), { x: 0.5, y, w: 9, h: 0.26, fontFace: BODY, fontSize: 10, bold: true,
      color: ORANGE, charSpacing: 3, margin: 0, isTextBox: true });
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

  // ---------- 1. Title ----------
  {
    const s = dark();
    logo(s, 0.5, 0.55, 0.8);
    s.addText("CEDARWORKS APPS", { x: 1.5, y: 0.72, w: 5, h: 0.45, fontFace: BODY, fontSize: 14, bold: true, color: CREAM, charSpacing: 4, margin: 0, isTextBox: true });
    s.addText("Practical AI apps,\nbuilt to do real work.", { x: 0.5, y: 1.65, w: 6.6, h: 1.7, fontFace: HEAD, fontSize: 40, bold: true, color: CREAM, margin: 0, valign: "top", isTextBox: true });
    s.addText("A small studio shipping focused, subscription AI apps: two live in the app stores today, and AppForge to build the next eight.", {
      x: 0.5, y: 3.45, w: 6.2, h: 0.75, fontFace: HEAD, italic: true, fontSize: 15, color: SAGE, margin: 0, valign: "top", isTextBox: true });
    // Ask chip
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 7.2, y: 1.7, w: 2.3, h: 1.55, rectRadius: 0.1, fill: { color: GREEN2 }, line: { color: ORANGE, width: 1.25 } });
    s.addText([
      { text: "PRE-SEED ROUND", options: { fontSize: 9, bold: true, color: ORANGE, charSpacing: 2, breakLine: true } },
      { text: "$250K", options: { fontSize: 36, bold: true, color: CREAM, fontFace: HEAD, breakLine: true } },
      { text: "Springfield, MO", options: { fontSize: 11, color: SAGE } },
    ], { x: 7.2, y: 1.8, w: 2.3, h: 1.35, fontFace: BODY, align: "center", valign: "middle", margin: 0, isTextBox: true });
    s.addText("James Chambers, Founder  ·  james@cedarworksapps.com  ·  417-370-2084  ·  cedarworksapps.com", {
      x: 0.5, y: 4.6, w: 9, h: 0.3, fontFace: BODY, fontSize: 11, color: CREAM, margin: 0, isTextBox: true });
    footer(s, true);
    s.addNotes("Cedarworks Apps is a Springfield, Missouri AI app studio. We build focused, subscription AI apps for specific professions. Two are live in the Apple App Store and Google Play today. We're raising a $250K pre-seed round to market them and use our in-house AppForge engine to launch eight more.");
  }

  // ---------- 2. Problem ----------
  {
    const s = light();
    eyebrow(s, "The problem");
    title(s, "Busy professionals get generic AI, not tools built for their job");
    const rows = [
      ["LuClock", "Pastors are stretched thin", "Sermons, Bible studies, devotionals, kids' ministry, social posts and translation fill the week. Most of it starts from a blank page, spread across a dozen tabs."],
      ["LuTarget", "Sales reps practice on real prospects", "Discovery and objection-handling are learned live, on the calls that matter. Every fumbled objection is a lost deal and lost commission."],
      ["LuBrain", "General chatbots don't close the gap", "A blank chat box doesn't know the workflow, the vocabulary or what 'good' looks like. People give up or spend more time prompting than working."],
    ];
    rows.forEach(([n, h, b], i) => {
      const y = 1.75 + i * 1.08;
      badge(s, n, 0.5, y, 0.62);
      T(s, h, { x: 1.35, y: y - 0.02, w: 7.9, h: 0.32, fontFace: HEAD, fontSize: 17, bold: true, color: GREEN });
      T(s, b, { x: 1.35, y: y + 0.33, w: 7.9, h: 0.6, fontSize: 12.5, color: MUTED });
    });
    footer(s);
    s.addNotes("Our customers are skilled professionals whose days are full of repeatable work that AI could help with. Generic AI tools make them do the translating. Pastors need ministry-shaped help. Sales reps need a safe place to practice before a real call.");
  }

  // ---------- 3. Solution ----------
  {
    const s = dark();
    eyebrow(s, "Our solution");
    title(s, "Focused AI apps that each do one job really well", true);
    T(s, "Cedarworks Apps builds niche, subscription AI apps: one profession, one clear job, a running start in seconds. Each app has its own workflow, prompts and guardrails, so users get results instead of a blank chat box.", {
      x: 0.5, y: 1.7, w: 4.3, h: 1.6, fontSize: 14, color: CREAM });
    T(s, "One studio, many apps. Shared code, shared infrastructure, one founder-led team.", {
      x: 0.5, y: 3.35, w: 4.3, h: 0.8, fontFace: HEAD, italic: true, fontSize: 14, color: SAGE });
    const pts = [["LuTarget", "Built for one job", "Purpose-built workflows, not generic chat"],
      ["LuSmartphone", "Everywhere users are", "iOS, Android and web"],
      ["LuRepeat", "Recurring revenue", "Monthly subscriptions with free trials"],
      ["LuLayers", "A portfolio, not a bet", "Many apps spread the risk and share costs"]];
    pts.forEach(([n, h, b], i) => {
      const y = 1.5 + i * 0.86;
      badge(s, n, 5.3, y, 0.55, true);
      T(s, h, { x: 6.05, y: y, w: 3.45, h: 0.28, fontSize: 14, bold: true, color: CREAM });
      T(s, b, { x: 6.05, y: y + 0.28, w: 3.45, h: 0.3, fontSize: 11.5, color: SAGE });
    });
    footer(s, true);
    s.addNotes("We don't try to build one app for everyone. We build many small apps that each serve one audience extremely well. That keeps each product simple and lets us reuse our tech across the whole portfolio.");
  }

  // ---------- 4. PastorAI ----------
  {
    const s = light();
    eyebrow(s, "Live product #1  ·  pastorai.io");
    title(s, "PastorAI: an AI ministry assistant for the whole week");
    card(s, 0.5, 1.6, 5.3, 3.4);
    T(s, "Features", { x: 0.8, y: 1.8, w: 4.8, h: 0.3, fontSize: 11, bold: true, color: ORANGE, charSpacing: 2 });
    const feats = ["Sermon and sermon-series planning", "Bible study and devotional writing", "Children's ministry lessons",
      "Church social media content", "Translation for multilingual congregations", "Output shaped to the pastor's own voice"];
    T(s, feats.map((f, i) => ({ text: f, options: { bullet: { indent: 14 }, breakLine: i < feats.length - 1 } })),
      { x: 0.8, y: 2.18, w: 4.8, h: 2.1, fontSize: 13.5, paraSpaceAfter: 5 });
    T(s, "Lightens the load without replacing the pastor: less busywork, more time with people.", {
      x: 0.8, y: 4.3, w: 4.8, h: 0.55, fontFace: HEAD, italic: true, fontSize: 12, color: GREEN });
    // stat column
    const stats = [["$47", "per month subscription"], ["~15", "trial members today"], ["iOS · Android · Web", "live in both app stores"]];
    stats.forEach(([big, small], i) => {
      const y = 1.6 + i * 1.17;
      card(s, 6.1, y, 3.4, 1.0, i === 0 ? GREEN : WHITE);
      T(s, big, { x: 6.35, y: y + 0.12, w: 3.0, h: 0.5, fontFace: HEAD, fontSize: i === 2 ? 17 : 28, bold: true, color: i === 0 ? CREAM : GREEN, valign: "middle" });
      T(s, small, { x: 6.35, y: y + 0.63, w: 3.0, h: 0.28, fontSize: 11, color: i === 0 ? SAGE : MUTED });
    });
    footer(s);
    s.addNotes("PastorAI puts the tools that fill a pastor's week in one place. It's live on iOS, Android and the web at $47 a month, with about 15 trial members today.");
  }

  // ---------- 5. Ready Room ----------
  {
    const s = light();
    eyebrow(s, "Live product #2  ·  tryreadyroom.com");
    title(s, "Ready Room: a flight simulator for sales calls");
    const stats = [["$29", "per month subscription"], ["~15", "trial members today"], ["iOS · Android · Web", "live in both app stores"]];
    stats.forEach(([big, small], i) => {
      const y = 1.6 + i * 1.17;
      card(s, 0.5, y, 3.4, 1.0, i === 0 ? GREEN : WHITE);
      T(s, big, { x: 0.75, y: y + 0.12, w: 3.0, h: 0.5, fontFace: HEAD, fontSize: i === 2 ? 17 : 28, bold: true, color: i === 0 ? CREAM : GREEN, valign: "middle" });
      T(s, small, { x: 0.75, y: y + 0.63, w: 3.0, h: 0.28, fontSize: 11, color: i === 0 ? SAGE : MUTED });
    });
    card(s, 4.2, 1.6, 5.3, 3.4);
    T(s, "How it works", { x: 4.5, y: 1.8, w: 4.8, h: 0.3, fontSize: 11, bold: true, color: ORANGE, charSpacing: 2 });
    const feats = ["A voice AI plays a realistic, often skeptical prospect", "It reacts to how you sell: warms up to good questions, shuts down when pushed",
      "Drills discovery and objections on price, timing and trust", "Easy, medium and hard levels",
      "An AI coach grades every session and says what to fix next"];
    T(s, feats.map((f, i) => ({ text: f, options: { bullet: { indent: 14 }, breakLine: i < feats.length - 1 } })),
      { x: 4.5, y: 2.18, w: 4.8, h: 2.2, fontSize: 13, paraSpaceAfter: 5 });
    T(s, "The best closers aren't born. They're rehearsed.", {
      x: 4.5, y: 4.45, w: 4.8, h: 0.4, fontFace: HEAD, italic: true, fontSize: 12.5, color: GREEN });
    footer(s);
    s.addNotes("Ready Room lets sales reps rehearse on an AI prospect that talks back out loud. Reps practice discovery and objection-handling, then get graded by an AI coach. It's live on iOS, Android and the web at $29 a month, with about 15 trial members today.");
  }

  // ---------- 6. AppForge ----------
  {
    const s = dark();
    eyebrow(s, "Our engine");
    title(s, "AppForge: our in-house system for finding and building the next app", true);
    const steps = [["LuSearch", "1  Analyze", "Scans the market for underserved niches with real demand and willingness to pay"],
      ["LuTarget", "2  Select", "Scores ideas on demand, competition and fit, and picks the strongest"],
      ["LuHammer", "3  Build", "Reuses the proven Cedarworks stack of AI, voice, billing and app-store shells"],
      ["LuRocket", "4  Launch", "Ships to iOS, Android and web with trials and subscriptions ready on day one"]];
    steps.forEach(([n, h, b], i) => {
      const x = 0.5 + i * 2.3;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 1.75, w: 2.05, h: 2.3, rectRadius: 0.08, fill: { color: GREEN2 }, line: { color: GREEN2 } });
      badge(s, n, x + 0.2, 1.95, 0.6, true);
      T(s, h, { x: x + 0.2, y: 2.68, w: 1.7, h: 0.32, fontSize: 15, bold: true, color: CREAM });
      T(s, b, { x: x + 0.2, y: 3.02, w: 1.7, h: 0.95, fontSize: 11, color: SAGE });
      if (i < 3) s.addText("›", { x: x + 2.05, y: 2.65, w: 0.25, h: 0.4, fontSize: 22, color: ORANGE, align: "center", margin: 0, isTextBox: true });
    });
    T(s, "Why it matters: each new app launches faster and cheaper than the last, and AppForge's market data picks where to build. Investor money goes into apps people already want.", {
      x: 0.5, y: 4.3, w: 9, h: 0.75, fontFace: HEAD, italic: true, fontSize: 13.5, color: CREAM });
    footer(s, true);
    s.addNotes("AppForge is our proprietary engine. It analyzes the market to find which apps are the best fit right now, then we build on our existing stack. PastorAI and Ready Room prove the build side works. The next eight apps will be chosen by AppForge's market analysis, not by guesswork.");
  }

  // ---------- 7. Traction ----------
  {
    const s = light();
    eyebrow(s, "Traction");
    title(s, "Built, shipped and live, before raising a dollar");
    const stats = [["2", "AI apps live on the\nApple App Store & Google Play"], ["~30", "trial members across\nPastorAI & Ready Room"],
      ["2", "subscription tiers\n($47 and $29 per month)"], ["1", "founder who built and\nlaunched both apps"]];
    stats.forEach(([big, small], i) => {
      const x = 0.5 + i * 2.3;
      card(s, x, 1.65, 2.05, 1.9);
      T(s, big, { x: x + 0.15, y: 1.8, w: 1.75, h: 0.85, fontFace: HEAD, fontSize: 44, bold: true, color: GREEN, align: "center", valign: "middle" });
      T(s, small, { x: x + 0.15, y: 2.7, w: 1.75, h: 0.7, fontSize: 11.5, color: MUTED, align: "center" });
    });
    badge(s, "LuTrendingUp", 0.5, 3.95, 0.6);
    T(s, "What's next: convert today's trial members to paid, then scale with paid acquisition. The product, stores and billing are already in place; distribution is the missing piece this round funds.", {
      x: 1.3, y: 3.92, w: 8.2, h: 0.95, fontSize: 13.5, color: INK });
    footer(s);
    s.addNotes("Cedarworks Apps LLC was formed in 2026. In that time we've shipped two complete AI apps to both major app stores with subscription billing, and have about 30 trial members. The product risk is largely retired. This round is about distribution.");
  }

  // ---------- 8. Market ----------
  {
    const s = light();
    eyebrow(s, "Market opportunity");
    title(s, "Two large markets to start, with room for eight more");
    card(s, 0.5, 1.6, 4.35, 3.35);
    badge(s, "LuChurch", 0.75, 1.82, 0.55);
    T(s, "PastorAI: churches", { x: 1.45, y: 1.9, w: 3.3, h: 0.35, fontFace: HEAD, fontSize: 16, bold: true, color: GREEN });
    T(s, "300,000+", { x: 0.75, y: 2.55, w: 3.9, h: 0.6, fontFace: HEAD, fontSize: 34, bold: true, color: ORANGE });
    T(s, "Protestant and independent congregations in the U.S. alone, most of them small, with few staff.", { x: 0.75, y: 3.18, w: 3.9, h: 0.55, fontSize: 12, color: MUTED });
    T(s, "At $47/mo ($564/yr), each 1% of U.S. churches is about $1.7M in annual recurring revenue.", { x: 0.75, y: 3.85, w: 3.9, h: 0.9, fontSize: 12.5, bold: true, color: INK });
    card(s, 5.15, 1.6, 4.35, 3.35);
    badge(s, "LuBriefcase", 5.4, 1.82, 0.55);
    T(s, "Ready Room: sales professionals", { x: 6.1, y: 1.9, w: 3.3, h: 0.35, fontFace: HEAD, fontSize: 16, bold: true, color: GREEN });
    T(s, "Millions", { x: 5.4, y: 2.55, w: 3.9, h: 0.6, fontFace: HEAD, fontSize: 34, bold: true, color: ORANGE });
    T(s, "of U.S. workers hold sales jobs, and new reps and small teams rarely get structured call practice.", { x: 5.4, y: 3.18, w: 3.9, h: 0.55, fontSize: 12, color: MUTED });
    T(s, "At $29/mo ($348/yr), every 10,000 reps is about $3.5M in annual recurring revenue.", { x: 5.4, y: 3.85, w: 3.9, h: 0.9, fontSize: 12.5, bold: true, color: INK });
    footer(s);
    s.addNotes("Sources to cite if asked: the Hartford Institute for Religion Research and the U.S. Religion Census estimate roughly 300,000 to 380,000 U.S. congregations. The Bureau of Labor Statistics counts millions of U.S. workers in sales occupations. Revenue figures are simple math on our current prices, not forecasts. AppForge adds new markets with each of the next eight apps.");
  }

  // ---------- 9. Business model ----------
  {
    const s = light();
    eyebrow(s, "Business model");
    title(s, "Recurring subscriptions across a growing portfolio");
    const cols = [["LuRepeat", "App subscriptions", "Monthly plans with free trials, billed through the App Store, Google Play and the web. PastorAI $47/mo · Ready Room $29/mo."],
      ["LuUsers", "Group plans", "Multi-seat plans for sales teams and church staffs: a higher contract value from one buyer."],
      ["LuWrench", "Custom AI apps", "Custom AI software built for local businesses: service revenue that helps fund the studio and brings in app ideas."]];
    cols.forEach(([n, h, b], i) => {
      const x = 0.5 + i * 3.07;
      card(s, x, 1.65, 2.85, 2.35);
      badge(s, n, x + 0.25, 1.85, 0.6);
      T(s, h, { x: x + 0.25, y: 2.6, w: 2.4, h: 0.32, fontFace: HEAD, fontSize: 15, bold: true, color: GREEN });
      T(s, b, { x: x + 0.25, y: 2.95, w: 2.4, h: 1.0, fontSize: 11.5, color: MUTED });
    });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 4.2, w: 9, h: 0.78, rectRadius: 0.08, fill: { color: TAN }, line: { color: TAN } });
    T(s, "Portfolio effect: every app shares one codebase, one AI stack and one founder's overhead. Each new app adds revenue with little added fixed cost.", {
      x: 0.75, y: 4.3, w: 8.5, h: 0.6, fontSize: 12.5, color: GREEN, valign: "middle" });
    footer(s);
    s.addNotes("Our core revenue is monthly subscriptions. Team and church plans raise the value of each sale. Custom app work for businesses is a secondary revenue line and a source of new app ideas. Because all the apps share infrastructure, margins improve as the portfolio grows.");
  }

  // ---------- 10. Go-to-market ----------
  {
    const s = light();
    eyebrow(s, "Go-to-market");
    title(s, "How we'll turn trials into paying customers");
    card(s, 0.5, 1.6, 4.35, 3.4);
    badge(s, "LuChurch", 0.75, 1.8, 0.5);
    T(s, "PastorAI", { x: 1.4, y: 1.88, w: 3.2, h: 0.35, fontFace: HEAD, fontSize: 16, bold: true, color: GREEN });
    const pa = ["Targeted Facebook & YouTube ads to pastors and church staff", "LinkedIn outreach to ministry leaders",
      "Partnerships with denominations, church networks & seminaries", "Pastor conferences, webinars and free sermon-prep resources"];
    T(s, pa.map((f, i) => ({ text: f, options: { bullet: { indent: 14 }, breakLine: i < pa.length - 1 } })),
      { x: 0.75, y: 2.5, w: 3.9, h: 2.4, fontSize: 12.5, paraSpaceAfter: 6 });
    card(s, 5.15, 1.6, 4.35, 3.4);
    badge(s, "LuMegaphone", 5.4, 1.8, 0.5);
    T(s, "Ready Room", { x: 6.05, y: 1.88, w: 3.2, h: 0.35, fontFace: HEAD, fontSize: 16, bold: true, color: GREEN });
    const rr = ["LinkedIn and short-form video ads showing live AI role-plays", "Direct outreach to sales managers for team plans",
      "Sales-trainer, coach and bootcamp affiliate partners", "App Store search optimization and referral rewards"];
    T(s, rr.map((f, i) => ({ text: f, options: { bullet: { indent: 14 }, breakLine: i < rr.length - 1 } })),
      { x: 5.4, y: 2.5, w: 3.9, h: 2.4, fontSize: 12.5, paraSpaceAfter: 6 });
    footer(s);
    s.addNotes("The biggest use of funds is customer acquisition. We'll start with small, measured ad tests on each channel, then double down on the ones with the best cost per paying subscriber. Free trials are already built into both apps.");
  }

  // ---------- 11. Competition ----------
  {
    const s = light();
    eyebrow(s, "Competitive landscape");
    title(s, "Why Cedarworks wins");
    const hdr = ["", "Generic AI chatbots", "Enterprise training / point tools", "Cedarworks apps"];
    const rows = [["Built for one profession", "No", "Partly", "Yes"], ["Ready to use with no prompting", "No", "Partly", "Yes"],
      ["Voice role-play & coaching (Ready Room)", "No", "Often costly", "Yes"], ["Price for individuals & small teams", "Low", "High", "Low ($29–$47)"],
      ["New products launched quickly", "N/A", "Slow", "AppForge"]];
    const cell = (t, o = {}) => ({ text: t, options: Object.assign({ fontFace: BODY, fontSize: 12, color: INK, valign: "middle" }, o) });
    const data = [hdr.map((h, i) => cell(h, { bold: true, color: i === 3 ? CREAM : GREEN, fill: { color: i === 3 ? GREEN : TAN }, align: i ? "center" : "left" }))]
      .concat(rows.map(r => r.map((c, i) => cell(c, { align: i ? "center" : "left", bold: i === 3, color: i === 3 ? GREEN : INK, fill: { color: i === 3 ? "E3EDE3" : WHITE } }))));
    s.addTable(data, { x: 0.5, y: 1.55, w: 9, colW: [3.3, 1.8, 2.1, 1.8], rowH: 0.47, border: { type: "solid", pt: 0.75, color: TAN }, margin: [0, 0.12, 0, 0.12] });
    T(s, "Our moat is speed and focus: a small, efficient studio that goes deep on niches the big platforms ignore.", {
      x: 0.5, y: 4.55, w: 9, h: 0.45, fontFace: HEAD, italic: true, fontSize: 13, color: GREEN });
    footer(s);
    s.addNotes("People can use general chatbots, but they have to do all the prompting and still get generic output. Enterprise sales-training platforms are priced for large companies. We sit in between: purpose-built, affordable, and quick to launch in new niches.");
  }

  // ---------- 12. Use of funds ----------
  {
    const s = light();
    eyebrow(s, "Use of funds");
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
    T(s, "About 18 months of runway on a lean, founder-led cost base.", { x: 4.5, y: 4.95 - 0.25, w: 5, h: 0.3, fontSize: 11, italic: true, color: MUTED });
    footer(s);
    s.addNotes("40% goes to customer acquisition for PastorAI and Ready Room, the fastest path to revenue. 30% funds building eight new apps with AppForge, including contract development, design and QA. The rest covers AI compute and hosting, legal and operations, and a founder runway reserve.");
  }

  // ---------- 13. Roadmap ----------
  {
    const s = light();
    eyebrow(s, "18-month roadmap");
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

  // ---------- 14. Team ----------
  {
    const s = light();
    eyebrow(s, "Team");
    title(s, "Founder-led, with a proven ability to ship");
    card(s, 0.5, 1.6, 4.4, 3.35);
    badge(s, "LuUser", 0.8, 1.85, 0.85);
    T(s, "James Chambers", { x: 1.85, y: 1.95, w: 2.9, h: 0.38, fontFace: HEAD, fontSize: 19, bold: true, color: GREEN });
    T(s, "Founder & CEO", { x: 1.85, y: 2.33, w: 2.9, h: 0.3, fontSize: 12, color: ORANGE, bold: true });
    const jb = ["Designed, built and launched PastorAI and Ready Room", "Shipped both to the Apple App Store and Google Play",
      "Created AppForge, the studio's market-analysis and build engine", "Leads product, AI, marketing and customer relationships"];
    T(s, jb.map((f, i) => ({ text: f, options: { bullet: { indent: 14 }, breakLine: i < jb.length - 1 } })),
      { x: 0.8, y: 2.95, w: 3.9, h: 1.9, fontSize: 12, paraSpaceAfter: 5 });
    T(s, "Hiring with this round", { x: 5.25, y: 1.65, w: 4.2, h: 0.3, fontSize: 11, bold: true, color: ORANGE, charSpacing: 2 });
    const hires = [["LuWrench", "Contract developer", "Speeds up launches of the next eight apps"],
      ["LuMegaphone", "Part-time growth marketer", "Runs paid acquisition and partnerships"],
      ["LuHandshake", "Advisors", "Ministry, sales-training and SaaS mentors"]];
    hires.forEach(([n, h, b], i) => {
      const y = 2.1 + i * 0.95;
      badge(s, n, 5.25, y, 0.55);
      T(s, h, { x: 6.0, y: y + 0.0, w: 3.5, h: 0.28, fontSize: 14, bold: true, color: GREEN });
      T(s, b, { x: 6.0, y: y + 0.29, w: 3.5, h: 0.3, fontSize: 11.5, color: MUTED });
    });
    footer(s);
    s.addNotes("I'm a solo founder who has already built and shipped two AI apps to both major app stores. This round lets me add contract development and part-time marketing help so I can go faster without a heavy payroll.");
  }

  // ---------- 15. The Ask ----------
  {
    const s = dark();
    logo(s, 0.5, 0.5, 0.7);
    eyebrow(s, "The ask", 1.45);
    s.addText("Raising $250K to take Cedarworks from 2 apps to 10", { x: 0.5, y: 1.75, w: 5.6, h: 1.3, fontFace: HEAD, fontSize: 28, bold: true, color: CREAM, margin: 0, valign: "top", isTextBox: true });
    const pts = ["Scale PastorAI and Ready Room with paid acquisition", "Launch 8 new market-picked apps with AppForge", "Reach ~1,000 paying subscribers in 18 months"];
    T(s, pts.map((f, i) => ({ text: f, options: { bullet: { indent: 14 }, breakLine: i < pts.length - 1 } })),
      { x: 0.5, y: 3.15, w: 5.6, h: 1.2, fontSize: 14, color: CREAM, paraSpaceAfter: 6 });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 6.5, y: 1.5, w: 3.0, h: 3.2, rectRadius: 0.1, fill: { color: GREEN2 }, line: { color: ORANGE, width: 1.25 } });
    T(s, "LET'S TALK", { x: 6.75, y: 1.7, w: 2.5, h: 0.3, fontSize: 10, bold: true, color: ORANGE, charSpacing: 3 });
    T(s, "James Chambers", { x: 6.75, y: 2.05, w: 2.5, h: 0.4, fontFace: HEAD, fontSize: 18, bold: true, color: CREAM });
    const contact = [["LuMail", "james@cedarworksapps.com"], ["LuPhone", "417-370-2084"], ["LuGlobe", "cedarworksapps.com"], ["LuSmartphone", "pastorai.io · tryreadyroom.com"]];
    contact.forEach(([n, t], i) => {
      const y = 2.65 + i * 0.47;
      s.addImage({ data: ic[n].o, x: 6.75, y: y + 0.04, w: 0.24, h: 0.24 });
      T(s, t, { x: 7.1, y, w: 2.3, h: 0.32, fontSize: 11, color: CREAM, valign: "middle" });
    });
    T(s, "Pre-seed · Terms open to discussion (SAFE or priced)", { x: 0.5, y: 4.55, w: 5.6, h: 0.3, fontSize: 11, italic: true, color: SAGE });
    footer(s, true);
    s.addNotes("We're raising $250K pre-seed. It lets us scale the two apps already in market and launch eight more that AppForge identifies. Thank you. I'd love to show you a live demo of PastorAI and Ready Room.");
  }

  await pres.writeFile({ fileName: OUT });
  console.log("wrote", OUT);
})();
