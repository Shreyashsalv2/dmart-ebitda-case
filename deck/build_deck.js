// Builds deck/DMart_EBITDA_Case.pptx from build/case.json (run scripts/case.py first).
// Usage: node deck/build_deck.js [path/to/apply_theme.js]
const path = require("path");
const fs = require("fs");
const pptxgen = require("pptxgenjs");

const ROOT = path.resolve(__dirname, "..");
const c = JSON.parse(fs.readFileSync(path.join(ROOT, "build", "case.json"), "utf8"));
const OUT = path.join(__dirname, "DMart_EBITDA_Case.pptx");

const THEME = {
  name: "DMart Case",
  headFontFace: "Cambria",
  bodyFontFace: "Calibri",
  colors: {
    dk1: "1B1B1B", lt1: "FFFFFF", dk2: "1B4332", lt2: "EEF5F0",
    accent1: "2D6A4F", accent2: "C0392B", accent3: "E9C46A", accent4: "74C69D",
    accent5: "40916C", accent6: "6C757D", hlink: "2D6A4F", folHlink: "1B4332",
  },
};
const HEX = THEME.colors;

// ---------- number formatting ----------
const n0 = (x) => Math.round(x).toLocaleString("en-IN");
const cr = (x) => `₹${n0(x)} cr`;
const pct = (x, d = 1) => `${x < 0 ? "−" : ""}${Math.abs(x * 100).toFixed(d)}%`;
const gpct = (x, d = 1) => `${x >= 0 ? "+" : "−"}${Math.abs(x * 100).toFixed(d)}%`;
const bps = (x) => `${x >= 0 ? "+" : "−"}${Math.abs(Math.round(x))} bps`;
const signed = (x) => `${x >= 0 ? "+" : "−"}${n0(Math.abs(x))}`;

const P25 = c.pnl.FY25, P26 = c.pnl.FY26, G = c.growth, O = c.operations;
const base = c.scenarios.Base, caut = c.scenarios.Cautious, stretch = c.scenarios.Stretch;
const dEmp = (P26.employee_pct - P25.employee_pct) * 1e4;
const dGm = (P26.gross_margin_pct - P25.gross_margin_pct) * 1e4;
const dOth = (P26.other_expenses_pct - P25.other_expenses_pct) * 1e4;
const dEbitda = (P26.ebitda_pct - P25.ebitda_pct) * 1e4;

async function main() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "Shreyash Singh Salv";
  pres.title = "DMart EBITDA improvement case";
  pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
  const C = pres.SchemeColor;

  pres.defineSlideMaster({
    title: "DARK_TITLE",
    background: { color: C.text2 },
    objects: [
      { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 1.35, w: 8.8, h: 1.4, fontSize: 36, bold: true, color: C.background1, valign: "bottom", align: "left" }, text: "" } },
      { placeholder: { options: { name: "body", type: "body", x: 0.6, y: 2.9, w: 8.8, h: 0.9, fontSize: 18, color: "95D5B2", valign: "top", align: "left" }, text: "" } },
    ],
  });
  pres.defineSlideMaster({
    title: "CONTENT",
    background: { color: C.background1 },
    margin: [0.5, 0.5, 0.6, 0.5],
    objects: [
      { placeholder: { options: { name: "title", type: "title", x: 0.5, y: 0.3, w: 9.0, h: 0.95, fontSize: 26, bold: true, color: C.text2, valign: "middle", align: "left" }, text: "" } },
      { text: { text: "DMart EBITDA case · standalone FY25–FY26 · from public filings", options: { x: 0.5, y: 5.2, w: 7.5, h: 0.3, fontSize: 10, color: C.accent6, fontFace: "Calibri" } } },
    ],
    slideNumber: { x: 9.0, y: 5.2, w: 0.5, h: 0.3, fontSize: 10, color: C.accent6, align: "right" },
  });

  const content = (section, titleText, notes) => {
    const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: section });
    s.addText(titleText, { placeholder: "title" });
    if (notes) s.addNotes(notes);
    return s;
  };
  const card = (s, x, y, w, h, name) =>
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.08, fill: { color: C.background2 }, line: { color: C.background2 }, objectName: name });
  const badge = (s, x, y, label, name) => {
    s.addShape(pres.shapes.OVAL, { x, y, w: 0.42, h: 0.42, fill: { color: C.accent1 }, line: { color: C.accent1 }, objectName: name });
    s.addText(label, { x, y, w: 0.42, h: 0.42, fontSize: 14, bold: true, color: C.background1, align: "center", valign: "middle", margin: 0, isTextBox: true, objectName: name + " label" });
  };
  const source = (s, text) =>
    s.addText(text, { x: 0.5, y: 4.88, w: 9.0, h: 0.3, fontSize: 10, italic: true, color: C.accent6, isTextBox: true, margin: 0, objectName: "source" });

  // 1 ---------- Title ----------
  pres.addSection({ title: "Opening" });
  const s1 = pres.addSlide({ masterName: "DARK_TITLE", sectionTitle: "Opening" });
  s1.addText("DMart grew 16% in FY26. Its EBITDA margin did not move.", { placeholder: "title" });
  s1.addText("An EBITDA improvement case built from public filings · Avenue Supermarts, standalone, FY25–FY26", { placeholder: "body" });
  s1.addText("Shreyash Singh Salv · November 2026", { x: 0.6, y: 4.3, w: 6, h: 0.35, fontSize: 14, color: C.background1, isTextBox: true, objectName: "author" });
  s1.addText("Independent student analysis of DMart's published filings. Not investment advice; not affiliated with or reviewed by DMart.", { x: 0.6, y: 4.75, w: 8.8, h: 0.4, fontSize: 10, italic: true, color: "95D5B2", isTextBox: true, objectName: "disclaimer" });
  s1.addNotes("Purpose: show how a consulting team would explain a flat margin and size improvement levers, using only public data. All numbers trace to DMart's NSE filings; lever sizes are my assumptions and are labelled as such.");

  // 2 ---------- Executive summary ----------
  pres.addSection({ title: "Summary" });
  const s2 = content("Summary", "Employee cost grew twice as fast as sales, cancelling out every other margin gain",
    "Read top to bottom: what happened, why, what to do. The base case is deliberately modest.");
  const findings = [
    `Revenue grew ${pct(G.revenue)} to ${cr(P26.revenue)} and EBITDA ${pct(G.ebitda)} to ${cr(P26.ebitda)}; the margin went from ${pct(P25.ebitda_pct, 2)} to ${pct(P26.ebitda_pct, 2)}.`,
    `Gross margin rose ${Math.round(dGm)} bps and other expenses fell ${Math.round(-dOth)} bps of revenue, but employee cost grew ${pct(G.employee)}, adding ${Math.round(dEmp)} bps.`,
    `DMart added ${O.stores_added_fy26} stores, ${O.stores_added_q4fy26} of them in Q4; revenue per sq ft of year-end space fell ${pct(-O.density_change_pct)}.`,
    `Three levers worth about ${cr(base.total)} in the base case would lift the FY26 margin to ${pct(base.proforma_margin, 2)}.`,
  ];
  findings.forEach((t, i) => {
    const y = 1.45 + i * 0.85;
    badge(s2, 0.5, y + 0.08, String(i + 1), `finding ${i + 1}`);
    s2.addText(t, { x: 1.1, y, w: 5.2, h: 0.75, fontSize: 14, color: C.text1, valign: "middle", isTextBox: true, margin: 0, objectName: `finding ${i + 1} text` });
  });
  card(s2, 6.7, 1.45, 2.8, 3.25, "headline card");
  s2.addText([
    { text: `+${Math.round(base.margin_uplift_bps)} bps`, options: { fontSize: 44, bold: true, color: C.accent1, breakLine: true } },
    { text: "base-case EBITDA margin uplift on FY26", options: { fontSize: 14, color: C.text1, breakLine: true } },
    { text: " ", options: { fontSize: 10, breakLine: true } },
    { text: cr(base.total), options: { fontSize: 28, bold: true, color: C.text2, breakLine: true } },
    { text: "of extra EBITDA from three levers", options: { fontSize: 14, color: C.text1 } },
  ], { x: 6.9, y: 1.6, w: 2.4, h: 2.95, valign: "middle", isTextBox: true, margin: 0, objectName: "headline numbers" });

  // 3 ---------- Situation ----------
  pres.addSection({ title: "Diagnosis" });
  const s3 = content("Diagnosis", "Stores and space grew about 20%; revenue grew 16%",
    "Scale grew faster than sales. That is normal in a year of heavy openings, because new stores take time to reach mature sales.");
  const stats = [
    [cr(P26.revenue), "Revenue from operations", `${gpct(G.revenue)} vs FY25`],
    [cr(P26.ebitda), "EBITDA", `${gpct(G.ebitda)} · margin ${pct(P26.ebitda_pct, 2)}`],
    [`${O.stores.FY25} → ${O.stores.FY26}`, "Stores at year end", gpct(O.stores.FY26 / O.stores.FY25 - 1)],
    [`${O.area_msqft.FY25} → ${O.area_msqft.FY26}`, "Retail area, million sq ft", gpct(O.area_growth_pct)],
  ];
  stats.forEach(([big, label, sub], i) => {
    const x = 0.5 + i * 2.3;
    card(s3, x, 1.55, 2.1, 2.6, `stat card ${i + 1}`);
    s3.addText([
      { text: big, options: { fontSize: 24, bold: true, color: C.text2, breakLine: true } },
      { text: label, options: { fontSize: 14, color: C.text1, breakLine: true } },
      { text: sub, options: { fontSize: 14, bold: true, color: C.accent1 } },
    ], { x: x + 0.15, y: 1.7, w: 1.8, h: 2.3, valign: "middle", isTextBox: true, margin: 0, objectName: `stat ${i + 1}` });
  });
  source(s3, "Source: DMart NSE filings and press release, FY25–FY26 (standalone). Growth figures are FY26 vs FY25.");

  // 4 ---------- Cost structure ----------
  const s4 = content("Diagnosis", "Gross margin and other costs improved; employee cost did not",
    "All lines are % of revenue. A small move matters: on FY26 revenue, 10 bps is about ₹67 crore of EBITDA.");
  const labels = ["Gross margin", "Employee cost", "Other expenses", "EBITDA"];
  const keys = ["gross_margin_pct", "employee_pct", "other_expenses_pct", "ebitda_pct"];
  s4.addChart(pres.charts.BAR, [
    { name: "FY25", labels, values: keys.map((k) => +(P25[k] * 100).toFixed(2)) },
    { name: "FY26", labels, values: keys.map((k) => +(P26[k] * 100).toFixed(2)) },
  ], {
    x: 0.5, y: 1.4, w: 5.6, h: 3.4, barDir: "col", barGrouping: "clustered", barGapWidthPct: 70,
    chartColors: [HEX.accent4, HEX.dk2], showLegend: true, legendPos: "t", legendFontSize: 12, legendFontFace: "+mn-lt",
    showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: '0.00"%"', dataLabelFontSize: 10, dataLabelFontFace: "+mn-lt", dataLabelColor: HEX.dk1,
    catAxisLabelFontSize: 11, catAxisLabelFontFace: "+mn-lt", catAxisLabelColor: HEX.dk1,
    valAxisHidden: true, valGridLine: { style: "none" }, catGridLine: { style: "none" },
    showTitle: true, title: "% of revenue", titleFontSize: 12, titleFontFace: "+mn-lt", titleColor: HEX.accent6,
    objectName: "cost structure chart",
  });
  const moves = [
    ["Gross margin", bps(dGm), C.accent1],
    ["Employee cost", bps(dEmp), C.accent2],
    ["Other expenses", bps(dOth), C.accent1],
    ["EBITDA margin", bps(dEbitda), C.text2],
  ];
  s4.addText("Change FY25 → FY26, bps of revenue", { x: 6.4, y: 1.45, w: 3.1, h: 0.35, fontSize: 14, bold: true, color: C.text2, isTextBox: true, margin: 0, objectName: "moves header" });
  moves.forEach(([k, v, col], i) => {
    const y = 1.95 + i * 0.68;
    card(s4, 6.4, y, 3.1, 0.55, `move ${i + 1}`);
    s4.addText(k, { x: 6.55, y, w: 1.7, h: 0.55, fontSize: 14, color: C.text1, valign: "middle", isTextBox: true, margin: 0, objectName: `move ${i + 1} label` });
    s4.addText(v, { x: 8.2, y, w: 1.2, h: 0.55, fontSize: 16, bold: true, color: col, align: "right", valign: "middle", isTextBox: true, margin: 0, objectName: `move ${i + 1} value` });
  });
  source(s4, "For costs, a positive change means a higher share of revenue (worse). Source: DMart NSE filings; author's calculations.");

  // 5 ---------- Bridge (drawn with shapes so every bar is labelled) ----------
  const s5 = content("Diagnosis", `A rising employee-cost share cost ${cr(-c.bridge[3].value)} and erased the margin gains`,
    "Bridge logic: revenue growth is valued at the FY25 margin; each other bar is the change in that line's share of revenue times FY26 revenue. The bars add up exactly to FY26 EBITDA. The axis starts at ₹4,000 crore.");
  const axisMin = 4000, axisMax = 5500, top = 1.55, bottom = 4.4;
  const yOf = (v) => bottom - ((v - axisMin) / (axisMax - axisMin)) * (bottom - top);
  const barW = 1.05, gap = 0.42, x0 = 0.75;
  let running = 0;
  const names = ["FY25 EBITDA", "Revenue growth at FY25 margin", "Gross margin rate", "Employee cost rate", "Other expenses rate", "FY26 EBITDA"];
  c.bridge.forEach((b, i) => {
    const x = x0 + i * (barW + gap);
    const isTotal = i === 0 || i === c.bridge.length - 1;
    let lo, hi, colour;
    if (isTotal) { lo = axisMin; hi = b.value; colour = C.text2; running = b.value; }
    else { const before = running; running += b.value; lo = Math.min(before, running); hi = Math.max(before, running); colour = b.value >= 0 ? C.accent1 : C.accent2; }
    s5.addShape(pres.shapes.RECTANGLE, { x, y: yOf(hi), w: barW, h: Math.max(yOf(lo) - yOf(hi), 0.02), fill: { color: colour }, line: { color: colour }, objectName: `bridge bar ${i + 1}` });
    const label = isTotal ? n0(b.value) : signed(b.value);
    s5.addText(label, { x: x - 0.1, y: yOf(hi) - 0.32, w: barW + 0.2, h: 0.3, fontSize: 12, bold: true, color: isTotal ? C.text2 : colour, align: "center", isTextBox: true, margin: 0, objectName: `bridge value ${i + 1}` });
    s5.addText(names[i], { x: x - 0.12, y: bottom + 0.05, w: barW + 0.24, h: 0.45, fontSize: 10, color: C.text1, align: "center", valign: "top", isTextBox: true, margin: 0, objectName: `bridge label ${i + 1}` });
  });
  s5.addShape(pres.shapes.LINE, { x: 0.6, y: bottom, w: 8.8, h: 0, line: { color: HEX.accent6, width: 0.75 }, objectName: "bridge axis" });
  source(s5, "INR crore; axis starts at ₹4,000 crore. Source: DMart NSE filings; author's calculations.");

  // 6 ---------- Hypothesis ----------
  const s6 = content("Diagnosis", "Hypothesis: new stores were staffed before they reached full sales",
    "This is a hypothesis from public data, not a finding. Testing it needs store-level sales and staffing by opening date, which only DMart has.");
  const evidence = [
    [`${O.stores_added_q4fy26} of ${O.stores_added_fy26}`, "new stores opened in Q4 FY26, leaving little time to trade"],
    [`${pct(O.density_change_pct)}`, `revenue per sq ft of year-end space (₹${n0(O.revenue_per_yearend_sqft.FY25)} → ₹${n0(O.revenue_per_yearend_sqft.FY26)})`],
    [gpct(G.employee), `employee cost growth vs revenue ${gpct(G.revenue)}; per store ₹${O.employee_cost_per_store_cr.FY25.toFixed(2)} → ₹${O.employee_cost_per_store_cr.FY26.toFixed(2)} cr`],
  ];
  evidence.forEach(([big, txt], i) => {
    const x = 0.5 + i * 3.1;
    card(s6, x, 1.5, 2.9, 2.0, `evidence ${i + 1}`);
    s6.addText([
      { text: big, options: { fontSize: 30, bold: true, color: i === 2 ? C.accent2 : C.text2, breakLine: true } },
      { text: txt, options: { fontSize: 14, color: C.text1 } },
    ], { x: x + 0.15, y: 1.6, w: 2.6, h: 1.8, valign: "top", isTextBox: true, margin: 0, objectName: `evidence ${i + 1} text` });
  });
  s6.addText([
    { text: "Below EBITDA, profit grew slower still: ", options: { bold: true } },
    { text: `depreciation ${gpct(G.depreciation)} and finance costs ${gpct(G.finance_costs)} took PAT margin from ${pct(P25.pat_pct, 2)} to ${pct(P26.pat_pct, 2)}. That is the cost of the store build-out; this case focuses on EBITDA.` },
  ], { x: 0.5, y: 3.75, w: 9.0, h: 0.9, fontSize: 14, color: C.text1, isTextBox: true, margin: 0, objectName: "below ebitda note" });
  source(s6, "Revenue per sq ft uses year-end area (author's measure; DMart reports its own on weighted-average area). Source: DMart filings.");

  // 7 ---------- Levers ----------
  pres.addSection({ title: "Levers" });
  const s7 = content("Levers", `Three levers could add about ${cr(base.total)} in the base case`,
    "Sizes are applied to the FY26 base (pro-forma), not forecast. Each assumption is a yellow cell in the Excel model and can be changed.");
  const leverCards = [
    ["Staff to the ramp", `Win back 10 of the ${Math.round(dEmp)} bps employee-cost rise by staffing new stores to a sales ramp curve and sharing staff across nearby stores.`, base.employee_productivity],
    ["Shift the mix", "Add 10 bps of gross margin by growing general merchandise and apparel, without raising everyday prices.", base.gross_margin_mix],
    ["Ramp new stores faster", `Close half the sales-density gap: ₹${n0(base.new_store_extra_revenue)} cr extra revenue at ${pct(base.flow_through_pct)} flow-through, as only half of operating cost is variable.`, base.new_store_ramp],
  ];
  leverCards.forEach(([head, body, value], i) => {
    const x = 0.5 + i * 3.1;
    card(s7, x, 1.45, 2.9, 3.3, `lever ${i + 1}`);
    badge(s7, x + 0.15, 1.6, String(i + 1), `lever ${i + 1} badge`);
    s7.addText(head, { x: x + 0.7, y: 1.6, w: 2.1, h: 0.42, fontSize: 16, bold: true, color: C.text2, valign: "middle", isTextBox: true, margin: 0, objectName: `lever ${i + 1} title` });
    s7.addText(body, { x: x + 0.15, y: 2.15, w: 2.6, h: 1.75, fontSize: 14, color: C.text1, valign: "top", isTextBox: true, margin: 0, objectName: `lever ${i + 1} body` });
    s7.addText(`+${cr(value)}`, { x: x + 0.15, y: 4.05, w: 2.6, h: 0.5, fontSize: 24, bold: true, color: C.accent1, isTextBox: true, margin: 0, objectName: `lever ${i + 1} value` });
  });

  // 8 ---------- Scenarios ----------
  const s8 = content("Levers", `Even the cautious case adds margin; the stretch case adds ${Math.round(stretch.margin_uplift_bps)} bps`,
    "Cautious assumes only a quarter of the new-store gap closes and no cost or mix gains. Stretch doubles the base assumptions for cost and mix.");
  const hdr = (t) => ({ text: t, options: { bold: true, color: HEX.lt1, fill: { color: HEX.dk2 }, align: "center" } });
  const row = (label, f, bold = false) => [
    { text: label, options: { bold, align: "left" } },
    ...[caut, base, stretch].map((s) => ({ text: f(s), options: { bold, align: "center" } })),
  ];
  s8.addTable([
    [{ text: "FY26 pro-forma", options: { bold: true, color: HEX.lt1, fill: { color: HEX.dk2 }, align: "left" } }, hdr("Cautious"), hdr("Base"), hdr("Stretch")],
    row("Employee cost recovered", (s) => `${s.assumptions.emp_bps} bps`),
    row("Gross margin gain from mix", (s) => `${s.assumptions.gm_bps} bps`),
    row("Share of new-store gap closed", (s) => pct(s.assumptions.density_recovery, 0)),
    row("Total lever impact", (s) => cr(s.total), true),
    row("Pro-forma EBITDA margin", (s) => pct(s.proforma_margin, 2), true),
    row("Uplift vs FY26 actual", (s) => bps(s.margin_uplift_bps), true),
  ], {
    x: 0.5, y: 1.45, w: 9.0, colW: [3.6, 1.8, 1.8, 1.8], rowH: 0.42, fontSize: 14, fontFace: "Calibri", color: HEX.dk1,
    border: { type: "solid", pt: 0.5, color: "CCCCCC" }, fill: { color: HEX.lt1 }, objectName: "scenario table",
  });
  source(s8, `FY26 actual EBITDA margin ${pct(P26.ebitda_pct, 2)}. Variable share of operating cost assumed 50% in every case. Author's assumptions.`);

  // 9 ---------- Risks and next steps ----------
  pres.addSection({ title: "Next steps" });
  const s9 = content("Next steps", "Validate with store-level data before acting on any lever",
    "The levers trade off against DMart's everyday-low-price model, so each needs a test before a roll-out.");
  const col = (x, head, items, name) => {
    card(s9, x, 1.45, 4.35, 3.3, name);
    s9.addText(head, { x: x + 0.2, y: 1.55, w: 3.95, h: 0.45, fontSize: 18, bold: true, color: C.text2, isTextBox: true, margin: 0, objectName: name + " title" });
    s9.addText(items.map((t, i) => ({ text: t, options: { bullet: true, breakLine: i < items.length - 1, paraSpaceAfter: 6 } })),
      { x: x + 0.2, y: 2.05, w: 3.95, h: 2.6, fontSize: 14, color: C.text1, valign: "top", isTextBox: true, margin: 0, objectName: name + " list" });
  };
  col(0.5, "Risks", [
    "Mix changes must not weaken DMart's low-price image",
    "Leaner staffing can hurt service, shrinkage and availability",
    "New-store ramp depends on location; some stores may never reach the average",
    "Lever sizes are assumptions, not tested with DMart data",
  ], "risks");
  col(5.15, "Next steps", [
    "Get store-level sales and staff hours by opening month",
    "Build a staffing-to-sales curve for new stores",
    "Check gross margin by category and region",
    "Pilot the staffing model in one city cluster, then measure",
  ], "next steps");

  // 10 ---------- Sources ----------
  pres.addSection({ title: "Appendix" });
  const s10 = content("Appendix", "Sources and method", "All sources are public. The Excel model's Checks sheet proves the numbers reconcile to the filings.");
  const srcRows = [
    ["FY26 P&L (standalone, audited)", "NSE integrated filing, 2 May 2026"],
    ["FY25 P&L (standalone, audited)", "NSE integrated filing, 3 May 2025 board meeting"],
    ["EBITDA, margins, stores, area FY26", "DMart press release, 2 May 2026"],
    ["Stores FY25 (415)", "DMart company update, 3 Apr 2025"],
    ["Retail area FY25 (17.2 mn sq ft)", "DMart Annual Report FY2024-25"],
    ["Method", "EBITDA = revenue − COGS − employee − other expenses; matches reported ₹5,255 cr"],
  ];
  s10.addTable([
    [{ text: "What", options: { bold: true, color: HEX.lt1, fill: { color: HEX.dk2 } } }, { text: "Source", options: { bold: true, color: HEX.lt1, fill: { color: HEX.dk2 } } }],
    ...srcRows.map(([a, b]) => [{ text: a }, { text: b }]),
  ], { x: 0.5, y: 1.45, w: 9.0, colW: [3.4, 5.6], rowH: 0.42, fontSize: 14, fontFace: "Calibri", color: HEX.dk1, border: { type: "solid", pt: 0.5, color: "CCCCCC" }, objectName: "sources table" });
  source(s10, "Links: github.com/Shreyashsalv2/dmart-ebitda-case (data/SOURCES.md).");

  await pres.writeFile({ fileName: OUT });
  const applyThemePath = process.argv[2];
  if (applyThemePath) {
    const { applyTheme } = require(applyThemePath);
    await applyTheme(OUT, THEME);
  }
  console.log(OUT);
}

main().catch((e) => { console.error(e); process.exit(1); });
