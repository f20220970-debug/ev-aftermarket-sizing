// Builds deck/ev2w_aftermarket_deck.pptx from deck/data.json (written by src/export_data.py).
const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");
let applyTheme = async () => console.log("(theme colours not applied: apply_theme helper not found)");
try { ({ applyTheme } = require(process.env.APPLY_THEME || "./apply_theme.js")); } catch (e) { /* optional */ }

const D = JSON.parse(fs.readFileSync(path.join(__dirname, "data.json"), "utf8"));
const OUT = path.join(__dirname, "ev2w_aftermarket_deck.pptx");
const THEME = {
  name: "EV Teal", headFontFace: "Georgia", bodyFontFace: "Calibri",
  colors: { dk1: "1B1B1B", lt1: "FFFFFF", dk2: "0B3C49", lt2: "EEF5F3", accent1: "028A74", accent2: "F2A541",
            accent3: "5B8C85", accent4: "C44536", accent5: "9DB4C0", accent6: "3D5A80", hlink: "028A74", folHlink: "3D5A80" },
};
const HEX = THEME.colors;
const cr = (v) => "₹" + Math.round(v).toLocaleString("en-IN") + " Cr";
const S = D.summary;
const base31 = S["Total aftermarket FY31 (Rs Cr)"][1], base26 = S["Total aftermarket FY26 (Rs Cr)"][1];

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.title = "India E2W aftermarket sizing"; pres.author = "Pavitra Seth";
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
const C = pres.SchemeColor;

pres.defineSlideMaster({ title: "DARK", background: { color: C.text2 },
  objects: [{ placeholder: { options: { name: "title", type: "title", x: 0.8, y: 1.5, w: 11.7, h: 2.6, fontSize: 36, bold: true, color: C.background1, fontFace: THEME.headFontFace, align: "left", valign: "bottom" }, text: "" } },
            { placeholder: { options: { name: "body", type: "body", x: 0.8, y: 4.5, w: 11.7, h: 1.2, fontSize: 18, color: C.accent2, align: "left", valign: "top" }, text: "" } }] });
pres.defineSlideMaster({ title: "CONTENT", background: { color: C.background1 },
  margin: [0.4, 0.6, 0.6, 0.6],
  objects: [{ placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.35, w: 12.1, h: 0.9, fontSize: 26, bold: true, color: C.text2, fontFace: THEME.headFontFace, valign: "top", align: "left" }, text: "" } },
            { text: { text: "India E2W aftermarket sizing · Pavitra Seth", options: { x: 0.6, y: 7.0, w: 8, h: 0.3, fontSize: 9, color: C.accent3 } } }],
  slideNumber: { x: 12.3, y: 7.0, w: 0.5, h: 0.3, fontSize: 9, color: C.accent3 } });

function content(title, section) {
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: section });
  s.addText(title, { placeholder: "title" });
  return s;
}
function stat(s, x, y, w, big, small, color) {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h: 1.9, fill: { color: C.background2 }, line: { color: C.background2 }, rectRadius: 0.12, objectName: "stat card" });
  s.addText(big, { x: x + 0.2, y: y + 0.2, w: w - 0.4, h: 0.9, fontSize: w < 3.6 ? 26 : 34, bold: true, color: color, fontFace: THEME.headFontFace, margin: 0, isTextBox: true });
  s.addText(small, { x: x + 0.2, y: y + 1.1, w: w - 0.4, h: 0.7, fontSize: 13, color: C.text1, margin: 0, isTextBox: true });
}
const axis = { catAxisLabelColor: HEX.dk2, valAxisLabelColor: HEX.dk2, catAxisLabelFontFace: "+mn-lt", valAxisLabelFontFace: "+mn-lt",
               valGridLine: { color: "DDE6E3", size: 0.5 }, catGridLine: { style: "none" }, titleFontFace: "+mn-lt", dataLabelFontFace: "+mn-lt", legendFontFace: "+mn-lt" };

// 1 Title
pres.addSection({ title: "Answer" });
let s = pres.addSlide({ masterName: "DARK", sectionTitle: "Answer" });
s.addText(`India's EV two-wheeler aftermarket: ${cr(base26)} today, about ${cr(base31)} by FY31`, { placeholder: "title" });
s.addText("Market sizing of servicing and battery replacement, FY22–FY31 · bottom-up model with bear / base / bull scenarios", { placeholder: "body" });

// 2 Executive summary
s = content("Executive summary: a small market today that grows about 6x in five years", "Answer");
stat(s, 0.6, 1.5, 3.9, cr(base26), "Base-case aftermarket in FY26 (service and wear parts only; batteries not yet due)", C.accent1);
stat(s, 4.7, 1.5, 3.9, cr(base31), `FY31 base case; bear ${cr(S["Total aftermarket FY31 (Rs Cr)"][0])}, bull ${cr(S["Total aftermarket FY31 (Rs Cr)"][2])}`, C.text2);
stat(s, 8.8, 1.5, 3.9, Math.round(100 * S["Battery replacement share FY31"][1]) + "%", "of the FY31 base market comes from battery replacements, a wave that starts in FY28", C.accent4);
s.addText([
  { text: "Growth comes from the fleet, not new sales alone: ", options: { bold: true, breakLine: false } },
  { text: `the fleet in use reaches about ${S["Fleet in use FY31 (million)"][1].toFixed(1)} million vehicles by FY31 in the base case, and every vehicle older than a year needs paid service.`, options: { breakLine: true } },
  { text: "The biggest swing factor is the battery wave: ", options: { bold: true } },
  { text: "FY23–FY25 sales (2.8 million vehicles) reach replacement age from FY29, creating a lumpy, high-value repair stream.", options: { breakLine: true } },
  { text: "Where to play: ", options: { bold: true } },
  { text: `independent garages could address about ${cr(S["Independent-garage addressable FY31 (Rs Cr)"][1])} by FY31; battery diagnostics, refurbishment and fleet contracts are the most attractive entry points.` },
], { x: 0.6, y: 3.7, w: 12.1, h: 2.9, fontSize: 15, color: C.text1, paraSpaceAfter: 10, valign: "top", isTextBox: true });

// 3 Context
pres.addSection({ title: "Market" });
s = content("Electric two-wheeler sales have grown about 6x since FY22", "Market");
s.addChart(pres.charts.BAR, [{ name: "E2W sales (lakh units)", labels: ["FY22", "FY23", "FY24", "FY25", "FY26 (est.)"], values: D.sales_hist.map((v) => +(v / 1e5).toFixed(1)) }],
  { x: 0.6, y: 1.4, w: 7.6, h: 5.3, barDir: "col", chartColors: [HEX.accent1], showValue: true, dataLabelPosition: "outEnd", dataLabelColor: HEX.dk2,
    showTitle: true, title: "E2W sales, lakh units", titleColor: HEX.dk2, titleFontSize: 13, showLegend: false, ...axis });
s.addText([
  { text: "What the data says", options: { bold: true, fontSize: 17, color: C.text2, breakLine: true } },
  { text: "2.5 lakh units in FY22 → 11.5 lakh in FY25 (Vahan-based government data).", options: { bullet: true, breakLine: true } },
  { text: "FY26 estimated at about 14.8 lakh: 2.55 million EV sales × ~58% two-wheeler share.", options: { bullet: true, breakLine: true } },
  { text: "Still only 6.3% of all two-wheeler sales in CY2025, so most growth is ahead.", options: { bullet: true, breakLine: true } },
  { text: "Growth slowed to +11% in CY2025: the base case assumes 20% a year, the bear case 10%.", options: { bullet: true } },
], { x: 8.5, y: 1.5, w: 4.2, h: 5.2, fontSize: 14, color: C.text1, paraSpaceAfter: 8, valign: "top", isTextBox: true });

// 4 Approach
s = content("Approach: build the market from the fleet up, then cross-check spend per vehicle", "Market");
const steps = [["1", "Sales by year", "Vahan-based E2W sales FY22–FY26; FY27–FY31 grown by scenario"],
               ["2", "Fleet in use", "Each year's fleet = last year's × (1 − 3% scrappage) + new sales"],
               ["3", "Service market", "Vehicles past their free first year × annual service and wear-parts spend"],
               ["4", "Battery wave", "Sales from 6 years earlier × share replacing × falling pack price"],
               ["5", "Cross-check", "₹2,000–5,000 a year per e-scooter, about half a petrol scooter's cost"]];
steps.forEach(([n, h, d], i) => {
  const x = 0.6 + i * 2.48;
  s.addShape(pres.shapes.OVAL, { x: x + 0.75, y: 1.6, w: 0.8, h: 0.8, fill: { color: i === 4 ? C.accent2 : C.accent1 }, line: { color: i === 4 ? C.accent2 : C.accent1 }, objectName: "step marker" });
  s.addText(n, { x: x + 0.75, y: 1.6, w: 0.8, h: 0.8, align: "center", valign: "middle", fontSize: 22, bold: true, color: C.background1, margin: 0, isTextBox: true });
  if (i < 4) s.addShape(pres.shapes.LINE, { x: x + 1.6, y: 2.0, w: 1.6, h: 0, line: { color: C.accent3, width: 1.5, endArrowType: "triangle" } });
  s.addText(h, { x, y: 2.6, w: 2.3, h: 0.5, align: "center", fontSize: 15, bold: true, color: C.text2, margin: 0, isTextBox: true });
  s.addText(d, { x, y: 3.1, w: 2.3, h: 1.6, align: "center", fontSize: 12, color: C.text1, margin: 0, valign: "top", isTextBox: true });
});
s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 5.0, w: 12.1, h: 1.6, fill: { color: C.background2 }, line: { color: C.background2 }, rectRadius: 0.1, objectName: "note card" });
s.addText("Why bottom-up: no published estimate exists for EV-specific aftermarket spend, and the EV fleet is young, so a top-down share of the petrol two-wheeler aftermarket would overstate it. Every input is in one Excel sheet, colour-coded as sourced or assumed.",
  { x: 0.9, y: 5.15, w: 11.5, h: 1.3, fontSize: 14, color: C.text1, valign: "middle", isTextBox: true });

// 5 Assumptions
s = content("Three scenarios bracket the uncertain inputs", "Market");
const tbl = [["Assumption", "Bear", "Base", "Bull", "Basis"].map((t) => ({ text: t, options: { bold: true, color: HEX.lt1, fill: { color: HEX.dk2 } } })),
  ["Annual E2W sales growth, FY27–31", "10%", "20%", "30%", "CY25 +11%; FY25 +21%"],
  ["Service + wear-parts spend per vehicle", "₹2,000", "₹3,000", "₹4,000", "Reported ₹2,000–5,000 a year"],
  ["Share replacing battery at age 6", "40%", "50%", "60%", "Assumption; warranties 3–5 years"],
  ["Scrappage per year", "3%", "3%", "3%", "Assumption"],
  ["Battery pack price (FY26), yearly fall", "₹55,000, −7%", "₹55,000, −7%", "₹55,000, −7%", "Reported ₹55,000–60,000"],
  ["Independent-garage share of spend", "25%", "25%", "25%", "Assumption"]];
s.addTable(tbl, { x: 0.6, y: 1.5, w: 12.1, colW: [3.9, 1.6, 1.6, 1.6, 3.4], fontSize: 13, color: HEX.dk1, border: { type: "solid", color: "DDE6E3", pt: 0.75 }, fill: { color: HEX.lt1 }, rowH: 0.55 });
s.addText("Biggest uncertainties to test first: spend per vehicle (no published average) and how many owners replace a battery rather than scrap the scooter.",
  { x: 0.6, y: 5.8, w: 12.1, h: 0.8, fontSize: 14, italic: true, color: C.accent3, isTextBox: true });

// 6 Market size
pres.addSection({ title: "Sizing" });
s = content(`The market reaches ${cr(S["Total aftermarket FY31 (Rs Cr)"][0])}–${cr(S["Total aftermarket FY31 (Rs Cr)"][2])} by FY31 across scenarios`, "Sizing");
s.addChart(pres.charts.LINE, ["Bear", "Base", "Bull"].map((n) => ({ name: n, labels: D.years, values: D.total[n].map((v) => Math.round(v)) })),
  { x: 0.6, y: 1.4, w: 8.4, h: 5.3, chartColors: [HEX.accent5, HEX.accent1, HEX.accent2], lineSize: 3, lineDataSymbol: "circle", lineDataSymbolSize: 6,
    showTitle: true, title: "Total aftermarket, ₹ Cr", titleColor: HEX.dk2, titleFontSize: 13, showLegend: true, legendPos: "b", ...axis });
stat(s, 9.3, 1.5, 3.4, Math.round(100 * S["CAGR FY26-FY31"][1]) + "% a year", "Base-case growth FY26–FY31: fleet growth plus the battery wave", C.accent1);
stat(s, 9.3, 3.7, 3.4, S["Fleet in use FY31 (million)"][1].toFixed(1) + " million", "E2Ws in use by FY31 (base), up from about 4.4 million in FY26", C.text2);

// 7 Mix
s = content("From FY28, battery replacements become a second, lumpier revenue stream", "Sizing");
s.addChart(pres.charts.BAR, [{ name: "Service + wear parts", labels: D.years, values: D.base_service.map((v) => Math.round(v)) },
                             { name: "Battery replacement", labels: D.years, values: D.base_battery.map((v) => Math.round(v)) }],
  { x: 0.6, y: 1.4, w: 8.4, h: 5.3, barDir: "col", barGrouping: "stacked", chartColors: [HEX.accent1, HEX.accent2], showValue: false,
    showTitle: true, title: "Base case, ₹ Cr", titleColor: HEX.dk2, titleFontSize: 13, showLegend: true, legendPos: "b", ...axis });
s.addText([
  { text: "Why the jump", options: { bold: true, fontSize: 17, color: C.text2, breakLine: true } },
  { text: "FY22's 2.5 lakh scooters reach age 6 in FY28; FY23's 7.3 lakh arrive in FY29.", options: { bullet: true, breakLine: true } },
  { text: "Each replacement is worth about ₹38,000–48,000 over FY28–FY31, versus ₹3,000 a year for routine service.", options: { bullet: true, breakLine: true } },
  { text: `Batteries are ${Math.round(100 * S["Battery replacement share FY31"][1])}% of the FY31 base market, so battery health checks and refurbishment matter.`, options: { bullet: true } },
], { x: 9.3, y: 1.5, w: 3.4, h: 5.2, fontSize: 14, color: C.text1, paraSpaceAfter: 8, valign: "top", isTextBox: true });

// 8 Sensitivity
s = content("Spend per vehicle moves the answer more than sales growth does", "Sizing");
const vals = D.sens.flat(); const lo = Math.min(...vals), hi = Math.max(...vals);
const shade = (v) => { const t = (v - lo) / (hi - lo); const mix = (a, b) => Math.round(a + (b - a) * t);
  const c = [mix(0xEE, 0x02), mix(0xF5, 0x8A), mix(0xF3, 0x74)]; return c.map((x) => x.toString(16).padStart(2, "0")).join("").toUpperCase(); };
const sens = [[{ text: "Spend ↓ / growth →", options: { bold: true, color: HEX.lt1, fill: { color: HEX.dk2 } } },
               ...D.sens_growth.map((g) => ({ text: Math.round(100 * g) + "%", options: { bold: true, color: HEX.lt1, fill: { color: HEX.dk2 }, align: "center" } }))]];
D.sens.forEach((r, i) => sens.push([{ text: "₹" + D.sens_spend[i].toLocaleString("en-IN"), options: { bold: true } },
  ...r.map((v) => { const t = (v - lo) / (hi - lo); return { text: Math.round(v).toLocaleString("en-IN"), options: { align: "center", fill: { color: shade(v) }, color: t > 0.55 ? HEX.lt1 : HEX.dk1, bold: D.sens_spend[i] === 3000 } }; })]));
s.addTable(sens, { x: 0.6, y: 1.5, w: 7.8, colW: [2.3, 1.1, 1.1, 1.1, 1.1, 1.1], fontSize: 14, rowH: 0.62, border: { type: "solid", color: "FFFFFF", pt: 1 } });
s.addText("FY31 total aftermarket, ₹ Cr; other inputs at base case. Battery replacements in FY31 come from FY25 sales, so they don't change with future growth.", { x: 0.6, y: 5.5, w: 7.8, h: 0.8, fontSize: 12, italic: true, color: C.accent3, isTextBox: true });
s.addText([
  { text: "Reading the grid", options: { bold: true, fontSize: 17, color: C.text2, breakLine: true } },
  { text: "Moving spend from ₹2,000 to ₹4,000 adds about ₹2,200–3,000 Cr.", options: { bullet: true, breakLine: true } },
  { text: "Moving growth from 10% to 30% adds about ₹800–1,600 Cr, because FY31's fleet is mostly vehicles already sold.", options: { bullet: true, breakLine: true } },
  { text: "So the first piece of primary research should be what owners actually spend.", options: { bullet: true } },
], { x: 8.8, y: 1.5, w: 3.9, h: 5.0, fontSize: 14, color: C.text1, paraSpaceAfter: 8, valign: "top", isTextBox: true });

// 9 Where to play
pres.addSection({ title: "Implications" });
s = content("Where to play: independent service networks have a real, growing slice", "Implications");
stat(s, 0.6, 1.5, 3.9, cr(S["Independent-garage addressable FY31 (Rs Cr)"][1]), "FY31 spend addressable by independent garages (base, 25% share)", C.accent1);
stat(s, 4.7, 1.5, 3.9, cr(S["Karnataka market FY31 (Rs Cr)"][1]), "FY31 Karnataka market (base; ~9% of national EV sales)", C.text2);
stat(s, 8.8, 1.5, 3.9, "~3,000", "Rs per vehicle per year: price point for an annual maintenance plan", C.accent2);
const plays = [["Battery diagnostics and refurbishment", "Highest value per job; owners need advice on repair vs replace as warranties expire."],
               ["Fleet contracts (delivery and rental)", "High-mileage scooters service more often; one contract covers many vehicles."],
               ["Multi-brand doorstep service", "OEM networks are likely thinner outside big cities (to verify); an on-site model avoids rent and serves several brands."]];
plays.forEach(([h, d], i) => {
  const y = 3.75 + i * 0.95;
  s.addShape(pres.shapes.OVAL, { x: 0.6, y: y + 0.1, w: 0.5, h: 0.5, fill: { color: C.accent1 }, line: { color: C.accent1 }, objectName: "play marker" });
  s.addText(String(i + 1), { x: 0.6, y: y + 0.1, w: 0.5, h: 0.5, align: "center", valign: "middle", fontSize: 14, bold: true, color: C.background1, margin: 0, isTextBox: true });
  s.addText([{ text: h + ": ", options: { bold: true, color: C.text2 } }, { text: d }], { x: 1.3, y, w: 11.4, h: 0.75, fontSize: 14, color: C.text1, valign: "middle", isTextBox: true });
});

// 10 Risks
s = content("What would change the answer", "Implications");
const risks = [["Batteries last longer", "If most packs outlive the scooter, the battery wave shrinks: the bear case already assumes only 40% replace."],
               ["OEMs lock in service", "Software diagnostics and parts control can keep work inside OEM networks, cutting the independent share below 25%."],
               ["Sales growth stalls", "Growth slowed to 11% in CY2025. At 10% a year (other inputs at base) FY31 is about 10% smaller; the full bear case is about a third smaller."],
               ["Data gaps", "FY26 sales are an estimate from two sources, and there is no published average spend per EV."]];
risks.forEach(([h, d], i) => {
  const x = 0.6 + (i % 2) * 6.15, y = 1.5 + Math.floor(i / 2) * 2.6;
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: 5.95, h: 2.35, fill: { color: C.background2 }, line: { color: C.background2 }, rectRadius: 0.1, objectName: "risk card" });
  s.addText(h, { x: x + 0.3, y: y + 0.2, w: 5.4, h: 0.5, fontSize: 17, bold: true, color: C.accent4, margin: 0, isTextBox: true });
  s.addText(d, { x: x + 0.3, y: y + 0.8, w: 5.4, h: 1.4, fontSize: 14, color: C.text1, margin: 0, valign: "top", isTextBox: true });
});

// 11 Next steps
s = content("Next steps: test the two inputs that matter most", "Implications");
const next = [["Owner survey (n ≈ 100)", "What do EV owners spend a year on service and parts, and where?"],
              ["10 garage and OEM-service interviews", "Job mix, prices, parts access, share of EV work today."],
              ["Battery health data", "Warranty claims and pack degradation by age from 1–2 OEMs or fleet operators."],
              ["Update the model", "Replace assumed inputs with survey medians; re-run the sensitivity grid."]];
next.forEach(([h, d], i) => {
  const y = 1.5 + i * 1.2;
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y, w: 12.1, h: 1.0, fill: { color: i % 2 ? C.background1 : C.background2 }, line: { color: C.background2 }, rectRadius: 0.08, objectName: "step card" });
  s.addText(String(i + 1), { x: 0.8, y: y + 0.15, w: 0.7, h: 0.7, fontSize: 26, bold: true, color: C.accent1, fontFace: THEME.headFontFace, margin: 0, valign: "middle", isTextBox: true });
  s.addText([{ text: h, options: { bold: true, color: C.text2, breakLine: true } }, { text: d }], { x: 1.6, y: y + 0.08, w: 10.9, h: 0.85, fontSize: 14, color: C.text1, valign: "middle", isTextBox: true });
});

// 12 Sources
pres.addSection({ title: "Appendix" });
s = content("Sources", "Appendix");
s.addText([
  "E2W sales FY22: Argus Media, citing Vahan (FY2021-22: 252,555 units).",
  "E2W sales FY23: PIB annexure, state and month-wise E-2W registrations (Vahan 4): 7,28,271.",
  "E2W sales FY24 and FY25: Ministry of Heavy Industries data via Times Drive (9,48,561 and 11,49,334).",
  "FY26 total EV sales: IESA India's Mobility Report FY2025-26 (2.55 million); E2W share ~58%: JMK Research.",
  "CY2025 E2W sales and penetration: EVreporter, Vahan dashboard (12,79,951; 6.3% of two-wheelers).",
  "Maintenance spend: MeraEV (2026), ₹2,000–5,000 a year for e-scooters vs ₹3,500–6,000 for petrol scooters.",
  "Battery replacement cost: e-vehicleinfo and BikeLeague (about ₹55,000–60,000).",
  "Karnataka share: IESA via IANS (about 9% of national EV sales in CY2025, all vehicle types).",
  "All other inputs are labelled assumptions in model/ev2w_service_market.xlsx (Inputs sheet).",
].map((t, i, a) => ({ text: t, options: { bullet: true, breakLine: i < a.length - 1 } })),
  { x: 0.6, y: 1.4, w: 12.1, h: 5.4, fontSize: 13, color: C.text1, paraSpaceAfter: 6, valign: "top", isTextBox: true });

pres.writeFile({ fileName: OUT }).then(async () => { await applyTheme(OUT, THEME); console.log("saved", OUT); });
