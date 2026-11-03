// Builds the Phase 1 PowerPoint deck for COMP 262 Team 2
const path = require("path");
const pptxgen = require("pptxgenjs");
const { applyTheme } = require("C:/Users/rishi/AppData/Local/Claude-Data/claude_pzs8sxrjxfjjc/local-agent-mode-sessions/skills-plugin/eda71d28-5986-460a-97fd-0320a6afe8c6/2f6516b7-f548-454a-816c-78c0fc049975/skills/pptx/scripts/apply_theme.js");

const PROJ = "C:/Users/rishi/OneDrive/Desktop/ARIES/COMP-262/project_phase1";
const FIG = (f) => path.join(PROJ, "figures", f);
const OUT = path.join(PROJ, "Team2_Phase1_Presentation.pptx");

// sentiment-coded palette that matches the charts: rose = negative, amber = neutral, teal = positive
const HEX = { ink: "15303A", teal: "00798C", amber: "EDAE49", rose: "D1495B", mist: "EEF5F6", grey: "5C6B70", line: "D5DEE0" };
const THEME = {
  name: "Team 2 Sentiment",
  headFontFace: "Cambria",
  bodyFontFace: "Calibri",
  colors: { dk1: "1A1A1A", lt1: "FFFFFF", dk2: HEX.ink, lt2: HEX.mist, accent1: HEX.teal, accent2: HEX.amber, accent3: HEX.rose,
            accent4: "7FC8D0", accent5: "F4D58D", accent6: HEX.grey, hlink: HEX.teal, folHlink: HEX.grey },
};

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9"; // 10 x 5.625 in
pres.title = "COMP 262 Project Phase 1 - Team 2";
pres.author = "COMP 262 Team 2";
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
const C = pres.SchemeColor;

// ---------- layouts ----------
pres.defineSlideMaster({
  title: "TITLE_DARK",
  background: { color: C.text2 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 1.35, w: 8.8, h: 1.5, fontFace: THEME.headFontFace, fontSize: 40, bold: true, color: C.background1, valign: "bottom", margin: 0 }, text: "" } },
    { placeholder: { options: { name: "body", type: "body", x: 0.6, y: 3.0, w: 8.8, h: 1.9, fontSize: 18, color: "CFE3E6", valign: "top", margin: 0 }, text: "" } },
  ],
});
pres.defineSlideMaster({
  title: "CONTENT",
  background: { color: C.background1 },
  margin: [0.5, 0.5, 0.55, 0.5],
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.5, y: 0.28, w: 9.0, h: 0.75, fontFace: THEME.headFontFace, fontSize: 30, bold: true, color: C.text2, valign: "middle", margin: 0 }, text: "" } },
    { text: { text: "COMP 262  |  Team 2  |  Phase 1", options: { x: 0.5, y: 5.22, w: 5, h: 0.25, fontSize: 10, color: C.accent6, margin: 0 } } },
  ],
  slideNumber: { x: 9.0, y: 5.22, w: 0.5, h: 0.25, fontSize: 10, color: C.accent6, align: "right", margin: 0 },
});

// ---------- helpers ----------
const shadow = () => ({ type: "outer", color: "000000", opacity: 0.12, blur: 6, offset: 2, angle: 90 });
function card(slide, x, y, w, h, name, fill = C.background2) {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.08, fill: { color: fill }, line: { color: fill }, objectName: name });
}
function stat(slide, x, y, w, value, label, color = C.accent1, name = "stat", size = 30) {
  card(slide, x, y, w, 1.25, `${name} card`);
  slide.addText(value, { x: x + 0.15, y: y + 0.12, w: w - 0.3, h: 0.65, fontFace: THEME.headFontFace, fontSize: size, bold: true, color, margin: 0, valign: "middle", isTextBox: true, objectName: `${name} value` });
  slide.addText(label, { x: x + 0.15, y: y + 0.76, w: w - 0.3, h: 0.42, fontSize: 12, color: C.text2, margin: 0, valign: "top", isTextBox: true, objectName: `${name} label` });
}
function bullets(slide, items, x, y, w, h, opts = {}) {
  slide.addText(items.map((t, i) => {
    const parts = Array.isArray(t) ? t : [t];
    return parts.map((p, j) => ({ text: typeof p === "string" ? p : p.text,
      options: { bold: typeof p !== "string" && p.bold, color: typeof p !== "string" && p.color ? p.color : undefined,
                 bullet: j === 0 ? true : undefined, breakLine: j === parts.length - 1 && i < items.length - 1 ? true : undefined,
                 paraSpaceAfter: 8 } }));
  }).flat(), { x, y, w, h, fontSize: opts.fontSize || 15, color: C.text1, valign: "top", margin: 0, isTextBox: true, objectName: opts.name || "bullets" });
}
const chartBase = (title) => ({
  showTitle: !!title, title, titleFontFace: "+mn-lt", titleFontSize: 13, titleColor: HEX.ink,
  catAxisLabelFontFace: "+mn-lt", valAxisLabelFontFace: "+mn-lt", dataLabelFontFace: "+mn-lt", legendFontFace: "+mn-lt",
  catAxisLabelColor: HEX.grey, valAxisLabelColor: HEX.grey, catAxisLabelFontSize: 11, valAxisLabelFontSize: 10,
  dataLabelFontSize: 10, dataLabelColor: HEX.ink, legendFontSize: 10,
  valGridLine: { color: HEX.line, size: 0.5 }, catGridLine: { style: "none" },
});
function content(title, section, notes) {
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: section });
  s.addText(title, { placeholder: "title" });
  if (notes) s.addNotes(notes);
  return s;
}

// ---------- slides ----------
pres.addSection({ title: "Introduction" });

// 1. Title
{
  const s = pres.addSlide({ masterName: "TITLE_DARK", sectionTitle: "Introduction" });
  s.addText("What 1,379 Beauty Reviews Say", { placeholder: "title" });
  s.addText([
    { text: "Lexicon sentiment analysis of Amazon All Beauty reviews: VADER vs TextBlob", options: { breakLine: true } },
    { text: " ", options: { breakLine: true, fontSize: 10 } },
    { text: "Team 2  |  COMP 262 Project Phase 1  |  November 5, 2026", options: { fontSize: 14, color: HEX.amber, breakLine: true } },
    { text: "Rishi Goyal  |  [Member 2]  |  [Member 3]  |  [Member 4]", options: { fontSize: 14 } },
  ], { placeholder: "body" });
  // three sentiment dots: the deck's motif
  [[HEX.rose, 0.6], [HEX.amber, 0.95], [HEX.teal, 1.3]].forEach(([col, x], i) =>
    s.addShape(pres.shapes.OVAL, { x, y: 0.75, w: 0.25, h: 0.25, fill: { color: col }, line: { color: col }, objectName: `motif dot ${i + 1}` }));
  s.addNotes("Presenter: [Member 1]. Introduce the team and the task: build a sentiment analyser for product reviews with two lexicon models. Our dataset is Amazon All Beauty (5-core) from Ni, Li and McAuley, 2019.");
}

// 2. Roadmap
{
  const s = content("Our Phase 1 pipeline in seven steps", "Introduction",
    "Presenter: [Member 1]. Walk through the seven steps; each presenter owns a block. Mention the single script runs everything in under a minute with seed 42.");
  const steps = [
    ["1", "Explore", "counts, distributions, duplicates", C.accent1],
    ["2", "Label", "4-5 Pos, 3 Neu, 1-2 Neg", C.accent1],
    ["3", "Lexicons", "VADER, TextBlob, SentiWordNet", C.accent2],
    ["4", "Pre-process", "per model, tested", C.accent2],
    ["5", "Sample", "1,000 random reviews", C.accent3],
    ["6", "Model", "VADER + TextBlob", C.accent3],
    ["7", "Validate", "metrics + comparison", C.accent3],
  ];
  const w = 1.18, gap = 0.12, x0 = 0.5;
  steps.forEach(([n, head, sub, col], i) => {
    const x = x0 + i * (w + gap);
    s.addShape(pres.shapes.OVAL, { x: x + w / 2 - 0.32, y: 1.45, w: 0.64, h: 0.64, fill: { color: col }, line: { color: col }, objectName: `step ${n} circle` });
    s.addText(n, { x: x + w / 2 - 0.32, y: 1.45, w: 0.64, h: 0.64, align: "center", valign: "middle", fontSize: 20, bold: true, color: C.background1, margin: 0, isTextBox: true, objectName: `step ${n} number` });
    s.addText(head, { x, y: 2.25, w, h: 0.4, align: "center", fontSize: 14, bold: true, color: C.text2, margin: 0, isTextBox: true, objectName: `step ${n} title` });
    s.addText(sub, { x, y: 2.65, w, h: 0.7, align: "center", fontSize: 11, color: C.accent6, margin: 0, valign: "top", isTextBox: true, objectName: `step ${n} detail` });
  });
  card(s, 0.5, 3.65, 9.0, 1.3, "presenters card");
  s.addText([
    { text: "Who presents what", options: { bold: true, color: HEX.ink, breakLine: true } },
    { text: "[Member 1]: intro and dataset  |  [Member 2]: exploration and duplicates  |  [Member 3]: pre-processing and lexicon choice  |  [Member 4]: models, results and live demo", options: { color: HEX.grey } },
  ], { x: 0.75, y: 3.78, w: 8.5, h: 1.05, fontSize: 13, valign: "top", margin: 0, isTextBox: true, objectName: "presenters text" });
}

pres.addSection({ title: "Data exploration" });

// 3. Dataset at a glance
{
  const s = content("The dataset at a glance", "Data exploration",
    "Presenter: [Member 2]. 5,269 rows, 85 products, 991 reviewers, average rating 4.77. 88% of rows are 5-star, so the data is heavily skewed positive. That skew drives our choice of metric later.");
  stat(s, 0.5, 1.25, 2.05, "5,269", "review rows, 12 columns", C.accent1, "rows");
  stat(s, 0.5, 2.65, 2.05, "85", "products (asin)", C.accent1, "products");
  stat(s, 2.75, 1.25, 2.05, "991", "reviewers", C.accent1, "reviewers");
  stat(s, 2.75, 2.65, 2.05, "4.77", "average star rating", C.accent2, "avg rating");
  s.addText("Also: 81% verified purchases, 10.5% have a helpful vote, 1.9% have images. Text fields are 99.9% complete.", { x: 0.5, y: 4.1, w: 4.3, h: 0.8, fontSize: 13, color: C.accent6, margin: 0, valign: "top", isTextBox: true, objectName: "extra facts" });
  s.addChart(pres.charts.BAR, [{ name: "Reviews", labels: ["1 star", "2 stars", "3 stars", "4 stars", "5 stars"], values: [115, 64, 109, 332, 4649] }], {
    x: 5.1, y: 1.2, w: 4.4, h: 3.75, barDir: "col", chartColors: [HEX.teal], showValue: true, dataLabelPosition: "outEnd",
    showLegend: false, valAxisHidden: true, valGridLine: { style: "none" }, ...chartBase("88% of reviews are 5 stars"), valGridLine: { style: "none" }, objectName: "rating chart",
  });
}

// 4. Duplicates
{
  const s = content("74% of the rows are duplicates", "Data exploration",
    "Presenter: [Member 2]. Amazon lists one review under every size or colour variant. Same reviewer, same text, same timestamp, different asin. 3,890 of 5,269 rows are copies; ratings never conflict, so we keep one copy. Only 1,379 unique reviews remain, and all modelling uses them. Without this, the same review could land in both the dev and test data.");
  card(s, 0.5, 1.25, 3.3, 3.7, "duplicate stat card", C.text2);
  s.addText("3,890", { x: 0.75, y: 1.45, w: 2.8, h: 0.9, fontFace: THEME.headFontFace, fontSize: 48, bold: true, color: C.accent2, margin: 0, isTextBox: true, objectName: "dup value" });
  s.addText("of 5,269 rows are the same review copied under another product variant", { x: 0.75, y: 2.35, w: 2.8, h: 0.9, fontSize: 14, color: C.background1, margin: 0, valign: "top", isTextBox: true, objectName: "dup label" });
  s.addText([
    { text: "1,379", options: { bold: true, fontSize: 28, color: HEX.amber, breakLine: true, fontFace: THEME.headFontFace } },
    { text: "unique reviews are left. Copies never disagree on the rating, so we safely keep one.", options: { fontSize: 13, color: "CFE3E6" } },
  ], { x: 0.75, y: 3.3, w: 2.8, h: 1.5, margin: 0, valign: "top", isTextBox: true, objectName: "unique text" });
  s.addImage({ path: FIG("02_duplicate_copies.png"), x: 4.1, y: 1.25, w: 5.4, h: 5.4 * 576 / 1027, objectName: "duplicate chart", altText: "Bar chart: number of product variants each unique review is copied to" });
  s.addText("Key = same reviewer + same text + same timestamp", { x: 4.1, y: 4.45, w: 5.4, h: 0.4, fontSize: 12, italic: true, color: C.accent6, margin: 0, isTextBox: true, objectName: "dup key note" });
}

// 5. Long tail
{
  const s = content("A few products and one-time reviewers", "Data exploration",
    "Presenter: [Member 2]. After de-duplication: 63 products, median 4 reviews each, max 480. The top 2 products hold 62.5% of reviews; 7 products hold 80%. Per user: the 5-core promise of 5+ reviews mostly comes from copies; 86% of reviewers actually wrote only one review.");
  s.addImage({ path: FIG("03_reviews_per_product.png"), x: 0.5, y: 1.2, w: 6.1, h: 6.1 * 604 / 1777, objectName: "products chart", altText: "Reviews per product histogram and cumulative share curve" });
  stat(s, 6.9, 1.2, 2.6, "62.5%", "of unique reviews belong to just 2 products", C.accent1, "top2");
  stat(s, 6.9, 2.6, 2.6, "86%", "of reviewers wrote only one unique review", C.accent3, "one-time");
  bullets(s, [
    [{ text: "Products: ", bold: true }, "median 4 reviews, max 480 (raw: 1,794)"],
    [{ text: "Users: ", bold: true }, "5-core in raw rows (4-12 each), but mean 1.39 unique reviews"],
    [{ text: "So what: ", bold: true }, "product averages from a few reviews are unreliable; matters for the Phase 2 recommender"],
  ], 0.5, 3.45, 6.1, 1.6, { fontSize: 13, name: "long tail bullets" });
}

// 6. Lengths
{
  const s = content("Mixed opinions take more words", "Data exploration",
    "Presenter: [Member 2]. Median review is 29 words; IQR fence is 161 words, giving 106 long outliers (7.7%). 2- and 3-star reviews are the longest (median about 115 words). Happy and angry customers are brief. This predicts that Neutral will be hardest for lexicons: long, mixed text.");
  s.addChart(pres.charts.BAR, [{ name: "Median words", labels: ["1 star", "2 stars", "3 stars", "4 stars", "5 stars"], values: [28, 119, 114, 75, 23] }], {
    x: 0.5, y: 1.2, w: 5.2, h: 3.8, barDir: "col", chartColors: [HEX.amber], showValue: true, dataLabelPosition: "outEnd", showLegend: false,
    valAxisHidden: true, ...chartBase("Median review length (words) by rating"), valGridLine: { style: "none" }, objectName: "length chart",
  });
  stat(s, 6.0, 1.2, 3.5, "29 words", "median review length (mean 56.8, max 962)", C.accent1, "median length");
  stat(s, 6.0, 2.6, 3.5, "106", "long outliers above the 161-word IQR fence: kept, and checked in testing", C.accent3, "outliers");
  s.addText("124 reviews have 3 words or fewer (\"Love it!\"): kept. 1 empty review: removed.", { x: 6.0, y: 4.0, w: 3.5, h: 0.9, fontSize: 13, color: C.accent6, margin: 0, valign: "top", isTextBox: true, objectName: "short note" });
}

// 7. Vocabulary + relationships
{
  const s = content("Complaints rarely use negative words", "Data exploration",
    "Presenter: [Member 2]. Top-15 negative words contain no negative word: like, really, hair, smell, water. Complaints are descriptive ('felt like water'). 'like' is top-3 in every class and VADER scores it positive (+0.36 for 'It smells like chemicals'). Other relationships: unverified reviews average 4.39 vs 4.78 verified; 1-star reviews are 7x more likely than 4-star to get a helpful vote.");
  s.addImage({ path: FIG("08_top_words_by_class.png"), x: 0.5, y: 1.2, w: 9.0, h: 9.0 * 696 / 2377, objectName: "top words chart", altText: "Top 15 words per sentiment class" });
  const rows = [
    ["\"like\" is top-3 in every class; VADER scores \"It smells like chemicals.\" +0.36", HEX.rose],
    ["Unverified buyers rate lower: 4.39 vs 4.78 average", HEX.amber],
    ["1-star reviews get helpful votes most often: 33% vs 4.6% for 4-star", HEX.teal],
  ];
  rows.forEach(([t, col], i) => {
    const y = 4.0 + i * 0.36;
    s.addShape(pres.shapes.OVAL, { x: 0.55, y: y + 0.08, w: 0.16, h: 0.16, fill: { color: col }, line: { color: col }, objectName: `insight dot ${i + 1}` });
    s.addText(t, { x: 0.85, y, w: 8.6, h: 0.32, fontSize: 13, color: C.text1, margin: 0, valign: "middle", isTextBox: true, objectName: `insight ${i + 1}` });
  });
}

pres.addSection({ title: "Pre-processing" });

// 8. Labels and columns
{
  const s = content("Labels, columns and the leak we removed", "Pre-processing",
    "Presenter: [Member 3]. Labels from rating: 91.3% Positive, 4.6% Neutral, 4.1% Negative. Input text = summary + review body. We remove Amazon's auto-generated 'Five Stars' titles (22% of reviews) because they restate the rating: label leakage. Rating is used only as ground truth; reviewerID + text + time is the de-dup key; metadata (verified, vote, image, style) is exploration only.");
  [["1,259", "Positive (4-5 stars)", "91.3%", C.accent1], ["64", "Neutral (3 stars)", "4.6%", C.accent2], ["56", "Negative (1-2 stars)", "4.1%", C.accent3]].forEach(([v, l, p, col], i) => {
    const y = 1.2 + i * 1.25;
    card(s, 0.5, y, 3.2, 1.08, `label card ${i + 1}`);
    s.addText(v, { x: 0.7, y: y + 0.12, w: 1.4, h: 0.85, fontFace: THEME.headFontFace, fontSize: 30, bold: true, color: col, margin: 0, valign: "middle", isTextBox: true, objectName: `label value ${i + 1}` });
    s.addText([{ text: l, options: { bold: true, breakLine: true } }, { text: `${p} of unique reviews`, options: { color: HEX.grey } }], { x: 2.05, y: y + 0.12, w: 1.6, h: 0.85, fontSize: 12, color: C.text2, margin: 0, valign: "middle", isTextBox: true, objectName: `label text ${i + 1}` });
  });
  const hdr = (t) => ({ text: t, options: { bold: true, color: "FFFFFF", fill: { color: HEX.teal } } });
  s.addTable([
    [hdr("Column"), hdr("Role"), hdr("Why")],
    ["reviewText", "Model input", "Main opinion text"],
    ["summary", "Model input", "Short, opinion-dense headline"],
    [{ text: "\"Five Stars\" titles", options: { color: HEX.rose, bold: true } }, { text: "Removed", options: { color: HEX.rose, bold: true } }, "Auto-generated from the rating: leaks the label"],
    ["overall", "Label only", "Ground truth, never shown to models"],
    ["reviewerID + time", "De-dup key", "Identifies one written review"],
    ["verified, vote, style", "Exploration", "No opinion words"],
  ], { x: 4.0, y: 1.2, w: 5.5, colW: [1.6, 1.15, 2.75], fontSize: 12, color: HEX.ink, border: { type: "solid", pt: 0.5, color: HEX.line }, rowH: 0.42, valign: "middle", objectName: "columns table" });
}

// 9. Lexicon study
{
  const s = content("Why VADER and TextBlob, not SentiWordNet", "Pre-processing",
    "Presenter: [Member 3]. We scored probe sentences with all three lexicons. SentiWordNet keeps 'not good' positive, ignores CAPS and !!!, and calls 'Not bad at all' strongly negative: no negation rules and it needs word-sense disambiguation. VADER handles negation, emphasis, emoticons and 'but'. TextBlob is adjective-based with negation and gives subjectivity, and it fails differently from VADER, which makes the comparison informative.");
  const hdr = (t) => ({ text: t, options: { bold: true, color: "FFFFFF", fill: { color: HEX.ink }, align: "center" } });
  const cell = (v) => {
    const n = Number(v);
    return { text: (n > 0 ? "+" : "") + n.toFixed(2), options: { align: "center", bold: true, color: n >= 0.05 ? HEX.teal : n <= -0.05 ? HEX.rose : HEX.grey } };
  };
  const probes = [
    ["This moisturizer is not good.", -0.341, -0.35, 0.083],
    ["This moisturizer is GOOD!!!", 0.671, 1.0, 0.438],
    ["The scent is nice, but it broke me out terribly.", -0.827, -0.2, 0.271],
    ["Love it :)", 0.802, 0.5, 0.5],
    ["Not bad at all.", 0.431, 0.35, -0.625],
  ];
  s.addTable([[hdr("Probe sentence"), hdr("VADER"), hdr("TextBlob"), hdr("SentiWordNet")],
    ...probes.map(([t, a, b, c]) => [t, cell(a), cell(b), cell(c)])],
    { x: 0.5, y: 1.2, w: 5.9, colW: [3.1, 0.9, 0.95, 0.95], fontSize: 12, color: HEX.ink, border: { type: "solid", pt: 0.5, color: HEX.line }, rowH: 0.5, valign: "middle", objectName: "probe table" });
  s.addText("Scores range from -1 to +1. Teal = positive, rose = negative.", { x: 0.5, y: 4.4, w: 5.9, h: 0.3, fontSize: 11, italic: true, color: C.accent6, margin: 0, isTextBox: true, objectName: "probe legend" });
  const picks = [
    ["VADER", "Built for short, informal opinions: negation, CAPS, !!!, emoticons, 'but'", C.accent1],
    ["TextBlob", "Adjective lexicon + subjectivity; makes different mistakes", C.accent1],
    ["SentiWordNet", "No negation or emphasis; needs sense disambiguation", C.accent3],
  ];
  picks.forEach(([h, t, col], i) => {
    const y = 1.2 + i * 1.12;
    card(s, 6.7, y, 2.8, 0.98, `pick card ${i + 1}`);
    s.addText([{ text: (i < 2 ? "Chosen: " : "Rejected: ") + h, options: { bold: true, color: col === C.accent3 ? HEX.rose : HEX.teal, breakLine: true } }, { text: t, options: { color: HEX.ink, fontSize: 11 } }],
      { x: 6.85, y: y + 0.08, w: 2.5, h: 0.85, fontSize: 13, margin: 0, valign: "top", isTextBox: true, objectName: `pick text ${i + 1}` });
  });
}

// 10. Pre-processing per model
{
  const s = content("Each lexicon needs its own pre-processing", "Pre-processing",
    "Presenter: [Member 3]. NLTK's stop-word list contains 'not', 'no', 'very', 'but', so the classic pipeline destroys lexicon signals. VADER: keep case, punctuation and stop words; only unescape HTML, strip URLs, squeeze 'sooooo'. TextBlob: lower-case, expand contractions, drop stop words but keep negations and intensifiers, lemmatise. Ablation on the 378-review development set (never the test set): VADER best with light cleaning; TextBlob best with P3 (0.506 vs 0.460). Removing negations hurts both.");
  const col = (x, title, color, items) => {
    card(s, x, 1.2, 2.85, 2.55, `${title} card`);
    s.addText(title, { x: x + 0.18, y: 1.3, w: 2.5, h: 0.4, fontSize: 16, bold: true, color, margin: 0, isTextBox: true, objectName: `${title} heading` });
    bullets(s, items, x + 0.18, 1.75, 2.55, 1.95, { fontSize: 12, name: `${title} steps` });
  };
  col(0.5, "VADER: light", C.accent1, ["Keep CAPS, !!!, emoticons", "Keep all stop words", "Unescape HTML, drop URLs", "\"sooooo\" -> \"soo\""]);
  col(3.5, "TextBlob: targeted", C.accent2, ["Lower-case, expand \"isn't\"", "Drop stop words, KEEP not / very / but", "Lemmatise, drop punctuation"]);
  s.addChart(pres.charts.BAR, [
    { name: "VADER", labels: ["Raw", "Light", "Lower, no punct", "Drop stop words, keep negations", "Drop ALL stop words"], values: [0.455, 0.455, 0.456, 0.437, 0.416] },
    { name: "TextBlob", labels: ["Raw", "Light", "Lower, no punct", "Drop stop words, keep negations", "Drop ALL stop words"], values: [0.46, 0.46, 0.464, 0.506, 0.492] },
  ], { x: 6.5, y: 1.15, w: 3.0, h: 3.9, barDir: "bar", chartColors: [HEX.teal, HEX.amber], showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "0.000",
       valAxisHidden: true, valAxisMinVal: 0.3, valAxisMaxVal: 0.56, showLegend: true, legendPos: "b", catAxisOrientation: "maxMin",
       ...chartBase("Dev-set macro-F1 by pipeline"), catAxisLabelFontSize: 9, dataLabelFontSize: 9, valGridLine: { style: "none" }, objectName: "ablation chart" });
  card(s, 0.5, 3.95, 5.85, 1.05, "ablation takeaway card", C.text2);
  s.addText([{ text: "Tested, not assumed: ", options: { bold: true, color: HEX.amber } },
    { text: "dropping \"not\" lowers macro-F1 for both models; TextBlob gains +0.046 from the targeted pipeline.", options: { color: "FFFFFF" } }],
    { x: 0.7, y: 4.03, w: 5.5, h: 0.9, fontSize: 13, margin: 0, valign: "middle", isTextBox: true, objectName: "ablation takeaway" });
}

pres.addSection({ title: "Modelling and results" });

// 11. Sampling and models
{
  const s = content("1,000-review test sample, two scorers", "Modelling and results",
    "Presenter: [Member 4]. 1,000 random unique reviews (seed 42) are the test sample: 911 Positive, 51 Neutral, 38 Negative. The other 378 reviews are a development set for pre-processing and threshold choices. VADER: compound score, >= 0.05 Positive, <= -0.05 Negative. TextBlob: averaged polarity, > 0 Positive, < 0 Negative. Tuning: grid-search thresholds on the dev set to maximise macro-F1.");
  s.addChart(pres.charts.DOUGHNUT, [{ name: "Test sample", labels: ["Positive", "Neutral", "Negative"], values: [911, 51, 38] }], {
    x: 0.4, y: 1.15, w: 3.2, h: 3.0, chartColors: [HEX.teal, HEX.amber, HEX.rose], holeSize: 60, showPercent: false, showValue: true, showLegend: true, legendPos: "b",
    dataLabelColor: "FFFFFF", ...chartBase("Test sample: 1,000 reviews"), dataLabelColor: "FFFFFF", objectName: "sample chart",
  });
  s.addText("+ 378-review development set for every modelling choice", { x: 0.5, y: 4.25, w: 3.0, h: 0.7, fontSize: 12, italic: true, color: C.accent6, margin: 0, valign: "top", isTextBox: true, objectName: "dev note" });
  const box = (x, title, color, lines) => {
    card(s, x, 1.2, 2.8, 3.75, `${title} model card`);
    s.addShape(pres.shapes.OVAL, { x: x + 0.2, y: 1.38, w: 0.22, h: 0.22, fill: { color }, line: { color }, objectName: `${title} dot` });
    s.addText(title, { x: x + 0.55, y: 1.3, w: 2.1, h: 0.4, fontSize: 18, bold: true, color: C.text2, margin: 0, isTextBox: true, objectName: `${title} title` });
    bullets(s, lines, x + 0.2, 1.85, 2.45, 3.0, { fontSize: 12, name: `${title} details` });
  };
  box(3.8, "VADER", HEX.teal, ["7,500-word human-rated lexicon (-4 to +4)", "Rules: !!!, CAPS, \"very\", \"but\", negation in a 3-word window", [{ text: "Compound score: ", bold: true }, "Pos >= 0.05, Neg <= -0.05"], [{ text: "Tuned: ", bold: true }, "+/- 0.40"]]);
  box(6.7, "TextBlob", HEX.amber, ["2,900 scored words, mostly adjectives", "Intensifier x1.3, negation x-0.5, then average", [{ text: "Polarity: ", bold: true }, "Pos > 0, Neg < 0, else Neutral"], [{ text: "Tuned: ", bold: true }, "Pos >= 0.075, Neg <= 0.05"]]);
}

// 12. Results
{
  const s = content("VADER wins, but neither finds Neutral", "Modelling and results",
    "Presenter: [Member 4]. Main metric is macro-F1 because 91% of reviews are positive. Always-Positive baseline: 91.1% accuracy but macro-F1 0.318. VADER default: 89.1% accuracy, macro-F1 0.469, Negative F1 0.43. TextBlob default: 86.6%, 0.404, Negative F1 0.23. Neutral F1 is about 0.02-0.04 for both. Tuning helps VADER to 0.490 (Neutral F1 0.18) but not TextBlob (0.391), because the dev set has only 13 neutral reviews.");
  const labels = ["Accuracy", "Macro-F1", "F1 Negative", "F1 Neutral"];
  s.addChart(pres.charts.BAR, [
    { name: "Always Positive", labels, values: [0.911, 0.318, 0, 0] },
    { name: "VADER", labels, values: [0.891, 0.469, 0.432, 0.023] },
    { name: "TextBlob", labels, values: [0.866, 0.404, 0.234, 0.044] },
    { name: "VADER tuned", labels, values: [0.859, 0.49, 0.361, 0.177] },
  ], { x: 0.4, y: 1.15, w: 6.0, h: 3.9, barDir: "col", barGapWidthPct: 60, chartColors: ["B8C4C7", HEX.teal, HEX.amber, "7FC8D0"], showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "0.00",
       valAxisMaxVal: 1, valAxisMinVal: 0, valAxisLabelFormatCode: "0.0", showLegend: true, legendPos: "b", ...chartBase("1,000-review test sample"), dataLabelFontSize: 8, objectName: "results chart" });
  stat(s, 6.7, 1.2, 2.8, "0.469", "VADER macro-F1 vs 0.404 for TextBlob (default thresholds)", C.accent1, "vader f1");
  stat(s, 6.7, 2.6, 2.8, "91.1%", "accuracy of always saying Positive: higher than both models", C.accent3, "baseline");
  s.addText("Accuracy rewards ignoring the 9% minority, so we rank models by macro-F1.", { x: 6.7, y: 4.0, w: 2.8, h: 0.95, fontSize: 13, color: C.accent6, margin: 0, valign: "top", isTextBox: true, objectName: "metric note" });
}

// 13. Confusion matrices
{
  const s = content("Most mistakes are \"Positive\"", "Modelling and results",
    "Presenter: [Member 4]. Read the rows: 73% of neutral reviews are called Positive by VADER, 84% by TextBlob. VADER catches 42% of negative reviews vs 24% for TextBlob. The models agree on 92% of reviews but Cohen's kappa is only 0.42: they disagree on the hard cases.");
  s.addImage({ path: path.join(__dirname, "cm_default.png"), x: 0.5, y: 1.15, w: 6.3, h: 6.3 * 700 / 1626, objectName: "confusion matrices", altText: "Confusion matrices for VADER and TextBlob with default thresholds" });
  stat(s, 7.05, 1.2, 2.45, "73-84%", "of neutral reviews predicted Positive", C.accent2, "neutral error");
  stat(s, 7.05, 2.6, 2.45, "42% vs 24%", "negative reviews caught: VADER vs TextBlob", C.accent3, "negative recall", 24);
  s.addText("Agreement 91.8%, Cohen's kappa 0.42, score correlation r = 0.27", { x: 0.5, y: 4.15, w: 6.3, h: 0.6, fontSize: 13, color: C.accent6, margin: 0, valign: "top", isTextBox: true, objectName: "agreement note" });
}

// 14. Where lexicons fail
{
  const s = content("Where lexicons break down", "Modelling and results",
    "Presenter: [Member 4]. Accuracy falls from about 95% on 11-30 word reviews to 74% on reviews over 150 words: long reviews are mixed. Error patterns: polite criticism full of positive words; sentiment only in the title; misspellings ('dissapointed', 'Unpleasent') missing from the lexicon; domain complaints with no negative word ('didn't smell like redcurrant'). On all 1,378 unique reviews the results are almost identical, so the sample is representative.");
  s.addChart(pres.charts.BAR, [
    { name: "VADER", labels: ["1-10", "11-30", "31-70", "71-150", "150+"], values: [0.897, 0.948, 0.891, 0.878, 0.737] },
    { name: "TextBlob", labels: ["1-10", "11-30", "31-70", "71-150", "150+"], values: [0.866, 0.922, 0.882, 0.829, 0.737] },
  ], { x: 0.4, y: 1.15, w: 4.6, h: 3.9, barDir: "col", chartColors: [HEX.teal, HEX.amber], showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "0%",
       valAxisMinVal: 0.5, valAxisMaxVal: 1, valAxisLabelFormatCode: "0%", showLegend: true, legendPos: "b", ...chartBase("Accuracy by review length (words)"), dataLabelFontSize: 8, objectName: "length accuracy chart" });
  const fails = [
    ["Polite criticism", "\"Not So Great for Volume... a bit smoother, but...\" VADER: +0.98 on a 2-star review", HEX.rose],
    ["Misspelled sentiment", "\"dissapointed\", \"Unpleasent\" are not in either lexicon", HEX.amber],
    ["Domain complaints", "\"Didn't smell like redcurrant and basil\" has no negative word", HEX.teal],
  ];
  fails.forEach(([h, t, col], i) => {
    const y = 1.2 + i * 1.27;
    card(s, 5.3, y, 4.2, 1.1, `failure card ${i + 1}`);
    s.addShape(pres.shapes.OVAL, { x: 5.48, y: y + 0.17, w: 0.2, h: 0.2, fill: { color: col }, line: { color: col }, objectName: `failure dot ${i + 1}` });
    s.addText([{ text: h, options: { bold: true, color: HEX.ink, breakLine: true } }, { text: t, options: { color: HEX.grey, fontSize: 12 } }],
      { x: 5.8, y: y + 0.1, w: 3.55, h: 0.92, fontSize: 14, margin: 0, valign: "top", isTextBox: true, objectName: `failure text ${i + 1}` });
  });
}

pres.addSection({ title: "Wrap-up" });

// 15. Live demo
{
  const s = content("Live demo: one script, end to end", "Wrap-up",
    "Presenter: [Member 4]. Run python team2_phase1_lexicon_sentiment.py. Show the console sections Step 1 to Step 7, then open figures/ and results/09_comparison_table.csv. Total runtime is under a minute.");
  card(s, 0.5, 1.2, 5.0, 3.2, "terminal card", C.text2);
  s.addText([
    { text: "> python team2_phase1_lexicon_sentiment.py", options: { color: HEX.amber, breakLine: true } },
    { text: " ", options: { breakLine: true } },
    { text: "STEP 1  Data exploration      -> 9 figures", options: { breakLine: true } },
    { text: "STEP 2  Labels, columns, outliers", options: { breakLine: true } },
    { text: "STEP 3  Lexicon probe sentences", options: { breakLine: true } },
    { text: "STEP 4  Per-model pre-processing", options: { breakLine: true } },
    { text: "STEP 5  Sample 1,000 + ablation", options: { breakLine: true } },
    { text: "STEP 6  VADER and TextBlob scores", options: { breakLine: true } },
    { text: "STEP 7  Metrics, confusion matrices", options: { breakLine: true } },
    { text: " ", options: { breakLine: true } },
    { text: "Done. 16 figures, 13 tables, run_log.txt", options: { color: "7FC8D0" } },
  ], { x: 0.75, y: 1.4, w: 4.6, h: 2.85, fontFace: "Courier New", fontSize: 12, color: "FFFFFF", margin: 0, valign: "top", isTextBox: true, objectName: "terminal text" });
  bullets(s, [
    [{ text: "Reproducible: ", bold: true }, "seed 42, downloads the data itself"],
    [{ text: "Documented: ", bold: true }, "sections follow assignment steps 1-7"],
    [{ text: "Outputs: ", bold: true }, "./figures (charts), ./results (CSV tables, full log)"],
    [{ text: "Deliverables: ", bold: true }, ".py source, analysis report with screenshots, this deck"],
  ], 5.85, 1.3, 3.65, 3.6, { fontSize: 14, name: "demo bullets" });
}

// 16. Conclusions
{
  const s = pres.addSlide({ masterName: "TITLE_DARK", sectionTitle: "Wrap-up" });
  s.addText("Takeaways and next steps", { placeholder: "title" });
  s.addNotes("Presenter: all members, one point each. Close with the Phase 2 plan, then take questions.");
  const pts = [
    ["Clean first", "74% of rows were copies; 1,379 real reviews", HEX.rose],
    ["Measure fairly", "Macro-F1, not accuracy, on 91% positive data", HEX.amber],
    ["VADER > TextBlob", "0.469 vs 0.404 macro-F1; Neutral is unsolved", HEX.teal],
    ["Phase 2", "ML models on the same 1,000 reviews, recommender, LLM tasks", "7FC8D0"],
  ];
  pts.forEach(([h, t, col], i) => {
    const x = 0.6 + i * 2.22;
    s.addShape(pres.shapes.OVAL, { x, y: 3.15, w: 0.3, h: 0.3, fill: { color: col }, line: { color: col }, objectName: `conclusion dot ${i + 1}` });
    s.addText([{ text: h, options: { bold: true, color: "FFFFFF", fontSize: 16, breakLine: true } }, { text: t, options: { color: "CFE3E6", fontSize: 12 } }],
      { x, y: 3.6, w: 2.0, h: 1.4, margin: 0, valign: "top", isTextBox: true, objectName: `conclusion ${i + 1}` });
  });
  s.addText("Questions?", { x: 0.6, y: 0.6, w: 4, h: 0.5, fontSize: 16, color: HEX.amber, margin: 0, isTextBox: true, objectName: "questions" });
}

(async () => {
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})();
