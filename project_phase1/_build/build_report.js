// Builds the Phase 1 analysis report (.docx) for COMP 262 Team 2
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, ImageRun, Table, TableRow, TableCell,
  HeadingLevel, AlignmentType, WidthType, ShadingType, BorderStyle, PageBreak,
  TableOfContents, LevelFormat, Footer, PageNumber, Header,
} = require("docx");

const PROJ = "C:\\Users\\rishi\\OneDrive\\Desktop\\ARIES\\COMP-262\\project_phase1";
const FIG = path.join(PROJ, "figures");
const T = JSON.parse(fs.readFileSync(path.join(__dirname, "tables.json"), "utf8"));
const SIZES = JSON.parse(fs.readFileSync(path.join(__dirname, "sizes.json"), "utf8"));

const FONT = "Calibri";
const ACCENT = "00798C";
const CONTENT_W = 9360; // 6.5" in DXA

// ---------- helpers ----------
// inline **bold** markup -> runs
function runs(text, opts = {}) {
  return text.split(/(\*\*[^*]+\*\*)/).filter(Boolean).map((s) =>
    s.startsWith("**") ? new TextRun({ text: s.slice(2, -2), bold: true, ...opts }) : new TextRun({ text: s, ...opts }));
}
const P = (text, opts = {}) => new Paragraph({ children: runs(text), spacing: { after: 120, line: 276 }, ...opts });
const H1 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun(t)] });
const H2 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun(t)] });
const H3 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun(t)] });
const B = (t, level = 0) => new Paragraph({ numbering: { reference: "bullets", level }, children: runs(t), spacing: { after: 60 } });
const N = (t, ref = "numbers") => new Paragraph({ numbering: { reference: ref, level: 0 }, children: runs(t), spacing: { after: 60 } });
const pageBreak = () => new Paragraph({ children: [new PageBreak()] });

let figNo = 0, tabNo = 0;
function figure(file, widthIn, caption) {
  const [w, h] = SIZES[file];
  const wPx = widthIn * 96, hPx = wPx * h / w;
  figNo += 1;
  return [
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 120, after: 40 }, keepNext: true,
      children: [new ImageRun({ type: "png", data: fs.readFileSync(path.join(FIG, file)),
        transformation: { width: Math.round(wPx), height: Math.round(hPx) },
        altText: { title: caption, description: caption, name: file } })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 200 },
      children: [new TextRun({ text: `Figure ${figNo}. `, bold: true, size: 18, color: "555555" }),
                 new TextRun({ text: caption, italics: true, size: 18, color: "555555" })] }),
  ];
}

const border = { style: BorderStyle.SINGLE, size: 4, color: "BFBFBF" };
const borders = { top: border, bottom: border, left: border, right: border };
// rows: array of arrays (first = header); widths in DXA summing to CONTENT_W (or less)
function table(rows, widths, caption, { highlight = () => false, fontSize = 18 } = {}) {
  const total = widths.reduce((a, b) => a + b, 0);
  tabNo += 1;
  const out = [];
  if (caption) out.push(new Paragraph({ spacing: { before: 160, after: 60 }, keepNext: true,
    children: [new TextRun({ text: `Table ${tabNo}. `, bold: true, size: 18, color: "555555" }),
               new TextRun({ text: caption, italics: true, size: 18, color: "555555" })] }));
  out.push(new Table({
    width: { size: total, type: WidthType.DXA }, columnWidths: widths,
    rows: rows.map((r, i) => new TableRow({
      tableHeader: i === 0, cantSplit: true,
      children: r.map((c, j) => new TableCell({
        borders, width: { size: widths[j], type: WidthType.DXA },
        margins: { top: 50, bottom: 50, left: 90, right: 90 },
        shading: i === 0 ? { fill: ACCENT, type: ShadingType.CLEAR, color: "auto" }
          : highlight(r, i) ? { fill: "E6F4F1", type: ShadingType.CLEAR, color: "auto" }
          : i % 2 === 0 ? { fill: "F7F7F7", type: ShadingType.CLEAR, color: "auto" } : undefined,
        children: [new Paragraph({ children: runs(String(c), { size: fontSize, bold: i === 0 ? true : undefined,
          color: i === 0 ? "FFFFFF" : undefined, font: FONT }) })],
      })),
    })),
  }));
  out.push(new Paragraph({ spacing: { after: 120 }, children: [] }));
  return out;
}
const tbl = (name) => [T.cols[name], ...T.rows[name]];
const f3 = (x) => Number(x).toFixed(3);

// ---------- content (essentials only) ----------
const children = [];
const TEAM = ["Manav Chaudhary", "Harsh Eary", "Steven Hao Gong", "Rishi Goyal", "Parth Kishorbhai Kevadiya", "Jasdish Singh"];

// Cover + contents
children.push(
  new Paragraph({ spacing: { before: 1800 }, alignment: AlignmentType.CENTER, children: [new TextRun({ text: "COMP 262: Natural Language Processing and Recommender Systems", size: 26, color: "555555" })] }),
  new Paragraph({ spacing: { before: 400, after: 200 }, alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Sentiment Analysis of Amazon All Beauty Reviews", size: 48, bold: true, color: ACCENT })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Phase 1: Lexicon-based Models (VADER and TextBlob)", size: 28 })] }),
  new Paragraph({ spacing: { before: 800, after: 120 }, alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Team 2", size: 32, bold: true })] }),
  ...TEAM.map((t) => new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: t, size: 24 })] })),
  new Paragraph({ spacing: { before: 800 }, alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Centennial College  |  Professor: [Name]  |  November 5, 2026", size: 22, color: "555555" })] }),
  pageBreak(),
  new Paragraph({ children: [new TextRun({ text: "Table of Contents", bold: true, size: 32, font: FONT, color: ACCENT })], spacing: { after: 200 } }),
  new TableOfContents("Table of Contents", { hyperlink: true, headingStyleRange: "1-1" }),
  pageBreak(),
);

// 1. Introduction
children.push(
  H1("1. Introduction"),
  B("**Goal:** label Amazon All Beauty reviews as Positive, Neutral or Negative using two lexicon models."),
  B("**Data:** Amazon Review Data (2018), All Beauty 5-core, published by Ni, Li and McAuley (UCSD)."),
  B("**Code:** team2_phase1_lexicon_sentiment.py runs every step and saves all charts and tables (seed 42)."),
);

// 2. Data exploration
children.push(
  H1("2. Data Exploration"),
  ...table([
    ["Rows", "Columns", "Products", "Reviewers", "Avg rating", "5-star", "Missing text"],
    ["5,269", "12", "85", "991", "4.77", "88.2%", "5 rows"],
  ], [1300, 1200, 1300, 1400, 1300, 1260, 1600], "Dataset overview"),
  B("**Duplicates:** 3,890 rows (73.8%) are the same review copied under different product sizes/colours. Only **1,379 unique reviews** remain."),
  B("**Imbalance:** 91% of unique reviews are positive, so we use **macro-F1**, not accuracy."),
  B("**Products:** median 4 reviews per product; 2 products hold 62.5% of all reviews."),
  B("**Users:** 86% of reviewers wrote only one unique review."),
  B("**Length:** median 29 words; 106 long outliers (> 161 words). 2-3 star reviews are the longest."),
  ...figure("01_rating_distribution.png", 3.9, "Rating distribution"),
  ...figure("02_duplicate_copies.png", 3.9, "Product variants each review is copied to"),
  ...figure("03_reviews_per_product.png", 6.0, "Reviews per product and their cumulative share"),
  ...figure("05_reviews_per_user.png", 6.0, "Reviews per user: raw vs unique"),
  ...figure("06_review_length.png", 6.5, "Review length, outliers and length by rating"),
);

// 3. Pre-processing
children.push(
  H1("3. Pre-processing"),
  ...table([["Label", "Ratings", "Reviews", "Share"], ["Positive", "4, 5", "1,259", "91.3%"], ["Neutral", "3", "64", "4.6%"], ["Negative", "1, 2", "56", "4.1%"]],
    [2340, 2340, 2340, 2340], "Labels from star ratings"),
  B("**Columns used:** reviewText + summary as input; overall as the label only. Other columns were used only for exploration."),
  B("**Removed:** duplicate copies, 1 empty review, and auto-generated \"Five Stars\" titles (they repeat the rating, so they leak the label)."),
  B("**Kept:** short and long reviews. They are real opinions, and long ones are mostly the minority classes."),
  ...table([
    ["Step", "VADER", "TextBlob", "Reason"],
    ["Clean HTML, URLs, \"sooooo\"", "Yes", "Yes", "Noise"],
    ["Lower-case, remove punctuation", "No", "Yes", "VADER uses CAPS and \"!\""],
    ["Remove stop words", "No", "Yes, keep \"not\"", "Negations carry sentiment"],
    ["Lemmatise", "No", "Yes", "Match TextBlob's base words"],
  ], [3000, 1300, 1900, 3160], "Pre-processing per model"),
  B("Tested on 378 held-out reviews: this setup gave TextBlob macro-F1 0.506 vs 0.460 with light cleaning. Removing \"not\" hurt both models."),
);

// 4. Models
children.push(
  H1("4. Models"),
  B("**Lexicons studied:** VADER, TextBlob, SentiWordNet. We chose **VADER** (handles negation, CAPS, \"!!!\", emoticons, \"but\") and **TextBlob** (adjective lexicon with subjectivity). **SentiWordNet** was rejected: it scores \"not good\" as positive."),
  ...table([
    ["", "VADER", "TextBlob"],
    ["How it works", "Rated word list + grammar rules; compound score -1 to +1", "Averages adjective scores; flips on negation; polarity -1 to +1"],
    ["Rule", "Pos >= 0.05, Neg <= -0.05", "Pos > 0, Neg < 0"],
    ["Tuned threshold", "+/- 0.40", "Pos >= 0.075, Neg <= 0.05"],
  ], [1900, 3730, 3730], "The two models"),
  B("**Test data:** 1,000 random unique reviews (911 Positive, 51 Neutral, 38 Negative). Thresholds were tuned only on the other 378 reviews."),
);

// 5. Results
const comp = tbl("09_comparison_table");
children.push(
  H1("5. Results"),
  ...table([["Model", "Accuracy", "Macro-F1", "F1 Neg", "F1 Neu", "F1 Pos"],
    ...comp.slice(1).map((r) => [r[0], f3(r[1]), f3(r[4]), f3(r[6]), f3(r[7]), f3(r[8])])],
    [3760, 1120, 1120, 1120, 1120, 1120], "Results on the 1,000-review test sample",
    { highlight: (r) => r[0].startsWith("VADER (default") }),
  ...figure("13_confusion_matrices.png", 5.0, "Confusion matrices (default and tuned thresholds)"),
  B("**VADER wins:** macro-F1 0.469 vs 0.404, and it finds far more negative reviews."),
  B("**Neutral is the weak spot:** 73-84% of neutral reviews are predicted Positive."),
  B("**Long reviews are harder:** accuracy falls from 95% (11-30 words) to 74% (150+ words)."),
  B("**Tuning** raised VADER to 0.490 but did not help TextBlob."),
);

// 6. Conclusion, assumptions, references
children.push(
  H1("6. Conclusion"),
  B("De-duplication was essential: 74% of rows were copies."),
  B("VADER is the better lexicon for these reviews."),
  B("Both lexicons miss neutral and mixed reviews, which Phase 2's machine-learning models will target on the same 1,000 reviews."),
  H1("7. Assumptions"),
  B("The star rating is the true sentiment."),
  B("Same reviewer + text + timestamp = the same review."),
  H1("8. References"),
  ...[
    "Ni, J., Li, J., & McAuley, J. (2019). Justifying recommendations using distantly-labeled reviews and fine-grained aspects. EMNLP-IJCNLP 2019. https://nijianmo.github.io/amazon/index.html",
    "Hutto, C. J., & Gilbert, E. (2014). VADER: A parsimonious rule-based model for sentiment analysis of social media text. ICWSM-14.",
    "Loria, S. TextBlob documentation. https://textblob.readthedocs.io",
    "Baccianella, S., Esuli, A., & Sebastiani, F. (2010). SentiWordNet 3.0. LREC 2010.",
    "Bird, S., Klein, E., & Loper, E. (2009). Natural Language Processing with Python. O'Reilly.",
  ].map((r) => new Paragraph({ numbering: { reference: "refs", level: 0 }, children: [new TextRun(r)], spacing: { after: 80 } })),
);

// Appendices
children.push(
  pageBreak(),
  H1("Appendix 1: Project Plan"),
  ...table([["Week", "Task", "Owner", "Status"],
    ["3", "Teams formed, dataset assigned", "All", "Done"],
    ["4-5", "Exploration and pre-processing", "Manav Chaudhary, Rishi Goyal", "Done"],
    ["6", "Lexicon study and per-model cleaning", "Steven Hao Gong, Harsh Eary", "Done"],
    ["7", "Models and validation", "Parth Kishorbhai Kevadiya, Jasdish Singh", "Done"],
    ["8", "Report and Phase 1 demo", "All", "In progress"],
    ["9-14", "Phase 2", "All", "Planned"]], [1300, 4860, 1700, 1500], "Project plan"),
  H1("Appendix 2: Meeting Register"),
  ...table([["Date / time", "Attendees", "Subjects", "Assignments"],
    ["[dd-mm-yyyy hh:mm]", "[names]", "Kick-off, dataset", "Manav, Rishi: exploration and pre-processing"],
    ["[dd-mm-yyyy hh:mm]", "[names]", "Exploration findings", "Steven, Harsh: lexicons and cleaning"],
    ["[dd-mm-yyyy hh:mm]", "[names]", "Lexicons and results", "Parth, Jasdish: models and results"],
    ["", "", "", ""]], [1900, 1900, 3060, 2500], "Meeting register"),
);

// ---------- document ----------
const doc = new Document({
  creator: "COMP 262 Team 2",
  title: "Phase 1: Lexicon Sentiment Analysis of Amazon All Beauty Reviews",
  styles: {
    default: { document: { run: { font: FONT, size: 22 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { font: FONT, size: 32, bold: true, color: ACCENT }, paragraph: { spacing: { before: 360, after: 160 }, keepNext: true, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { font: FONT, size: 26, bold: true, color: "333333" }, paragraph: { spacing: { before: 240, after: 100 }, keepNext: true, outlineLevel: 1 } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { font: FONT, size: 22, bold: true }, paragraph: { spacing: { before: 160, after: 80 }, keepNext: true, outlineLevel: 2 } },
    ],
  },
  numbering: { config: [
    { reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "\u2022", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 540, hanging: 270 } } } },
                                     { level: 1, format: LevelFormat.BULLET, text: "\u25E6", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 1080, hanging: 270 } } } }] },
    { reference: "numbers", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 540, hanging: 300 } } } }] },
    { reference: "numbers2", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 540, hanging: 300 } } } }] },
    { reference: "refs", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "[%1]", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 600, hanging: 600 } } } }] },
  ] },
  sections: [{
    properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 } }, titlePage: true },
    headers: { default: new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "COMP 262 Project | Team 2 | Phase 1", size: 16, color: "888888" })] })] }) },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ children: ["Page ", PageNumber.CURRENT], size: 16, color: "888888" })] })] }) },
    children,
  }],
});

Packer.toBuffer(doc).then((buf) => {
  const out = path.join(PROJ, "Team2_Phase1_Analysis_Report.docx");
  fs.writeFileSync(out, buf);
  console.log("wrote", out);
});
