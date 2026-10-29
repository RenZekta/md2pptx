// JS engine tests for md2pptx (the engine core in webapp/index.html).
//
// Run from the repo root (A:\a\md2pptx) AFTER run_tests.py:
//     node Tests/run_tests_js.js
//
// Extracts the engine core from webapp/index.html (everything up to the
// Mode B UI marker), builds the same deck from Tests/test_structure.md +
// MFUA-Template.pptx as run_tests.py, and checks:
//   - slide1 topic = "На тему: <topic>" (format preserved), bold, font kept
//   - numbered body lines and plan items -> auto-numbered paragraphs
//     (a:buAutoNum), the number not stored in the text; run/startAt/scheme
//     rules and the preview's numberLabels, on Tests/test_numbering.md
//   - plan item sizes, closing title size, no sz on body runs
//   - inline marks: *italic* -> i="1", ++underline++ -> u="sng", **bold** ->
//     b="1", and no marker characters left in the text
//   - pipe table -> p:graphicFrame with a:tbl, one tr per row, one td per cell,
//     head row bold, cells stamped at body_size - 1, separator alignment, frame
//     placed after the body placeholder inside the shape tree
//   - byte-identical part-by-part parity vs Tests/out/ru_deck_py.pptx
//   - byte-identical parity vs Tests/out/numbering_deck_py.pptx (an embedded
//     .xlsx is compared by its entries)
//   - layouts / table options / merged cells / charts on the slides appended
//     to test_structure.md, their error messages, the preview's slideSpecs
//   - lang front matter: every generated rPr/endParaRPr stamped with
//     lang + dirty="0", invalid lang rejected, and byte-identical parity
//     vs Tests/out/lang_deck_py.pptx (both produced by run_tests.py)
//
// Requires: Node, and jszip 3.x. JSZIP_PATH points at the local Windows copy;
// on other machines the plain `require("jszip")` fallback is used (install
// jszip in a node_modules next to this test, e.g. at the repo root).

const fs = require("fs");
const path = require("path");

const HERE = __dirname;
const REPO = path.resolve(HERE, "..");
const JSZIP_PATH =
  "C:/Users/Ren/AppData/Roaming/npm/node_modules/openclaw/node_modules/jszip";
let JSZip;
try {
  JSZip = require(JSZIP_PATH);
} catch (e) {
  JSZip = require("jszip");
}

const html = fs.readFileSync(path.join(REPO, "webapp", "index.html"), "utf8");
// Engine core = the section between the shared-core header and the Mode B
// UI marker (the bundled JSZip/jsPDF scripts and the DOM UI are outside it).
const START = "// md2pptx deterministic engine";
const MARK = "// md2pptx web app (Mode B)";
const sIdx = html.indexOf(START);
const eIdx = html.indexOf(MARK);
if (sIdx < 0 || eIdx < 0 || eIdx <= sIdx) {
  throw new Error("engine core markers not found in webapp/index.html");
}
const core = html.slice(sIdx, eIdx);

// Evaluate the engine core with its own `module.exports` guard intact.
const mod = { exports: {} };
const api = new Function("module", "exports", "JSZip",
  core + "\nreturn module.exports;")(mod, mod.exports, JSZip);
if (!api || typeof api.buildZip !== "function") {
  throw new Error("engine core did not export buildZip");
}

const structText = fs.readFileSync(path.join(HERE, "test_structure.md"), "utf8");
const tplBuf = fs.readFileSync(path.join(REPO, "MFUA-Template.pptx"));
const tplEnBuf = fs.readFileSync(path.join(REPO, "English-Template.pptx"));
const outDir = path.join(HERE, "out");
fs.mkdirSync(outDir, { recursive: true });
const jsOut = path.join(outDir, "ru_deck_js.pptx");
const pyOut = path.join(outDir, "ru_deck_py.pptx");
if (!fs.existsSync(pyOut)) {
  throw new Error("run_tests.py has not produced " + pyOut);
}

let failures = [];
function check(name, cond, detail) {
  console.log("  " + (cond ? "PASS" : "FAIL") + "  " + name +
    (cond ? "" : " -> " + detail));
  if (!cond) failures.push(name);
}
function texts(xml) {
  return [...xml.matchAll(/<a:t>([\s\S]*?)<\/a:t>/g)].map((m) => m[1]);
}
function runOf(xml, text) {
  const i = xml.indexOf(text);
  if (i < 0) return "";
  return xml.slice(xml.lastIndexOf("<a:r>", i), xml.indexOf("</a:r>", i) + 6);
}
function region(xml, marker) {
  const i = xml.indexOf(marker);
  if (i < 0) return "";
  return xml.slice(i, xml.indexOf("</p:sp>", i) + 7);
}

// Names of the parts whose bytes differ between two packages, or that only one
// of them has. An embedded .xlsx is itself a zip whose wrapper bytes depend on
// the zip implementation (Python zipfile vs JSZip), so two workbooks are equal
// when their entries are.
async function partsDiff(za, zb) {
  const names = (z2) => Object.keys(z2.files).filter((n) => !z2.files[n].dir);
  const out = [];
  for (const n of new Set([...names(za), ...names(zb)])) {
    const fa = za.file(n), fb = zb.file(n);
    if (!fa || !fb) { out.push((fa ? "missing in py: " : "missing in js: ") + n); continue; }
    const a = await fa.async("nodebuffer"), b = await fb.async("nodebuffer");
    if (n.endsWith(".xlsx")) {
      const xa = await JSZip.loadAsync(a), xb = await JSZip.loadAsync(b);
      const inner = await partsDiff(xa, xb);
      if (inner.length) out.push(n + " (" + inner.join(", ") + ")");
    } else if (!a.equals(b)) {
      out.push(n);
    }
  }
  return out;
}

(async () => {
  const j = JSON.stringify;
  const doc = api.parseMd(structText);
  const { zip } = await api.buildZip(tplBuf, doc);
  fs.writeFileSync(jsOut, await zip.generateAsync({ type: "nodebuffer" }));

  const load = (p) => JSZip.loadAsync(fs.readFileSync(p));
  const z = await load(jsOut);
  const part = (z2, n) => z2.file(n).async("string");
  const topic = doc.title;

  const s1 = await part(z, "ppt/slides/slide1.xml");
  const t1 = texts(s1);
  check("slide1 topic = 'На тему: ' + topic",
    t1.includes("На тему: " + topic), JSON.stringify(t1));
  const tr = runOf(s1, "На тему: " + topic);
  check("topic run is bold", tr.includes('b="1"'), tr);
  check("topic typeface kept (Times New Roman)",
    tr.includes('typeface="Times New Roman"'), tr);

  const s2 = await part(z, "ppt/slides/slide2.xml");
  const t2 = texts(s2);
  check("plan items carry no typed numbers",
    JSON.stringify(t2.slice(1, 5)) === JSON.stringify(doc.plan.map((t) => t[0])) &&
    !t2.slice(1, 5).some((t) => /^[0-9]+[.)] /.test(t)), JSON.stringify(t2));
  check("plan items are auto-numbered (arabicPeriod, no startAt)",
    (region(s2, '<p:ph idx="1"/>').match(/<a:buAutoNum type="arabicPeriod"\/>/g) || []).length === 4);
  check("plan items stamped at body_size - 1",
    (region(s2, '<p:ph idx="1"/>').match(/sz="1700"/g) || []).length === 4);

  const s3 = await part(z, "ppt/slides/slide3.xml");
  check("numbered body line text carries no typed number",
    texts(s3).includes("Англия закрепляет монопольный характер патента") &&
    !texts(s3).some((t) => /^[0-9]+[.)] /.test(t)));
  const i3 = s3.indexOf("Англия закрепляет");
  const numberedPpr = s3.slice(s3.lastIndexOf("<a:p>", i3), i3);
  check("numbered body line is level 1 and auto-numbered",
    numberedPpr.includes('lvl="1"') &&
    numberedPpr.includes('<a:buFontTx/><a:buAutoNum type="arabicPeriod"/>'), numberedPpr);
  const v3 = s3.indexOf("Венеция");
  check("bullet line is not auto-numbered",
    !s3.slice(s3.lastIndexOf("<a:p>", v3), v3).includes("buAutoNum"));
  check("body runs carry no sz",
    !region(s3, '<p:ph idx="1"/>').includes('sz="'));

  const s4 = await part(z, "ppt/slides/slide4.xml");
  check("closing title stamped at template size (3600)",
    region(s4, '<p:ph type="title"/>').includes('sz="3600"'));

  const t3 = texts(s3);
  const flat3 = t3.join(" ");
  const ir = runOf(s3, "italic");
  check("italic run stamped i=\"1\"", ir.includes('i="1"'), ir);
  const ur = runOf(s3, "underlined");
  check("underlined run stamped u=\"sng\"", ur.includes('u="sng"'), ur);
  check("bold still stamped alongside", runOf(s3, "Италии").includes('b="1"'));
  check("no stray markers in the text", !flat3.includes("*") && !flat3.includes("++"),
    JSON.stringify(flat3));
  const s5 = await part(z, "ppt/slides/slide5.xml");
  const t5 = texts(s5);
  check("table cells present",
    ["Cell A", "Cell B", "Cell C", "6"].every(function (x) { return t5.includes(x); }),
    JSON.stringify(t5));
  check("table is a p:graphicFrame", s5.includes("<p:graphicFrame>"), s5.slice(0, 200));
  check("a:tbl with one row per line",
    (s5.match(/<a:tr /g) || []).length === 3, s5);
  check("one a:tc per cell (3 rows x 3 cols)",
    (s5.match(/<a:tc>/g) || []).length === 9, s5);
  check("head row cells inherit the table's bold",
    runOf(s5, "Cell A").includes('b="1"'), runOf(s5, "Cell A"));
  check("table cells stamped at body_size - 1",
    (s5.match(/sz="1700"/g) || []).length === 9, s5);
  check("separator alignment applied",
    s5.includes('algn="ctr"') && s5.includes('algn="r"'), s5);
  check("frame comes after the body placeholder",
    s5.indexOf("<p:graphicFrame>") > s5.indexOf('<p:ph idx="1"/>'), s5.slice(0, 200));
  check("frame inside the shape tree",
    s5.indexOf("<p:graphicFrame>") < s5.lastIndexOf("</p:spTree>"), s5.slice(-200));

  // Byte-identical part-by-part parity with the Python build.
  const py = await load(pyOut);
  const mism = await partsDiff(z, py);
  check("byte-identical parts vs Python build", mism.length === 0,
    mism.join(", "));

  // lang front matter: stamped on every generated run + end mark.
  const langText = structText.replace("---\nbody_size: 18\n---",
    "---\nbody_size: 18\nlang: ru-RU\n---");
  const langPyOut = path.join(outDir, "lang_deck_py.pptx");
  let langDoc = null;
  let parseErr = null;
  try {
    langDoc = api.parseMd(langText);
  } catch (e) {
    parseErr = e;
  }
  check("lang front matter parsed", !!langDoc && langDoc.lang === "ru-RU",
    parseErr ? parseErr.message : JSON.stringify(langDoc && langDoc.lang));

  if (langDoc) {
    const langJsOut = path.join(outDir, "lang_deck_js.pptx");
    const { zip: lz } = await api.buildZip(tplBuf, langDoc);
    fs.writeFileSync(langJsOut, await lz.generateAsync({ type: "nodebuffer" }));

    const bad = [];
    const stampedBad = (xml) => {
      for (const m of xml.matchAll(/<a:(?:rPr|endParaRPr)\b[^>]*>/g)) {
        const tag = m[0];
        const lang = tag.match(/lang="([^"]+)"/);
        const dirty = tag.match(/dirty="([^"]+)"/);
        if (!lang || lang[1] !== "ru-RU" || !dirty || dirty[1] !== "0") {
          bad.push(tag);
        }
      }
    };
    const s1 = await part(lz, "ppt/slides/slide1.xml");
    stampedBad(region(s1, "На тему: " + langDoc.title));
    const s2 = await part(lz, "ppt/slides/slide2.xml");
    stampedBad(region(s2, '<p:ph type="title"/>'));
    stampedBad(region(s2, '<p:ph idx="1"/>'));
    for (let i = 3; i <= 4; i++) {
      const x = await part(lz, "ppt/slides/slide" + i + ".xml");
      stampedBad(region(x, '<p:ph type="title"/>'));
      stampedBad(region(x, '<p:ph idx="1"/>'));
    }
    check("all generated rPr/endParaRPr carry lang=ru-RU dirty=0",
      bad.length === 0, bad.slice(0, 4).join("; "));

    if (fs.existsSync(langPyOut)) {
      const lp = await load(langPyOut);
      const mism2 = await partsDiff(lz, lp);
      check("byte-identical parts vs Python lang build", mism2.length === 0,
        mism2.join(", "));
    }
  }

  // Numbered lists: runs, startAt, schemes, tables, plan (Tests/test_numbering.md).
  {
    const numText = fs.readFileSync(path.join(HERE, "test_numbering.md"), "utf8");
    const numDoc = api.parseMd(numText);
    const P = "arabicPeriod", R = "arabicParenR";
    const j = JSON.stringify;
    const m1 = numDoc.slides[0][1][0].map((it) => it[2]);
    check("runs: a bullet, a plain line or a restart ends a run; first number is the start",
      j(m1) === j([[P, 1], [P, 1], null, [P, 1], null, [P, 5], [P, 5]]), j(m1));
    const m2 = numDoc.slides[1][1][0].filter((it) => typeof it[0] === "number").map((it) => it[2]);
    check("a scheme change starts a new run; a table does not split a run",
      j(m2) === j([[R, 1], [R, 1], [P, 1], [P, 1]]), j(m2));
    check("plan run keeps its written start",
      j(numDoc.plan.map((t) => t[1])) === j([[P, 3], [P, 3]]), j(numDoc.plan));
    check("1) is the arabicParenR scheme, 1. the arabicPeriod scheme",
      api.parseNumbered("1) a")[0] === R && api.parseNumbered("1. a")[0] === P);
    check("not list items: no space after the marker, or a decimal",
      api.parseNumbered("1.5 million") === null &&
      api.parseNumbered("2026.") === null && api.parseNumbered("- 1. a") === null);
    check("numberLabels count a run from its start and restart after a break",
      j(api.numberLabels(m1)) === j(["1.", "2.", null, "1.", null, "5.", "6."]) &&
      j(api.numberLabels(m2)) === j(["1)", "2)", "1.", "2."]),
      j(api.numberLabels(m1)) + " " + j(api.numberLabels(m2)));

    const { zip: nz } = await api.buildZip(tplBuf, numDoc);
    fs.writeFileSync(path.join(outDir, "numbering_deck_js.pptx"),
      await nz.generateAsync({ type: "nodebuffer" }));
    const n2 = await part(nz, "ppt/slides/slide2.xml");
    const n3 = await part(nz, "ppt/slides/slide3.xml");
    const n5 = await part(nz, "ppt/slides/slide5.xml");
    const cnt = (x, re) => (x.match(re) || []).length;
    check("plan items stamped startAt=3 on every item of the run",
      cnt(region(n2, '<p:ph idx="1"/>'), /<a:buAutoNum type="arabicPeriod" startAt="3"\/>/g) === 2);
    check("startAt is stamped only on the run that starts at 5",
      cnt(n3, /startAt="5"/g) === 2 && cnt(n3, /startAt=/g) === 2);
    check("runs starting at 1 carry no startAt", !n5.includes("startAt"));
    check("slide 2 body: two arabicParenR and two arabicPeriod paragraphs",
      cnt(n5, /type="arabicParenR"/g) === 2 && cnt(n5, /type="arabicPeriod"/g) === 2);
    const kb = n3.indexOf("A bullet breaks");
    check("a bullet inside the list stays a plain level-1 paragraph",
      n3.slice(n3.lastIndexOf("<a:p>", kb), kb).startsWith('<a:p><a:pPr lvl="1"/>'));
    check("every numbered paragraph has buFontTx before buAutoNum (schema order)",
      cnt(n3, /<a:buFontTx\/><a:buAutoNum/g) === cnt(n3, /<a:buAutoNum/g));

    const numPy = path.join(outDir, "numbering_deck_py.pptx");
    if (fs.existsSync(numPy)) {
      const np = await load(numPy);
      const mism3 = await partsDiff(nz, np);
      check("byte-identical parts vs Python numbering build", mism3.length === 0,
        mism3.join(", "));
    }
  }


  // Layouts, table options, merged cells, charts (the slides appended to
  // Tests/test_structure.md), their error messages and the fit estimate.
  {
    const cnt = (x, re) => (x.match(re) || []).length;
    const buildFrom = async (text) => {
      const d = api.parseMd(text);
      const built = await api.buildZip(tplBuf, d);
      return { d: d, built: built };
    };
    const throwsWith = async (text, needle) => {
      try { await buildFrom(text); } catch (e) { return String(e.message).includes(needle); }
      return false;
    };
    const slideDeck = (body) => "# title\nT\n\n# slide\n## S\n" + body + "\n";
    const sx = (n) => part(z, "ppt/slides/slide" + n + ".xml");
    const rl = (n) => part(z, "ppt/slides/_rels/slide" + n + ".xml.rels");

    check("two columns point their slide at the two-content layout (slideLayout4)",
      (await rl(6)).includes("slideLayout4.xml") && !(await rl(6)).includes("slideLayout2.xml"));
    check("left column fills the layout's idx 1 placeholder",
      (await sx(6)).includes('<p:ph sz="half" idx="1"/>') && (await sx(6)).includes("Монополия короны"));
    check("a column without text loses its body placeholder (no empty txBody)",
      !(await sx(6)).includes('<p:ph sz="half" idx="2"/>'));
    check("the chart sits in the right half of the slide",
      /<p:xfrm><a:off x="4648200" y="[0-9]+"\/><a:ext cx="4038600"/.test(await sx(6)));
    check("chart part, relationship and workbook exist",
      (await rl(6)).includes('Target="../charts/chart1.xml"') &&
      !!z.file("ppt/charts/chart1.xml") && !!z.file("ppt/embeddings/Microsoft_Excel_Sheet1.xlsx"));
    const ctx = await part(z, "[Content_Types].xml");
    check("content types: chart override and xlsx default registered once",
      cnt(ctx, /\/ppt\/charts\/chart1\.xml/g) === 1 && cnt(ctx, /Extension="xlsx"/g) === 1);
    const c1 = await part(z, "ppt/charts/chart1.xml");
    check("pie: first series only, per-point theme colours, percent labels, legend right",
      cnt(c1, /<c:ser>/g) === 1 && cnt(c1, /<c:dPt>/g) === 3 &&
      c1.includes('<c:showPercent val="1"/>') && c1.includes('<c:legendPos val="r"/>'));

    const s7 = await sx(7);
    check("table options: widths 3,2,2,2 of the full body width",
      j(s7.match(/<a:gridCol w="(\d+)"\/>/g).map((m) => /(\d+)/.exec(m)[1])) ===
      j(["2743200", "1828800", "1828800", "1828800"]));
    check("table options: size 15 stamped on every cell run",
      j([...new Set([...s7.matchAll(/<a:rPr[^>]* sz="(\d+)"/g)].map((m) => m[1]))]) === j(["1500"]));
    check("table options: firstRow + firstCol + bandRow, grid style",
      s7.includes('<a:tblPr firstRow="1" firstCol="1" bandRow="1">') &&
      s7.includes("{5940675A-B579-460E-94D1-54222C63F5DA}"));
    check("merged cells: one colspan, three rowspans, covered cells flagged",
      cnt(s7, /gridSpan="2"/g) === 1 && cnt(s7, /rowSpan="2"/g) === 3 &&
      cnt(s7, /hMerge="1"/g) === 1 && cnt(s7, /vMerge="1"/g) === 3);
    check("a table-only column drops its body placeholder, text column keeps it",
      s7.includes('<p:ph idx="1"/>') && s7.includes("Заявки по юрисдикциям"));

    const c2 = await part(z, "ppt/charts/chart2.xml");
    check("column chart: clustered, two series, data labels, legend bottom, title",
      c2.includes('<c:barDir val="col"/><c:grouping val="clustered"/>') &&
      cnt(c2, /<c:ser>/g) === 2 && c2.includes('<c:showVal val="1"/>') &&
      c2.includes('<c:legendPos val="b"/>') && c2.includes("Заявки по годам"));
    check("chart values: '1 540' -> 1540, '1 700,5' -> 1700.5",
      c2.includes("<c:v>1540</c:v>") && c2.includes("<c:v>1700.5</c:v>"));
    check("chart text uses the master body typeface and the default size (table - 4)",
      c2.includes('typeface="Times New Roman"') && c2.includes('<a:defRPr sz="1300">'));

    const c3 = await part(z, "ppt/charts/chart3.xml"), c4 = await part(z, "ppt/charts/chart4.xml");
    check("stacked bar: stacked + overlap 100, first category on top, axis at the bottom",
      c3.includes('<c:barDir val="bar"/><c:grouping val="stacked"/>') &&
      c3.includes('<c:overlap val="100"/>') && c3.includes('<c:orientation val="maxMin"/>') &&
      c3.includes('<c:crosses val="max"/>'));
    check("line chart: markers, size option 12, no legend for one series",
      c4.includes("<c:lineChart>") && c4.includes('<a:defRPr sz="1200">') && !c4.includes("<c:legend>"));
    const s9 = await sx(9);
    check("chart-only columns drop both body placeholders",
      s9.includes("<p:ph ") && !s9.includes('<p:ph sz="half"'));
    const hs = [...s9.matchAll(/<a:ext cx="4038600" cy="(\d+)"\/><\/p:xfrm><a:graphic>/g)]
      .map((m) => parseInt(m[1], 10)).sort((a, b) => a - b);
    check("`height: 3.5` is 3200400 EMU; the other chart fills what is left",
      hs[0] === 3200400 && hs[1] > 3200400, j(hs));

    const s10 = await sx(10);
    const sizes = [...new Set([...s10.matchAll(/<a:rPr[^>]* sz="(\d+)"/g)].map((m) => parseInt(m[1], 10)))];
    check("long table shrinks below 17 pt but not below TABLE_MIN_PT",
      sizes.length === 1 && sizes[0] >= 1000 && sizes[0] < 1700, j(sizes));
    check("banded: off drops bandRow; layout line keeps slideLayout2",
      !s10.includes("bandRow") && (await rl(10)).includes("slideLayout2.xml"));
    check("a table-only slide has no body placeholder", !s10.includes('<p:ph idx="1"/>'));

    const book = await JSZip.loadAsync(await z.file("ppt/embeddings/Microsoft_Excel_Sheet2.xlsx").async("nodebuffer"));
    const sheet = await book.file("xl/worksheets/sheet1.xml").async("string");
    check("workbook: header row, categories, numbers",
      sheet.includes('<c r="B1" t="inlineStr"><is><t xml:space="preserve">США</t></is></c>') &&
      sheet.includes('<c r="A3" t="inlineStr"><is><t xml:space="preserve">2022</t></is></c>') &&
      sheet.includes('<c r="C4"><v>1700.5</v></c>'));

    check("parseNumber: decimal comma, thousands comma, spaces, percent",
      api.parseNumber("1,5") === 1.5 && api.parseNumber("1,234.5") === 1234.5 &&
      api.parseNumber("1 234") === 1234 && api.parseNumber("12%") === 12 && api.parseNumber("-3") === -3);
    check("tableSpans: block merges resolve to one origin",
      j(api.tableSpans([["A", "<<"], ["^^", "<<"]])) ===
      j([[[2, 2, false, false], [1, 1, true, false]], [[1, 1, false, true], [1, 1, true, true]]]));
    check("markers with nothing to join stay text",
      j(api.tableSpans([["<<", "^^"], ["a", "b"]])) ===
      j([[[1, 1, false, false], [1, 1, false, false]], [[1, 1, false, false], [1, 1, false, false]]]));
    const tOpts = api.parseOptions("{widths: 1,2; header: off; style: plain}");
    check("parseOptions reads keys and values",
      j(tOpts) === j({ widths: "1,2", header: "off", style: "plain" }));
    const tb = api.parseTable(["| a | b |", "|--|--|", "| 1 | 2 |"], tOpts);
    check("header: off drops the bold row and firstRow; plain style key kept",
      !tb.header && !tb.firstRow && tb.style === "plain" && j(tb.widths) === j([1, 2]) &&
      !tb.cells[0].some((cell) => cell.some((run) => run[1])));

    check("unknown table option is rejected",
      await throwsWith(slideDeck("{colour: red}\n| a |\n| b |"), "unknown table option"));
    check("unknown chart option is rejected",
      await throwsWith(slideDeck("{chart: bar; colour: red}\n| a | b |\n| x | 1 |"), "unknown chart option"));
    check("options line without a table is rejected",
      await throwsWith(slideDeck("{size: 12}\nplain text"), "pipe table"));
    check("widths count must match the columns",
      await throwsWith(slideDeck("{widths: 1,2,3}\n| a | b |\n| x | 1 |"), "widths"));
    check("non-numeric chart value is rejected",
      await throwsWith(slideDeck("{chart: bar}\n| a | b |\n| x | many |"), "not a number"));
    check("unknown chart type is rejected",
      await throwsWith(slideDeck("{chart: donut}\n| a | b |\n| x | 1 |"), "chart expects"));
    check("stacked line chart is rejected",
      await throwsWith(slideDeck("{chart: line; stacked: on}\n| a | b |\n| x | 1 |"), "stacked"));
    check("unknown layout name lists the template's layouts",
      await throwsWith(slideDeck("layout: Nope\ntext"), "not found in the template"));
    check("a one-body layout cannot take two columns",
      await throwsWith(slideDeck("layout: Заголовок и объект\n::: left\na\n:::\n::: right\nb\n:::"),
        "1 body placeholder"));
    check("text outside the column blocks is rejected",
      await throwsWith(slideDeck("stray\n::: left\na\n:::\n::: right\nb\n:::"), "outside"));
    check("blocks must be left then right",
      await throwsWith(slideDeck("::: right\na\n:::\n::: left\nb\n:::"), "left block"));

    // The fit estimate uses the column's own placeholder size; the preview
    // gets the flow buildZip computed.
    const many = Array.from({ length: 14 }, (_, i) =>
      "- punkt " + i + " dlinnyj tekst punkta chtoby zanyat mesto").join("\n");
    const fit = await buildFrom(slideDeck("::: left\n" + many + "\n:::\n::: right\nshort\n:::"));
    const lay = fit.built.layouts[0];
    check("layouts carry each column's box, placeholder size and flow",
      lay.length === 2 && lay[0].pt === 28 && lay[0].box[2] === 4038600 &&
      lay[0].flow.hEmu > lay[0].box[3] && lay[1].flow.hEmu <= lay[1].box[3], j(lay.map((l) => [l.pt, l.flow.hEmu])));
    const specs = api.slideSpecs(fit.d, fit.built.bodyPt, fit.built.layouts);
    check("slideSpecs: content spec has two columns with their flows",
      specs[1].kind === "content" && specs[1].columns.length === 2 && !!specs[1].columns[0].flow);
  }


  // English template: ids derived from the template, plan title, template
  // body size, list styles, sentinel-index layouts, paragraph spacing, text
  // colour, and byte parity with the Python build.
  {
    const cnt = (x, re) => (x.match(re) || []).length;
    const enText = fs.readFileSync(path.join(HERE, "test_structure_en.md"), "utf8");
    const enDoc = api.parseMd(enText);
    const enBuilt = await api.buildZip(tplEnBuf, enDoc);
    const ez = enBuilt.zip;
    fs.writeFileSync(path.join(outDir, "en_deck_js.pptx"), await ez.generateAsync({ type: "nodebuffer" }));
    const pres = await part(ez, "ppt/presentation.xml");
    const rels = await part(ez, "ppt/_rels/presentation.xml.rels");
    const rids = rels.match(/\bId="rId[0-9]+"/g);
    check("clones take free relationship ids (the template already uses rId11)",
      rids.length === new Set(rids).size, j(rids));
    const ids = [...pres.matchAll(/<p:sldId id="([0-9]+)"/g)].map((m) => m[1]);
    check("the template's own slide ids are kept; clones continue after the highest",
      j(ids.slice(0, 3)) === j(["400", "385", "386"]) && ids[ids.length - 1] === "346" &&
      j(ids.slice(3, -1)) === j(Array.from({ length: ids.length - 4 }, (_, i) => String(401 + i))), j(ids));
    const xe = (n) => part(ez, "ppt/slides/slide" + n + ".xml");
    check("plan title is the template's own (no Russian default)",
      (await xe(2)).includes("Agenda") && !(await xe(2)).includes("План"));
    check("a template hint in parentheses is dropped from the plan title",
      api.planTitleText('<p:sp><p:ph type="title"/><a:t>Contents (if any)</a:t></p:sp>') === "Contents" &&
      api.planTitleText(await part(await JSZip.loadAsync(tplBuf), "ppt/slides/slide2.xml")) === "План");
    const overDoc = api.parseMd(enText.replace("# plan\n", "# plan\n## Contents\n"));
    check("`## <title>` in # plan overrides the title",
      overDoc.planTitle === "Contents" && overDoc.plan[0][0] === "Origins");
    check("body_size defaults to the template's size (14 pt): plan 13 pt, tables 13 pt",
      enDoc.bodySize === 14 && enDoc.TOC === 13 && enDoc.tablePt === 13 &&
      (await xe(2)).includes('sz="1300"') && (await xe(7)).includes('sz="1300"'));
    check("an explicit body_size overrides the template's size",
      api.parseMd(enText.replace("lang: en", "lang: en\nbody_size: 20")).bodySize === 20);
    const s3e = await xe(3);
    check("`- ` items get a bullet and a hanging indent the template lacks",
      s3e.includes('<a:pPr lvl="1" marL="685800" indent="-228600"><a:buFont typeface="Arial"/>' +
        '<a:buChar char="\u2022"/></a:pPr>'));
    check("numbered items hang off their number",
      s3e.includes('<a:pPr lvl="1" marL="685800" indent="-228600"><a:buFontTx/>' +
        '<a:buAutoNum type="arabicPeriod"/></a:pPr>'));
    check("plain lines carry no bullet override", s3e.includes('<a:pPr lvl="0"/>'));
    check("plan numbers hang off their number",
      (await xe(2)).includes('<a:pPr lvl="0" marL="228600" indent="-228600"><a:buFontTx/>'));
    const ru3 = await part(z, "ppt/slides/slide3.xml");
    check("MFUA's own bullets and indents are left alone (no stamped bullet)",
      !ru3.includes("<a:buChar") && !ru3.includes("marL="));

    const geoEn = await api.loadGeometryParts(ez);
    const blk = api.layoutCatalog(geoEn).filter((c) => c[1] === "Title, Content 1 (Max)")[0][3][0][3];
    const spc = api.effectiveSpacing(blk, geoEn["ppt/slideMasters/slideMaster1.xml"], 14);
    check("effective spacing: 120 % lines, 10 pt before level 0, 5 pt before level 1",
      j(spc) === j({ 0: [120, 127000, 0], 1: [120, 63500, 0] }), j(spc));
    const geoRu = await api.loadGeometryParts(z);
    check("a template with no spacing keeps the plain estimate (MFUA)",
      j(api.effectiveSpacing(null, geoRu["ppt/slideMasters/slideMaster1.xml"], 18)) ===
      j({ 0: [100, 0, 0], 1: [100, 0, 0] }));
    const items = [0, 1, 2, 3].map(() => [0, [["a line", false, false, false]], null]);
    check("spacing makes the text estimate taller",
      api.columnFlow(items, [0, 0, 8229600, 4000000], 14, 13, spc).hEmu >
      api.columnFlow(items, [0, 0, 8229600, 4000000], 14, 13).hEmu);

    const s4e = await xe(5);
    check("two columns on a layout whose placeholders share the sentinel idx",
      (await part(ez, "ppt/slides/_rels/slide5.xml.rels")).includes("slideLayout43.xml"));
    check("each sentinel placeholder gets its own unique idx",
      s4e.includes('<p:ph type="body" idx="4294967295"/>') && !s4e.includes("4294967294") &&
      cnt(s4e, /<p:ph /g) >= 2);
    check("the sentinel placeholder carries its explicit box and list style",
      s4e.includes('<a:off x="838203" y="2147459"/><a:ext cx="5181603" cy="3878628"/>') &&
      s4e.includes('<a:lvl1pPr marL="0" indent="0">'));
    check("an explicit `layout:` on a sentinel layout takes its box and style",
      (await part(ez, "ppt/slides/_rels/slide8.xml.rels")).includes("slideLayout15.xml") &&
      (await xe(8)).includes('<a:off x="5282415" y="2386584"/>'));
    const two = api.layoutCatalog(geoEn)
      .filter((c) => api.isTwoContent(c, api.masterBodyBox(geoEn))).map((c) => c[1]);
    check("only the genuine two-column layout is detected", j(two) === j(["Two Content"]), j(two));

    const ce = await part(ez, "ppt/charts/chart1.xml");
    check("chart text takes the master's body colour and typeface",
      ce.includes('<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:latin typeface="Poppins"/>'));
    check("cells of a fill-less table style carry the body text colour",
      cnt(await xe(7), /<a:solidFill><a:srgbClr val="FFFFFF"\/><\/a:solidFill>/g) >= 6);
    check("the default table style keeps its own cell colours",
      !(await xe(9)).split("<a:tbl>")[1].includes("<a:solidFill>"));

    check("the preview takes the slide size from the template (16:9 here, 4:3 for MFUA)",
      j(enBuilt.slideSize) === j([12192000, 6858000]) &&
      j((await api.buildZip(tplBuf, api.parseMd(structText))).slideSize) === j([9144000, 6858000]));
    api.setSlideSize('<p:sldSz cx="12192000" cy="6858000" type="custom"/>');
    check("setSlideSize reads cx and cy from <p:sldSz>",
      j(api.getSlideSize()) === j([12192000, 6858000]));

    const enPy = path.join(outDir, "en_deck_py.pptx");
    if (fs.existsSync(enPy)) {
      const mism4 = await partsDiff(ez, await load(enPy));
      check("byte-identical parts vs Python English build", mism4.length === 0, mism4.join(", "));
    }
  }

  let threw = false;
  try {
    api.parseMd(structText.replace("body_size: 18",
      "body_size: 18\nlang: 123abc"));
  } catch (e) {
    threw = /invalid lang/.test(e.message);
  }
  check("invalid lang raises", threw);

  if (failures.length) {
    console.log("FAILED: " + failures.length + " check(s)");
    process.exit(1);
  }
  console.log("ALL JS TESTS PASSED");
})().catch((e) => {
  console.error("ERROR: " + (e && e.stack || e));
  process.exit(1);
});
