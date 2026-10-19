// JS engine tests for md2pptx (the engine core in webapp/index.html).
//
// Run from the repo root (A:\a\md2pptx) AFTER run_tests.py:
//     node Tests/run_tests_js.js
//
// Extracts the engine core from webapp/index.html (everything up to the
// Mode B UI marker), builds the same deck from Tests/test_structure.md +
// MFUA-Template.pptx as run_tests.py, and checks:
//   - slide1 topic = "На тему: <topic>" (format preserved), bold, font kept
//   - numbered body lines -> level 1, number kept in the text
//   - plan item sizes, closing title size, no sz on body runs
//   - inline marks: *italic* -> i="1", ++underline++ -> u="sng", **bold** ->
//     b="1", and no marker characters left in the text
//   - pipe table -> p:graphicFrame with a:tbl, one tr per row, one td per cell,
//     head row bold, cells stamped at body_size - 1, separator alignment, frame
//     placed after the body placeholder inside the shape tree
//   - byte-identical part-by-part parity vs Tests/out/ru_deck_py.pptx
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

(async () => {
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
  check("plan items keep the numbers",
    t2[1].startsWith("1. ") && t2[4].startsWith("4. "), JSON.stringify(t2));
  check("plan items stamped at body_size - 1",
    (region(s2, '<p:ph idx="1"/>').match(/sz="1700"/g) || []).length === 4);

  const s3 = await part(z, "ppt/slides/slide3.xml");
  check("numbered body line keeps the number",
    s3.includes("1. Англия закрепляет монопольный характер патента"));
  const i3 = s3.indexOf("Англия закрепляет");
  check("numbered body line is level 1",
    s3.slice(s3.lastIndexOf("<a:p>", i3), i3).includes('lvl="1"'));
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
  const mism = [];
  for (const name of Object.keys(z.files)) {
    if (z.files[name].dir) continue;
    const a = await z.file(name).async("nodebuffer");
    const b = py.file(name);
    if (!b) { mism.push("missing in py: " + name); continue; }
    const bb = await b.async("nodebuffer");
    if (!a.equals(bb)) mism.push(name);
  }
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
      const mism2 = [];
      for (const name of Object.keys(lz.files)) {
        if (lz.files[name].dir) continue;
        const a = await lz.file(name).async("nodebuffer");
        const b = lp.file(name);
        if (!b) { mism2.push("missing in py: " + name); continue; }
        const bb = await b.async("nodebuffer");
        if (!a.equals(bb)) mism2.push(name);
      }
      check("byte-identical parts vs Python lang build", mism2.length === 0,
        mism2.join(", "));
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
