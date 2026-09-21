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
//   - byte-identical part-by-part parity vs Tests/out/ru_deck_py.pptx
//
// Requires: Node, and jszip 3.x at the path below (local copy).

const fs = require("fs");
const path = require("path");

const HERE = __dirname;
const REPO = path.resolve(HERE, "..");
const JSZIP_PATH =
  "C:/Users/Ren/AppData/Roaming/npm/node_modules/openclaw/node_modules/jszip";
const JSZip = require(JSZIP_PATH);

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

  if (failures.length) {
    console.log("FAILED: " + failures.length + " check(s)");
    process.exit(1);
  }
  console.log("ALL JS TESTS PASSED");
})().catch((e) => {
  console.error("ERROR: " + (e && e.stack || e));
  process.exit(1);
});
