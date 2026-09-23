# -*- coding: utf-8 -*-
"""Python engine tests for md2pptx.

Run from the repo root (A:\\a\\md2pptx):
    python Tests/run_tests.py

Builds a deck from Tests/test_structure.md against MFUA-Template.pptx and
checks:
  1. RU anchor: slide1 topic = "На тему: <topic>" (format preserved), bold,
     font kept (lang + Times New Roman typeface of the anchor run).
  2. EN anchor: on a "Topic"-anchored template copy, slide1 topic is the
     bare topic (the whole line is replaced).
  3. Missing anchor: build raises SystemExit.
  4. Numbered body lines -> level-1 paragraphs, number kept in the text.
  5. Font rules: no sz on body runs, plan items stamped at body_size - 1,
     closing title stamped at the template-declared size (36 pt -> 3600).
  6. Structural validation PASSED.

Writes Tests/out/ru_deck_py.pptx and Tests/out/lang_deck_py.pptx (the
parity inputs for run_tests_js.js) and the template variants under
Tests/out/.
"""
import importlib.util
import contextlib
import io
import os
import re
import sys
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(HERE)
OUT = os.path.join(HERE, "out")
TEMPLATE = os.path.join(REPO, "MFUA-Template.pptx")
STRUCT = os.path.join(HERE, "test_structure.md")
ENGINE = os.path.join(REPO, "md2pptx", "engine.py")

spec = importlib.util.spec_from_file_location("engine", ENGINE)
engine = importlib.util.module_from_spec(spec)
spec.loader.exec_module(engine)

os.makedirs(OUT, exist_ok=True)
ru_out = os.path.join(OUT, "ru_deck_py.pptx")
en_tpl = os.path.join(OUT, "en_template.pptx")
en_out = os.path.join(OUT, "en_deck_py.pptx")
na_tpl = os.path.join(OUT, "noanchor_template.pptx")
na_out = os.path.join(OUT, "noanchor_out.pptx")

with open(STRUCT, "r", encoding="utf-8") as f:
    struct_text = f.read()
TOPIC = engine.parse_md(struct_text)["title"]

fails = []


def check(name, cond, detail=""):
    print("  %s  %s%s" % ("PASS" if cond else "FAIL", name,
                          "" if cond else " -> " + detail))
    if not cond:
        fails.append(name)


def variant_template(src, dst, sub_old, sub_new):
    with zipfile.ZipFile(src) as zin, \
            zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "ppt/slides/slide1.xml":
                data = data.decode("utf-8").replace(sub_old, sub_new)
                data = data.encode("utf-8")
            zout.writestr(item, data)


def texts(xml):
    return re.findall(r"<a:t>(.*?)</a:t>", xml)


def run_of(xml, text):
    i = xml.index(text)
    s = xml.rfind("<a:r>", 0, i)
    e = xml.index("</a:r>", i)
    return xml[s:e + 6]


def region(xml, marker):
    i = xml.index(marker)
    return xml[i:xml.find("</p:sp>", i) + 7]


def stamped_lang(xml, wanted="ru-RU"):
    """Bad tags among rPr/endParaRPr in `xml`: missing/wrong lang or dirty."""
    bad = []
    for m in re.finditer(r"<a:(?:rPr|endParaRPr)\b[^>]*>", xml):
        tag = m.group(0)
        lang = re.search(r'\blang="([^"]+)"', tag)
        dirty = re.search(r'\bdirty="([^"]+)"', tag)
        if (not lang or lang.group(1) != wanted
                or not dirty or dirty.group(1) != "0"):
            bad.append(tag)
    return bad


print("[1] RU template: build + checks")
doc, ncontent, body_pt = engine.build(struct_text, TEMPLATE, ru_out)
probs = engine.validate(ru_out)
check("validation PASSED", not probs, "; ".join(probs))

z = zipfile.ZipFile(ru_out)
s1 = z.read("ppt/slides/slide1.xml").decode("utf-8")
t1 = texts(s1)
check("slide1 topic = 'На тему: ' + topic", "На тему: " + TOPIC in t1,
      repr(t1))
tr = run_of(s1, "На тему: " + TOPIC)
check("topic run is bold", 'b="1"' in tr, tr)
check("topic typeface kept (Times New Roman)",
      'typeface="Times New Roman"' in tr, tr)

s2 = z.read("ppt/slides/slide2.xml").decode("utf-8")
t2 = texts(s2)
check("plan items keep the numbers",
      t2[1].startswith("1. ") and t2[4].startswith("4. "), repr(t2))
check("plan items stamped at body_size - 1",
      region(s2, '<p:ph idx="1"/>').count('sz="1700"') == 4)

s3 = z.read("ppt/slides/slide3.xml").decode("utf-8")
t3 = texts(s3)
check("numbered body lines keep the number",
      "1. Англия закрепляет монопольный характер патента" in t3, repr(t3))
i = s3.index("Англия закрепляет")
check("numbered body line is level 1",
      'lvl="1"' in s3[s3.rfind("<a:p>", 0, i):i])
check("body runs carry no sz", 'sz="' not in region(s3, '<p:ph idx="1"/>'))

s4 = z.read("ppt/slides/slide4.xml").decode("utf-8")
check("closing title stamped at template size (3600)",
      'sz="3600"' in region(s4, '<p:ph type="title"/>'))

print("[2] EN 'Topic' template: whole line replaced")
variant_template(TEMPLATE, en_tpl, "На тему", "Topic")
engine.build(struct_text, en_tpl, en_out)
z = zipfile.ZipFile(en_out)
s1 = z.read("ppt/slides/slide1.xml").decode("utf-8")
t1 = texts(s1)
check("slide1 topic = bare topic", TOPIC in t1, repr(t1))

print("[3] Missing anchor: build must fail")
variant_template(TEMPLATE, na_tpl, "На тему", "Заголовок")
try:
    engine.build(struct_text, na_tpl, na_out)
    check("missing anchor raises SystemExit", False, "no exception raised")
except SystemExit as e:
    check("missing anchor raises SystemExit", "topic anchor" in str(e), str(e))

print("[4] lang front matter: stamped on every generated run + end mark")
lang_text = struct_text.replace("---\nbody_size: 18\n---",
                               "---\nbody_size: 18\nlang: ru-RU\n---")
lang_out = os.path.join(OUT, "lang_deck_py.pptx")
engine.build(lang_text, TEMPLATE, lang_out)
probs = engine.validate(lang_out)
check("lang deck validation PASSED", not probs, "; ".join(probs))

z = zipfile.ZipFile(lang_out)
bad = []
s1 = z.read("ppt/slides/slide1.xml").decode("utf-8")
bad += stamped_lang(region(s1, "На тему: " + TOPIC))
s2 = z.read("ppt/slides/slide2.xml").decode("utf-8")
bad += stamped_lang(region(s2, '<p:ph type="title"/>'))
bad += stamped_lang(region(s2, '<p:ph idx="1"/>'))
for i in (3, 4):
    x = z.read("ppt/slides/slide%d.xml" % i).decode("utf-8")
    bad += stamped_lang(region(x, '<p:ph type="title"/>'))
    bad += stamped_lang(region(x, '<p:ph idx="1"/>'))
check("all generated rPr/endParaRPr carry lang=ru-RU dirty=0",
      not bad, "; ".join(bad[:4]))
z.close()

# without the key the generated body runs stay bare (template default)
z0 = zipfile.ZipFile(ru_out)
s3 = z0.read("ppt/slides/slide3.xml").decode("utf-8")
body_region = region(s3, '<p:ph idx="1"/>')
check("no lang: generated body runs stay bare (template default)",
      'lang="' not in body_region, body_region[:160])
z0.close()

# an invalid BCP-47 tag is rejected before any build
bad_text = struct_text.replace("body_size: 18", "body_size: 18\nlang: 123abc")
try:
    engine.build(bad_text, TEMPLATE, os.path.join(OUT, "bad_lang.pptx"))
    check("invalid lang raises SystemExit", False, "no exception raised")
except SystemExit as e:
    check("invalid lang raises SystemExit", "invalid lang" in str(e), str(e))

# the CLI report shows the lang line
lang_struct = os.path.join(OUT, "lang_struct.md")
with open(lang_struct, "w", encoding="utf-8") as f:
    f.write(lang_text)
old_argv = sys.argv
sys.argv = ["engine.py", lang_struct, TEMPLATE, lang_out]
buf = io.StringIO()
try:
    with contextlib.redirect_stdout(buf):
        rc = engine.main()
finally:
    sys.argv = old_argv
check("main() reports 'lang: ru-RU (stamped, dirty=0)'",
      rc == 0 and "  lang: ru-RU (stamped, dirty=0)" in buf.getvalue(),
      buf.getvalue().replace("\n", " | "))

if fails:
    print("FAILED: %d check(s)" % len(fails))
    sys.exit(1)
print("ALL PYTHON TESTS PASSED")
