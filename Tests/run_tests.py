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
  4. Numbered body lines and plan items are auto-numbered paragraphs
     (<a:buAutoNum>), the number not stored in the text.
  5. Font rules: no sz on body runs, plan items stamped at body_size - 1,
     closing title stamped at the template-declared size (36 pt -> 3600).
  6. Structural validation PASSED (parts, relationships, content types).
  7. Two-column layouts, table options, merged cells and charts (the slides
     appended to test_structure.md), their error messages and the fit check.
  8. The English template (English-Template.pptx, test_structure_en.md): ids
     derived from the template, plan title, template body size, list styles,
     sentinel-index layouts, paragraph spacing, text colour, duplicate-id
     validation.

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
TEMPLATE_EN = os.path.join(REPO, "English-Template.pptx")
STRUCT_EN = os.path.join(HERE, "test_structure_en.md")
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
check("plan items carry no typed numbers",
      t2[1:5] == list(doc["plan"][k][0] for k in range(4))
      and not any(re.match(r"^[0-9]+[.)] ", t) for t in t2[1:5]), repr(t2))
check("plan items are auto-numbered (arabicPeriod, no startAt)",
      region(s2, '<p:ph idx="1"/>').count('<a:buAutoNum type="arabicPeriod"/>') == 4)
check("plan items stamped at body_size - 1",
      region(s2, '<p:ph idx="1"/>').count('sz="1700"') == 4)

s3 = z.read("ppt/slides/slide3.xml").decode("utf-8")
t3 = texts(s3)
check("numbered body line text carries no typed number",
      "Англия закрепляет монопольный характер патента" in t3
      and not any(re.match(r"^[0-9]+[.)] ", t) for t in t3), repr(t3))
i = s3.index("Англия закрепляет")
numbered_ppr = s3[s3.rfind("<a:p>", 0, i):i]
check("numbered body line is level 1 and auto-numbered",
      'lvl="1"' in numbered_ppr
      and '<a:buFontTx/><a:buAutoNum type="arabicPeriod"/>' in numbered_ppr,
      numbered_ppr)
j = s3.index("Венеция")
check("bullet line is not auto-numbered",
      "buAutoNum" not in s3[s3.rfind("<a:p>", 0, j):j])
check("body runs carry no sz", 'sz="' not in region(s3, '<p:ph idx="1"/>'))

s4 = z.read("ppt/slides/slide4.xml").decode("utf-8")
check("closing title stamped at template size (3600)",
      'sz="3600"' in region(s4, '<p:ph type="title"/>'))

print("[4b] inline marks (*italic*, ++underline++) and pipe tables")
flat = " ".join(t3)
check("italic run stamped i=\"1\"", 'i="1"' in run_of(s3, "italic"),
      run_of(s3, "italic"))
check("underlined run stamped u=\"sng\"", 'u="sng"' in run_of(s3, "underlined"),
      run_of(s3, "underlined"))
check("bold still stamped alongside", 'b="1"' in run_of(s3, "Италии"),
      run_of(s3, "Италии"))
check("no stray markers in the text",
      "*" not in flat and "++" not in flat, repr(flat))
s5 = z.read("ppt/slides/slide5.xml").decode("utf-8")
t5 = texts(s5)
check("table cells present", all(x in t5 for x in ("Cell A", "Cell B", "Cell C", "6")),
      repr(t5))
check("table is a p:graphicFrame", "<p:graphicFrame>" in s5, s5[:200])
check("a:tbl with one row per line", s5.count("<a:tr ") == 3, s5)
check("one a:tc per cell (3 rows x 3 cols)", s5.count("<a:tc>") == 9, s5)
check("head row cells inherit the table's bold",
      'b="1"' in run_of(s5, "Cell A"), run_of(s5, "Cell A"))
check("table cells stamped at body_size - 1",
      s5.count('sz="1700"') == 9, s5)
check("separator alignment applied",
      'algn="ctr"' in s5 and 'algn="r"' in s5, s5)
check("frame comes after the body placeholder",
      s5.index("<p:graphicFrame>") > s5.index('<p:ph idx="1"/>'), s5[:200])
check("frame inside the shape tree",
      s5.index("<p:graphicFrame>") < s5.rindex("</p:spTree>"), s5[-200:])

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

print("[5] numbered lists: runs, startAt, schemes, tables, plan")
num_path = os.path.join(HERE, "test_numbering.md")
with open(num_path, "r", encoding="utf-8") as f:
    num_text = f.read()
num_out = os.path.join(OUT, "numbering_deck_py.pptx")
num_doc, _n, _b = engine.build(num_text, TEMPLATE, num_out)
probs = engine.validate(num_out)
check("numbering deck validation PASSED", not probs, "; ".join(probs))

P, R = "arabicPeriod", "arabicParenR"
marks = [it[2] for it in num_doc["slides"][0][1][0]]
check("runs: a bullet, a plain line or a restart ends a run; first number is the start",
      marks == [(P, 1), (P, 1), None, (P, 1), None, (P, 5), (P, 5)], repr(marks))
marks = [it[2] for it in num_doc["slides"][1][1][0] if it[0] in (0, 1)]
check("a scheme change starts a new run; a table does not split a run",
      marks == [(R, 1), (R, 1), (P, 1), (P, 1)], repr(marks))
check("plan run keeps its written start",
      [m for _t, m in num_doc["plan"]] == [(P, 3), (P, 3)],
      repr(num_doc["plan"]))
check("1) is the arabicParenR scheme, 1. the arabicPeriod scheme",
      engine.parse_numbered("1) a")[0] == R
      and engine.parse_numbered("1. a")[0] == P)
check("not list items: no space after the marker, or a decimal",
      engine.parse_numbered("1.5 million") is None
      and engine.parse_numbered("2026.") is None
      and engine.parse_numbered("- 1. a") is None)

zn = zipfile.ZipFile(num_out)
n2 = zn.read("ppt/slides/slide2.xml").decode("utf-8")
n3 = zn.read("ppt/slides/slide3.xml").decode("utf-8")
n5 = zn.read("ppt/slides/slide5.xml").decode("utf-8")
check("plan items stamped startAt=3 on every item of the run",
      region(n2, '<p:ph idx="1"/>').count(
          '<a:buAutoNum type="arabicPeriod" startAt="3"/>') == 2)
check("startAt is stamped only on the run that starts at 5",
      n3.count('startAt="5"') == 2 and n3.count("startAt=") == 2, n3.count("startAt="))
check("runs starting at 1 carry no startAt", "startAt" not in n5)
check("slide 2 body: two arabicParenR and two arabicPeriod paragraphs",
      n5.count('type="arabicParenR"') == 2 and n5.count('type="arabicPeriod"') == 2)
k = n3.index("A bullet breaks")
check("a bullet inside the list stays a plain level-1 paragraph",
      n3[n3.rfind("<a:p>", 0, k):k].startswith('<a:p><a:pPr lvl="1"/>'))
check("every numbered paragraph has buFontTx before buAutoNum (schema order)",
      n3.count("<a:buFontTx/><a:buAutoNum") == n3.count("<a:buAutoNum"))
zn.close()

print("[6] layouts, table options, merged cells, charts (test_structure.md)")


def build_from(text, name):
    out_path = os.path.join(OUT, name)
    with contextlib.redirect_stdout(io.StringIO()):
        doc_, _n, _b = engine.build(text, TEMPLATE, out_path)
    return doc_, zipfile.ZipFile(out_path)


def fails_with(text, needle):
    """True when building `text` stops with a SystemExit mentioning `needle`."""
    try:
        engine.parse_md(text)
        build_from(text, "scratch.pptx")
    except SystemExit as e:
        return needle in str(e)
    return False


def slide_deck(body):
    return "# title\nT\n\n# slide\n## S\n" + body + "\n"


zf = zipfile.ZipFile(ru_out)
rels = lambda n: zf.read("ppt/slides/_rels/slide%d.xml.rels" % n).decode("utf-8")
sx = lambda n: zf.read("ppt/slides/slide%d.xml" % n).decode("utf-8")

# slide6: two columns (text + pie) on the two-content layout
check("two columns point their slide at the two-content layout (slideLayout4)",
      "slideLayout4.xml" in rels(6) and "slideLayout2.xml" not in rels(6), rels(6))
check("left column fills the layout's idx 1 placeholder",
      '<p:ph sz="half" idx="1"/>' in sx(6) and "Монополия короны" in sx(6))
check("a column without text loses its body placeholder (no empty txBody)",
      '<p:ph sz="half" idx="2"/>' not in sx(6))
check("the chart sits in the right half of the slide",
      re.search(r'<p:xfrm><a:off x="4648200" y="[0-9]+"/><a:ext cx="4038600"', sx(6)))
check("chart part, relationship and workbook exist",
      'Target="../charts/chart1.xml"' in rels(6)
      and "ppt/charts/chart1.xml" in zf.namelist()
      and "ppt/embeddings/Microsoft_Excel_Sheet1.xlsx" in zf.namelist())
ct = zf.read("[Content_Types].xml").decode("utf-8")
check("content types: chart override and xlsx default registered once",
      ct.count("/ppt/charts/chart1.xml") == 1 and ct.count('Extension="xlsx"') == 1)
c1 = zf.read("ppt/charts/chart1.xml").decode("utf-8")
check("pie: first series only, per-point theme colours, percent labels, legend right",
      c1.count("<c:ser>") == 1 and c1.count("<c:dPt>") == 3
      and '<c:showPercent val="1"/>' in c1 and '<c:legendPos val="r"/>' in c1)

# slide7: widths / size / first-col / style + merged cells
s7 = sx(7)
check("table options: widths 3,2,2,2 of the full body width",
      re.findall(r'<a:gridCol w="(\d+)"/>', s7) == ["2743200", "1828800", "1828800", "1828800"])
check("table options: size 15 stamped on every cell run",
      set(re.findall(r'<a:rPr[^>]* sz="(\d+)"', s7)) == {"1500"})
check("table options: firstRow + firstCol + bandRow, grid style",
      '<a:tblPr firstRow="1" firstCol="1" bandRow="1">' in s7
      and "{5940675A-B579-460E-94D1-54222C63F5DA}" in s7)
check("merged cells: one colspan, three rowspans, covered cells flagged",
      s7.count('gridSpan="2"') == 1 and s7.count('rowSpan="2"') == 3
      and s7.count('hMerge="1"') == 1 and s7.count('vMerge="1"') == 3,
      "%d %d %d %d" % (s7.count('gridSpan="2"'), s7.count('rowSpan="2"'),
                       s7.count('hMerge="1"'), s7.count('vMerge="1"')))
check("a table-only column drops its body placeholder, text column keeps it",
      '<p:ph idx="1"/>' in s7 and "Заявки по юрисдикциям" in s7)

# slide8: clustered column chart; numbers with spaces and decimal comma
c2 = zf.read("ppt/charts/chart2.xml").decode("utf-8")
check("column chart: clustered, two series, data labels, legend bottom, title",
      '<c:barDir val="col"/><c:grouping val="clustered"/>' in c2
      and c2.count("<c:ser>") == 2 and '<c:showVal val="1"/>' in c2
      and '<c:legendPos val="b"/>' in c2 and "Заявки по годам" in c2)
check("chart values: '1 540' -> 1540, '1 700,5' -> 1700.5",
      "<c:v>1540</c:v>" in c2 and "<c:v>1700.5</c:v>" in c2)
check("chart text uses the master body typeface and the default size (table - 4)",
      'typeface="Times New Roman"' in c2 and '<a:defRPr sz="1300">' in c2)

# slide9: two charts side by side; both placeholders dropped
c3 = zf.read("ppt/charts/chart3.xml").decode("utf-8")
c4 = zf.read("ppt/charts/chart4.xml").decode("utf-8")
check("stacked bar: stacked + overlap 100, first category on top, axis at the bottom",
      '<c:barDir val="bar"/><c:grouping val="stacked"/>' in c3
      and '<c:overlap val="100"/>' in c3 and '<c:orientation val="maxMin"/>' in c3
      and '<c:crosses val="max"/>' in c3)
check("line chart: markers, size option 12, no legend for one series",
      "<c:lineChart>" in c4 and '<a:defRPr sz="1200">' in c4
      and "<c:legend>" not in c4)
check("chart-only columns drop both body placeholders",
      "<p:ph " in sx(9) and '<p:ph sz="half"' not in sx(9))
h = [int(v) for v in re.findall(r'<a:ext cx="4038600" cy="(\d+)"/></p:xfrm><a:graphic>', sx(9))]
check("`height: 3.5` is 3200400 EMU; the other chart fills what is left",
      sorted(h)[0] == 3200400 and sorted(h)[1] > 3200400, repr(h))

# slide10: explicit layout line, auto-shrink, banded off
s10 = sx(10)
sizes = {int(v) for v in re.findall(r'<a:rPr[^>]* sz="(\d+)"', s10)}
check("long table shrinks below 17 pt but not below TABLE_MIN_PT",
      len(sizes) == 1 and 1000 <= min(sizes) < 1700, repr(sizes))
check("banded: off drops bandRow; layout line keeps slideLayout2",
      'bandRow' not in s10 and "slideLayout2.xml" in rels(10))
check("a table-only slide has no body placeholder",
      '<p:ph idx="1"/>' not in s10)
check("fit check passes for the feature deck", engine.fit_check(doc) == [],
      repr(engine.fit_check(doc)))

# workbook behind a chart
wb = zipfile.ZipFile(io.BytesIO(zf.read("ppt/embeddings/Microsoft_Excel_Sheet2.xlsx")))
sheet = wb.read("xl/worksheets/sheet1.xml").decode("utf-8")
check("workbook: header row, categories, numbers",
      '<c r="B1" t="inlineStr"><is><t xml:space="preserve">США</t></is></c>' in sheet
      and '<c r="A3" t="inlineStr"><is><t xml:space="preserve">2022</t></is></c>' in sheet
      and '<c r="C4"><v>1700.5</v></c>' in sheet)
check("workbook bytes depend only on the data",
      engine.workbook_bytes(engine.parse_md(slide_deck(
          "{chart: pie}\n| a | b |\n|--|--|\n| x | 1 |"))["slides"][0][1][0][0][1])
      == engine.workbook_bytes(engine.parse_md(slide_deck(
          "{chart: pie}\n| a | b |\n|--|--|\n| x | 1 |"))["slides"][0][1][0][0][1]))

# parsing units
check("parse_number: decimal comma, thousands comma, spaces, percent",
      engine.parse_number("1,5") == 1.5 and engine.parse_number("1,234.5") == 1234.5
      and engine.parse_number("1 234") == 1234 and engine.parse_number("12%") == 12
      and engine.parse_number("-3") == -3)
check("table_spans: block merges resolve to one origin",
      engine.table_spans([["A", "<<"], ["^^", "<<"]])
      == [[(2, 2, False, False), (1, 1, True, False)],
          [(1, 1, False, True), (1, 1, True, True)]])
check("markers with nothing to join stay text",
      engine.table_spans([["<<", "^^"], ["a", "b"]])
      == [[(1, 1, False, False)] * 2, [(1, 1, False, False)] * 2])
t_opts = engine.parse_options("{widths: 1,2; header: off; style: plain}")
check("parse_options reads keys and values",
      t_opts == {"widths": "1,2", "header": "off", "style": "plain"})
tb = engine.parse_table(["| a | b |", "|--|--|", "| 1 | 2 |"], t_opts)
check("header: off drops the bold row and firstRow; plain style key kept",
      not tb["header"] and not tb["first_row"] and tb["style"] == "plain"
      and tb["widths"] == [1, 2]
      and not any(run[1] for cell in tb["cells"][0] for run in cell))

# errors are explicit
check("unknown table option is rejected",
      fails_with(slide_deck("{colour: red}\n| a |\n| b |"), "unknown table option"))
check("unknown chart option is rejected",
      fails_with(slide_deck("{chart: bar; colour: red}\n| a | b |\n| x | 1 |"),
                 "unknown chart option"))
check("options line without a table is rejected",
      fails_with(slide_deck("{size: 12}\nplain text"), "pipe table"))
check("widths count must match the columns",
      fails_with(slide_deck("{widths: 1,2,3}\n| a | b |\n| x | 1 |"), "widths"))
check("non-numeric chart value is rejected",
      fails_with(slide_deck("{chart: bar}\n| a | b |\n| x | many |"), "not a number"))
check("unknown chart type is rejected",
      fails_with(slide_deck("{chart: donut}\n| a | b |\n| x | 1 |"), "chart expects"))
check("stacked line chart is rejected",
      fails_with(slide_deck("{chart: line; stacked: on}\n| a | b |\n| x | 1 |"),
                 "stacked"))
check("unknown layout name lists the template's layouts",
      fails_with(slide_deck("layout: Nope\ntext"), "not found in the template"))
check("a one-body layout cannot take two columns",
      fails_with(slide_deck("layout: Заголовок и объект\n::: left\na\n:::\n::: right\nb\n:::"),
                 "1 body placeholder"))
check("text outside the column blocks is rejected",
      fails_with(slide_deck("stray\n::: left\na\n:::\n::: right\nb\n:::"),
                 "outside"))
check("blocks must be left then right",
      fails_with(slide_deck("::: right\na\n:::\n::: left\nb\n:::"), "left block"))

# fit estimate uses the column's own placeholder size and names the column
many = "\n".join("- punkt %d dlinnyj tekst punkta chtoby zanyat mesto" % i for i in range(14))
doc_fit, _z = build_from(slide_deck("::: left\n" + many + "\n:::\n::: right\nshort\n:::"),
                        "fit_deck.pptx")
over = engine.fit_check(doc_fit)
check("overflowing left column is reported by name",
      len(over) == 1 and over[0][0] == "S (left)", repr(over))

print("[7] English template: ids, plan title, body size, list styles, sentinel layouts")
with open(STRUCT_EN, "r", encoding="utf-8") as f:
    en_text = f.read()
en_out = os.path.join(OUT, "en_deck_py.pptx")
with contextlib.redirect_stdout(io.StringIO()):
    en_doc, _n, en_body = engine.build(en_text, TEMPLATE_EN, en_out)
check("English deck validation PASSED (parts, relationships, duplicate ids)",
      engine.validate(en_out) == [], "; ".join(engine.validate(en_out)))
ze = zipfile.ZipFile(en_out)
pres_e = ze.read("ppt/presentation.xml").decode("utf-8")
rels_e = ze.read("ppt/_rels/presentation.xml.rels").decode("utf-8")
rids = re.findall(r'\bId="(rId[0-9]+)"', rels_e)
check("clones take free relationship ids (the template already uses rId11)",
      len(rids) == len(set(rids)), repr(sorted(rids)))
ids = re.findall(r'<p:sldId id="([0-9]+)"', pres_e)
check("the template's own slide ids are kept; clones continue after the highest",
      ids[:3] == ["400", "385", "386"] and ids[-1] == "346"
      and ids[3:-1] == [str(i) for i in range(401, 401 + len(ids) - 4)], repr(ids))
check("clone parts continue after the template's last slide file",
      "ppt/slides/slide5.xml" in ze.namelist() and "ppt/slides/slide9.xml" in ze.namelist()
      and "ppt/slides/slide10.xml" not in ze.namelist())

xe = lambda n: ze.read("ppt/slides/slide%d.xml" % n).decode("utf-8")
check("plan title is the template's own (no Russian default)",
      "Agenda" in xe(2) and "План" not in xe(2))
check("a template hint in parentheses is dropped from the plan title",
      engine.plan_title_text('<p:sp><p:ph type="title"/><a:t>Contents (if any)</a:t></p:sp>')
      == "Contents"
      and engine.plan_title_text(zipfile.ZipFile(TEMPLATE).read("ppt/slides/slide2.xml").decode("utf-8"))
      == "План")
over_doc = engine.parse_md(en_text.replace("# plan\n", "# plan\n## Contents\n"))
check("`## <title>` in # plan overrides the title", over_doc["plan_title"] == "Contents"
      and over_doc["plan"][0][0] == "Origins")

check("body_size defaults to the template's size (14 pt): plan 13 pt, tables 13 pt",
      en_doc["body_size"] == 14 and en_doc["TOC"] == 13 and en_doc["table_pt"] == 13
      and 'sz="1300"' in xe(2) and 'sz="1300"' in xe(7))
check("an explicit body_size overrides the template's size",
      engine.parse_md(en_text.replace("lang: en", "lang: en\nbody_size: 20"))["body_size"] == 20)

# lists: the content layout declares no bullet and indent 0 at every level
s3e = xe(3)
check("`- ` items get a bullet and a hanging indent the template lacks",
      '<a:pPr lvl="1" marL="685800" indent="-228600"><a:buFont typeface="Arial"/>'
      '<a:buChar char="•"/></a:pPr>' in s3e)
check("numbered items hang off their number",
      '<a:pPr lvl="1" marL="685800" indent="-228600"><a:buFontTx/>'
      '<a:buAutoNum type="arabicPeriod"/></a:pPr>' in s3e)
check("plain lines carry no bullet override", '<a:pPr lvl="0"/>' in s3e)
check("plan numbers hang off their number",
      '<a:pPr lvl="0" marL="228600" indent="-228600"><a:buFontTx/>' in xe(2))
check("MFUA's own bullets and indents are left alone (no stamped bullet)",
      "<a:buChar" not in zipfile.ZipFile(ru_out).read("ppt/slides/slide3.xml").decode("utf-8")
      and "marL=" not in zipfile.ZipFile(ru_out).read("ppt/slides/slide3.xml").decode("utf-8"))

# paragraph spacing of the layout counts in the estimate
cat = engine.layout_catalog({n: ze.read(n) for n in ze.namelist()})
blk = [c for c in cat if c[1] == "Title, Content 1 (Max)"][0][3][0][3]
spc = engine.effective_spacing(blk, ze.read(engine.MASTER).decode("utf-8"), 14)
check("effective spacing: 120 % lines, 10 pt before level 0, 5 pt before level 1",
      spc == {0: (120, 127000, 0), 1: (120, 63500, 0)}, repr(spc))
check("a template with no spacing keeps the plain estimate (MFUA)",
      engine.effective_spacing(None, zipfile.ZipFile(TEMPLATE).read(engine.MASTER).decode("utf-8"), 18)
      == {0: (100, 0, 0), 1: (100, 0, 0)})
text_items = [(0, [("a line", False, False, False)], None)] * 4
check("spacing makes the text estimate taller",
      engine.column_flow(text_items, (0, 0, 8229600, 4000000), 14, 13, spc)["h_emu"]
      > engine.column_flow(text_items, (0, 0, 8229600, 4000000), 14, 13)["h_emu"])

# sentinel-index layouts: two columns and an explicit single column
s4e = xe(5)
check("two columns on a layout whose placeholders share the sentinel idx",
      "slideLayout43.xml" in ze.read("ppt/slides/_rels/slide5.xml.rels").decode("utf-8"))
check("each sentinel placeholder gets its own unique idx",
      '<p:ph type="body" idx="4294967295"/>' in s4e and "4294967294" not in s4e
      and s4e.count("<p:ph ") >= 2)
check("the sentinel placeholder carries its explicit box and list style",
      '<a:off x="838203" y="2147459"/><a:ext cx="5181603" cy="3878628"/>' in s4e
      and '<a:lvl1pPr marL="0" indent="0">' in s4e)
s8e = xe(8)
check("an explicit `layout:` on a sentinel layout takes its box and style",
      "slideLayout15.xml" in ze.read("ppt/slides/_rels/slide8.xml.rels").decode("utf-8")
      and '<a:off x="5282415" y="2386584"/>' in s8e)
cat_e = engine.layout_catalog({n: ze.read(n) for n in ze.namelist()})
two = [c[1] for c in cat_e if engine.is_two_content(c, engine.master_body_box({n: ze.read(n) for n in ze.namelist()}))]
check("only the genuine two-column layout is detected", two == ["Two Content"], repr(two))

# text colour where the template's text is not the theme's tx1
ce = ze.read("ppt/charts/chart1.xml").decode("utf-8")
check("chart text takes the master's body colour and typeface",
      '<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:latin typeface="Poppins"/>' in ce)
s6e = xe(7)
check("cells of a fill-less table style carry the body text colour",
      s6e.count('<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>') >= 6)
check("the default table style keeps its own cell colours",
      '<a:solidFill>' not in xe(9).split("<a:tbl>")[1])

# validation reports a duplicate relationship id
dup_path = os.path.join(OUT, "dup_rid.pptx")
with zipfile.ZipFile(en_out) as zin, zipfile.ZipFile(dup_path, "w") as zout:
    for item in zin.infolist():
        data = zin.read(item.filename)
        if item.filename == "ppt/_rels/presentation.xml.rels":
            data = data.replace(b'Id="rId12"', b'Id="rId11"', 1)
        zout.writestr(item, data)
check("validation reports a duplicate relationship id",
      any("duplicate relationship id" in p_ for p_ in engine.validate(dup_path)))
check("fit check passes for the English deck", engine.fit_check(en_doc) == [],
      repr(engine.fit_check(en_doc)))

if fails:
    print("FAILED: %d check(s)" % len(fails))
    sys.exit(1)
print("ALL PYTHON TESTS PASSED")
