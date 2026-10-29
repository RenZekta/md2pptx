#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""md2pptx deterministic engine.

Reads a structure.md and a template .pptx, produces a .pptx.

    python engine.py <structure.md> <template.pptx> <output.pptx>

Design notes
------------
* OOXML is a ZIP of XML. All edits are targeted raw-string replacements on the
  raw XML (NO ElementTree round-trip -- a round-trip would rewrite namespace
  prefixes and corrupt the deck). ElementTree is used ONLY to *parse-check*
  well-formedness at validation time.
* Text is XML-escaped (&, <, >).
* Inline marks: `**bold**`, `*italic*`/`_italic_`, `++underline++` become the
  run properties b="1" / i="1" / u="sng" (parse_inline documents the boundary
  rules; an unbalanced marker stays plain text).
* A numbered item (`1. text` or `1) text`) is a PowerPoint auto-numbered
  paragraph: the number is not stored in the text, the paragraph carries
  <a:buAutoNum> and PowerPoint draws (and renumbers) the numerals.
  Consecutive items of one scheme form a run; a run written starting at
  N != 1 carries startAt="N". The same applies to `# plan` items.
* A body line starting with "|" is a pipe table: a <p:graphicFrame><a:tbl>
  shape appended to the slide's <p:spTree> (a table cannot live in a text
  placeholder), placed below the column's text lines under the same
  average-character height estimate fit_check uses. An options line
  `{widths: 3,1,1; size: 14; header: off; first-col: on; banded: off;
  style: plain|grid}` directly above the table customises it; a cell that is
  exactly `<<` merges into the cell to its left, `^^` into the cell above. A
  table that would overflow its box shrinks its font down to TABLE_MIN_PT
  unless `size` is set.
* `{chart: column|bar|line|pie; title: ..; legend: top|bottom|left|right|off;
  labels: on; stacked: on; size: 12; height: 3.5}` above a pipe table makes
  the table a native chart (<p:graphicFrame> -> ppt/charts/chartN.xml with an
  embedded workbook, so "Edit Data" works): first column = categories, the
  header row = series names, the other cells = numbers (a comma is a decimal
  comma). Without `height` the chart fills the space left in its column.
* `::: left` ... `::: right` ... `:::` blocks give a slide two columns, backed
  by the template's two-content layout (a title plus two side-by-side body
  placeholders); `layout: <name>` picks any template layout by name. Each
  column has its own box, read from the layout placeholders (placeholder ->
  layout -> master), and tables/charts flow inside their own column.
* The template must contain: a title slide (slide1, freeform, topic line
  anchored by TITLE_ANCHOR_RU or TITLE_ANCHOR_EN), a plan slide (slide2), a content slide
  (slide3, the clone prototype, with <p:ph type="title"/> and <p:ph idx="1"/>),
  and a closing slide (slide4).
* Content slides are cloned from slide3 (each clone is registered in
  [Content_Types].xml, presentation.xml.rels and the sldIdLst), then the
  sldIdLst is reordered so the closing slide is last. A slide that selects
  another layout points its relationship at that layout file.
* Font sizes: body text (levels 0/1) and content-slide titles stamp NO size --
  the runs inherit the template fonts (placeholder -> layout -> master ->
  theme default 18 pt), so the template controls the deck's body font.
  Plan-slide items are stamped at `body_size - 1` (the engine's plan override);
  the closing title is stamped at the size declared in the template's
  closing-slide title placeholder (fallback 36 pt).
* Proofing language: generated runs carry no lang unless the structure.md
  front matter declares `lang: <BCP-47>`; without it PowerPoint falls back to
  its default proofing language and flags non-English text as "misspelled".
  When `lang` is declared, every generated run and endParaRPr gets
  lang="<tag>" dirty="0" (the same stamp PowerPoint puts on pasted runs).

The web app (Mode B) in webapp/index.html implements the same pipeline as a
self-contained page and shares the template font preservation behavior
(kept original-run rPr typefaces/lang, sizes stamped in hundredths of a
point).
"""
import sys
import os
import re
import math
import io
import posixpath
import zipfile
from xml.etree import ElementTree as ET  # parse-check only

# ----------------------------- geometry / config -----------------------------
EMU_PER_INCH = 914400.0
SLIDE_W = 9144000.0                  # EMU  (4:3, 10in x 7.5in)
SLIDE_H = 6858000.0                  # EMU
# Body box (x, y, cx, cy in EMU) used when the template declares no body
# placeholder geometry anywhere in the slide -> layout -> master chain.
DEFAULT_BOX = (457200, 1600200, 8229600, 4525963)
MASTER = "ppt/slideMasters/slideMaster1.xml"

# Tables: a <p:graphicFrame> carrying an <a:tbl> (the OOXML form PowerPoint
# itself writes; see python-pptx's graphicFrame output).
TABLE_URI = "http://schemas.openxmlformats.org/drawingml/2006/table"
TABLE_STYLE_DEFAULT = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"  # Office default
TABLE_GAP = 137160                   # gap above a table (0.15 in, EMU)
TABLE_CELL_PAD = 91440               # cell top+bottom inset (2 * 0.05 in)
TABLE_MIN_COL_CHARS = 6              # column width floor, in characters
TABLE_MIN_PT = 10                    # smallest size a table shrinks to
# `style:` option values; None keeps the template's default table style.
TABLE_STYLES = {"default": None,
                "plain": "{2D5ABB26-0587-4C30-8999-92F81FD0307C}",  # No Style, No Grid
                "grid": "{5940675A-B579-460E-94D1-54222C63F5DA}"}   # No Style, Table Grid

# Charts: a <p:graphicFrame> whose <c:chart> points at ppt/charts/chartN.xml.
CHART_URI = "http://schemas.openxmlformats.org/drawingml/2006/chart"
CHART_TYPES = ("column", "bar", "line", "pie")   # bar = horizontal bars
CHART_MIN_H = 1828800                # smallest height of a chart (2.0 in)
CHART_PT_BELOW_TABLE = 4             # default chart text = table size - 4 pt
MAX_SERIES = 25                      # workbook columns B..Z

CONTENT_SLIDE = "ppt/slides/slide3.xml"          # clone prototype
CONTENT_RELS = "ppt/slides/_rels/slide3.xml.rels"
TITLE_SLIDE = "ppt/slides/slide1.xml"
PLAN_SLIDE = "ppt/slides/slide2.xml"
CLOSING_SLIDE = "ppt/slides/slide4.xml"
CLOSE_SLIDE = "ppt/slides/slide4.xml"

TITLE_ANCHOR_RU = "На тему"         # topic-line anchor in the title slide
TITLE_ANCHOR_EN = "Topic"           # English title-slide topic-line anchor
TITLE_MARKER = '<p:ph type="title"/>'   # title placeholder
BODY_MARKER = '<p:ph idx="1"/>'       # body placeholder


def topic_anchor(title_xml):
    """Anchor present in the title slide: Russian first, then English;
    raises when neither is found."""
    if TITLE_ANCHOR_RU in title_xml:
        return TITLE_ANCHOR_RU
    if TITLE_ANCHOR_EN in title_xml:
        return TITLE_ANCHOR_EN
    raise SystemExit("ERROR: no title-slide topic anchor found (expected "
                     "%r or %r)" % (TITLE_ANCHOR_RU, TITLE_ANCHOR_EN))


TITLE_PT = 36                        # fallback closing-slide title size
LINE_SPACING = 1.2                    # em, for the fit estimate
AVG_CHAR_EM = 0.5                   # em, average char width for the fit estimate


# ----------------------------- xml helpers -----------------------------------
def esc(t):
    return t.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def parse_inline(text):
    """Split text into (text, bold, italic, underline) runs on the inline
    marks: `**bold**`, `*italic*` and `_italic_`, `++underline++`.

    A marker toggles its style only at a real boundary: when opening, the
    character after it must not be a space (or the end of the line); when
    closing, the character before it must not be a space. On either side the
    neighbouring character must not repeat the marker, so a stray `a ** b`
    stays literal and `__x__` is not a mark (bold is `**`). `_` is stricter
    still: it marks only at word boundaries (`isalnum` on either side
    disqualifies it), so `snake_case` and file names stay plain text. `**` is
    considered before `*`, which keeps `**bold**` bold and leaves the outer
    stars of `***x***` literal. Marks must balance within a line;
    an unbalanced opening mark simply styles to the end of the line."""
    n = len(text)

    def alnum(p):
        return 0 <= p < n and (text[p].isalnum() or text[p] == "_")

    def blank(p):
        return 0 <= p < n and (text[p] == " " or text[p] == "\t")

    state = [False, False, False]        # bold, italic, underline
    runs = []
    buf = []

    def toggles(i, ln, ch, idx, word_strict):
        """May the marker at `i` open or close its style?"""
        nxt, prv = i + ln, i - 1
        if state[idx]:                                   # closing
            if prv < 0 or blank(prv) or text[prv] == ch:
                return False
            if word_strict and alnum(nxt):
                return False
            return True
        if nxt >= n or blank(nxt) or text[nxt] == ch:     # opening
            return False
        if word_strict and alnum(prv):
            return False
        return True

    i = 0
    while i < n:
        ch = text[i]
        two = text[i:i + 2]
        hit = None
        if two == "**" and toggles(i, 2, "*", 0, False):
            hit = (2, 0)
        elif two == "++" and toggles(i, 2, "+", 2, False):
            hit = (2, 2)
        elif ch == "*" and toggles(i, 1, "*", 1, False):
            hit = (1, 1)
        elif ch == "_" and toggles(i, 1, "_", 1, True):
            hit = (1, 1)
        if hit:
            if buf:
                runs.append(("".join(buf), state[0], state[1], state[2]))
                buf = []
            state[hit[1]] = not state[hit[1]]
            i += hit[0]
        else:
            buf.append(ch)
            i += 1
    if buf:
        runs.append(("".join(buf), state[0], state[1], state[2]))
    if not runs:
        runs = [("", False, False, False)]
    return runs


def _rpr(fmt, sz=None, lang=None):
    """Run properties. `fmt` is the (bold, italic, underline) triple; `sz` is
    in hundredths of a point (omitted: inherit the template size); `lang`
    (BCP-47, from the structure.md front matter) stamps lang + dirty="0" so
    PowerPoint does not fall back to its default proofing language and flag the
    text as misspelled. Italic is `i="1"`, underline `u="sng"`; `fmt` None
    means no marks."""
    bold, italic, underline = fmt if fmt else (False, False, False)
    attrs = []
    if sz:
        attrs.append('sz="%d"' % sz)
    if bold:
        attrs.append('b="1"')
    if italic:
        attrs.append('i="1"')
    if underline:
        attrs.append('u="sng"')
    if lang:
        attrs.append('lang="%s"' % lang)
        attrs.append('dirty="0"')
    return "<a:rPr%s/>" % ((" " + " ".join(attrs)) if attrs else "")


# Full <a:rPr .../> (self-closing) or <a:rPr ...>children</a:rPr> — the
# children (a:latin/a:cs/a:ea typeface elements) must be kept, otherwise
# replaced runs fall back to the theme font.
RPR_RE = re.compile(
    r"<a:rPr\b[^>]*/>|<a:rPr\b[^>]*>(?:(?!</a:rPr>).)*</a:rPr>", re.S)


def first_rpr(xml, marker):
    """First run rPr (full XML) inside the <p:sp> that contains `marker`,
    or None. Used as the font base for the replacement runs."""
    i = xml.index(marker)
    end = xml.find("</p:sp>", i) + len("</p:sp>")
    m = RPR_RE.search(xml[i:end])
    return m.group(0) if m else None


def merge_rpr(base, fmt=None, sz=None, lang=None):
    """Build an rPr from the original run's rPr (`base`): keeps its typeface
    child elements; stamps `fmt` (the bold/italic/underline triple, None = no
    marks) and `sz` (hundredths of a point), overriding any stamped in `base`
    — including a template run that was already italic or underlined, which
    must not leak into the new text. When `lang` is given, any `lang`/
    `dirty` attrs stamped in `base` are replaced by `lang` + dirty="0" so
    PowerPoint proves the run in the declared language instead of its
    default."""
    bold, italic, underline = fmt if fmt else (False, False, False)
    if not base:
        return _rpr(fmt, sz, lang)
    if base.endswith("/>"):
        attrs = base[6:-2].strip()
        children = ""
    else:
        head = base[6:]
        gi = head.index(">")
        attrs = head[:gi].strip()
        children = head[gi + 1:-len("</a:rPr>")]
    attrs = re.sub(r'\bsz="\d+"', "", attrs)
    attrs = re.sub(r'\bb="1"', "", attrs)
    attrs = re.sub(r'\bi="1"', "", attrs)
    attrs = re.sub(r'\bu="[^"]*"', "", attrs)
    extra = []
    if sz:
        extra.append('sz="%d"' % sz)
    if bold:
        extra.append('b="1"')
    if italic:
        extra.append('i="1"')
    if underline:
        extra.append('u="sng"')
    if lang:
        attrs = re.sub(r'\blang="[^"]*"', "", attrs)
        attrs = re.sub(r'\bdirty="[^"]*"', "", attrs)
        extra.append('lang="%s"' % lang)
        extra.append('dirty="0"')
    a = (attrs + (" " + " ".join(extra) if extra else "")).strip()
    prefix = "<a:rPr" + (" " + a if a else "")
    if children:
        return prefix + ">" + children + "</a:rPr>"
    return prefix + "/>"


def endpara_rpr(lang=None):
    """<a:endParaRPr/> end mark of a paragraph. When `lang` is given it
    carries lang + dirty="0" too (PowerPoint stamps both on paste; a bare
    end mark would leave the paragraph's proofing language at the
    default)."""
    if lang:
        return '<a:endParaRPr lang="%s" dirty="0"/>' % lang
    return "<a:endParaRPr/>"


# Numbered list items: "1. text" or "1) text". The scheme is the OOXML
# <a:buAutoNum type="..."> value; [0-9] keeps the match ASCII-only, the same
# as the JS engine.
NUMBERED_RE = re.compile(r"^([0-9]+)([.)])\s+(.*)$")
NUM_SCHEME = {".": "arabicPeriod", ")": "arabicParenR"}


def parse_numbered(line):
    """(scheme, number, text) when `line` is a numbered list item, else None."""
    m = NUMBERED_RE.match(line)
    if not m:
        return None
    return NUM_SCHEME[m.group(2)], int(m.group(1)), m.group(3)


def group_numbering(raw):
    """Turn per-item marks into per-run marks.

    `raw` lists, in paragraph order, None (not numbered) or (scheme, number)
    as written. A run is a maximal stretch of consecutive numbered items of
    one scheme; PowerPoint numbers a run on its own, counting up from its
    first item, so every item of a run gets (scheme, start) where `start` is
    the number the run's first item was written with. Later numbers in a run
    are ignored (CommonMark semantics)."""
    out = []
    prev = None
    start = None
    for mark in raw:
        if mark is None:
            prev = None
            out.append(None)
            continue
        scheme, number = mark
        if scheme != prev:
            start = number
        prev = scheme
        out.append((scheme, start))
    return out


def number_slide_items(parsed):
    """Replace the as-written numbers of a slide's text items with run marks.
    Tables and charts are separate shapes, so they do not split a run: the
    paragraphs PowerPoint sees in the body placeholder are contiguous."""
    idx = [k for k, it in enumerate(parsed) if it[0] in (0, 1)]
    marks = group_numbering([parsed[k][2] for k in idx])
    for k, mark in zip(idx, marks):
        parsed[k] = (parsed[k][0], parsed[k][1], mark)


# Hanging indent (EMU, 0.25 in) given to a bullet or number the template's
# list style leaves flush against its text.
HANG = 228600


def level_style(source, level):
    """(marL, indent, bullet) a list-style block declares for 0-based `level`:
    each None when it declares nothing; bullet is "none", "char", "autonum"
    or "blip". `source` is the XML of an <a:lstStyle> or a <p:bodyStyle>."""
    n = level + 1
    m = re.search(r"<a:lvl%dpPr\b([^>]*?)(?:/>|>(.*?)</a:lvl%dpPr>)" % (n, n),
                  source, re.S)
    if not m:
        return None, None, None
    head, inner = m.group(1), m.group(2) or ""
    ml = re.search(r'\bmarL="(-?[0-9]+)"', head)
    ind = re.search(r'\bindent="(-?[0-9]+)"', head)
    bullet = None
    for tag, kind in (("<a:buNone/>", "none"), ("<a:buChar", "char"),
                      ("<a:buAutoNum", "autonum"), ("<a:buBlip", "blip")):
        if tag in inner:
            bullet = kind
            break
    return (int(ml.group(1)) if ml else None,
            int(ind.group(1)) if ind else None, bullet)


def effective_style(block, master_xml, level):
    """Effective (marL, indent, bullet) of `level` in a body placeholder: the
    placeholder's own <a:lstStyle> (`block` is its layout or slide <p:sp>),
    then the master's bodyStyle; the first declaration of each wins."""
    sources = []
    if block:
        m = re.search(r"<a:lstStyle>.*?</a:lstStyle>", block, re.S)
        sources.append(m.group(0) if m else "")
    if master_xml:
        m = re.search(r"<p:bodyStyle>.*?</p:bodyStyle>", master_xml, re.S)
        sources.append(m.group(0) if m else "")
    out = [None, None, None]
    for src in sources:
        got = level_style(src, level)
        for k in range(3):
            if out[k] is None:
                out[k] = got[k]
    return tuple(out)


def level_spacing(source, level):
    """(line, before, after) a list-style block declares for 0-based `level`,
    each None when absent or ("pct", n) / ("pts", n) as written (n in the
    file's thousandths of a percent / hundredths of a point)."""
    n = level + 1
    m = re.search(r"<a:lvl%dpPr\b[^>]*?(?:/>|>(.*?)</a:lvl%dpPr>)" % (n, n),
                  source, re.S)
    inner = (m.group(1) or "") if m else ""
    out = []
    for tag in ("lnSpc", "spcBef", "spcAft"):
        v = re.search(r'<a:%s><a:spc(Pct|Pts) val="(-?[0-9]+)"/></a:%s>' % (tag, tag),
                      inner)
        out.append((v.group(1).lower(), int(v.group(2))) if v else None)
    return tuple(out)


def effective_spacing(block, master_xml, pt):
    """{level: (line %, before EMU, after EMU)} for levels 0 and 1 of a body
    placeholder: the first declaration in its own <a:lstStyle>, then the
    master's bodyStyle (line spacing 100 % and no space around paragraphs when
    neither declares one). A percentage of space before/after is of the font
    size `pt`; an exact line height is expressed as a percentage of the
    estimate's single line."""
    sources = []
    if block:
        m = re.search(r"<a:lstStyle>.*?</a:lstStyle>", block, re.S)
        sources.append(m.group(0) if m else "")
    if master_xml:
        m = re.search(r"<p:bodyStyle>.*?</p:bodyStyle>", master_xml, re.S)
        sources.append(m.group(0) if m else "")
    out = {}
    for level in (0, 1):
        got = [None, None, None]
        for src in sources:
            for k, v in enumerate(level_spacing(src, level)):
                if got[k] is None:
                    got[k] = v

        def emu(v):
            if v is None:
                return 0
            return v[1] * pt * 127 // 1000 if v[0] == "pct" else v[1] * 127
        line = got[0]
        if line is None:
            pct = 100
        elif line[0] == "pct":
            pct = line[1] // 1000
        else:
            pct = line[1] * 127 * 100 // line_emu(pt)
        out[level] = (pct, emu(got[1]), emu(got[2]))
    return out


def list_fixes(block, master_xml):
    """The <a:pPr> additions ((attrs, children) strings) that make a body
    placeholder's lists read as the structure.md asks, where its effective
    list style does not already:

        plain   a level-0 line carries no bullet (a template whose level 1 has
                one gets buNone and no indent)
        bullet  a `- ` item shows a bullet (a template with none gets "•")
        num     {level: ...} a numbered item hangs off its number

    A marker the template leaves with no hanging indent (indent 0 or
    undeclared) gets one of HANG EMU, the text moving right of the position
    the template intended for the marker."""
    l0 = effective_style(block, master_xml, 0)
    l1 = effective_style(block, master_xml, 1)

    def hang(style):
        marl, indent, _b = style
        if indent in (None, 0):
            return ' marL="%d" indent="%d"' % ((marl or 0) + HANG, -HANG)
        return ""

    plain = ("", "")
    if l0[2] in ("char", "autonum", "blip"):
        plain = (' marL="0" indent="0"', "<a:buNone/>")
    bullet = (hang(l1), "")
    if l1[2] in (None, "none"):
        bullet = (hang(l1), '<a:buFont typeface="Arial"/><a:buChar char="•"/>')
    return {"plain": plain, "bullet": bullet,
            "num": {0: (hang(l0), ""), 1: (hang(l1), "")}}


def para_props(level, num, add=("", "")):
    """<a:pPr> of a body or plan paragraph. `add` is (attrs, children) from
    list_fixes. A numbered paragraph swaps the inherited bullet for
    auto-numbering in the text's own font (the indents come from the template
    unless `add` gives them); `startAt` is written only when the run does not
    start at 1."""
    attrs, kids = add
    if num:
        scheme, start = num
        start_at = ' startAt="%d"' % start if start != 1 else ""
        kids += '<a:buFontTx/><a:buAutoNum type="%s"%s/>' % (scheme, start_at)
    if not kids:
        return '<a:pPr lvl="%d"%s/>' % (level, attrs)
    return '<a:pPr lvl="%d"%s>%s</a:pPr>' % (level, attrs, kids)


def para(level, runs, sz=None, base=None, lang=None, num=None, add=("", "")):
    """Body paragraph at `level` (0 = plain line, 1 = bullet or numbered
    item; `num` is the (scheme, start) run mark of a numbered item). `runs` are the
    (text, bold, italic, underline) triples of parse_inline. The font size
    is stamped only when `sz` is given; otherwise runs inherit the
    template's body font. `base` is the original paragraph's first-run rPr:
    its typeface elements (a:latin/a:cs/a:ea) are kept on every new run so
    the replaced text keeps the template's font. `lang` (BCP-47) is stamped
    on every run and the paragraph end mark so PowerPoint does not flag the
    text."""
    ppr = para_props(level, num, add)
    rs = "".join('<a:r>%s<a:t>%s</a:t></a:r>'
                 % (merge_rpr(base, r[1:], sz, lang), esc(r[0])) for r in runs)
    return '<a:p>%s%s%s</a:p>' % (ppr, rs, endpara_rpr(lang))


def title_para(text, sz=None, base=None, lang=None):
    """`sz` is in points (stamped as hundredths, like the template does).
    A title carries no inline marks of its own; slide titles take the
    template's title style."""
    return ('<a:p><a:r>%s<a:t>%s</a:t></a:r>%s</a:p>'
            % (merge_rpr(base, None, sz * 100 if sz else None, lang),
               esc(text), endpara_rpr(lang)))


def toc_para(text, sz, base=None, lang=None, num=None, add=("", "")):
    """Plan-slide item at `sz` pt; `num` is the (scheme, start) run mark when
    the item is numbered."""
    return ('<a:p>%s<a:r>%s<a:t>%s</a:t></a:r>%s</a:p>'
            % (para_props(0, num, add), merge_rpr(base, None, sz * 100, lang),
               esc(text), endpara_rpr(lang)))


def chars_per_line(width_in, pt):
    """Characters that fit on one line of `width_in` inches at `pt` — the
    average-char-width heuristic fit_check and the table layout share."""
    return max(1, int(width_in / (AVG_CHAR_EM * (pt / 72.0))))


def est_lines(text, width_in, pt):
    return max(1, math.ceil(len(text) / chars_per_line(width_in, pt)))


def line_emu(pt):
    """One line's height in EMU at `pt` (integer: the heuristic is exact for
    whole point sizes, so Python and JS agree byte-for-byte)."""
    return int(LINE_SPACING * (pt / 72.0) * EMU_PER_INCH + 0.5)


def runs_text(runs):
    return "".join(r[0] for r in runs)


# ----------------------------- options lines ---------------------------------
TABLE_OPTION_KEYS = ("widths", "size", "header", "first-col", "banded", "style")
CHART_OPTION_KEYS = ("chart", "title", "legend", "labels", "stacked", "size",
                     "height")


def is_options_line(line):
    s = line.strip()
    return len(s) >= 2 and s.startswith("{") and s.endswith("}")


def parse_options(line):
    """Key/value pairs of an options line `{key: value; key: value}`.

    The keys are those of a table, or of a chart when `chart` is present;
    anything else is an error, so a misspelt option never silently does
    nothing. Keys are lower-cased, values keep their case (a chart title)."""
    opts = {}
    for part in line.strip()[1:-1].split(";"):
        part = part.strip()
        if not part:
            continue
        if ":" not in part:
            raise SystemExit("ERROR: option %r in %r needs the form key: value"
                             % (part, line.strip()))
        k, v = part.split(":", 1)
        opts[k.strip().lower()] = v.strip()
    chart = "chart" in opts
    allowed = CHART_OPTION_KEYS if chart else TABLE_OPTION_KEYS
    for k in opts:
        if k not in allowed:
            raise SystemExit("ERROR: unknown %s option %r (allowed: %s)"
                             % ("chart" if chart else "table", k,
                                ", ".join(allowed)))
    return opts


def opt_flag(opts, key, default):
    """An on/off option; `default` when absent."""
    v = opts.get(key)
    if v is None:
        return default
    if v.lower() in ("on", "off"):
        return v.lower() == "on"
    raise SystemExit("ERROR: option %s expects on or off, got %r" % (key, v))


def opt_int(opts, key, lo, hi):
    """A whole-number option in [lo, hi]; None when absent."""
    v = opts.get(key)
    if v is None:
        return None
    if not re.match(r"^[0-9]+$", v) or not lo <= int(v) <= hi:
        raise SystemExit("ERROR: option %s expects a whole number from %d to %d, "
                         "got %r" % (key, lo, hi, v))
    return int(v)


# ----------------------------- tables ----------------------------------------
def split_table_rows(lines):
    """Pipe-table lines -> (rows of cell strings, alignments, header).

    A line is a row when it starts with "|"; cells are split on "|" and
    trimmed, and the outer pipes are optional. A row of only |, -, : and
    spaces is the header separator: it is dropped, marks the row above as the
    header and sets each column's alignment from its colons (`:--` left,
    `--:` right, `:-:` centre). Short rows are padded to the widest one."""
    rows, aligns, header = [], [], False
    for ln in lines:
        cells = [c.strip() for c in ln.strip().strip("|").split("|")]
        if (rows and any(cells)
                and all(re.match(r"^:?-+:?$", c) for c in cells if c)):
            header = True
            al = []
            for c in cells:
                if c.startswith(":") and c.endswith(":"):
                    al.append("ctr")
                elif c.endswith(":"):
                    al.append("r")
                elif c.startswith(":"):
                    al.append("l")
                else:
                    al.append(None)
            aligns = al
            continue
        rows.append(cells)
    width = max(len(r) for r in rows)
    for r in rows:
        del r[width:]                       # ragged wide rows: extra cells drop
        while len(r) < width:
            r.append("")
    return rows, aligns + [None] * (width - len(aligns)), header


def table_spans(rows):
    """Merge structure of a table whose merged cells are written `<<` (joins
    the cell to its left) and `^^` (joins the cell above).

    Returns spans[r][c] = (gridSpan, rowSpan, hMerge, vMerge). A merge chain
    resolves to one origin cell (a `<<` under a `^^` joins the same block),
    which carries the spans; the cells it covers are flagged hMerge and/or
    vMerge, as PowerPoint writes a merged range. A marker with no cell to join
    (first column / first row) is ordinary text."""
    nrow, ncol = len(rows), len(rows[0])
    origin = [[None] * ncol for _ in range(nrow)]
    for r in range(nrow):
        for c in range(ncol):
            if rows[r][c] == "<<" and c > 0:
                origin[r][c] = origin[r][c - 1]
            elif rows[r][c] == "^^" and r > 0:
                origin[r][c] = origin[r - 1][c]
            else:
                origin[r][c] = (r, c)
    grid, rowsp = {}, {}
    for r in range(nrow):
        for c in range(ncol):
            r0, c0 = origin[r][c]
            grid[(r0, c0)] = max(grid.get((r0, c0), 1), c - c0 + 1)
            rowsp[(r0, c0)] = max(rowsp.get((r0, c0), 1), r - r0 + 1)
    spans = []
    for r in range(nrow):
        row = []
        for c in range(ncol):
            r0, c0 = origin[r][c]
            if (r0, c0) == (r, c):
                row.append((grid[(r, c)], rowsp[(r, c)], False, False))
            else:
                row.append((1, 1, c0 < c, r0 < r))
        spans.append(row)
    return spans


def parse_widths(opts, ncol):
    """`widths: 3,1,1` -> relative column weights (whole numbers > 0), one per
    column; None when absent."""
    v = opts.get("widths")
    if v is None:
        return None
    parts = [p.strip() for p in v.split(",")]
    if (len(parts) != ncol or not all(re.match(r"^[0-9]+$", p) and int(p) > 0
                                      for p in parts)):
        raise SystemExit("ERROR: option widths expects %d positive whole numbers "
                         "separated by commas (one per column), got %r"
                         % (ncol, v))
    return [int(p) for p in parts]


def parse_table(lines, opts=None):
    """Parse pipe-table lines (see split_table_rows) into a table dict:

        header     first row is bold (a separator row, or `header: on`)
        first_row  tblPr firstRow flag (the style's header formatting)
        align      per-column "l" / "ctr" / "r" / None
        cells      [[run-tuples]], runs carry the same inline marks as body lines
        spans      table_spans(), parallel to cells
        widths     `widths:` weights or None
        pt         `size:` in points or None (None = the table may shrink)
        first_col, banded, style   tblPr flags and the `style:` key

    `opts` is the table's options line, see parse_options."""
    opts = opts or {}
    rows, aligns, header = split_table_rows(lines)
    if "header" in opts:
        header = opt_flag(opts, "header", header)
    style = opts.get("style", "default").lower()
    if style not in TABLE_STYLES:
        raise SystemExit("ERROR: option style expects one of %s, got %r"
                         % (", ".join(TABLE_STYLES), opts["style"]))
    spans = table_spans(rows)
    cells = []
    for ri, row in enumerate(rows):
        rr = []
        for ci, cell in enumerate(row):
            if spans[ri][ci][2] or spans[ri][ci][3]:
                runs = []                   # covered by a merge
            else:
                runs = parse_inline(cell)
                if header and ri == 0:
                    runs = [(t, True, it, ul) for (t, _b, it, ul) in runs]
            rr.append(runs)
        cells.append(rr)
    return {"header": header, "first_row": header if "header" in opts else True,
            "align": aligns, "cells": cells, "spans": spans,
            "widths": parse_widths(opts, len(aligns)),
            "pt": opt_int(opts, "size", 8, 40),
            "first_col": opt_flag(opts, "first-col", False),
            "banded": opt_flag(opts, "banded", True), "style": style}


# ----------------------------- charts ----------------------------------------
def parse_number(text):
    """A chart value: an optional sign, digits and one decimal separator. A
    comma is a decimal comma ("1,5"); with both a comma and a point the commas
    are thousands separators; spaces and a trailing % are ignored."""
    t = text.replace("\u00a0", "").replace(" ", "")
    if t.endswith("%"):
        t = t[:-1]
    if "," in t and "." in t:
        t = t.replace(",", "")
    elif "," in t:
        t = t.replace(",", ".")
    if not re.match(r"^[+-]?[0-9]+(\.[0-9]+)?$", t):
        raise SystemExit("ERROR: chart value %r is not a number" % text)
    v = float(t)
    if abs(v) >= 1e15 or (v != int(v) and abs(v) < 1e-4):
        raise SystemExit("ERROR: chart value %r is out of the supported range"
                         % text)
    return v


def fmt_num(v):
    """A number as written into chart XML and the workbook."""
    return str(int(v)) if v == int(v) else repr(v)


def parse_chart(lines, opts):
    """Chart dict from pipe-table lines and the chart's options line.

    The first row names the series, the first column holds the categories and
    every other cell is a number. A pie uses only the first series."""
    kind = opts["chart"].lower()
    if kind not in CHART_TYPES:
        raise SystemExit("ERROR: option chart expects one of %s, got %r"
                         % (", ".join(CHART_TYPES), opts["chart"]))
    rows, _aligns, _header = split_table_rows(lines)
    if len(rows) < 2 or len(rows[0]) < 2:
        raise SystemExit("ERROR: a chart table needs a header row, at least one "
                         "category row and one series column")
    if len(rows[0]) - 1 > MAX_SERIES:
        raise SystemExit("ERROR: a chart supports at most %d series" % MAX_SERIES)

    def plain(cell):
        return runs_text(parse_inline(cell))
    cats = [plain(r[0]) for r in rows[1:]]
    series = [(plain(rows[0][j]), [parse_number(r[j]) for r in rows[1:]])
              for j in range(1, len(rows[0]))]
    legend = opts.get("legend")
    if legend is not None and legend.lower() not in ("top", "bottom", "left",
                                                     "right", "off"):
        raise SystemExit("ERROR: option legend expects top, bottom, left, right "
                         "or off, got %r" % legend)
    stacked = opt_flag(opts, "stacked", False)
    if stacked and kind not in ("column", "bar"):
        raise SystemExit("ERROR: option stacked applies to column and bar charts")
    height = None
    if "height" in opts:
        if (not re.match(r"^[0-9]+(\.[0-9]+)?$", opts["height"])
                or not 1 <= float(opts["height"]) <= 7):
            raise SystemExit("ERROR: option height expects inches from 1 to 7, "
                             "got %r" % opts["height"])
        height = int(float(opts["height"]) * EMU_PER_INCH + 0.5)
    return {"type": kind, "title": opts.get("title"), "cats": cats,
            "series": series if kind != "pie" else series[:1],
            "legend": legend.lower() if legend else None,
            "labels": opt_flag(opts, "labels", False), "stacked": stacked,
            "pt": opt_int(opts, "size", 8, 40), "height": height}


# ----------------------------- template geometry -----------------------------
XFRM_RE = re.compile(r'<a:off x="(-?[0-9]+)" y="(-?[0-9]+)"/>'
                     r'<a:ext cx="([0-9]+)" cy="([0-9]+)"/>')


def sp_block(xml, marker):
    """The <p:sp>...</p:sp> containing `marker`, or None."""
    i = xml.find(marker)
    if i == -1:
        return None
    start = xml.rfind("<p:sp>", 0, i)
    return xml[start:xml.find("</p:sp>", i) + len("</p:sp>")]


def sp_box(block):
    """(x, y, cx, cy) in EMU declared in a shape's spPr, or None when the
    shape inherits its geometry."""
    if block is None:
        return None
    m = XFRM_RE.search(block.split("<p:txBody>")[0])
    return tuple(int(g) for g in m.groups()) if m else None


# A layout placeholder with idx at or above this carries PowerPoint's "no
# index" sentinel (4294967295): no slide placeholder can be linked to it by
# idx, and several of them can share the value.
SENTINEL_IDX = 4294967000


def body_phs(layout_xml):
    """Body placeholders of a layout as (idx, <p:ph> tag, box or None, <p:sp>
    block), ordered by idx and then left to right, top to bottom (placeholders
    that share the sentinel idx are told apart by position). Footer, date,
    slide number, picture and vertical-text placeholders are not body
    placeholders."""
    out = []
    for sp in re.findall(r"<p:sp>.*?</p:sp>", layout_xml, re.S):
        m = re.search(r"<p:ph\b[^>]*/>", sp)
        if not m:
            continue
        tag = m.group(0)
        idx = re.search(r'\bidx="([0-9]+)"', tag)
        typ = re.search(r'\btype="([^"]+)"', tag)
        if not idx or (typ and typ.group(1) not in ("body", "obj")):
            continue
        if 'orient="vert"' in tag:
            continue
        out.append((int(idx.group(1)), tag, sp_box(sp), sp))
    out.sort(key=lambda p: (p[0], p[2][0] if p[2] else 0, p[2][1] if p[2] else 0))
    return out


def master_body_box(parts):
    """The slide master's body placeholder box, the end of the geometry chain."""
    xml = parts.get(MASTER)
    if xml is None:
        return DEFAULT_BOX
    return (sp_box(sp_block(xml.decode("utf-8"), '<p:ph type="body" idx="1"/>'))
            or DEFAULT_BOX)


def layout_file(rels_xml):
    """The slide layout part a slide's rels point at, or None."""
    m = re.search(r'Target="[^"]*(slideLayout[0-9]+\.xml)"', rels_xml)
    return "ppt/slideLayouts/" + m.group(1) if m else None


def layout_body_block(parts, rels_xml, idx):
    """The <p:sp> of the layout (named by a slide's rels) body placeholder
    with this idx, or None."""
    lay = parts.get(layout_file(rels_xml) or "")
    if lay is not None:
        for i, _tag, _box, block in body_phs(lay.decode("utf-8")):
            if i == idx:
                return block
    return None


def default_body(parts, content_xml, content_rels):
    """(box, layout block) of the clone prototype's body placeholder. The box
    is the slide's own geometry, else its layout's, else the master's; the
    block is the layout's placeholder (None without one)."""
    block = layout_body_block(parts, content_rels, 1)
    box = sp_box(sp_block(content_xml, BODY_MARKER)) or sp_box(block)
    return box or master_body_box(parts), block


def layout_catalog(parts):
    """(file, name, has_title, body_phs) of every layout, in numeric order."""
    files = sorted((n for n in parts
                    if re.match(r"^ppt/slideLayouts/slideLayout[0-9]+\.xml$", n)),
                   key=lambda n: int(re.search(r"([0-9]+)\.xml$", n).group(1)))
    out = []
    for f in files:
        x = parts[f].decode("utf-8")
        nm = re.search(r'<p:cSld name="([^"]*)"', x)
        out.append((f, nm.group(1) if nm else "", '<p:ph type="title"/>' in x,
                    body_phs(x)))
    return out


def is_two_content(entry, master_box):
    """A title plus exactly two body placeholders sitting side by side."""
    _f, _name, has_title, phs = entry
    if not has_title or len(phs) != 2:
        return False
    a, b = phs[0][2] or master_box, phs[1][2] or master_box
    return a[1] == b[1] and a[0] + a[2] <= b[0]


def plan_slide(parts, default, body_pt, layout_key, ncols):
    """Where a slide's columns go: {"file": layout part or None for the clone
    prototype's own layout, "phs": one dict per column}. A column dict has:

        tag     the <p:ph> tag the slide's placeholder carries
        box     (x, y, cx, cy) EMU of the column
        pt      the column's body text size for the fit estimate: the size the
                layout placeholder declares, else `body_pt`
        block   the layout's <p:sp> for the placeholder (its list style)
        orphan  True when the layout placeholder has the sentinel idx; the
                slide placeholder then gets its own unique sentinel idx, the
                box as an explicit xfrm and a copy of the list style, because
                nothing can link it to the layout placeholder

    `default` is default_body()'s (box, block). Each column dict also has
    `spacing`, effective_spacing() at the column's size."""
    master_xml = parts[MASTER].decode("utf-8") if MASTER in parts else ""
    if layout_key is None and ncols == 1:
        return {"file": None, "phs": [{"tag": BODY_MARKER, "box": default[0],
                                       "pt": body_pt, "block": default[1],
                                       "orphan": False,
                                       "spacing": effective_spacing(default[1], master_xml, body_pt)}]}
    catalog = layout_catalog(parts)
    master_box = master_body_box(parts)
    if layout_key is None or layout_key.lower() == "two-content":
        found = [c for c in catalog if is_two_content(c, master_box)]
        if not found:
            raise SystemExit("ERROR: the template has no two-content layout (a "
                             "title plus two side-by-side body placeholders)")
    else:
        found = [c for c in catalog if c[1].lower() == layout_key.lower()]
        if not found:
            raise SystemExit("ERROR: layout %r not found in the template "
                             "(layouts: %s)" % (layout_key,
                                                ", ".join(repr(c[1]) for c in catalog)))
    f, name, has_title, phs = found[0]
    if not has_title or len(phs) != ncols:
        raise SystemExit("ERROR: layout %r has %s and %d body placeholder(s), "
                         "but the slide has %d column(s) (two columns need "
                         "::: left and ::: right blocks)"
                         % (name, "a title" if has_title else "no title",
                            len(phs), ncols))
    out = []
    orphans = 0
    for _i, tag, box, block in phs:
        orphan = _i >= SENTINEL_IDX
        if orphan:
            tag = re.sub(r'\bidx="[0-9]+"', 'idx="%d"' % (4294967295 - orphans), tag)
            orphans += 1
        pt = block_sz(block) or body_pt
        out.append({"tag": tag, "box": box or master_box, "pt": pt,
                    "block": block, "orphan": orphan,
                    "spacing": effective_spacing(block, master_xml, pt)})
    return {"file": f, "phs": out}


# ----------------------------- layout flow -----------------------------------
def table_geom(tbl, x, total, y, pt):
    """Geometry (EMU) of one table `total` EMU wide at (x, y): column widths
    from the `widths:` weights, else proportional to the longest cell of each
    column (floored at TABLE_MIN_COL_CHARS characters; merged cells do not
    count), leftover EMU to the leftmost columns; row heights from the
    wrapped-line estimate at `pt` plus the cell insets. A cell spanning
    columns wraps at the width of the columns it covers; a cell spanning rows
    does not set a row height."""
    ncol = len(tbl["align"])
    spans = tbl["spans"]
    weights = tbl["widths"]
    if not weights:
        weights = []
        for ci in range(ncol):
            w = TABLE_MIN_COL_CHARS
            for ri, row in enumerate(tbl["cells"]):
                if spans[ri][ci][0] == 1:
                    w = max(w, len(runs_text(row[ci])))
            weights.append(w)
    s = sum(weights)
    col_w = [total * w // s for w in weights]
    for ci in range(total - sum(col_w)):
        col_w[ci] += 1
    row_h = []
    for ri, row in enumerate(tbl["cells"]):
        lines = 1
        for ci, cell in enumerate(row):
            gs, rs, hm, vm = spans[ri][ci]
            if hm or vm or rs > 1:
                continue
            lines = max(lines, est_lines(runs_text(cell),
                                         sum(col_w[ci:ci + gs]) / EMU_PER_INCH, pt))
        row_h.append(lines * line_emu(pt) + TABLE_CELL_PAD)
    return {"x": x, "y": y, "cx": sum(col_w), "cy": sum(row_h),
            "col_w": col_w, "row_h": row_h}


def column_flow(items, box, body_pt, table_pt, spacing=None):
    """Vertical flow of one column inside its body `box` (x, y, cx, cy).

    A text placeholder flows all of its own paragraphs from the top of the
    box, so tables and charts - separate <p:graphicFrame> shapes - sit below
    them: the text height is estimated first (the same average-character
    heuristic as fit_check), then the blocks stack in document order, each
    TABLE_GAP under the previous one. Tables without a `size:` shrink one
    point at a time (down to TABLE_MIN_PT) while the column overflows; charts
    without a `height:` then share the space that is left, at least
    CHART_MIN_H each. `spacing` is effective_spacing()'s {level: (line %,
    before EMU, after EMU)}: each paragraph's lines are scaled by the line
    spacing and followed by the space around it. Returns {"h_emu": total
    height, "pt": table size used, "frames": [(kind, table or chart, geom)]},
    the estimate fit_check compares with the box."""
    bx, by, bw, bh = box
    width_in = bw / EMU_PER_INCH
    spacing = spacing or {0: (100, 0, 0), 1: (100, 0, 0)}
    text_emu = 0
    blocks = []
    for it in items:
        if it[0] in (0, 1):
            pct, before, after = spacing[it[0]]
            text_emu += (est_lines(runs_text(it[1]), width_in, body_pt)
                         * (line_emu(body_pt) * pct // 100) + before + after)
        else:
            blocks.append(it)
    flex = sum(1 for k, o in blocks if k == "chart" and o["height"] is None)
    shrinkable = any(k == "table" and o["pt"] is None for k, o in blocks)

    def place(pt, flex_h):
        y = by + text_emu
        h = text_emu
        frames = []
        for kind, obj in blocks:
            y += TABLE_GAP
            h += TABLE_GAP
            if kind == "table":
                geom = table_geom(obj, bx, bw, y, obj["pt"] or pt)
                geom["pt"] = obj["pt"] or pt
            else:
                geom = {"x": bx, "y": y, "cx": bw, "cy": obj["height"] or flex_h}
            frames.append((kind, obj, geom))
            y += geom["cy"]
            h += geom["cy"]
        return h, frames

    pt = table_pt
    h, frames = place(pt, CHART_MIN_H)
    while shrinkable and h > bh and pt > TABLE_MIN_PT:
        pt -= 1
        h, frames = place(pt, CHART_MIN_H)
    if flex and h < bh:
        h, frames = place(pt, CHART_MIN_H + (bh - h) // flex)
    return {"h_emu": h, "pt": pt, "frames": frames}


def set_block_paras(xml, marker, new_paras):
    """Replace all <a:p>...</a:p> inside the <p:sp> that contains `marker`."""
    i = xml.index(marker)
    start = xml.rfind("<p:sp>", 0, i)
    end = xml.find("</p:sp>", i) + len("</p:sp>")
    block = xml[start:end]
    ps = block.find("<a:p>")
    pe = block.rfind("</a:p>") + len("</a:p>")
    block = block[:ps] + new_paras + block[pe:]
    return xml[:start] + block + xml[end:]


def replace_para(xml, anchor, new_para):
    """Replace the <a:p>...</a:p> containing `anchor` with new_para."""
    i = xml.index(anchor)
    ps = xml.rfind("<a:p>", 0, i)
    pe = xml.find("</a:p>", i) + len("</a:p>")
    return xml[:ps] + new_para + xml[pe:]


def placeholder_sz(xml, marker):
    """Font size in points declared in the <p:sp> block containing `marker`
    (any sz in the placeholder's txBody), or None when nothing is declared
    (the runs inherit the template default)."""
    i = xml.find(marker)
    if i == -1:
        return None
    start = xml.rfind("<p:sp>", 0, i)
    end = xml.find("</p:sp>", i)
    m = re.search(r'sz="(\d+)"', xml[start:end])
    return int(m.group(1)) // 100 if m else None


def block_sz(block):
    """Font size in points declared anywhere in a <p:sp> block, or None."""
    m = re.search(r'sz="(\d+)"', block or "")
    return int(m.group(1)) // 100 if m else None


def add_fill(rpr, fill):
    """`rpr` (an <a:rPr> element) with `fill` (an <a:solidFill>) as its first
    child; run properties that already carry a fill are kept as they are."""
    if not fill or "<a:solidFill>" in rpr:
        return rpr
    if rpr.endswith("/>"):
        return rpr[:-2] + ">" + fill + "</a:rPr>"
    i = rpr.index(">") + 1
    return rpr[:i] + fill + rpr[i:]


def cell_para(runs, algn, sz, base=None, lang=None, fill=""):
    """One paragraph of a table cell. `runs` are (text, bold, italic,
    underline) triples; `sz` is in points and always stamped — a cell has no
    placeholder to inherit a size from. `algn` is stamped only for "ctr"/"r"
    (left is the OOXML default). `fill` is a text colour (<a:solidFill>) for
    cells whose table style gives none. An empty cell is the bare <a:p/>."""
    if not any(r[0] for r in runs):
        return "<a:p/>"
    ppr = '<a:pPr algn="%s"/>' % algn if algn else ""
    rs = "".join('<a:r>%s<a:t>%s</a:t></a:r>'
                 % (add_fill(merge_rpr(base, r[1:], sz * 100 if sz else None, lang),
                              fill), esc(r[0])) for r in runs if r[0])
    return '<a:p>%s%s%s</a:p>' % (ppr, rs, endpara_rpr(lang))


def cell_xml(cell, algn, sz, base, lang, span=(1, 1, False, False), fill=""):
    """One <a:tc> cell: text body first, then the (empty) cell properties -
    the order the TableCellBase type requires. `span` is the cell's entry in
    table_spans: the origin of a merge carries gridSpan/rowSpan, the cells it
    covers carry hMerge/vMerge (and stay in the grid, empty)."""
    gs, rs, hm, vm = span
    attrs = ""
    if gs > 1:
        attrs += ' gridSpan="%d"' % gs
    if rs > 1:
        attrs += ' rowSpan="%d"' % rs
    if hm:
        attrs += ' hMerge="1"'
    if vm:
        attrs += ' vMerge="1"'
    return ('<a:tc%s><a:txBody><a:bodyPr/><a:lstStyle/>%s</a:txBody>'
            '<a:tcPr/></a:tc>' % (attrs, cell_para(cell, algn, sz, base, lang, fill)))


def table_frame_xml(tbl, geom, style_id, sz, base, lang, shape_id, name_no,
                    fill=""):
    """One <p:graphicFrame> carrying an <a:tbl> - the OOXML form PowerPoint
    writes for a slide table (a table can never live inside a text
    placeholder, so it is a separate shape in <p:spTree>)."""
    rows = tbl["cells"]
    aligns = tbl["align"]
    grid = "".join('<a:gridCol w="%d"/>' % w for w in geom["col_w"])
    trs = []
    for ri, row in enumerate(rows):
        cells = "".join(cell_xml(cell, aligns[ci] if ci < len(aligns) else None,
                                 sz, base, lang, tbl["spans"][ri][ci], fill)
                        for ci, cell in enumerate(row))
        trs.append('<a:tr h="%d">%s</a:tr>' % (geom["row_h"][ri], cells))
    flags = ""
    if tbl["first_row"]:
        flags += ' firstRow="1"'
    if tbl["first_col"]:
        flags += ' firstCol="1"'
    if tbl["banded"]:
        flags += ' bandRow="1"'
    return ('<p:graphicFrame><p:nvGraphicFramePr>'
            '<p:cNvPr id="%d" name="Table %d"/>'
            '<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/>'
            '</p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>'
            '<p:xfrm><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></p:xfrm>'
            '<a:graphic><a:graphicData uri="%s"><a:tbl>'
            '<a:tblPr%s><a:tableStyleId>%s'
            '</a:tableStyleId></a:tblPr><a:tblGrid>%s</a:tblGrid>%s'
            '</a:tbl></a:graphicData></a:graphic></p:graphicFrame>'
            % (shape_id, name_no, geom["x"], geom["y"], geom["cx"], geom["cy"],
               TABLE_URI, flags, style_id, grid, "".join(trs)))


# ----------------------------- chart parts -----------------------------------
def col_letter(j):
    """Workbook column letter for a 0-based column index (A..Z)."""
    return chr(ord("A") + j)


def c_str_cache(vals):
    return ('<c:strCache><c:ptCount val="%d"/>%s</c:strCache>'
            % (len(vals), "".join('<c:pt idx="%d"><c:v>%s</c:v></c:pt>'
                                  % (i, esc(v)) for i, v in enumerate(vals))))


def c_num_cache(vals):
    return ('<c:numCache><c:formatCode>General</c:formatCode>'
            '<c:ptCount val="%d"/>%s</c:numCache>'
            % (len(vals), "".join('<c:pt idx="%d"><c:v>%s</c:v></c:pt>'
                                  % (i, fmt_num(v)) for i, v in enumerate(vals))))


def c_series(ch, j, name, vals):
    """One <c:ser>. Colours are theme accents (accent1..accent6, cycling), so
    the template's palette applies; the child order is the one CT_BarSer /
    CT_LineSer / CT_PieSer require."""
    n = len(ch["cats"])
    letter = col_letter(j + 1)
    kind = ch["type"]
    accent = '<a:solidFill><a:schemeClr val="accent%d"/></a:solidFill>' % (j % 6 + 1)
    out = ['<c:idx val="%d"/><c:order val="%d"/>' % (j, j),
           '<c:tx><c:strRef><c:f>Sheet1!$%s$1</c:f>%s</c:strRef></c:tx>'
           % (letter, c_str_cache([name]))]
    if kind in ("column", "bar"):
        out.append('<c:spPr>%s</c:spPr><c:invertIfNegative val="0"/>' % accent)
    elif kind == "line":
        out.append('<c:spPr><a:ln w="28575" cap="rnd">%s<a:round/></a:ln></c:spPr>'
                   '<c:marker><c:symbol val="circle"/><c:size val="7"/>'
                   '<c:spPr>%s<a:ln w="9525">%s</a:ln></c:spPr></c:marker>'
                   % (accent, accent, accent))
    else:
        out.extend('<c:dPt><c:idx val="%d"/><c:bubble3D val="0"/><c:spPr>'
                   '<a:solidFill><a:schemeClr val="accent%d"/></a:solidFill>'
                   '<a:ln w="19050"><a:solidFill><a:schemeClr val="bg1"/>'
                   '</a:solidFill></a:ln></c:spPr></c:dPt>' % (i, i % 6 + 1)
                   for i in range(n))
    if ch["labels"]:
        pie = kind == "pie"
        out.append('<c:dLbls><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>'
                   '<c:showLegendKey val="0"/><c:showVal val="%d"/>'
                   '<c:showCatName val="0"/><c:showSerName val="0"/>'
                   '<c:showPercent val="%d"/><c:showBubbleSize val="0"/></c:dLbls>'
                   % (0 if pie else 1, 1 if pie else 0))
    out.append('<c:cat><c:strRef><c:f>Sheet1!$A$2:$A$%d</c:f>%s</c:strRef></c:cat>'
               % (n + 1, c_str_cache(ch["cats"])))
    out.append('<c:val><c:numRef><c:f>Sheet1!$%s$2:$%s$%d</c:f>%s</c:numRef></c:val>'
               % (letter, letter, n + 1, c_num_cache(vals)))
    if kind == "line":
        out.append('<c:smooth val="0"/>')
    return "<c:ser>%s</c:ser>" % "".join(out)


def c_plot_area(ch):
    """<c:plotArea>: the chart group and, except for a pie, the two axes. A
    horizontal bar chart reverses its category axis so the first row is on
    top, and its value axis crosses at the far end to stay at the bottom."""
    kind = ch["type"]
    sers = "".join(c_series(ch, j, nm, v) for j, (nm, v) in enumerate(ch["series"]))
    axes = '<c:axId val="111111111"/><c:axId val="222222222"/>'
    if kind in ("column", "bar"):
        stacked = ch["stacked"]
        group = ('<c:barChart><c:barDir val="%s"/><c:grouping val="%s"/>'
                 '<c:varyColors val="0"/>%s<c:gapWidth val="%d"/>%s%s</c:barChart>'
                 % ("col" if kind == "column" else "bar",
                    "stacked" if stacked else "clustered", sers,
                    100 if stacked else 80,
                    '<c:overlap val="100"/>' if stacked else "", axes))
    elif kind == "line":
        group = ('<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>'
                 '%s<c:marker val="1"/>%s</c:lineChart>' % (sers, axes))
    else:
        return ('<c:plotArea><c:layout/><c:pieChart><c:varyColors val="1"/>%s'
                '<c:firstSliceAng val="0"/></c:pieChart></c:plotArea>' % sers)
    horiz = kind == "bar"
    cat_ax = ('<c:catAx><c:axId val="111111111"/><c:scaling><c:orientation '
              'val="%s"/></c:scaling><c:delete val="0"/><c:axPos val="%s"/>'
              '<c:numFmt formatCode="General" sourceLinked="1"/>'
              '<c:majorTickMark val="none"/><c:minorTickMark val="none"/>'
              '<c:tickLblPos val="nextTo"/><c:spPr><a:ln w="9525"><a:solidFill>'
              '<a:srgbClr val="BFBFBF"/></a:solidFill></a:ln></c:spPr>'
              '<c:crossAx val="222222222"/><c:crosses val="autoZero"/>'
              '<c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/>'
              '<c:noMultiLvlLbl val="0"/></c:catAx>'
              % ("maxMin" if horiz else "minMax", "l" if horiz else "b"))
    val_ax = ('<c:valAx><c:axId val="222222222"/><c:scaling><c:orientation '
              'val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="%s"/>'
              '<c:majorGridlines><c:spPr><a:ln w="9525"><a:solidFill>'
              '<a:srgbClr val="D9D9D9"/></a:solidFill></a:ln></c:spPr>'
              '</c:majorGridlines><c:numFmt formatCode="General" sourceLinked="1"/>'
              '<c:majorTickMark val="none"/><c:minorTickMark val="none"/>'
              '<c:tickLblPos val="nextTo"/><c:spPr><a:ln><a:noFill/></a:ln></c:spPr>'
              '<c:crossAx val="111111111"/><c:crosses val="%s"/>'
              '<c:crossBetween val="between"/></c:valAx>'
              % ("b" if horiz else "l", "max" if horiz else "autoZero"))
    return "<c:plotArea><c:layout/>%s%s%s</c:plotArea>" % (group, cat_ax, val_ax)


def chart_xml(ch, pt, lang, font, color=""):
    """The ppt/charts/chartN.xml part. `pt` is the text size in points, `font`
    the typeface of the template's body text (None = the theme font), `color`
    its colour (an <a:solidFill>, "" = the chart default), `lang` the declared
    proofing language. The chart background is transparent so the slide
    background shows through."""
    lang_attr = ' lang="%s"' % lang if lang else ""
    face = ('<a:latin typeface="%s"/><a:cs typeface="%s"/>' % (esc(font), esc(font))
            if font else "")
    if ch["title"]:
        title = ('<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr>'
                 '<a:defRPr sz="%d" b="1">%s</a:defRPr></a:pPr><a:r><a:rPr%s sz="%d" b="1">%s</a:rPr>'
                 '<a:t>%s</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/>'
                 '</c:title><c:autoTitleDeleted val="0"/>'
                 % ((pt + 2) * 100, color, lang_attr, (pt + 2) * 100, color,
                    esc(ch["title"])))
    else:
        title = '<c:autoTitleDeleted val="1"/>'
    legend = ch["legend"]
    if legend is None:
        legend = ("right" if ch["type"] == "pie"
                  else "bottom" if len(ch["series"]) > 1 else "off")
    legend_xml = ("" if legend == "off" else
                  '<c:legend><c:legendPos val="%s"/><c:overlay val="0"/></c:legend>'
                  % {"top": "t", "bottom": "b", "left": "l", "right": "r"}[legend])
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<c:chartSpace xmlns:c="%s" '
            'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '<c:date1904 val="0"/><c:roundedCorners val="0"/>'
            '<c:chart>%s%s%s<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/>'
            '</c:chart><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>'
            '<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="%d">%s'
            '</a:defRPr></a:pPr><a:endParaRPr%s/></a:p></c:txPr>'
            '<c:externalData r:id="rId1"><c:autoUpdate val="0"/></c:externalData>'
            '</c:chartSpace>'
            % (CHART_URI, title, c_plot_area(ch), legend_xml, pt * 100,
               color + face, lang_attr))


def chart_frame_xml(geom, rid, shape_id, name_no):
    """The <p:graphicFrame> that places a chart part on the slide."""
    return ('<p:graphicFrame><p:nvGraphicFramePr>'
            '<p:cNvPr id="%d" name="Chart %d"/>'
            '<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/>'
            '</p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>'
            '<p:xfrm><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></p:xfrm>'
            '<a:graphic><a:graphicData uri="%s">'
            '<c:chart xmlns:c="%s" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
            'r:id="%s"/></a:graphicData></a:graphic></p:graphicFrame>'
            % (shape_id, name_no, geom["x"], geom["y"], geom["cx"], geom["cy"],
               CHART_URI, CHART_URI, rid))


def workbook_bytes(ch):
    """The minimal .xlsx behind a chart (what PowerPoint opens for "Edit
    Data"): A2.. categories, B1.. series names, B2.. values on Sheet1, text as
    inline strings. Entries carry a fixed timestamp so the bytes depend only
    on the data."""
    def text_cell(ref, t):
        return ('<c r="%s" t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>'
                % (ref, esc(t)))
    rows = ['<row r="1">%s</row>' % "".join(
        text_cell("%s1" % col_letter(j + 1), nm)
        for j, (nm, _v) in enumerate(ch["series"]))]
    for i, cat in enumerate(ch["cats"]):
        cells = text_cell("A%d" % (i + 2), cat) + "".join(
            '<c r="%s%d"><v>%s</v></c>' % (col_letter(j + 1), i + 2, fmt_num(v[i]))
            for j, (_nm, v) in enumerate(ch["series"]))
        rows.append('<row r="%d">%s</row>' % (i + 2, cells))
    main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    pkg = "http://schemas.openxmlformats.org/package/2006/relationships"
    xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    files = [
        ("[Content_Types].xml", xml +
         '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
         '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
         '<Default Extension="xml" ContentType="application/xml"/>'
         '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
         '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
         '</Types>'),
        ("_rels/.rels", xml + '<Relationships xmlns="%s"><Relationship Id="rId1" '
         'Type="%s/officeDocument" Target="xl/workbook.xml"/></Relationships>'
         % (pkg, rel)),
        ("xl/workbook.xml", xml + '<workbook xmlns="%s" xmlns:r="%s"><sheets>'
         '<sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>'
         % (main, rel)),
        ("xl/_rels/workbook.xml.rels", xml + '<Relationships xmlns="%s">'
         '<Relationship Id="rId1" Type="%s/worksheet" Target="worksheets/sheet1.xml"/>'
         '</Relationships>' % (pkg, rel)),
        ("xl/worksheets/sheet1.xml", xml + '<worksheet xmlns="%s"><sheetData>%s'
         '</sheetData></worksheet>' % (main, "".join(rows))),
    ]
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in files:
            info = zipfile.ZipInfo(name, (1980, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(info, data.encode("utf-8"))
    return buf.getvalue()


def body_color(parts):
    """Colour of the master's level-1 body text as an <a:solidFill> element,
    or "" when the master declares none."""
    xml = parts.get(MASTER)
    if xml is None:
        return ""
    x = xml.decode("utf-8")
    i = x.find("<p:bodyStyle>")
    j = x.find("<a:lvl1pPr", i) if i != -1 else -1
    if j == -1:
        return ""
    d = re.search(r"<a:defRPr\b[^>]*>(.*?)</a:defRPr>",
                  x[j:x.find("</a:lvl1pPr>", j)], re.S)
    f = re.search(r"<a:solidFill>.*?</a:solidFill>", d.group(1), re.S) if d else None
    return f.group(0) if f else ""


def body_font(parts):
    """Typeface of the master's level-1 body text, or None."""
    xml = parts.get(MASTER)
    if xml is None:
        return None
    x = xml.decode("utf-8")
    i = x.find("<p:bodyStyle>")
    j = x.find("<a:lvl1pPr", i) if i != -1 else -1
    m = re.search(r'<a:latin typeface="([^"]+)"', x[j:x.find("</a:lvl1pPr>", j)]) if j != -1 else None
    return m.group(1) if m and not m.group(1).startswith("+") else None


def next_shape_id(xml):
    """One past the largest <p:cNvPr id=".."> in a slide — shape ids must be
    unique per slide."""
    ids = [int(m) for m in re.findall(r'<p:cNvPr\b[^>]*\bid="(\d+)"', xml)]
    return (max(ids) if ids else 1) + 1


def insert_graphic_frames(xml, frames):
    """Append graphic frames at the end of the slide's <p:spTree>."""
    i = xml.rindex("</p:spTree>")
    return xml[:i] + "".join(frames) + xml[i:]


def table_style_id(table_styles_xml):
    """The default table style the template declares
    (<a:tblStyleLst def="{GUID}">), else the Office default GUID."""
    m = re.search(r'<a:tblStyleLst\b[^>]*\bdef="([^"]+)"', table_styles_xml)
    return m.group(1) if m else TABLE_STYLE_DEFAULT


# ----------------------------- slide surgery ---------------------------------
def remove_sp(xml, marker):
    """Drop the <p:sp> containing `marker` (an empty body placeholder; an
    <p:txBody> with no <a:p> is not valid)."""
    block = sp_block(xml, marker)
    return xml.replace(block, "", 1)


def add_body_clone(xml, marker, new_tag):
    """Insert after the body shape containing `marker` a copy of it for the
    placeholder `new_tag`, with its own shape id and without the creationId
    extension (which must be unique)."""
    block = sp_block(xml, marker)
    clone = block.replace(marker, new_tag, 1)
    clone = re.sub(r'(<p:cNvPr id=")[0-9]+(")',
                   lambda m: m.group(1) + str(next_shape_id(xml)) + m.group(2),
                   clone, 1)
    clone = re.sub(r'(<p:cNvPr [^>]*>)<a:extLst>.*?</a:extLst>(</p:cNvPr>)',
                   r'\1\2', clone, 1, flags=re.S)
    clone = re.sub(r'(<p:cNvPr id="[0-9]+" name=")([^"]*)(")',
                   r'\1\2 (2)\3', clone, 1)
    return xml.replace(block, block + clone, 1)


def stamp_orphan(xml, tag, box, block):
    """Give the slide placeholder `tag` its layout placeholder's geometry and
    list style as its own: an explicit <a:xfrm> from `box` and a copy of the
    layout block's <a:lstStyle>. A placeholder with a sentinel idx is linked
    to nothing, so it inherits neither."""
    sp = sp_block(xml, tag)
    new = sp.replace("<p:spPr/>",
                     '<p:spPr><a:xfrm><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/>'
                     "</a:xfrm></p:spPr>" % tuple(box), 1)
    lst = re.search(r"<a:lstStyle>.*?</a:lstStyle>", block or "", re.S)
    if lst:
        new = new.replace("<a:lstStyle/>", lst.group(0), 1)
    return xml.replace(sp, new, 1)


def rels_name(slide_name):
    base = slide_name.split("/")[-1]            # "slide5.xml"
    return "ppt/slides/_rels/" + base[:-4] + ".xml.rels"


# ----------------------------- structure.md parser --------------------------
BLOCK_RE = re.compile(r"^:::\s*(left|right)\s*$", re.I)
LAYOUT_RE = re.compile(r"^layout:\s*(.+)$", re.I)


def parse_column(lines):
    """Items of one column: (level, runs, num) text items, ("table", table)
    and ("chart", chart). An options line `{...}` applies to the pipe table
    that follows it and is an error anywhere else."""
    parsed = []
    pending = None
    j = 0
    while j < len(lines):
        s = lines[j].strip()
        if not s:
            j += 1
            continue
        if is_options_line(s):
            if pending is not None:
                raise SystemExit("ERROR: two options lines in a row (%r)" % s)
            pending = parse_options(s)
            j += 1
            continue
        if s.startswith("|"):
            rows = []
            while j < len(lines) and lines[j].strip().startswith("|"):
                rows.append(lines[j])
                j += 1
            if pending is not None and "chart" in pending:
                parsed.append(("chart", parse_chart(rows, pending)))
            else:
                parsed.append(("table", parse_table(rows, pending)))
            pending = None
            continue
        if pending is not None:
            raise SystemExit("ERROR: an options line must be followed by a "
                             "pipe table, got %r" % s)
        nm = parse_numbered(s)
        if s.startswith("- "):
            parsed.append((1, parse_inline(s[2:]), None))
        elif nm:
            # the as-written number is replaced by its run mark below
            parsed.append((1, parse_inline(nm[2]), (nm[0], nm[1])))
        else:
            parsed.append((0, parse_inline(s), None))
        j += 1
    if pending is not None:
        raise SystemExit("ERROR: an options line must be followed by a pipe table")
    number_slide_items(parsed)
    return parsed


def parse_slide(body):
    """(title, columns, layout) of one `# slide` section. `columns` holds one
    item list, or two when the section has `::: left` and `::: right` blocks
    (a block ends at `:::`, at the next block or at the end of the section);
    `layout` is the `layout: <name>` line or None."""
    title = None
    layout = None
    top = []
    blocks = []
    inside = False
    for b in body:
        if b.startswith("## "):
            title = b[3:].strip()
            continue
        s = b.strip()
        m = BLOCK_RE.match(s)
        if m:
            blocks.append((m.group(1).lower(), []))
            inside = True
        elif s == ":::":
            inside = False
        elif not inside and LAYOUT_RE.match(s):
            layout = LAYOUT_RE.match(s).group(1).strip()
        else:
            (blocks[-1][1] if inside else top).append(b)
    if not blocks:
        return (title, [parse_column(top)], layout)
    if any(t.strip() for t in top):
        raise SystemExit("ERROR: slide %r has text outside its ::: left / "
                         "::: right blocks" % title)
    if [n for n, _l in blocks] != ["left", "right"]:
        raise SystemExit("ERROR: slide %r needs exactly one ::: left block "
                         "followed by one ::: right block" % title)
    return (title, [parse_column(l) for _n, l in blocks], layout)


def parse_md(text):
    lines = text.splitlines()
    n = len(lines)
    i = 0
    body_size = None
    lang = None
    if n and lines[0].strip() == "---":
        i = 1
        fm = {}
        while i < n and lines[i].strip() != "---":
            if ":" in lines[i]:
                k, v = lines[i].split(":", 1)
                fm[k.strip()] = v.strip()
            i += 1
        if i < n:
            i += 1
        try:
            body_size = int(fm["body_size"]) if "body_size" in fm else None
        except ValueError:
            body_size = None
        lang_in = fm.get("lang", "").strip()
        if lang_in:
            if not re.match(r"^[A-Za-z]{2,8}(-[A-Za-z0-9]{1,8})*$", lang_in):
                raise SystemExit(
                    "ERROR: invalid lang %r in structure.md front matter "
                    "(expected a BCP-47 tag like ru-RU)" % lang_in)
            lang = lang_in
    # body_size drives the plan slide (items at body_size - 1), the table
    # cells (also body_size - 1: a table has no placeholder to inherit a size
    # from) and the default chart text size; without it the template's own
    # body size is used (build fills it in). Body text takes its size from the
    # template (resolve_body_size) and stamps no size of its own. `lang`, when
    # declared, is stamped on every generated run and end mark.
    TOC = body_size - 1 if body_size else None

    sections = []
    cur = None
    for ln in lines[i:]:
        if ln.startswith("# "):
            if cur is not None:
                sections.append(cur)
            cur = (ln[2:].strip(), [])
        elif cur is not None:
            cur[1].append(ln)
    if cur is not None:
        sections.append(cur)

    doc = {"body_size": body_size, "TOC": TOC, "table_pt": TOC, "lang": lang,
           "title": None, "plan": None, "plan_title": None, "slides": [],
           "closing": None}
    for header, body in sections:
        if header == "title":
            doc["title"] = " ".join(b.strip() for b in body if b.strip()).strip()
        elif header == "plan":
            texts, raw = [], []
            for b in (b.strip() for b in body):
                if not b:
                    continue
                if b.startswith("## "):
                    doc["plan_title"] = b[3:].strip()
                    continue
                nm = parse_numbered(b)
                texts.append(nm[2] if nm else b)
                raw.append((nm[0], nm[1]) if nm else None)
            doc["plan"] = list(zip(texts, group_numbering(raw)))
        elif header == "closing":
            doc["closing"] = " ".join(b.strip() for b in body if b.strip()).strip()
        elif header == "slide":
            doc["slides"].append(parse_slide(body))
    return doc


# ----------------------------- template font -------------------------------
def resolve_body_size(parts):
    """Effective body font size in points of the template's content slide:
    slide3 body placeholder -> its slide layout -> OOXML default 18 pt."""
    xml = parts[CONTENT_SLIDE].decode("utf-8")
    sz = placeholder_sz(xml, BODY_MARKER)
    if sz:
        return sz
    rels = parts[CONTENT_RELS].decode("utf-8")
    m = re.search(r'Target="([^"]*slideLayout\d+\.xml)"', rels)
    if m:
        layout_name = "ppt/slideLayouts/" + m.group(1).split("/")[-1]
        layout = parts.get(layout_name)
        if layout:
            lx = layout.decode("utf-8")
            pm = re.search(r'<p:ph[^>]*idx="1"[^>]*/>', lx)
            if pm:
                start = lx.rfind("<p:sp>", 0, pm.start())
                end = lx.find("</p:sp>", pm.start())
                sm = re.search(r'sz="(\d+)"', lx[start:end])
                if sm:
                    return int(sm.group(1)) // 100
    return 18


# ----------------------------- package slots ---------------------------------
def unesc(text):
    """Text of an XML text node with its character references resolved."""
    return (text.replace("&lt;", "<").replace("&gt;", ">")
            .replace("&quot;", '"').replace("&apos;", "'").replace("&amp;", "&"))


def package_slots(parts, pres, pres_rels):
    """Where the template's four slides sit in presentation.xml and where the
    clones may go: ({slide file: (sldId id, rId)} for slide1..slide4,
    first free sldId, first free rId number, first free slide file number).
    Everything is read from the template, so a template with its own ids and
    relationships (or more parts) is cloned without collisions."""
    slots = {}
    for n in (1, 2, 3, 4):
        fn = "ppt/slides/slide%d.xml" % n
        rel = re.search(r'<Relationship\b[^>]*Target="slides/slide%d\.xml"[^>]*>' % n,
                        pres_rels)
        rid = re.search(r'\bId="([^"]+)"', rel.group(0)) if rel else None
        sld = (re.search(r'<p:sldId\b[^>]*\br:id="%s"[^>]*>' % re.escape(rid.group(1)), pres)
               if rid else None)
        sid = re.search(r'\bid="([0-9]+)"', sld.group(0)) if sld else None
        if not sid:
            raise SystemExit("ERROR: template slide%d.xml is not listed in "
                             "presentation.xml (the first four slides must be "
                             "title, plan, content, closing)" % n)
        slots[fn] = (int(sid.group(1)), rid.group(1))
    sids = [int(v) for v in re.findall(r'<p:sldId\b[^>]*\bid="([0-9]+)"', pres)]
    rids = [int(v) for v in re.findall(r'\bId="rId([0-9]+)"', pres_rels)]
    files = [int(m.group(1)) for m in
             (re.match(r"^ppt/slides/slide([0-9]+)\.xml$", n) for n in parts) if m]
    return (slots, max(sids + [255]) + 1, max(rids + [0]) + 1,
            max(files + [4]) + 1)


def plan_title_text(plan_xml):
    """The plan slide's title as the template has it, minus a trailing
    parenthesised remark (a template hint such as "Plan (if any)")."""
    block = sp_block(plan_xml, TITLE_MARKER) or ""
    text = unesc("".join(re.findall(r"<a:t>([^<]*)</a:t>", block))).strip()
    return re.sub(r"\s*\([^()]*\)\s*$", "", text).strip() or "Plan"


# ----------------------------- engine ---------------------------------------
def build(structure_text, template_path, out_path):
    doc = parse_md(structure_text)
    n = len(doc["slides"])
    if n < 1:
        raise SystemExit("ERROR: no '# slide' sections found in structure.md")

    with zipfile.ZipFile(template_path, "r") as z:
        parts = {name: z.read(name) for name in z.namelist()}

    pres = parts["ppt/presentation.xml"].decode("utf-8")
    pres_rels = parts["ppt/_rels/presentation.xml.rels"].decode("utf-8")
    ct = parts["[Content_Types].xml"].decode("utf-8")
    content_xml = parts[CONTENT_SLIDE].decode("utf-8")
    content_rels = parts[CONTENT_RELS].decode("utf-8")
    body_pt = resolve_body_size(parts)
    if doc["body_size"] is None:
        doc["body_size"] = body_pt
        doc["TOC"] = doc["table_pt"] = body_pt - 1
    slots, next_sid, next_rid, next_file = package_slots(parts, pres, pres_rels)
    master_xml = parts[MASTER].decode("utf-8") if MASTER in parts else ""

    # Where each slide's columns go (layout file, placeholder tags, boxes).
    # Resolved before anything is written so a bad layout fails early.
    default = default_body(parts, content_xml, content_rels)
    plans = [plan_slide(parts, default, body_pt, layout_key, len(columns))
             for _t, columns, layout_key in doc["slides"]]
    doc["geom"] = [[(ph["box"], ph["pt"], ph["spacing"]) for ph in plan["phs"]]
                   for plan in plans]

    # 1) clone the content slide (n-1 clones -> slide5 .. slide(n+3))
    clones = []  # (filename, sldId, rId)
    for i in range(1, n):
        fn = "ppt/slides/slide%d.xml" % (next_file + i - 1)
        rid = "rId%d" % (next_rid + i - 1)
        sid = next_sid + i - 1
        clones.append((fn, sid, rid))
        parts[fn] = content_xml.encode("utf-8")
        parts[rels_name(fn)] = content_rels.encode("utf-8")

    # 2) register clones in [Content_Types].xml and presentation.xml.rels
    ct_slide = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"
    rel_slide = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"
    ct_add = "".join('<Override PartName="/%s" ContentType="%s"/>' % (fn, ct_slide)
                     for fn, _sid, _rid in clones)
    rel_add = "".join('<Relationship Id="%s" Type="%s" Target="%s"/>'
                      % (rid, rel_slide, fn[len("ppt/"):])
                      for fn, _sid, rid in clones)
    if ct_add:
        ct = ct.replace("</Types>", ct_add + "</Types>")
    if rel_add:
        pres_rels = pres_rels.replace("</Relationships>", rel_add + "</Relationships>")

    # 3) rebuild sldIdLst: title, plan?, content..., closing? (closing last)
    def sld(fn):
        return '<p:sldId id="%d" r:id="%s"/>' % slots[fn]
    sldids = [sld(TITLE_SLIDE)]
    if doc["plan"] is not None:
        sldids.append(sld(PLAN_SLIDE))
    sldids.append(sld(CONTENT_SLIDE))
    for _fn, sid, rid in clones:
        sldids.append('<p:sldId id="%d" r:id="%s"/>' % (sid, rid))
    if doc["closing"] is not None:
        sldids.append(sld(CLOSING_SLIDE))
    pres = re.sub(r"<p:sldIdLst>.*?</p:sldIdLst>",
                  "<p:sldIdLst>%s</p:sldIdLst>" % "".join(sldids), pres)

    TOC = doc["TOC"]

    # 4) fill the title slide (topic line; bold, size inherited from the
    #    template; font of the original anchor run is kept). Russian standard:
    #    with the RU anchor the engine keeps the "На тему: <тема>" format by
    #    prepending "На тему: "; with the EN anchor the whole line is
    #    replaced by the topic.
    title_xml = parts[TITLE_SLIDE].decode("utf-8")
    topic = doc["title"] or ""
    anchor = topic_anchor(title_xml)
    topic_text = (anchor + ": " + topic) if anchor == TITLE_ANCHOR_RU else topic
    topic_base = first_rpr(title_xml, anchor)
    topic_para = ('<a:p><a:r>%s<a:t>%s</a:t></a:r>%s</a:p>'
                  % (merge_rpr(topic_base, (True, False, False), None, doc["lang"]),
                     esc(topic_text), endpara_rpr(doc["lang"])))
    title_xml = replace_para(title_xml, anchor, topic_para)
    parts[TITLE_SLIDE] = title_xml.encode("utf-8")

    # 5) fill the plan slide (items stamped at body_size - 1)
    if doc["plan"] is not None:
        plan_xml = parts[PLAN_SLIDE].decode("utf-8")
        plan_title_base = first_rpr(plan_xml, TITLE_MARKER)
        plan_xml = set_block_paras(
            plan_xml, TITLE_MARKER,
            title_para(doc["plan_title"] or plan_title_text(plan_xml),
                       base=plan_title_base, lang=doc["lang"]))
        plan_body_base = first_rpr(plan_xml, BODY_MARKER)
        plan_fix = list_fixes(
            layout_body_block(parts, parts[rels_name(PLAN_SLIDE)].decode("utf-8"), 1),
            master_xml)
        plan_xml = set_block_paras(plan_xml, BODY_MARKER,
                                  "".join(toc_para(t, TOC, base=plan_body_base,
                                                 lang=doc["lang"], num=num,
                                                 add=plan_fix["num"][0] if num else ("", ""))
                                         for t, num in doc["plan"]))
        parts[PLAN_SLIDE] = plan_xml.encode("utf-8")

    # 6) fill the content slides (slide3 + clones, in order). Body runs
    #    inherit the template body font (no sz stamped); the original runs'
    #    typeface/lang are kept via first_rpr. Each column fills its own body
    #    placeholder (dropped when the column has no text); its tables and
    #    charts become <p:graphicFrame> shapes below the text, table cells at
    #    the size column_flow chose, charts wired to a chart part, its
    #    embedded workbook and a slide relationship.
    content_order = [CONTENT_SLIDE] + [fn for fn, _sid, _rid in clones]
    style_default = table_style_id(
        parts.get("ppt/tableStyles.xml", b"").decode("utf-8", "ignore"))
    font = body_font(parts)
    color = body_color(parts)
    chart_no = 0
    ct_extra = ""
    for (stitle, columns, _lk), plan, fn in zip(doc["slides"], plans, content_order):
        x = content_xml
        phs = plan["phs"]
        rels = content_rels
        if plan["file"]:
            x = x.replace(BODY_MARKER, phs[0]["tag"], 1)
            if len(phs) == 2:
                x = add_body_clone(x, phs[0]["tag"], phs[1]["tag"])
            for ph in phs:
                if ph["orphan"]:
                    x = stamp_orphan(x, ph["tag"], ph["box"], ph["block"])
            rels = re.sub(r'Target="\.\./slideLayouts/slideLayout[0-9]+\.xml"',
                          'Target="../slideLayouts/%s"' % plan["file"].split("/")[-1],
                          rels, 1)
        ct_base = first_rpr(x, TITLE_MARKER)
        x = set_block_paras(x, TITLE_MARKER,
                           title_para(stitle or "", base=ct_base,
                                     lang=doc["lang"]))
        flows, bases = [], []
        for items, ph in zip(columns, phs):
            tag = ph["tag"]
            flows.append(column_flow(items, ph["box"], ph["pt"], doc["table_pt"],
                                     ph["spacing"]))
            base = first_rpr(x, tag)
            bases.append(base)
            fix = list_fixes(ph["block"], master_xml)

            def add_for(it):
                if it[2]:
                    return fix["num"][it[0]]
                return fix["bullet"] if it[0] == 1 else fix["plain"]
            text = "".join(para(it[0], it[1], base=base, lang=doc["lang"],
                                num=it[2], add=add_for(it))
                           for it in items if it[0] in (0, 1))
            x = set_block_paras(x, tag, text) if text else remove_sp(x, tag)
        nid = next_shape_id(x)
        rid_no = 1 + max([int(r) for r in re.findall(r'Id="rId([0-9]+)"', rels)] or [0])
        frames, tno, cno, rel_add = [], 0, 0, ""
        for flow, base in zip(flows, bases):
            for kind, obj, geom in flow["frames"]:
                if kind == "table":
                    tno += 1
                    # a style without fills leaves the text colour to the
                    # slide, which for a table is not the body text's
                    frames.append(table_frame_xml(
                        obj, geom, TABLE_STYLES[obj["style"]] or style_default,
                        geom["pt"], base, doc["lang"], nid, tno,
                        color if TABLE_STYLES[obj["style"]] else ""))
                else:
                    chart_no += 1
                    cno += 1
                    rid = "rId%d" % rid_no
                    rid_no += 1
                    cpt = obj["pt"] or max(TABLE_MIN_PT,
                                           doc["table_pt"] - CHART_PT_BELOW_TABLE)
                    name = "chart%d.xml" % chart_no
                    book = "Microsoft_Excel_Sheet%d.xlsx" % chart_no
                    parts["ppt/charts/" + name] = chart_xml(
                        obj, cpt, doc["lang"], font, color).encode("utf-8")
                    parts["ppt/charts/_rels/%s.rels" % name] = (
                        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" '
                        'Target="../embeddings/%s"/></Relationships>' % book).encode("utf-8")
                    parts["ppt/embeddings/" + book] = workbook_bytes(obj)
                    ct_extra += ('<Override PartName="/ppt/charts/%s" ContentType='
                                 '"application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>'
                                 % name)
                    rel_add += ('<Relationship Id="%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" '
                                'Target="../charts/%s"/>' % (rid, name))
                    frames.append(chart_frame_xml(geom, rid, nid, cno))
                nid += 1
        if frames:
            x = insert_graphic_frames(x, frames)
        if rel_add:
            rels = rels.replace("</Relationships>", rel_add + "</Relationships>")
        parts[fn] = x.encode("utf-8")
        parts[rels_name(fn)] = rels.encode("utf-8")

    if chart_no:
        if 'Extension="xlsx"' not in ct:
            ct_extra = ('<Default Extension="xlsx" ContentType="application/'
                        'vnd.openxmlformats-officedocument.spreadsheetml.sheet"/>'
                        + ct_extra)
        ct = ct.replace("</Types>", ct_extra + "</Types>")

    # 7) fill the closing slide (title stamped at the template's declared size)
    if doc["closing"] is not None:
        close_xml = parts[CLOSE_SLIDE].decode("utf-8")
        close_sz = placeholder_sz(close_xml, TITLE_MARKER) or TITLE_PT
        close_base = first_rpr(close_xml, TITLE_MARKER)
        close_xml = set_block_paras(close_xml, TITLE_MARKER,
                                    title_para(doc["closing"], sz=close_sz,
                                              base=close_base,
                                              lang=doc["lang"]))
        parts[CLOSE_SLIDE] = close_xml.encode("utf-8")

    # put the modified top-level parts back into `parts`
    parts["ppt/presentation.xml"] = pres.encode("utf-8")
    parts["ppt/_rels/presentation.xml.rels"] = pres_rels.encode("utf-8")
    parts["[Content_Types].xml"] = ct.encode("utf-8")

    # 8) re-zip ([Content_Types].xml first, forward-slash arcnames)
    if os.path.exists(out_path):
        os.remove(out_path)
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct.encode("utf-8"))
        for name in parts:
            if name == "[Content_Types].xml":
                continue
            z.writestr(name, parts[name])

    return doc, len(content_order), body_pt


# ----------------------------- checks ---------------------------------------
def fit_check(doc):
    """Heuristic per-column height estimate at the column's template body
    size: its text lines plus the tables and charts column_flow places in its
    box. Returns (slide, est_in, box_in) for every column that overflows; a
    two-column slide's title says which column."""
    out = []
    for (stitle, columns, _lk), geoms in zip(doc["slides"], doc["geom"]):
        for ci, (items, (box, pt, spacing)) in enumerate(zip(columns, geoms)):
            h_in = (column_flow(items, box, pt, doc["table_pt"], spacing)["h_emu"]
                    / EMU_PER_INCH)
            box_in = box[3] / EMU_PER_INCH
            if h_in > box_in:
                label = stitle if len(columns) == 1 else "%s (%s)" % (
                    stitle, ("left", "right")[ci])
                out.append((label, round(h_in, 2), round(box_in, 2)))
    return out


def validate(out_path):
    """Return a list of problem strings (empty == pass)."""
    problems = []
    with zipfile.ZipFile(out_path, "r") as z:
        names = set(z.namelist())
        for name in z.namelist():
            if name.endswith(".xml") or name.endswith(".rels"):
                try:
                    ET.fromstring(z.read(name))
                except Exception as e:
                    problems.append("%s: %s" % (name, e))
        pres = z.read("ppt/presentation.xml").decode("utf-8")
        rels = z.read("ppt/_rels/presentation.xml.rels").decode("utf-8")
        rid_target = dict(re.findall(r'Id="(rId\d+)"\s+Type="[^"]+"\s+Target="([^"]+)"', rels))
        for rid in re.findall(r'r:id="(rId\d+)"', pres):
            tgt = rid_target.get(rid)
            if tgt is None:
                problems.append("sldId rId %s has no relationship" % rid)
            elif "ppt/" + tgt not in names:
                problems.append("sldId %s -> %s (missing part)" % (rid, tgt))
        # every internal relationship target and every content-type override
        # must name a part that exists
        for name in z.namelist():
            if not name.endswith(".rels"):
                continue
            base = posixpath.dirname(posixpath.dirname(name))
            for rel in re.findall(r"<Relationship\b[^>]*>", z.read(name).decode("utf-8")):
                tm = re.search(r'Target="([^"]+)"', rel)
                if not tm or 'TargetMode="External"' in rel:
                    continue
                tgt = tm.group(1)
                full = (tgt.lstrip("/") if tgt.startswith("/")
                        else posixpath.normpath(posixpath.join(base, tgt)))
                if full not in names:
                    problems.append("%s -> %s (missing part)" % (name, full))
        # ids that must be unique: relationship ids per .rels file, sldId ids,
        # shape ids per slide
        for name in z.namelist():
            if name.endswith(".rels"):
                ids = re.findall(r'<Relationship\b[^>]*\bId="([^"]+)"',
                                 z.read(name).decode("utf-8"))
                for dup in sorted({i for i in ids if ids.count(i) > 1}):
                    problems.append("%s: duplicate relationship id %s" % (name, dup))
            elif re.match(r"^ppt/slides/slide[0-9]+\.xml$", name):
                ids = re.findall(r'<p:cNvPr id="([0-9]+)"',
                                 z.read(name).decode("utf-8"))
                for dup in sorted({i for i in ids if ids.count(i) > 1}):
                    problems.append("%s: duplicate shape id %s" % (name, dup))
        sld_ids = re.findall(r'<p:sldId\b[^>]*\bid="([0-9]+)"', pres)
        for dup in sorted({i for i in sld_ids if sld_ids.count(i) > 1}):
            problems.append("presentation.xml: duplicate sldId %s" % dup)
        ct = z.read("[Content_Types].xml").decode("utf-8")
        for part in re.findall(r'<Override\b[^>]*PartName="/([^"]+)"', ct):
            if part not in names:
                problems.append("[Content_Types].xml override %s (missing part)" % part)
    return problems


# ----------------------------- main -----------------------------------------
def main():
    if len(sys.argv) != 4:
        print("usage: python engine.py <structure.md> <template.pptx> <output.pptx>")
        return 2
    struct_path, template_path, out_path = sys.argv[1:4]
    with open(struct_path, "r", encoding="utf-8") as f:
        structure_text = f.read()

    doc, ncontent, body_pt = build(structure_text, template_path, out_path)
    problems = validate(out_path)
    over = fit_check(doc)

    print("built %s" % out_path)
    print("  slides: %d (title + %s + %d content + %s)"
          % (2 + ncontent + (1 if doc["closing"] else 0),
             ("1 plan" if doc["plan"] is not None else "no plan"),
             ncontent,
             ("1 closing" if doc["closing"] is not None else "no closing")))
    print("  fonts: body = template (est. %d pt), plan items = %d pt (body_size=%d)"
          % (body_pt, doc["TOC"], doc["body_size"]))
    items = [it for _t, columns, _lk in doc["slides"]
             for column in columns for it in column]
    n_tables = sum(1 for it in items if it[0] == "table")
    n_charts = sum(1 for it in items if it[0] == "chart")
    n_cols = sum(1 for _t, columns, _lk in doc["slides"] if len(columns) == 2)
    print("  tables: %d (cells at %d pt)" % (n_tables, doc["table_pt"]))
    if n_charts:
        print("  charts: %d" % n_charts)
    if n_cols:
        print("  two-column slides: %d" % n_cols)
    print("  lang: %s" % ((doc["lang"] + " (stamped, dirty=0)") if doc["lang"]
                           else "not set (template default)"))
    if problems:
        print("VALIDATION FAILED:")
        for p in problems:
            print("  - " + p)
        return 1
    print("  validation: PASSED")
    if over:
        print("  fit warnings (est. text + table + chart height > box at %d pt):"
              % body_pt)
        for stitle, h, box in over:
            print("    - %r: est %0.2fin > %0.2fin box" % (stitle, h, box))
    else:
        print("  fit: all content slides within the box (estimate)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
