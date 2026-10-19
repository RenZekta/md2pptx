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
* A body line starting with "|" is a pipe table: it becomes a real
  <p:graphicFrame><a:tbl> shape appended to the slide's <p:spTree> (a table
  cannot live in a text placeholder), placed below the slide's text lines
  under the same average-character height estimate fit_check uses.
* The template must contain: a title slide (slide1, freeform, topic line
  anchored by the string TITLE_ANCHOR), a plan slide (slide2), a content slide
  (slide3, the clone prototype, with <p:ph type="title"/> and <p:ph idx="1"/>),
  and a closing slide (slide4).
* Content slides are cloned from slide3 (new parts registered in
  [Content_Types].xml, presentation.xml.rels and the sldIdLst), then the
  sldIdLst is reordered so the closing slide is last.
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
import zipfile
from xml.etree import ElementTree as ET  # parse-check only

# ----------------------------- geometry / config -----------------------------
EMU_PER_INCH = 914400.0
SLIDE_W = 9144000.0                  # EMU  (4:3, 10in x 7.5in)
SLIDE_H = 6858000.0                  # EMU
BODY_W = 8229600.0                   # body box width  (EMU)
BODY_H = 4525963.0                   # body box height (EMU)
BODY_W_IN = BODY_W / EMU_PER_INCH    # 9.0 in
BODY_H_IN = BODY_H / EMU_PER_INCH     # 4.95 in
BODY_OFF_X = 457200                  # body box offset (EMU, reference template)
BODY_OFF_Y = 1600200

# Tables: a <p:graphicFrame> carrying an <a:tbl> (the OOXML form PowerPoint
# itself writes; see python-pptx's graphicFrame output).
TABLE_URI = "http://schemas.openxmlformats.org/drawingml/2006/table"
TABLE_STYLE_DEFAULT = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"  # Office default
TABLE_GAP = 137160                   # gap above a table (0.15 in, EMU)
TABLE_CELL_PAD = 91440               # cell top+bottom inset (2 * 0.05 in)
TABLE_MIN_COL_CHARS = 6              # column width floor, in characters

CONTENT_SLIDE = "ppt/slides/slide3.xml"          # clone prototype
CONTENT_RELS = "ppt/slides/_rels/slide3.xml.rels"
TITLE_SLIDE = "ppt/slides/slide1.xml"
PLAN_SLIDE = "ppt/slides/slide2.xml"
CLOSE_SLIDE = "ppt/slides/slide4.xml"

TITLE_ANCHOR_RU = "На тему"         # topic-line anchor in the title slide
TITLE_ANCHOR_EN = "Topic"           # English title-slide topic-line anchor
TITLE_ANCHOR = TITLE_ANCHOR_RU      # kept for compatibility
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

SLIDE_ID_BASE = 274                  # new sldId ids (existing: 270,271,272,273)
RID_BASE = 11                        # new rIds (existing: rId1..rId10)

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
    stars of `***x***` literal, as before. Marks must balance within a line;
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
    text as misspelled. Italic is `i="1"`, underline `u="sng"`."""
    bold, italic, underline = fmt
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


def para(level, runs, sz=None, base=None, lang=None):
    """Body paragraph at `level` (0 = plain line, 1 = bullet). `runs` are the
    (text, bold, italic, underline) triples of parse_inline. The font size
    is stamped only when `sz` is given; otherwise runs inherit the
    template's body font. `base` is the original paragraph's first-run rPr:
    its typeface elements (a:latin/a:cs/a:ea) are kept on every new run so
    the replaced text keeps the template's font. `lang` (BCP-47) is stamped
    on every run and the paragraph end mark so PowerPoint does not flag the
    text."""
    ppr = '<a:pPr lvl="%d"/>' % level
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


def toc_para(text, sz, base=None, lang=None):
    return ('<a:p><a:pPr lvl="0"/>'
            '<a:r>%s<a:t>%s</a:t></a:r>%s</a:p>'
            % (merge_rpr(base, None, sz * 100, lang), esc(text),
               endpara_rpr(lang)))


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


def parse_table(lines):
    """Parse pipe-table lines into {"header": bool, "align": [...],
    "cells": [[run-tuples]]}.

    A line is a row when it starts with "|"; cells are split on "|" and
    trimmed, and the outer pipes are optional. A row of only |, -, : and
    spaces is the header separator: it is dropped, marks the row above as the
    header (rendered bold) and sets each column's alignment from its colons
    (`:--` left, `--:` right, `:-:` centre). Short rows are padded to the
    widest one; cells carry the same inline marks as body lines."""
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
    cells = []
    for ri, row in enumerate(rows):
        rr = []
        for cell in row:
            runs = parse_inline(cell)
            if header and ri == 0:
                runs = [(t, True, it, ul) for (t, _b, it, ul) in runs]
            rr.append(runs)
        cells.append(rr)
    return {"header": header, "align": aligns + [None] * (width - len(aligns)),
            "cells": cells}


def table_geom(tbl, y, table_pt):
    """Geometry (EMU) of one table in the body box: full body width,
    column widths proportional to the longest cell of each column (floored at
    TABLE_MIN_COL_CHARS characters, leftover EMU to the leftmost columns),
    row heights from the wrapped-line estimate at the table text size plus the
    cell insets. `y` is where the table starts in the body box."""
    ncol = len(tbl["align"])
    weights = []
    for ci in range(ncol):
        w = TABLE_MIN_COL_CHARS
        for row in tbl["cells"]:
            w = max(w, len(runs_text(row[ci])))
        weights.append(w)
    total = int(BODY_W)
    s = sum(weights)
    col_w = [total * w // s for w in weights]
    for ci in range(total - sum(col_w)):
        col_w[ci] += 1
    row_h = []
    for row in tbl["cells"]:
        lines = 1
        for ci, cell in enumerate(row):
            lines = max(lines, est_lines(runs_text(cell),
                                         col_w[ci] / EMU_PER_INCH, table_pt))
        row_h.append(lines * line_emu(table_pt) + TABLE_CELL_PAD)
    return {"x": BODY_OFF_X, "y": y, "cx": sum(col_w), "cy": sum(row_h),
            "col_w": col_w, "row_h": row_h}


def slide_layout(parsed, body_pt, table_pt):
    """Vertical flow of one content slide, from the body-box top.

    A text placeholder flows all of its own paragraphs from the top of the
    body box, so a table — a separate <p:graphicFrame> shape — always sits
    below them: the text height is estimated first (the same
    average-character heuristic as fit_check), then the tables stack in
    document order, each TABLE_GAP under the previous block. Returns
    {"h_emu": total height, "tables": [(tbl, geom)]}, the estimate fit_check
    compares with the body box."""
    text_emu = 0
    for item in parsed:
        if item[0] != "table":
            text_emu += (est_lines(runs_text(item[1]), BODY_W_IN, body_pt)
                         * line_emu(body_pt))
    y = int(BODY_OFF_Y) + text_emu
    h = text_emu
    tables = []
    for item in parsed:
        if item[0] == "table":
            y += TABLE_GAP
            h += TABLE_GAP
            geom = table_geom(item[1], y, table_pt)
            tables.append((item[1], geom))
            y += geom["cy"]
            h += geom["cy"]
    return {"h_emu": h, "tables": tables}


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


def cell_para(runs, algn, sz, base=None, lang=None):
    """One paragraph of a table cell. `runs` are (text, bold, italic,
    underline) triples; `sz` is in points and always stamped — a cell has no
    placeholder to inherit a size from. `algn` is stamped only for "ctr"/"r"
    (left is the OOXML default). An empty cell is the bare <a:p/>."""
    if not any(r[0] for r in runs):
        return "<a:p/>"
    ppr = '<a:pPr algn="%s"/>' % algn if algn else ""
    rs = "".join('<a:r>%s<a:t>%s</a:t></a:r>'
                 % (merge_rpr(base, r[1:], sz * 100 if sz else None, lang),
                    esc(r[0])) for r in runs if r[0])
    return '<a:p>%s%s%s</a:p>' % (ppr, rs, endpara_rpr(lang))


def cell_xml(cell, algn, sz, base, lang):
    """One <a:tc> cell: text body first, then the (empty) cell properties —
    the order the TableCellBase type requires."""
    return ('<a:tc><a:txBody><a:bodyPr/><a:lstStyle/>%s</a:txBody>'
            '<a:tcPr/></a:tc>' % cell_para(cell, algn, sz, base, lang))


def table_frame_xml(tbl, geom, style_id, sz, base, lang, shape_id, name_no):
    """One <p:graphicFrame> carrying an <a:tbl> — the OOXML form PowerPoint
    writes for a slide table (a table can never live inside a text
    placeholder, so it is a separate shape in <p:spTree>)."""
    rows = tbl["cells"]
    aligns = tbl["align"]
    grid = "".join('<a:gridCol w="%d"/>' % w for w in geom["col_w"])
    trs = []
    for ri, row in enumerate(rows):
        cells = "".join(cell_xml(cell, aligns[ci] if ci < len(aligns) else None,
                                 sz, base, lang)
                        for ci, cell in enumerate(row))
        trs.append('<a:tr h="%d">%s</a:tr>' % (geom["row_h"][ri], cells))
    return ('<p:graphicFrame><p:nvGraphicFramePr>'
            '<p:cNvPr id="%d" name="Table %d"/>'
            '<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/>'
            '</p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>'
            '<p:xfrm><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></p:xfrm>'
            '<a:graphic><a:graphicData uri="%s"><a:tbl>'
            '<a:tblPr firstRow="1" bandRow="1"><a:tableStyleId>%s'
            '</a:tableStyleId></a:tblPr><a:tblGrid>%s</a:tblGrid>%s'
            '</a:tbl></a:graphicData></a:graphic></p:graphicFrame>'
            % (shape_id, name_no, geom["x"], geom["y"], geom["cx"], geom["cy"],
               TABLE_URI, style_id, grid, "".join(trs)))


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


def rels_name(slide_name):
    base = slide_name.split("/")[-1]            # "slide5.xml"
    return "ppt/slides/_rels/" + base[:-4] + ".xml.rels"


# ----------------------------- structure.md parser --------------------------
def parse_md(text):
    lines = text.splitlines()
    n = len(lines)
    i = 0
    body_size = 18
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
            body_size = int(fm.get("body_size", "18"))
        except ValueError:
            body_size = 18
        lang_in = fm.get("lang", "").strip()
        if lang_in:
            if not re.match(r"^[A-Za-z]{2,8}(-[A-Za-z0-9]{1,8})*$", lang_in):
                raise SystemExit(
                    "ERROR: invalid lang %r in structure.md front matter "
                    "(expected a BCP-47 tag like ru-RU)" % lang_in)
            lang = lang_in
    # body_size now only drives the plan slide (items at body_size - 1) and the
    # table cells (also body_size - 1: a table has no placeholder to inherit a
    # size from). Body text takes its size from the template (resolve_body_size)
    # and stamps no size of its own. `lang`, when declared, is stamped on every
    # generated run and end mark.
    TOC = body_size - 1

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
           "title": None, "plan": None, "slides": [], "closing": None}
    for header, body in sections:
        if header == "title":
            doc["title"] = " ".join(b.strip() for b in body if b.strip()).strip()
        elif header == "plan":
            doc["plan"] = [b.strip() for b in body if b.strip()]
        elif header == "closing":
            doc["closing"] = " ".join(b.strip() for b in body if b.strip()).strip()
        elif header == "slide":
            stitle = None
            parsed = []
            j = 0
            while j < len(body):
                b = body[j]
                if b.startswith("## "):
                    stitle = b[3:].strip()
                    j += 1
                    continue
                s = b.strip()
                if not s:
                    j += 1
                    continue
                if s.startswith("|"):
                    rows = []
                    while (j < len(body) and body[j].strip()
                           and body[j].strip().startswith("|")):
                        rows.append(body[j])
                        j += 1
                    parsed.append(("table", parse_table(rows)))
                    continue
                if s.startswith("- "):
                    parsed.append((1, parse_inline(s[2:])))
                elif re.match(r"^\d+\.\s", s):
                    # numbered list item: rendered as a level-1 bullet, keeping
                    # the number in the text (like the plan slide)
                    parsed.append((1, parse_inline(s)))
                else:
                    parsed.append((0, parse_inline(s)))
                j += 1
            doc["slides"].append((stitle, parsed))
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

    # 1) clone the content slide (n-1 clones -> slide5 .. slide(n+3))
    clones = []  # (filename, sldId, rId)
    for i in range(1, n):
        fn = "ppt/slides/slide%d.xml" % (4 + i)
        rid = "rId%d" % (RID_BASE + i - 1)
        sid = SLIDE_ID_BASE + i - 1
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
    sldids = ['<p:sldId id="270" r:id="rId2"/>']  # title (slide1)
    if doc["plan"] is not None:
        sldids.append('<p:sldId id="273" r:id="rId3"/>')  # plan (slide2)
    sldids.append('<p:sldId id="271" r:id="rId4"/>')  # content 1 (slide3)
    for _fn, sid, rid in clones:
        sldids.append('<p:sldId id="%d" r:id="%s"/>' % (sid, rid))
    if doc["closing"] is not None:
        sldids.append('<p:sldId id="272" r:id="rId5"/>')  # closing (slide4)
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
        plan_xml = set_block_paras(plan_xml, TITLE_MARKER,
                                  title_para("План", base=plan_title_base,
                                            lang=doc["lang"]))
        plan_body_base = first_rpr(plan_xml, BODY_MARKER)
        plan_xml = set_block_paras(plan_xml, BODY_MARKER,
                                  "".join(toc_para(t, TOC, base=plan_body_base,
                                                 lang=doc["lang"])
                                         for t in doc["plan"]))
        parts[PLAN_SLIDE] = plan_xml.encode("utf-8")

    # 6) fill the content slides (slide3 + clones, in order); body runs
    #    inherit the template body font (no sz stamped); the original
    #    runs' typeface/lang are kept via first_rpr; table blocks become
    #    <p:graphicFrame> shapes below the slide's text, cells stamped at
    #    body_size - 1 (a table has no placeholder to inherit a size from)
    content_order = [CONTENT_SLIDE] + [fn for fn, _sid, _rid in clones]
    style_id = table_style_id(
        parts.get("ppt/tableStyles.xml", b"").decode("utf-8", "ignore"))
    for (stitle, parsed), fn in zip(doc["slides"], content_order):
        x = parts[fn].decode("utf-8")
        ct_base = first_rpr(x, TITLE_MARKER)
        x = set_block_paras(x, TITLE_MARKER,
                           title_para(stitle or "", base=ct_base,
                                     lang=doc["lang"]))
        cb_base = first_rpr(x, BODY_MARKER)
        lay = slide_layout(parsed, body_pt, doc["table_pt"])
        x = set_block_paras(x, BODY_MARKER,
                           "".join(para(lv, runs, base=cb_base, lang=doc["lang"])
                                  for (lv, runs) in parsed if lv != "table"))
        if lay["tables"]:
            nid = next_shape_id(x)
            x = insert_graphic_frames(
                x, [table_frame_xml(tbl, geom, style_id, doc["table_pt"], cb_base,
                                    doc["lang"], nid + k, k + 1)
                    for k, (tbl, geom) in enumerate(lay["tables"])])
        parts[fn] = x.encode("utf-8")

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
def fit_check(doc, body_pt):
    """Heuristic per-content-slide height estimate at the template body size
    (`body_pt`): the text lines plus any table geometry slide_layout places in
    the body box; returns (slide, est_in, box_in) warnings."""
    out = []
    for stitle, parsed in doc["slides"]:
        h_in = slide_layout(parsed, body_pt, doc["table_pt"])["h_emu"] / EMU_PER_INCH
        if h_in > BODY_H_IN:
            out.append((stitle, round(h_in, 2), round(BODY_H_IN, 2)))
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
    over = fit_check(doc, body_pt)

    print("built %s" % out_path)
    print("  slides: %d (title + %s + %d content + %s)"
          % (2 + ncontent + (1 if doc["closing"] else 0),
             ("1 plan" if doc["plan"] is not None else "no plan"),
             ncontent,
             ("1 closing" if doc["closing"] is not None else "no closing")))
    print("  fonts: body = template (est. %d pt), plan items = %d pt (body_size=%d)"
          % (body_pt, doc["TOC"], doc["body_size"]))
    n_tables = sum(1 for _st, parsed in doc["slides"]
                   for it in parsed if it[0] == "table")
    print("  tables: %d (cells at %d pt)" % (n_tables, doc["table_pt"]))
    print("  lang: %s" % ((doc["lang"] + " (stamped, dirty=0)") if doc["lang"]
                           else "not set (template default)"))
    if problems:
        print("VALIDATION FAILED:")
        for p in problems:
            print("  - " + p)
        return 1
    print("  validation: PASSED")
    if over:
        print("  fit warnings (est. text + table height > %0.2fin box at %d pt):"
              % (BODY_H_IN, body_pt))
        for stitle, h, box in over:
            print("    - %r: est %0.2fin" % (stitle, h))
    else:
        print("  fit: all content slides within the box (estimate)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
