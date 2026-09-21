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


def parse_bold(text):
    """Split text into (text, bold) runs on **...**."""
    runs = []
    for part in re.split(r"(\*\*[^*]+\*\*)", text):
        if not part:
            continue
        if part.startswith("**") and part.endswith("**"):
            runs.append((part[2:-2], True))
        else:
            runs.append((part, False))
    if not runs:
        runs = [("", False)]
    return runs


def _rpr(bold, sz=None):
    """Run properties. `sz` is in hundredths of a point; when omitted the
    run inherits the template font size. No lang is stamped: language is
    inherited from the template too."""
    attrs = []
    if sz:
        attrs.append('sz="%d"' % sz)
    if bold:
        attrs.append('b="1"')
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


def merge_rpr(base, bold=None, sz=None):
    """Build an rPr from the original run's rPr (`base`): keeps its typeface
    child elements and lang; stamps `bold`/`sz` (sz in hundredths of a
    point), overriding any stamped in `base`."""
    if not base:
        return _rpr(bold, sz)
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
    extra = []
    if sz:
        extra.append(' sz="%d"' % sz)
    if bold:
        extra.append(' b="1"')
    a = (attrs + " ".join(extra)).strip()
    prefix = "<a:rPr" + (" " + a if a else "")
    if children:
        return prefix + ">" + children + "</a:rPr>"
    return prefix + "/>"


def para(level, runs, sz=None, base=None):
    """Body paragraph at `level` (0 = plain line, 1 = bullet). The font size
    is stamped only when `sz` is given; otherwise runs inherit the
    template's body font. `base` is the original paragraph's first-run rPr:
    its typeface elements (a:latin/a:cs/a:ea) and lang are kept on every
    new run so the replaced text keeps the template's font."""
    ppr = '<a:pPr lvl="%d"/>' % level
    rs = "".join('<a:r>%s<a:t>%s</a:t></a:r>' % (merge_rpr(base, b, sz), esc(t))
                  for (t, b) in runs)
    return '<a:p>%s%s<a:endParaRPr/></a:p>' % (ppr, rs)


def title_para(text, sz=None, base=None):
    """`sz` is in points (stamped as hundredths, like the template does)."""
    return ('<a:p><a:r>%s<a:t>%s</a:t></a:r><a:endParaRPr/></a:p>'
            % (merge_rpr(base, False, sz * 100 if sz else None), esc(text)))


def toc_para(text, sz, base=None):
    return ('<a:p><a:pPr lvl="0"/>'
            '<a:r>%s<a:t>%s</a:t></a:r>'
            '<a:endParaRPr/></a:p>' % (merge_rpr(base, False, sz * 100), esc(text)))


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


def rels_name(slide_name):
    base = slide_name.split("/")[-1]            # "slide5.xml"
    return "ppt/slides/_rels/" + base[:-4] + ".xml.rels"


# ----------------------------- structure.md parser --------------------------
def parse_md(text):
    lines = text.splitlines()
    n = len(lines)
    i = 0
    body_size = 18
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
    # body_size now only drives the plan slide: items render at body_size - 1.
    # Body text takes its size from the template (see resolve_body_size).
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

    doc = {"body_size": body_size, "TOC": TOC,
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
            for b in body:
                if b.startswith("## "):
                    stitle = b[3:].strip()
                else:
                    s = b.strip()
                    if not s:
                        continue
                    if s.startswith("- "):
                        parsed.append((1, parse_bold(s[2:])))
                    elif re.match(r"^\d+\.\s", s):
                        # numbered list item: rendered as a level-1 bullet,
                        # keeping the number in the text (like the plan slide)
                        parsed.append((1, parse_bold(s)))
                    else:
                        parsed.append((0, parse_bold(s)))
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
    topic_para = ('<a:p><a:r>%s<a:t>%s</a:t></a:r><a:endParaRPr/></a:p>'
                  % (merge_rpr(topic_base, True), esc(topic_text)))
    title_xml = replace_para(title_xml, anchor, topic_para)
    parts[TITLE_SLIDE] = title_xml.encode("utf-8")

    # 5) fill the plan slide (items stamped at body_size - 1)
    if doc["plan"] is not None:
        plan_xml = parts[PLAN_SLIDE].decode("utf-8")
        plan_title_base = first_rpr(plan_xml, TITLE_MARKER)
        plan_xml = set_block_paras(plan_xml, TITLE_MARKER,
                                  title_para("План", base=plan_title_base))
        plan_body_base = first_rpr(plan_xml, BODY_MARKER)
        plan_xml = set_block_paras(plan_xml, BODY_MARKER,
                                  "".join(toc_para(t, TOC, base=plan_body_base)
                                         for t in doc["plan"]))
        parts[PLAN_SLIDE] = plan_xml.encode("utf-8")

    # 6) fill the content slides (slide3 + clones, in order); body runs
    #    inherit the template body font (no sz stamped); the original
    #    runs' typeface/lang are kept via first_rpr
    content_order = [CONTENT_SLIDE] + [fn for fn, _sid, _rid in clones]
    for (stitle, parsed), fn in zip(doc["slides"], content_order):
        x = parts[fn].decode("utf-8")
        ct_base = first_rpr(x, TITLE_MARKER)
        x = set_block_paras(x, TITLE_MARKER,
                           title_para(stitle or "", base=ct_base))
        cb_base = first_rpr(x, BODY_MARKER)
        x = set_block_paras(x, BODY_MARKER,
                           "".join(para(lv, runs, base=cb_base)
                                  for lv, runs in parsed))
        parts[fn] = x.encode("utf-8")

    # 7) fill the closing slide (title stamped at the template's declared size)
    if doc["closing"] is not None:
        close_xml = parts[CLOSE_SLIDE].decode("utf-8")
        close_sz = placeholder_sz(close_xml, TITLE_MARKER) or TITLE_PT
        close_base = first_rpr(close_xml, TITLE_MARKER)
        close_xml = set_block_paras(close_xml, TITLE_MARKER,
                                    title_para(doc["closing"], sz=close_sz,
                                              base=close_base))
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
    (`body_pt`); returns (slide, est_in, box_in) warnings."""
    out = []
    for stitle, parsed in doc["slides"]:
        h = 0.0
        for _lv, runs in parsed:
            text = "".join(t for (t, _b) in runs)
            char_w_in = AVG_CHAR_EM * (body_pt / 72.0)
            cpl = max(1, int(BODY_W_IN / char_w_in))
            nlines = max(1, math.ceil(len(text) / cpl))
            h += nlines * (LINE_SPACING * body_pt / 72.0)
        if h > BODY_H_IN:
            out.append((stitle, round(h, 2), round(BODY_H_IN, 2)))
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
    if problems:
        print("VALIDATION FAILED:")
        for p in problems:
            print("  - " + p)
        return 1
    print("  validation: PASSED")
    if over:
        print("  fit warnings (est. body height > %0.2fin box at %d pt):"
              % (BODY_H_IN, body_pt))
        for stitle, h, box in over:
            print("    - %r: est %0.2fin" % (stitle, h))
    else:
        print("  fit: all content slides within the box (estimate)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
