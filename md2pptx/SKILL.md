---
name: md2pptx
description: "Use this skill whenever a .pptx deck must be built from a structure outline and a PowerPoint template — creating slide decks from a rough outline or draft structure.md, filling a template's slide types with content, and verifying the result. Trigger whenever the user asks to turn an outline or structure.md plus a template into a finished .pptx, or references building slides from a structure outline."
---

# md2pptx

Build a PowerPoint deck from a structured structure and a template, using
a deterministic engine. This skill is the content side: you refine the
`structure.md`, run the engine, and verify the result.

## When to use
You have a rough outline (or a draft `structure.md`) and a PowerPoint template,
and you need to produce the finished `.pptx`.

## Workflow
1. Take a rough outline or a draft `structure.md`.
2. Refine the content: wording, references/citations, consistency, slide count,
   following the conventions below.
3. Run the deterministic engine with the `structure.md` and the template
   → `.pptx` (exact invocation below; the template is passed to the engine,
   it is not part of the `structure.md`).
4. Read the check report (validation pass/fail, per-slide fit, overflow flags).
5. Inspect the rendered result and fix content if needed; re-run.
6. Deliver the final `.pptx`.

## Running the engine
Exact invocation, arguments in this order:

    python <skill-dir>/engine.py <structure.md> <template.pptx> <output>.pptx

- `<skill-dir>` — this skill's folder (where `SKILL.md` and `engine.py` live).
- Name the output file after the deck topic — `<topic>.pptx` (the `# title`
  topic text, sanitized), not a generic `output.pptx`.
- Exit code `0` = success, `1` = validation failure.
- The report always goes to stdout; read it every run:

      built <topic>.pptx
        slides: N (title + 1 plan + N content + 1 closing)
        fonts: body = template (est. 18 pt), plan items = 17 pt (body_size=18)
        lang: ru-RU (stamped, dirty=0)
        validation: PASSED
        fit: all content slides within the box (estimate)

  (the lang line reads `lang: not set (template default)` when the key is
  absent)

  `validation: FAILED` lists one line per problem. If any slide overflows the
  body box, the last block becomes `fit warnings (...)` with one line per
  overflowing slide.
- Never edit `engine.py`; it is deterministic and complete. All inputs go
  through `structure.md` and the template. If it fails on a valid template
  and valid `structure.md`, that is a real engine bug to fix — not a content
  problem.

## Template contract
The engine assumes the template's first four slides are exactly:

- **slide1 — title**: a shape (a text box or a placeholder) with a line
  containing the anchor string `На тему` (Russian standard — the engine keeps
  the `На тему: <тема>` format by prepending `На тему: ` to the `# title`
  topic) or `Topic` (English — the whole line is replaced by the `# title`
  topic). Either way the new paragraph is bold and keeps the original run's
  font/lang.
- **slide2 — plan**: title placeholder `<p:ph type="title"/>` and body
  placeholder `<p:ph idx="1"/>`. The title's text is the default plan title
  (a trailing `(...)` hint is dropped); a `## <title>` line in `# plan`
  overrides it.
- **slide3 — content**: the content-slide prototype, same two placeholders.
  Every additional `# slide` clones slide3 as a new part (registered in
  `[Content_Types].xml`, `ppt/_rels/presentation.xml.rels` and `sldIdLst`);
  the closing slide is reordered to be last. A slide with `::: left` /
  `::: right` blocks or a `layout:` line points its clone at that layout
  instead; the template must contain it (a two-content layout needs a title and
  exactly two side-by-side body placeholders).
- **slide4 — closing**: title placeholder `<p:ph type="title"/>`.

Everything else is read from the template, so templates of any slide size and
design work: the `sldId` / relationship ids and slide file numbers of the
template (the clones continue after the highest), the geometry of every body
placeholder (slide → layout → master), the effective list style of each
placeholder (a template whose lists have no bullet, or no hanging indent, gets a
`•` and a 0.25 in hanging indent for `- ` and numbered items; a template whose
level 0 shows bullets gets none on plain lines), paragraph spacing (counted in
the height estimate), the body text colour and typeface (stamped on chart text
and on cells of the `plain` / `grid` table styles, which have no fills of their
own) and the body size. A layout placeholder with the no-index sentinel
`idx="4294967295"` (several may share it) cannot be linked to by index: the
slide's placeholder gets its own unique sentinel index, an explicit position and
a copy of the layout's list style.

If a template does not match this contract the engine raises a clear error —
do not work around it; fix/replace the template or report the mismatch.

## The `structure.md` format
A single source of truth. The engine reads it and produces the deck.

**Front matter** (optional keys in a leading block):
- `body_size: <pt>` — baseline font size for the plan-slide items and the
  table cells (both at `body_size - 1` pt). Without it the template's own body
  size is used (read from the content slide's body placeholder), so leave it out
  unless the plan and tables should differ. Body text takes no size from
  `structure.md` — it inherits the template's body font.
- `lang: <BCP-47>` — proofing language for all generated text (e.g. `ru-RU`).
  Without it, generated runs carry no lang and PowerPoint proves them in its
  default language (English), flagging non-English text as "misspelled"; with
  it, every generated run and paragraph end mark gets `lang="<tag>"`
  `dirty="0"` (the same stamp PowerPoint puts on pasted runs).

**Slide sections**, in presentation order:
- `# title` — the title slide (freeform topic line; the topic text only —
  the engine adds the `На тему: ` prefix itself for Russian-anchored
  templates). Exactly one.
- `# plan` — the plan / table-of-contents slide. Optional, at most one.
  Written as a **numbered list**: `1.`, `2.`, `3.` items.
- `# slide` — a content slide (cloned from the template's content-slide type).
  Repeated once per content slide, in order.
- `# closing` — the closing slide. Optional, at most one.

**Inside a `# slide`:**
- `## <text>` — the slide's title placeholder text.
- Body lines:
  - a plain line (column 0, no bullet) → **level 0** (label / intro, no bullet).
  - a `- ` list item → **level 1** (bulleted).
  - a numbered list item (`1. <text>` or `1) <text>`) → **level 1**, a
    PowerPoint auto-numbered paragraph. The number is not stored in the text:
    PowerPoint draws and renumbers it. Consecutive items form one list; a `- `
    line, a plain line or a change of `.`/`)` ends it. The first number is where
    the list starts (`3.` starts at 3); later numbers are ignored, as in
    CommonMark. A table does not end a list.
- Inline marks (also inside table cells):
  - `**bold**` — a bold run (`b="1"`).
  - `*single asterisks*` and `_underscores_` — an italic run (`i="1"`); each
    marker must sit at a word boundary, so a lone `*` inside a word (a footnote
    or a multiplication sign) and `a_b_c` stay literal.
  - `++plus signs++` — an underlined run (`u="sng"`); a single `+` is never
    markup, so `1 + 2` and `C++` pass through.
- Pipe tables: consecutive lines each starting with `|` are one table — the
  first line is the head row (bold), the second is the separator
  (`| --- | :--: | ---: |`, which also sets column alignment), the rest are body
  rows. Outer pipes are optional on every row; short rows are padded. A table
  is a `<p:graphicFrame>` shape below the column's text (a table cannot live
  inside a text placeholder), so its cells take an explicit size: `body_size - 1`
  pt. A table that would overflow its box shrinks one point at a time, down to
  10 pt.
- Table options: a line `{key: value; key: value}` directly above the table.
  - `widths: 3,1,1` — relative column widths, one whole number per column
    (default: proportional to the longest cell of each column).
  - `size: 14` — the cell font size in pt (8–40); a fixed size does not shrink.
  - `header: off` — the first row is not a header (no bold, no header
    formatting from the table style); `header: on` makes the first row a header
    even without a separator row.
  - `first-col: on` — the table style's first-column formatting.
  - `banded: off` — no alternating row shading.
  - `style: default|plain|grid` — the template's default table style, no style
    (no fills, no rules), or no fills with a table grid.
- Merged cells: a cell that is exactly `<<` joins the cell to its left, `^^`
  the cell above; a `<<` under a `^^` joins the same block. The range shows the
  content of its top-left cell. A marker with nothing to join (first column or
  first row) is plain text.
- Charts: an options line with `chart:` above a pipe table makes the table a
  native, editable PowerPoint chart (`ppt/charts/chartN.xml` with an embedded
  workbook, so *Edit Data* works). The first column holds the categories, the
  head row the series names and every other cell a number (spaces and a trailing
  `%` are ignored; a comma is a decimal comma: `1,5`; a pie uses the first
  series only). Colours are the template's theme accents; text is the body
  typeface.
  - `chart: column|bar|line|pie` — vertical columns, horizontal bars, lines
    with markers, or a pie.
  - `title: <text>` — a chart title (no `;` in it).
  - `legend: top|bottom|left|right|off` — default: right for a pie, bottom for
    several series, none for one.
  - `labels: on` — data labels (percentages on a pie).
  - `stacked: on` — stacked columns or bars.
  - `size: 12` — chart text in pt (default: the table size minus 4).
  - `height: 3.5` — chart height in inches (1–7); without it the chart fills
    the space left in its column, at least 2 in.
- Columns and layouts: `::: left` … `::: right` … `:::` blocks (one of each, in
  that order; only the `## ` title and a `layout:` line may sit outside them)
  give the slide two columns on the template's two-content layout (a title plus
  two side-by-side body placeholders). Each column has its own text, tables and
  charts and its own box, read from the layout's placeholders; a column without
  text has no text placeholder. `layout: <name>` selects any template layout by
  its name (case-insensitive) instead of the content layout; the slide's columns
  must match the layout's body placeholders (`layout: two-content` is the
  two-content layout of any template).

**Inside a `# plan`:**
- one numbered item per line: `1. <text>`, `2. <text>`, ... (auto-numbered
  paragraphs, like body numbered items).

## Content conventions
The rules to follow when producing the `structure.md`:
- **Slide types** — use only `# title`, `# plan`, `# slide`, `# closing`.
- **Levels** — plain line = level 0 (label / intro); `- ` = level 1 (bullet);
  `1. ` = level 1 (numbered).
- **Bullets** — do not overuse bullet points. Keep `- ` for genuinely parallel
  items; prefer a short plain line for a single point, a numbered list for a
  genuine sequence or ranking, and a table for structured data. A slide that is
  all bullets is a sign the content needs another shape.
- **Plan** — a numbered list (`1.`, `2.`, ...), not bullets.
- **Bold** — mark key labels with `**bold**`.
- **Italic** — `*italic*` for as-of values, dates, foreign or quoted terms;
  not for whole paragraphs.
- **Underline** — `++underlined++` only where the source itself underlines
  (a term defined on first use); never for decoration.
- **Tables** — a pipe table for structured data (figures, statutory points,
  comparisons); put it after the lines it belongs to. Up to about 5 rows fit
  under a few lines of text at the full table size; a longer table shrinks (down
  to 10 pt), so keep it to roughly 12 rows or split it. Use `widths` when one
  column carries most of the text, `<<` / `^^` for grouped headings, and
  `style: grid` for a ruled table.
- **Charts** — a chart when the point is a trend, a ranking or a share of a
  whole (`line`, `column` / `bar`, `pie`); a table when the exact values are
  the point. One message per chart: a `title`, at most about 6 categories on a
  pie, and `labels: on` when the numbers matter more than the axis.
- **Columns** — `::: left` / `::: right` to set text beside a table or chart (or
  two charts side by side); the layout's own text size applies, so keep the
  text in a column short.
- **Citations** — cite a statute/article by number *and* by point where it
  enumerates points (e.g. `ч. 1 ст. 56 УК РФ`, `п. 2 ст. 401 ГК РФ`, U.S. Const. art. I, para. 8); do not cite a bare
  article when a specific part is meant.
- **No hyperlinks** — drop any URLs from the source content; fold references into
  the body text instead. Use hyperlinks only if explicitly tasked to.
- **Point numbers** — if the reference is a bare article that enumerates
  points/parts, add the point numbers (`no. N` / `ч. N`).

## Font sizes
- **Body text (level 0 and level 1) and content-slide titles** — the engine
  stamps no size on these runs; they inherit the template's fonts
  (placeholder → slide layout → master → theme default; 18 pt in the
  reference template). To change the body font, edit the template, not
  `structure.md`.
- Replaced runs keep the typeface elements (a:latin/a:cs) of the original
  run's rPr, so an explicitly set template font (e.g. Times New Roman) is
  preserved instead of falling back to the theme font. Their lang is the
  declared `lang` when the front matter gives one, otherwise the original
  run's lang (or none, when the template run is bare).
- **Plan-slide items** — stamped at `body_size - 1` pt (the template's body
  size less one when `body_size` is not given).
- **Table cells** — stamped at `body_size - 1` pt too (the report shows
  `tables: N (cells at X pt)`); a cell has no placeholder to inherit a size
  from, so the size is always explicit. A table that would overflow shrinks to
  fit (not below 10 pt) unless it sets `size`.
- **Chart text** — `size` or the table size minus 4 pt, in the body typeface of
  the master.
- **Column text** — inherits the size of the layout's body placeholder (28 pt in
  the reference template's two-content layout).
- **Closing-slide title** — stamped at the size declared in the template's
  closing-slide title placeholder (36 pt in the reference template).
- The fit report estimates each column's height (text + tables + charts) at
  the size of its body placeholder (read from the placeholder → its layout;
  fallback 18 pt) against that placeholder's box.

## The engine (internals)
The deterministic engine (`engine.py` in this skill folder) edits the raw OOXML
by targeted string replacement (no ElementTree round-trip; ElementTree is used
only to parse-check well-formedness), clones/reorders the template slide types
per the contract above, fills the placeholders, re-zips, and runs the checks
(structural well-formedness + per-slide fit estimate).

## Verification and delivery
- The engine's report is a structural check — necessary but not sufficient.
- The real gate is PowerPoint: open the deck. If PowerPoint offers to
  "repair" the file, do **not** deliver it — the package is damaged (in
  practice caused by invalid run properties, i.e. an engine bug); investigate
  before re-delivery.
- Optional visual QA: render to PDF (e.g. LibreOffice headless) and inspect the
  pages.
