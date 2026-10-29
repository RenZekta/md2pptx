# md2pptx

A deterministic **`.md` → `.pptx`** pipeline for slide decks.

Write (or have a model write) a structured `structure.md`; the engine reads it,
clones and reorders the slide types of a PowerPoint template, fills the
placeholders with the structured content, re-zips the result, and verifies it.
One engine, two front-ends.

## Why

The slow, error-prone part of building a deck from a template is the *mechanics*:
cloning slide types, reordering, filling placeholders, re-zipping, and verifying
the result. This project makes those mechanics a single **deterministic program**.
The only non-deterministic part is *content* — writing or refining the
`structure.md`. Splitting mechanics from content is what makes the workflow fast
and reliable.

## Two modes

| Mode | Where | Front-end | Agent? |
|------|-------|-----------|--------|
| **A — agentic skill** | `md2pptx/` | an agent refines the `.md`, runs the engine, reads the check report, inspects the result | yes |
| **B — standalone web app** | `webapp/` | a single self-contained `.html`; no backend, no agent, fully client-side | no |

Both modes drive the **same deterministic engine**; they differ only in who
produces and verifies the content.

### Mode A — agentic skill
An agent takes a rough outline, refines the `structure.md` (wording, references,
consistency, slide count), runs the engine, and checks the result. Higher content
quality; needs an agent runtime. See `md2pptx/SKILL.md`.

### Mode B — standalone web app
No backend, no agent. A single self-contained `.html` reads a `structure.md` and
a `template.pptx` entirely in the browser and outputs a downloadable `.pptx`
named after the deck topic (`<topic>.pptx`), with an optional in-browser PDF
preview. A **Copy prompt for a model** button copies the format + deck-quality
rules together with the current `structure.md`, so the outline can be pasted
into a chat model *with* guidance instead of as a bare file. The model's only
role is *learning the conventions* — producing a well-formed `structure.md`;
the parser enforces the format deterministically.

## The `structure.md` format (the contract)

A single source of truth. The engine reads it and produces the deck.

**Front matter** (optional keys in a leading block):
- `body_size: <pt>` — baseline font size for the plan-slide items and the
  table cells (both at `body_size - 1` pt). Without it the template's own body
  size is used (read from the content slide's body placeholder: 18 pt in the
  Russian reference template, 14 pt in the English one). Body text takes its
  size from the template's body placeholder; the engine stamps no body size of
  its own.
- `lang: <BCP-47>` — proofing language for all generated text (e.g. `ru-RU`).
  Without it, generated runs carry no lang and PowerPoint proves them in its
  default language (English), flagging non-English text as "misspelled"; with
  it, every generated run and paragraph end mark gets `lang="<tag>"`
  `dirty="0"` (the same stamp PowerPoint puts on pasted runs).

The template is **not** part of `structure.md` — it is supplied to the engine
separately (a tool-call argument in the skill, a file input in the web app), so
the `.md` stays portable across templates.

**Slide sections**, in presentation order:
- `# title` — the title slide (freeform topic line; the topic text only — the
  engine adds the `На тему: ` prefix itself for Russian-anchored templates).
  Exactly one.
- `# plan` — the plan / table-of-contents slide. Optional, at most one, written
  as a numbered list. Its title is the template's own plan-slide title (a
  trailing parenthesised hint such as `План (если есть)` is dropped); a
  `## <title>` line in the section sets another one.
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
- Inline marks (same rules inside cells):
  - `**bold**` — a bold run within a line (`b="1"`).
  - `*single asterisks*` — an italic run (`i="1"`); a lone `*` inside a word
    (a footnote or a multiplication sign) stays literal.
  - `_underscores_` — an italic run too, word-boundary only, so `a_b_c` is
    literal.
  - `++plus signs++` — an underlined run (`u="sng"`); `+` is never markup on
    its own, so `1 + 2` and `C++` pass through.
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

## Folder layout

```
md2pptx/
├── README.md            # this file
├── MFUA-Template.pptx   # reference template: Russian, 4:3, light background
├── English-Template.pptx # reference template: English, 16:9, dark background
├── md2pptx/               # Mode A: the agentic skill
│   ├── SKILL.md         # the skill (workflow + conventions)
│   └── engine.py        # the deterministic engine (CLI)
├── webapp/              # Mode B: the standalone web app
│   ├── structure.md     # generalized example of the input format
│   └── index.html       # the self-contained web app (JSZip + jsPDF inlined)
└── Tests/               # engine tests (Python + JS)
    ├── test_structure.md    # shared sample outline (exercises all features)
    ├── test_numbering.md    # shared outline for numbered-list runs
    ├── test_structure_en.md # outline built against English-Template.pptx
    ├── run_tests.py         # Python engine tests (anchors, numbering, fonts, inline marks, tables, layouts, charts, lang, both templates)
    ├── run_tests_js.js      # JS engine tests + byte-parity vs Python (same coverage)
    └── out/                 # build artifacts (regenerated on each run)
```

## Status

Both modes are implemented. `md2pptx/engine.py` is the deterministic engine (Mode A
CLI); `webapp/index.html` is the same engine running in the browser (Mode B).
The engine is verified: it reproduces the template deck deterministically and its
output parses as valid OOXML. `webapp/structure.md` is the generalized format
example.

## Trade-offs / notes
- **Inline markup and tables are engine-side, not template-side** — `**bold**`,
  `*italic*`, `++underline++` and pipe tables are parsed identically in both
  engines and stamped as `b="1"` / `i="1"` / `u="sng"` runs plus an
  `<a:tbl>`-bearing `<p:graphicFrame>`; `Tests/` asserts the two builds are
  byte-identical, so the CLI deck and the browser deck carry the same marks and
  the same table geometry.
- **Table geometry is engine-computed** — column widths come from the `widths`
  weights or are proportional to the longest cell of each column, row heights
  from the wrapped-line estimate at the table size, and the frame is placed
  below the column's text inside the column's box (the `fit_check` report counts
  text plus table plus chart height per column). Borders and fills come from the
  template's table style (`<a:tableStyleId>`: the default one, or `plain` /
  `grid`), not from the engine.
- **Boxes come from the template** — a column's box is its body placeholder's
  `<a:xfrm>`, resolved slide → layout → master, and its text size for the fit
  estimate is the size the placeholder declares (28 pt in the reference
  template's two-content layout, 18 pt in its content layout).
- **Templates are read, not assumed** — slide size, `sldId` and relationship
  ids, slide file numbers, list styles, paragraph spacing, text colour and the
  body size all come from the template, so a 16:9 template with its own ids and
  a dark background works the same as the 4:3 reference one. The web app's
  preview takes its slide size, text sizes, fonts and colours from the template
  too.
- **Charts are native parts** — each chart is a chart part plus an embedded
  `.xlsx` (inline strings, fixed timestamps) and a slide relationship. The two
  engines write identical chart XML and workbook entries; the zip wrapper of the
  workbook differs by implementation, so `Tests/` compares workbooks by their
  entries.
- **Body font comes from the template** — both modes stamp no font size on
  body text or slide titles (runs inherit the template's body font; 18 pt in
  the reference template). Only the plan-slide items are engine-stamped
  (`body_size - 1`) and the closing title takes the size declared in the
  template's closing-slide title placeholder. Replaced runs also keep the
  typeface elements (a:latin/a:cs) of the original run's rPr, so a template
  font set explicitly (e.g. Times New Roman on the title slide) is preserved;
  their proofing language is the declared `lang` when the front matter gives
  one (otherwise the original run's lang, or none).
- **Web app PDF preview** is an approximate canvas render (browser font, layout
  math), not a true PowerPoint render — open the `.pptx` in PowerPoint for the
  final look.
- **No full XSD `validate.py` gate** client-side — the web app does best-effort
  structural checks; opening the file in PowerPoint is the real gate.
- **Content accuracy** (e.g., legal references) is the content producer's
  responsibility in Mode B; Mode A's agent verifies it.
