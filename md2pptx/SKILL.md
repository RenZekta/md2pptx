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
- Optional settings follow the three paths: `--lang`, `--body-size`,
  `--table-theme-fill on|off`, `--topic-anchors`, `--title-marker`,
  `--body-marker`, `--title-slide`, `--plan-slide`, `--content-slide`,
  `--closing-slide`. They are defaults for the front-matter keys of the same name,
  which win when both are given.
- `--detect` instead of an output file prints the roles, positions, unused
  slides, master and layout names this structure would get from this template,
  and writes nothing. Use it on any template you have not built with before.
- Name the output file after the deck topic — `<topic>.pptx` (the `# title`
  topic text, sanitized), not a generic `output.pptx`.
- Exit code `0` = success, `1` = validation failure.
- The report always goes to stdout; read it every run:

      built <topic>.pptx
        template roles:
          title:   slide1.xml (position 1, ppt/slides/slide1.xml: anchor 'На тему')
          plan:    slide2.xml (position 2, ppt/slides/slide2.xml: title, body)
          content: slide3.xml (position 3, ppt/slides/slide3.xml: title, body)
          closing: slide4.xml (position 4, ppt/slides/slide4.xml: title, body)
          unused   slide7.xml, slide8.xml (dropped from the deck)
          master   slideMaster1.xml
        slides: N (title + 1 plan + N content + 1 closing)
        fonts: body = template (est. 18 pt), plan items = 17 pt (body_size=18)
        lang: ru-RU (stamped, dirty=0)
        validation: PASSED
        fit: all content slides within the box (estimate)

  (read the `template roles:` block first on a template you have not built with:
  it names the slide each part took, its position in the deck, why — the topic
  anchor found, and whether the title/body placeholder was matched by the exact
  `title_marker` / `body_marker` tag or only by the general rule (`title*` /
  `body*`) — plus the slides it dropped and the master it read the geometry from.
  A part reading `not used` means the deck has no such section. When a `*` marks
  the part you need to fill, set the marker to the tag the template really uses.)
  (the lang line reads `lang: not set (template default)` when the key is
  absent; a `tables:` line notes how many tables were given the theme-coloured
  default style, and each `warning:` line names a front-matter problem such as
  an unknown key)

  `validation: FAILED` lists one line per problem. If any slide overflows the
  body box, the last block becomes `fit warnings (...)` with one line per
  overflowing slide.
- Never edit `engine.py`; it is deterministic and complete. All inputs go
  through `structure.md` and the template. If it fails on a valid template
  and valid `structure.md`, that is a real engine bug to fix — not a content
  problem.

## Template roles
The engine does not assume any slide sits at a fixed position or file name: it
reads which slide plays each of the four parts, and every reading can be
overruled. Detection walks the slides in the order they play (`sldIdLst` → their
relationships), not in the order they are named:

- **title slide** — the first slide carrying a topic anchor: `На тему` (Russian
  standard — the engine keeps the `На тему: <тема>` format by prepending
  `На тему: ` to the `# title` topic) or `Topic` (English — the whole line is
  replaced by the topic). Anchors, and whether each stays as a prefix, are the
  `topic_anchors` setting; the new paragraph is bold and keeps the original
  run's font and language. With no anchor anywhere, the title slide is the first
  slide with a title placeholder and no body placeholder *that comes before any
  slide having both*, failing that the first slide — and such a slide takes the
  topic into its title placeholder.
- **plan slide, then content slide** — the first two slides after the title that
  have both a title placeholder and a body placeholder: the first becomes the
  plan / table-of-contents slide, the second is the content prototype. Its title
  is the template's own plan title (a trailing `(...)` hint is dropped) unless
  `# plan` carries a `## <title>`. Every additional `# slide` clones that
  prototype into a new part (registered in `[Content_Types].xml`,
  `ppt/_rels/presentation.xml.rels` and `sldIdLst`); the closing slide is ordered
  last. A slide with `::: left` / `::: right` blocks or a `layout:` line points
  its clone at that layout instead; the template must contain it (a two-content
  layout needs a title and exactly two side-by-side body placeholders).
- **closing slide** — the last remaining slide with a title placeholder and no
  body placeholder; failing that, the first remaining slide with a title.
- Every other slide — the template's own showcase or example pages, and any role
  slide the deck has no section for — is dropped from the output together with
  its relationships, content type and note parts.

`title_marker` / `body_marker` (defaults `<p:ph type="title"/>` /
`<p:ph idx="1"/>`) are the tags that decide "has a title / body placeholder" for
detection, for the fills and for the geometry: `layout_catalog` looks for the
marker (or any `<p:ph … type="title"…>`), and the box a table, chart or picture
is laid out in is read through the body marker's `idx`, walking slide → layout →
master, where the master is the one the content slide's layout actually
references rather than an assumed `slideMaster1.xml`. So a template whose body
placeholder is `<p:ph type="body" idx="2"/>` gets its frames inside that
placeholder's box, not at a fallback position.

`title_slide`, `plan_slide`, `content_slide` and `closing_slide` (front matter,
`--title-slide` … on the command line, *Template keywords* in the web app)
overrule the detection. Each takes a 1-based deck position (`3` counts the
slides in the order they play, not their file names), a part name
(`slide7.xml`), or — for `plan_slide` and `closing_slide` only — `none`. The
content slide is the clone prototype and can never be `none`.

Ask the engine before trusting it:

    python3 md2pptx/engine.py structure.md Template.pptx --detect

prints the role assignments with their deck positions, the slides the deck will
not use, the master, and every layout name with its body-placeholder count — the
names a `layout:` line can take. The web app shows the same table under
**Template** as soon as a `.pptx` is loaded, and the build log ends with the
slides it took from the template.

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

When the roles come out wrong, do not edit the template's XML: run `--detect`,
then name the slides with the `*_slide` settings — that is exactly what they are
for. Only when a part genuinely has no slide to serve it (no slide carries the
placeholders it needs) is the template the problem: report the mismatch or fix
and replace it. The engine raises a clear error naming the setting in that case.

## The `structure.md` format
A single source of truth. The engine reads it and produces the deck.

**Front matter** (optional keys in a leading block, written between two `---`
lines). The engine also reads the keys when the fences are missing (`key: value`
lines before the first `# ` heading), and tolerates a byte-order mark or blank
lines at the top of the file; an unknown key or stray text there is reported as
a warning rather than ignored silently. Every key can also be given as a default
(command-line option, web-app Settings panel); a value in the front matter wins.
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
- `table_theme_fill: on|off` (default `on`) — a table whose style paints no
  cell (`style: plain` or `grid`, or a template whose default table style has no
  fills) is given the template's default table style instead, so its colours come
  from the theme: an accent-coloured header row and filled body rows. `off` keeps
  such tables as written. Command line: `--table-theme-fill on|off`.
- `topic_anchors: <text>[ (keep)], ...` (default `На тему (keep), Topic`) — the
  text on the title slide that marks the topic line, tried in order. `(keep)`
  keeps the anchor in front of the topic (`На тему: <тема>`); any other anchor's
  whole line is replaced by the topic. Command line: `--topic-anchors`.
- `title_marker: <tag>` (default `<p:ph type="title"/>`) and `body_marker: <tag>`
  (default `<p:ph idx="1"/>`) — the placeholder tags, as written in the
  template's slide XML, that identify the title and the body: of the title, plan,
  content and closing slides, of the two-content layouts a `::: left` /
  `::: right` slide may use, and of the box a table, chart or picture is laid out
  in. Set them when a custom template's placeholders carry other attributes (for
  example `<p:ph type="body" idx="2"/>`); the `idx` in `body_marker` is what
  selects that box, so leaving it wrong places the frames at a fallback position.
  The engine stops with a message naming the setting when a slide has no match.
  Command line: `--title-marker`, `--body-marker`.
- `title_slide`, `plan_slide`, `content_slide`, `closing_slide`:
  `<position>|<part-name>|none` (default: detected, see *Template roles*) —
  which slide of the template plays each part: a 1-based position in the deck
  (`3` counts the slides in the order they play, not their file names) or a part
  name such as `slide7.xml`. `plan_slide` and `closing_slide` also take `none`
  for a template with no such slide; `title_slide` and `content_slide` cannot,
  since a deck always needs a cover and a slide type to clone. Two roles naming
  one slide, or a position outside the deck, raises an error naming the setting.
  Command line: `--title-slide`, `--plan-slide`, `--content-slide`,
  `--closing-slide`.

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
    (no fills, no rules), or no fills with a table grid. With `table_theme_fill`
    on (the default) `plain` and `grid` give way to the theme-coloured default
    style; set `table_theme_fill: off` to keep them.
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
only to parse-check well-formedness), detects the template's roles (or takes the
ones the settings name), clones/reorders those slide types, fills the
placeholders, drops the slides the deck does not use, re-zips, and runs the checks
(structural well-formedness + per-slide fit estimate).

## Verification and delivery
- The engine's report is a structural check — necessary but not sufficient.
- The real gate is PowerPoint: open the deck. If PowerPoint offers to
  "repair" the file, do **not** deliver it — the package is damaged (in
  practice caused by invalid run properties, i.e. an engine bug); investigate
  before re-delivery.
- Optional visual QA: render to PDF (e.g. LibreOffice headless) and inspect the
  pages.
