# Tests

Tests for the md2pptx engine — the Python CLI engine (`md2pptx/engine.py`)
and the JS engine core in `webapp/index.html`. Run from the repo root:

    python Tests/run_tests.py
    node Tests/run_tests_js.js

- `test_numbering.md` — shared outline for numbered lists: a plan starting at
  3, runs broken by a bullet or a plain line, a run starting at 5, a `1)` run,
  a table inside a list.
- `test_structure_en.md` — English outline built against
  `English-Template.pptx` (16:9, its own `sldId` / relationship ids, a content
  layout with no bullets and no hanging indent, layouts whose placeholders share
  the no-index sentinel): bullets and numbers, two columns on the "Two Content"
  layout, a pie and a column chart, a merged `grid` table, an explicit
  `layout: Title, Content 2`, and a 16-row table that has to shrink.
- `test_structure.md` — shared sample outline exercising all features
  (Russian topic, plan, level-0 lines, `- ` bullets, numbered body lines,
  bold/italic/underline runs, a 3×3 pipe table with centred and right-aligned
  columns, closing). Its last content slides cover the layout and chart
  features: two columns (text + pie), a table with `widths` / `size` /
  `first-col` / `style` and merged cells, a column chart with spaces and a
  decimal comma in the numbers, two charts side by side (stacked bar with
  `legend`, line with `height` and `size`), and a 14-row table under an explicit
  `layout:` line that has to shrink to fit.
- `run_tests.py` — Python engine: the RU anchor preserves the
  `На тему: <тема>` format; the EN `Topic` anchor replaces the whole line;
  a missing anchor fails; numbered body lines and plan items are
  auto-numbered paragraphs (`<a:buAutoNum>`, no number in the text); font rules (no `sz` on body runs, plan items
  at `body_size - 1`, closing title at the template-declared size); inline
  marks (`*italic*` → `i="1"`, `++underline++` → `u="sng"`, `**bold**` →
  `b="1"`, no marker characters left in the text); tables (a `p:graphicFrame`
  carrying an `a:tbl`, one `a:tr` per row, one `a:tc` per cell, head row bold,
  cells stamped at `body_size - 1`, the separator's alignment stamped as
  `algn="ctr"`/`algn="r"`, the frame placed after the body placeholder inside
  the shape tree); structural validation. The lang section: `lang: ru-RU` in
  the front matter stamps `lang` + `dirty="0"` on every generated run and end
  mark (checked per slide region), absent key leaves generated body runs bare,
  an invalid BCP-47 tag is rejected, and the CLI report carries the lang line.
  Numbering section: run grouping, `startAt` stamped only on runs not starting
  at 1, the `.` / `)` schemes, tables not splitting a run, `buFontTx` before
  `buAutoNum` (schema order). Layout / table / chart section: the two-content
  layout relationship, dropped empty body placeholders, column boxes read from
  the layout, `widths` / `size` / `style` / flags, merge attributes, chart part,
  relationship, content type and workbook, every chart type's XML markers,
  auto-shrink bounds, the `height` / fill-the-rest rule, per-column fit reports,
  and each option / layout / column error message.
  English-template section: relationship and slide ids read from the template,
  the plan title (and its `## title` override), the template's body size as the
  default, stamped bullets and hanging indents (MFUA's own left alone),
  paragraph spacing in the estimate, sentinel-index layouts (unique idx,
  explicit box and list style), body text colour on charts and fill-less
  tables, and the duplicate-id validation.
  Settings section: the front matter is read with fences, without fences, after
  a byte-order mark and after blank lines (each stamps `lang` on every content
  run), a front matter value wins over the defaults passed in, unknown keys and
  stray text become warnings, an invalid `table_theme_fill` is rejected;
  `style: grid` / `plain` tables take the default style when `table_theme_fill`
  is on (and keep the header and banding flags) and keep their own style when it
  is off, `style_has_fills` on built-in and template-defined styles; the
  template keywords (default anchors, `(keep)` parsing, invalid marker, an
  unknown anchor, a body marker the slide does not match, a custom anchor with
  and without `(keep)`, custom markers on a template copy, keywords given as
  defaults); the command-line options and a structure file saved with a
  byte-order mark.
  Template-roles section: a template whose title, plan, content and closing
  slides are `slide3..slide6` (two decoy showcase slides prepended, its
  `[Content_Types].xml` deliberately stale) is detected and built with; the
  topic, plan and content land on the detected slides; the unused slides, their
  relationships and their stale content types are gone, and every kept slide
  declares one; a deck with no plan section drops the plan slide; `plan_slide` /
  `content_slide` overrule the detected pair by position and by part name; the
  five role-setting errors (switched off while needed, two roles on one slide,
  position out of range, a part that is not a slide, the title slide off); a
  template with no closing slide serves a deck that needs none and names
  `closing_slide` when the deck asks for one; a body placeholder moved to
  `idx="2"` still gives a table the same frame as the untouched template (and
  not the fallback box); a title placeholder carrying `idx="0"` keeps
  two-content detection working through `title_marker`; a cover with no topic
  anchor text takes the topic into its title placeholder; `--detect` reports the
  roles and layouts, exits 0 and writes no file.
  Writes `out/ru_deck_py.pptx`, `out/lang_deck_py.pptx`,
  `out/numbering_deck_py.pptx` and `out/en_deck_py.pptx` (the parity inputs)
  plus `out/shifted_template.pptx`, `out/noclose_template.pptx`,
  `out/en_body_idx2_template.pptx` and `out/ru_title_idx0_template.pptx` (the
  role fixtures the JS suite builds against) and the role builds
  `out/shift.pptx`, `out/noclose.pptx`, `out/geom_moved.pptx`, `out/t0.pptx`,
  `out/fallback.pptx`.
- `run_tests_js.js` — the same checks on the JS core extracted from
  `webapp/index.html` (including the preview's `numberLabels` and slide-size handling, and
  the same settings section as `run_tests.py`), plus byte-identical part-by-part parity against
  `out/ru_deck_py.pptx`, `out/lang_deck_py.pptx`,
  `out/numbering_deck_py.pptx` and `out/en_deck_py.pptx` (all produced by
  `run_tests.py` — run that first). An embedded chart workbook is a zip whose
  wrapper bytes depend on the zip implementation, so workbooks are compared by
  their entries and every other part byte for byte, in both directions.
  The template-roles section mirrors the Python one on the same fixtures
  (`out/shifted_template.pptx` and friends), asserts the same detections,
  overrides, errors, pruning and geometry, and compares each resulting package
  byte for byte against the matching Python build (`out/shift.pptx`,
  `out/noclose.pptx`, `out/geom_moved.pptx`, `out/t0.pptx`,
  `out/fallback.pptx`), so the two engines cannot diverge on a custom template.
  Requires Node and jszip 3.x: the Windows copy path is tried first, then a
  plain `require("jszip")`, and failing both the JSZip copy inlined in
  `webapp/index.html` is loaded, so the suite runs with no install.
- `out/` — build artifacts, regenerated on every run (do not edit).
