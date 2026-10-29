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
  Writes `out/ru_deck_py.pptx`, `out/lang_deck_py.pptx`,
  `out/numbering_deck_py.pptx` and `out/en_deck_py.pptx` (the parity inputs).
- `run_tests_js.js` — the same checks on the JS core extracted from
  `webapp/index.html` (including the preview's `numberLabels` and slide-size handling), plus byte-identical part-by-part parity against
  `out/ru_deck_py.pptx`, `out/lang_deck_py.pptx`,
  `out/numbering_deck_py.pptx` and `out/en_deck_py.pptx` (all produced by
  `run_tests.py` — run that first). An embedded chart workbook is a zip whose
  wrapper bytes depend on the zip implementation, so workbooks are compared by
  their entries and every other part byte for byte, in both directions. Requires Node and jszip 3.x: the Windows
  copy path is tried first, then a plain `require("jszip")` (install jszip in
  a `node_modules` visible from this test).
- `out/` — build artifacts, regenerated on every run (do not edit).
