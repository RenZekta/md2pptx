# Tests

Tests for the md2pptx engine — the Python CLI engine (`md2pptx/engine.py`)
and the JS engine core in `webapp/index.html`. Run from the repo root:

    python Tests/run_tests.py
    node Tests/run_tests_js.js

- `test_structure.md` — shared sample outline exercising all features
  (Russian topic, plan, level-0 lines, `- ` bullets, numbered body lines,
  bold runs, closing).
- `run_tests.py` — Python engine: the RU anchor preserves the
  `На тему: <тема>` format; the EN `Topic` anchor replaces the whole line;
  a missing anchor fails; numbered body lines become level-1 bullets with
  the number kept in the text; font rules (no `sz` on body runs, plan items
  at `body_size - 1`, closing title at the template-declared size);
  structural validation. The lang section: `lang: ru-RU` in the front matter
  stamps `lang` + `dirty="0"` on every generated run and end mark (checked
  per slide region), absent key leaves generated body runs bare, an invalid
  BCP-47 tag is rejected, and the CLI report carries the lang line. Writes
  `out/ru_deck_py.pptx` and `out/lang_deck_py.pptx` (the parity inputs).
- `run_tests_js.js` — the same checks on the JS core extracted from
  `webapp/index.html`, plus byte-identical part-by-part parity against
  `out/ru_deck_py.pptx` and `out/lang_deck_py.pptx` (both produced by
  `run_tests.py` — run that first). Requires Node and jszip 3.x (local path
  hardcoded at the top of the script).
- `out/` — build artifacts, regenerated on every run (do not edit).
