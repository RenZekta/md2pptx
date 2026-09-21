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
  structural validation. Writes `out/ru_deck_py.pptx` (the parity input).
- `run_tests_js.js` — the same checks on the JS core extracted from
  `webapp/index.html`, plus byte-identical part-by-part parity against
  `out/ru_deck_py.pptx`. Requires Node and jszip 3.x (local path hardcoded
  at the top of the script).
- `out/` — build artifacts, regenerated on every run (do not edit).
