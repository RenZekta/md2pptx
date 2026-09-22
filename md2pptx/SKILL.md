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
        validation: PASSED
        fit: all content slides within the box (estimate)

  `validation: FAILED` lists one line per problem. If any slide overflows the
  body box, the last block becomes `fit warnings (...)` with one line per
  overflowing slide.
- Never edit `engine.py`; it is deterministic and complete. All inputs go
  through `structure.md` and the template. If it fails on a valid template
  and valid `structure.md`, that is a real engine bug to fix — not a content
  problem.

## Template contract
The engine assumes the template's first four slides are exactly:

- **slide1 — title**: a freeform topic line containing the anchor string
  `На тему` (Russian standard — the engine keeps the `На тему: <тема>`
  format by prepending `На тему: ` to the `# title` topic) or `Topic`
  (English — the whole line is replaced by the `# title` topic). Either way
  the new paragraph is bold and keeps the original run's font/lang.
- **slide2 — plan**: title placeholder `<p:ph type="title"/>` and body
  placeholder `<p:ph idx="1"/>`.
- **slide3 — content**: the content-slide prototype, same two placeholders.
  Every additional `# slide` clones slide3 as a new part (registered in
  `[Content_Types].xml`, `ppt/_rels/presentation.xml.rels` and `sldIdLst`);
  the closing slide is reordered to be last.
- **slide4 — closing**: title placeholder `<p:ph type="title"/>`.

If a template does not match this contract the engine raises a clear error —
do not work around it; fix/replace the template or report the mismatch.

## The `structure.md` format
A single source of truth. The engine reads it and produces the deck.

**Front matter** (optional keys in a leading block):
- `body_size: <pt>` — baseline font size for the plan-slide items (rendered at
  `body_size - 1` pt). Body text takes no size from `structure.md` — it
  inherits the template's body font.

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
  - a numbered list item (`1. <text>`) → **level 1** (bulleted; the number is
    kept in the text, like the plan slide).
- `**bold**` — a bold run within a line.

**Inside a `# plan`:**
- one numbered item per line: `1. <text>`, `2. <text>`, ...

## Content conventions
The rules to follow when producing the `structure.md`:
- **Slide types** — use only `# title`, `# plan`, `# slide`, `# closing`.
- **Levels** — plain line = level 0 (label / intro); `- ` = level 1 (bullet).
- **Plan** — a numbered list (`1.`, `2.`, ...), not bullets.
- **Bold** — mark key labels with `**bold**`.
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
- Replaced runs keep the typeface elements (a:latin/a:cs) and lang of the
  original run's rPr, so an explicitly set template font (e.g. Times New
  Roman) is preserved instead of falling back to the theme font.
- **Plan-slide items** — stamped at `body_size - 1` pt (the only place
  `body_size` acts).
- **Closing-slide title** — stamped at the size declared in the template's
  closing-slide title placeholder (36 pt in the reference template).
- The fit report estimates body height at the template's effective body size
  (read from the content-slide body placeholder → its layout; fallback 18 pt).

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
