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
- `body_size: <pt>` — baseline font size for the plan-slide items (rendered at
  `body_size - 1` pt). Body text takes its size from the template's body
  placeholder (18 pt in the reference template); the engine stamps no body
  size of its own.
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
  as a numbered list.
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

## Folder layout

```
md2pptx/
├── README.md            # this file
├── md2pptx/               # Mode A: the agentic skill
│   ├── SKILL.md         # the skill (workflow + conventions)
│   └── engine.py        # the deterministic engine (CLI)
├── webapp/              # Mode B: the standalone web app
│   ├── structure.md     # generalized example of the input format
│   └── index.html       # the self-contained web app (JSZip + jsPDF inlined)
└── Tests/               # engine tests (Python + JS)
    ├── test_structure.md    # shared sample outline (exercises all features)
    ├── run_tests.py         # Python engine tests (anchors, numbering, fonts, lang)
    ├── run_tests_js.js      # JS engine tests + byte-parity vs Python (incl. lang)
    └── out/                 # build artifacts (regenerated on each run)
```

## Status

Both modes are implemented. `md2pptx/engine.py` is the deterministic engine (Mode A
CLI); `webapp/index.html` is the same engine running in the browser (Mode B).
The engine is verified: it reproduces the template deck deterministically and its
output parses as valid OOXML. `webapp/structure.md` is the generalized format
example.

## Trade-offs / notes
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
