---
name: powerpoint
description: Generate PowerPoint (.pptx) presentations from a structured spec or freeform notes. Use this skill whenever the user asks to create, build, generate, or assemble a PowerPoint, .pptx, slide deck, presentation, or "slides" — and also when they hand over an outline, doc, or notes and want them turned into slides, even if they don't say "PowerPoint" explicitly. Output is a real .pptx that opens in PowerPoint, Keynote, and Google Slides.
---

# PowerPoint Generation

This skill turns a JSON spec into a polished `.pptx` using `python-pptx`. The generator is bundled at `scripts/generate_pptx.py` and supports the slide types most decks need: title, section divider, bullets, two-column, image, table, chart, quote, and closing.

## When to use

Trigger on any of these signals:

- Words: "PowerPoint", "PPT", "pptx", "slides", "deck", "presentation", "slideshow"
- Phrases: "make slides", "turn this into a deck", "build a presentation", "summarize this as slides"
- The user shares an outline, doc, transcript, or set of notes and the implied next step is a slide deck

If a non-PowerPoint format (Google Slides natively, Keynote `.key`) is explicitly requested, this skill is not the right tool — flag it and ask.

## Workflow

1. **Lock the outline first.** If the user only gave a topic, propose a 6–10 slide outline (titles + 1–2 word notes per slide) and confirm before generating. If they handed over a doc, summarize how you'll group it into slides and confirm.
2. **Translate to a spec.** Build a JSON spec following `references/spec_reference.md`. Keep slides focused: one idea per slide, ≤6 bullets per slide, ≤18 words per bullet. See `references/design_guide.md` for the content rules and `references/example_spec.json` for a complete example.
3. **Generate.** Run the script:
   ```bash
   python scripts/generate_pptx.py <path/to/spec.json> <path/to/output.pptx>
   ```
   The script resolves any `image_path` values relative to the spec file's directory.
4. **Report.** Tell the user the output path and slide count. If the user wants tweaks, edit the spec and re-run — the spec is the source of truth, not the `.pptx`.

## Spec at a glance

```json
{
  "title": "Q1 Review",
  "author": "Optional, written to file metadata",
  "slides": [
    {"type": "title", "title": "Q1 Review", "subtitle": "April 2026"},
    {"type": "bullets", "title": "Highlights", "bullets": ["Revenue +18%", "NPS 62"]},
    {"type": "two_column", "title": "Now vs. Next", "left": ["..."], "right": ["..."]},
    {"type": "image", "title": "Architecture", "image_path": "diagram.png"},
    {"type": "table", "title": "Metrics", "headers": ["KPI","Value"], "rows": [["DAU","1.2M"]]},
    {"type": "chart", "title": "Growth", "chart_type": "column", "categories": ["Q1","Q2"], "series": [{"name":"Rev","values":[100,120]}]},
    {"type": "section", "title": "Part 2: Roadmap"},
    {"type": "quote", "text": "Make it work, then make it right.", "attribution": "Kent Beck"},
    {"type": "closing", "title": "Questions?"}
  ]
}
```

Add `"notes": "..."` to any slide for speaker notes.

Full field reference: `references/spec_reference.md`.

## Dependency

Requires `python-pptx`. If the import fails, install once:

```bash
pip install python-pptx
```

The generator handles charts (bar/column/line/pie) and tables natively — it does not shell out to PowerPoint or LibreOffice, so it runs in any Python environment.

## Pitfalls to avoid

- **Wall of text.** If a bullet is longer than ~18 words or a slide has more than 6, split it. The `design_guide.md` lays out the rules.
- **Missing image files.** When the spec references `image_path`, verify the file exists before running the generator — otherwise the script raises `FileNotFoundError` and writes nothing.
- **Wrong extension.** Output must end in `.pptx`. The generator does not produce the legacy `.ppt` binary format.
- **Overstuffed charts.** A chart with 8+ categories or 5+ series turns into noise. Aggregate or split.
- **Skipping the outline step.** Going directly from a vague topic to a generated deck almost always produces something the user wants to throw out. Spend the round-trip on the outline first.
