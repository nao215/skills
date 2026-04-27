---
name: word
description: Generate or edit Microsoft Word (.docx) documents from a structured spec or freeform notes. Use this skill whenever the user asks to create, build, generate, draft, or edit a Word doc, .docx, document, report, memo, or letter — and also when they hand over notes, an outline, or markdown they want turned into a Word file. Output is a real .docx that opens in Word, Pages, and Google Docs.
---

# Word Generation

This skill turns a JSON spec into a polished `.docx` using `python-docx`. The generator is bundled at `scripts/generate_docx.py` and supports the block types most documents need: heading, paragraph, bulleted/numbered lists, table, image, quote, code, page break, and divider. It can also edit an existing `.docx` (find/replace plus appending blocks).

## When to use

Trigger on any of these signals:

- Words: "Word", "Word doc", "docx", "document", "report", "memo", "letter", "write-up"
- Phrases: "draft a doc", "turn this into a Word file", "write a report", "edit this docx"
- The user shares notes, an outline, or markdown and the implied next step is a Word document

If a non-Word format (Google Docs natively, plain markdown, PDF) is explicitly requested, this skill is not the right tool — flag it and ask.

## Workflow

1. **Lock the outline first.** If the user only gave a topic, propose a section outline (headings + 1–2 line summaries) and confirm before generating. If they handed over a doc or notes, summarize how you'll structure it and confirm.
2. **Translate to a spec.** Build a JSON spec following `references/spec_reference.md`. Keep sections focused: meaningful headings, paragraphs of 2–5 sentences, lists for parallel items. See `references/design_guide.md` for the content rules and `references/example_spec.json` for a complete example.
3. **Generate or edit.** Run the script:
   ```bash
   # Generate from scratch
   python scripts/generate_docx.py <path/to/spec.json> <path/to/output.docx>

   # Edit an existing file: apply replacements then append blocks
   python scripts/generate_docx.py <path/to/spec.json> <path/to/output.docx> --base <path/to/existing.docx>
   ```
   The script resolves any `image_path` values relative to the spec file's directory.
4. **Report.** Tell the user the output path, block count, and (in edit mode) how many replacements were applied. If they want tweaks, edit the spec and re-run — the spec is the source of truth, not the `.docx`.

## Spec at a glance

```json
{
  "title": "Q1 Review",
  "author": "Optional, written to file metadata",
  "blocks": [
    {"type": "heading", "level": 1, "text": "Q1 Review"},
    {"type": "paragraph", "text": "Revenue grew 18% QoQ, beating plan."},
    {"type": "heading", "level": 2, "text": "Highlights"},
    {"type": "bullets", "items": ["Revenue +18%", "NPS 62"]},
    {"type": "numbered", "items": ["Ship SSO", "Cut p95 latency"]},
    {"type": "table", "headers": ["KPI","Value"], "rows": [["DAU","1.2M"]]},
    {"type": "image", "image_path": "diagram.png", "caption": "Architecture"},
    {"type": "quote", "text": "Make it work, then make it right.", "attribution": "Kent Beck"},
    {"type": "code", "text": "print('hello')"},
    {"type": "page_break"},
    {"type": "divider"}
  ]
}
```

For inline formatting inside a paragraph, use `runs` instead of `text`:

```json
{"type": "paragraph", "runs": [
  {"text": "Revenue grew "},
  {"text": "18%", "bold": true},
  {"text": " QoQ."}
]}
```

For edit mode, add a `replacements` array — applied to the base doc before any new blocks are appended:

```json
{
  "replacements": [
    {"find": "{{quarter}}", "replace": "Q1 2026"},
    {"find": "DRAFT", "replace": "FINAL"}
  ],
  "blocks": [
    {"type": "heading", "level": 2, "text": "Addendum"},
    {"type": "paragraph", "text": "Added after review."}
  ]
}
```

Full field reference: `references/spec_reference.md`.

## Dependency

Requires `python-docx` (imported as `docx`). If the import fails, install once:

```bash
pip install python-docx
```

The generator handles tables, images, and lists natively — it does not shell out to Word or LibreOffice, so it runs in any Python environment.

## Pitfalls to avoid

- **Wall of text.** A paragraph longer than ~5 sentences usually wants to be split or turned into bullets. The `design_guide.md` lays out the rules.
- **Missing image files.** When the spec references `image_path`, verify the file exists before running the generator — otherwise the script raises `FileNotFoundError` and writes nothing.
- **Wrong extension.** Output must end in `.docx`. The generator does not produce the legacy `.doc` binary format.
- **Find/replace across runs.** Replacement collapses each matched paragraph into a single run, losing inline formatting in that paragraph. Use unique placeholders (`{{name}}`) and avoid replacing partial words inside formatted text.
- **Skipping the outline step.** Going directly from a vague topic to a generated doc almost always produces something the user wants to throw out. Spend the round-trip on the outline first.
