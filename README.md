# DOCX/WPS Redline Toolkit

Agent-ready toolkit for creating **real Word/WPS tracked changes and comments** in `.docx` files by deterministic OOXML patching.

`docx-wps-redline` is designed for agents, automation workflows, and command-line use. It extracts document structure, accepts a structured edit plan, writes true `<w:del>`, `<w:ins>`, and `word/comments.xml` records, and validates the resulting package. It does not simulate redlines with colored text, inline brackets, or body notes.

The repository includes a Codex-compatible `SKILL.md`, but the core scripts and schema are platform neutral. Any agent platform can orchestrate the same workflow.

Author note: if a user requests a specific review/revision/comment author, pass that exact name through `--author` when writing and validating. This controls the OOXML author fields shown by Word/WPS, not just the JSON log.

## Capabilities

- Create a new reviewed DOCX without overwriting the source file.
- Add real deletion revisions (`<w:del>`) and insertion revisions (`<w:ins>`).
- Add real Word/WPS comments in `word/comments.xml`.
- Enable revision display settings for Word/WPS review views.
- Preserve existing comments and tracked changes while avoiding ID collisions.
- Validate output package structure, comments, revisions, authors, relationships, content types, and timestamp metadata.
- Mark ambiguous, missing, repeated, or hash-mismatched targets as `unresolved` instead of guessing.

## Supported Actions

- `comment_only`: add a real comment to a uniquely located paragraph or table-cell paragraph.
- `replace_sentence`: replace a complete sentence or exact substring that appears once in the target paragraph.
- `replace_clause`: replace a complete paragraph, table-cell paragraph, or contiguous paragraph range.
- `insert_sentence_after`: insert text after a unique anchor sentence or exact target text.
- `delete_sentence`: delete a complete sentence or exact substring that appears once in the target paragraph.

## Install

```bash
python3 -m venv .venv
source .venv/bin/activate
python3 -m pip install -r requirements-dev.txt
```

Runtime use only needs `requirements.txt`. `requirements-dev.txt` adds test dependencies.

## Agent Workflow

1. Extract DOCX structure:

```bash
python3 scripts/extract_docx_structure.py "input.docx" --output "input_structure.json"
```

2. Build `edit_plan.json` from the extracted structure. Use exact `target_text` and `normalized_text_hash`.

3. Dry run when target confidence is uncertain:

```bash
python3 scripts/docx_redline_writer.py "input.docx" "edit_plan.json" \
  --dry-run \
  --log "redline_log.dry-run.json" \
  --author "<review-author>"
```

4. Write the reviewed DOCX:

```bash
python3 scripts/docx_redline_writer.py "input.docx" "edit_plan.json" \
  --output "input_redlined.docx" \
  --log "input_redline_log.json" \
  --timestamp-mode synthetic_spread \
  --spread-minutes 120 \
  --author "<review-author>"
```

5. Validate the output:

```bash
python3 scripts/validate_docx_redline.py "input_redlined.docx" \
  --log "input_redline_log.json" \
  --report "input_validation_report.json" \
  --author "<review-author>"
```

6. Open the result in WPS or Word and follow `docs/WPS_MANUAL_TEST_CHECKLIST.md`.

For a platform-neutral agent contract, see `docs/AGENT_INTEGRATION.md`.

## Edit Plan Example

```json
{
  "document_id": "sample-document",
  "generated_at": "2026-06-02T12:00:00+08:00",
  "actions": [
    {
      "action_id": "A001",
      "action_type": "replace_sentence",
      "target": {
        "container_type": "paragraph",
        "paragraph_index": 1,
        "table_path": null,
        "target_text": "The draft must be delivered by Friday.",
        "context_before": "",
        "context_after": "",
        "normalized_text_hash": "..."
      },
      "replacement_text": "The draft must be delivered by Monday.",
      "comment": "Deadline changed after review."
    }
  ]
}
```

Full field rules are in `schemas/edit-plan.schema.json`.

## Repository Contents

- `SKILL.md`: Codex adapter instructions for this toolkit.
- `agents/`: example agent prompts/configuration.
- `scripts/extract_docx_structure.py`: extracts paragraphs, table cells, run text, and hashes.
- `scripts/docx_redline_writer.py`: applies edit-plan actions as true OOXML revisions/comments.
- `scripts/validate_docx_redline.py`: validates generated DOCX redline packages.
- `schemas/edit-plan.schema.json`: edit plan schema.
- `references/redline-workflow.md`: workflow and failure-mode notes.
- `docs/WPS_MANUAL_TEST_CHECKLIST.md`: manual WPS acceptance checklist.
- `tests/fixtures/public_sample.docx`: public synthetic test sample.

## Tests

```bash
python3 -B -m unittest discover -s tests -v
```

Tests cover ordinary paragraphs, table cells, contiguous paragraph ranges, real revisions/comments, author override, preservation of existing comments/revisions, missing targets, repeated targets, hash mismatch, cross-parent paragraph ranges, and overwrite prevention.

## Limitations

The first version focuses on stable deterministic OOXML writing. It does not support nested table targeting, semantic cross-paragraph inference, accepting or rejecting existing revisions, complex numbering repair, or word-level intelligent diffing.

Automatic validation means the generated package has the expected OOXML structure. It is not a blanket guarantee that every WPS version and every complex document layout is accepted. Always perform the manual WPS check before claiming WPS compatibility for a delivered document.

## License

MIT. See `LICENSE`.
