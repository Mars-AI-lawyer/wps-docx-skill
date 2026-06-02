---
name: docx-wps-redline
description: Codex adapter for an agent-ready DOCX redline toolkit that creates WPS-compatible review copies with real Word tracked changes and real comments by deterministic OOXML patching. Use when Codex needs to orchestrate DOCX redlines, insertions, deletions, replacements, or margin comments while preserving the original file, avoiding simulated colored text, validating package structure, and preparing output for WPS or Microsoft Word review workflows.
---

# DOCX/WPS Redline Codex Adapter

## Overview

Use this Codex adapter to orchestrate the platform-neutral DOCX/WPS redline toolkit. The toolkit creates a new `.docx` review copy with true OOXML tracked changes and comments. The scripts write `<w:del>`, `<w:ins>`, and `word/comments.xml`, enable revision display settings, and validate the resulting package.

Do not directly edit DOCX XML by hand for redlines. Generate an `edit_plan.json`, let the writer locate exact targets, and treat unresolved actions as manual-review items.

## Author Handling

Default the review author to `Reviewer`. If the user asks to use a specific reviewer, redline author, revision author, comment author, or phrasing such as "use XX as the reviser", pass that exact name to every writer, dry-run, and validation command with `--author "XX"`. Do not only mention the requested author in notes or logs. The same `--author` value must become the OOXML `w:author` for both tracked revisions and comments so Word/WPS displays the requested name in the review UI.

## Workflow

1. Extract document structure:

```bash
python3 scripts/extract_docx_structure.py "input.docx" --output "input_structure.json"
```

2. Build `edit_plan.json` using exact text from the structure output. Include stable target fields: `container_type`, `paragraph_index` or `table_path` or `paragraph_range`, `target_text`, and `normalized_text_hash`.

3. Run a dry run when the action list is large or target confidence is uncertain:

```bash
python3 scripts/docx_redline_writer.py "input.docx" "edit_plan.json" \
  --dry-run \
  --log "redline_log.dry-run.json" \
  --author "<review-author>"
```

4. Write a new reviewed copy:

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

6. Ask the user to open the result in WPS or Word and follow `docs/WPS_MANUAL_TEST_CHECKLIST.md` before claiming WPS acceptance.

## Supported Actions

- `comment_only`: add a real Word/WPS comment to a uniquely located paragraph or table-cell paragraph.
- `replace_sentence`: replace a complete sentence or exact substring that appears once in the target paragraph.
- `replace_clause`: replace a complete paragraph, table-cell paragraph, or contiguous paragraph range.
- `insert_sentence_after`: insert one complete sentence after a unique anchor sentence or text.
- `delete_sentence`: delete a complete sentence or exact substring that appears once in the target paragraph.

The writer does not support semantic cross-paragraph inference, nested table targeting, accepting or rejecting existing revisions, micro-diff edits, or complex numbering repair.

## Edit Plan Rules

Read `schemas/edit-plan.schema.json` before generating an edit plan. Each action must include:

- `action_id`
- `action_type`
- `target`
- `comment`
- `replacement_text` for `replace_sentence`, `replace_clause`, and `insert_sentence_after`

Use exact `target_text` copied from extracted structure. If the target is missing, repeated, mismatched, or has a different `normalized_text_hash`, the writer must log the action as `unresolved` and not guess.

For `paragraph_range`, use `container_type: "paragraph_range"`, set `paragraph_index` and `table_path` to `null`, and prefer explicit `paragraph_range.start_paragraph_index` and `paragraph_range.end_paragraph_index`.

## Output Guarantees

- Never overwrite the source DOCX.
- Preserve existing comments and existing revisions elsewhere in the document.
- Allocate new revision and comment IDs after existing IDs.
- Create comments as real `word/comments.xml` comments, not inline body notes.
- Add the comments relationship and content type when missing.
- Enable tracked-revision display settings in `word/settings.xml`.
- Use the `CommentReference` run style for comment references.

## Resources

- `scripts/extract_docx_structure.py`: extract paragraphs, table cells, run text, and hashes.
- `scripts/docx_redline_writer.py`: apply confirmed edit-plan actions as true OOXML revisions/comments.
- `scripts/validate_docx_redline.py`: validate package parts, revisions, comments, authors, and timestamp audit fields.
- `schemas/edit-plan.schema.json`: edit plan schema.
- `references/redline-workflow.md`: detailed workflow and failure-mode notes.
- `docs/WPS_MANUAL_TEST_CHECKLIST.md`: manual WPS acceptance checklist.
