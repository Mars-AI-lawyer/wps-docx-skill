# Agent Integration Guide

`docx-wps-redline` can be driven by any agent platform that can read files, write JSON, and run command-line tools. The agent should treat the scripts as deterministic tools and the edit plan as the contract.

## Agent Contract

Inputs:

- Source `.docx`.
- `edit_plan.json` that follows `schemas/edit-plan.schema.json`.
- Review author name.

Outputs:

- Redlined `.docx`.
- Redline `log.json`.
- `validation_report.json`.

The agent must not overwrite the source DOCX. The agent must not simulate revisions with colored text, bracketed notes, or body comments. The writer is expected to create real OOXML revisions and comments.

## Author Contract

Default the review author to `Reviewer`. If the user specifies a reviewer, redline author, revision author, comment author, reviser, or similar role, the agent must pass that exact name to:

- `docx_redline_writer.py --author`
- `validate_docx_redline.py --author`

The requested author must not be recorded only in the log, delivery note, or surrounding conversation. It must be written into the generated DOCX so Word/WPS displays the requested author for both tracked revisions and comments. In OOXML terms, all generated `<w:ins>`, `<w:del>`, and `<w:comment>` elements should carry the requested `w:author`.

## Required Workflow

1. Extract structure from the source DOCX.
2. Build an edit plan using exact target text and hashes from the extracted structure.
3. Run a dry run when target confidence is uncertain.
4. Write the reviewed DOCX to a new path.
5. Validate the output with the log and the same author value used for writing.
6. Ask for WPS or Word manual confirmation when WPS compatibility matters.

## Targeting Rules

- Use `container_type: "paragraph"` plus `paragraph_index` for body paragraphs.
- Use `container_type: "table_cell"` plus `table_path` for table-cell paragraphs.
- Use `container_type: "paragraph_range"` plus `paragraph_range` for contiguous body paragraph ranges.
- Copy `target_text` exactly from extracted structure.
- Include `normalized_text_hash` so stale or mismatched plans fail safely.

## Unresolved Actions

The writer logs actions as `unresolved` when targets are missing, repeated, mismatched, stale, or outside supported boundaries. An agent must report unresolved actions as manual-review items and must not guess a nearby replacement.

Common reasons include:

- `target_text_not_found`
- `target_text_not_unique`
- `target_text_hash_mismatch`
- `paragraph_index_text_mismatch`
- `replace_clause_requires_full_paragraph_target`
- `paragraph_range_text_mismatch`
- `paragraph_range_crosses_parent_boundary`

## Validation and WPS

`validate_docx_redline.py` checks package structure, revision/comment XML, authors, relationships, content types, comment references, and timestamp audit fields. Passing validation means the generated DOCX has the expected OOXML structure.

WPS acceptance still requires opening the generated DOCX manually and following `docs/WPS_MANUAL_TEST_CHECKLIST.md`.
