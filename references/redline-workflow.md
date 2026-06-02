# DOCX Redline Workflow Reference

## Current Support

The writer supports deterministic OOXML edits for ordinary paragraphs, table-cell paragraphs, and contiguous paragraph ranges:

- `comment_only`
- `replace_sentence`
- `replace_clause`
- `insert_sentence_after`
- `delete_sentence`

It preserves existing comments and tracked changes, creates missing comment package parts, and enables revision display settings. It intentionally avoids semantic guessing and partial word-level diffing.

## Planning Actions

Always extract the structure first. Use exact `target_text` and `normalized_text_hash` from the extracted JSON. Prefer explicit indexes:

- ordinary paragraph: `container_type: "paragraph"` and `paragraph_index`
- table cell: `container_type: "table_cell"` and `table_path`
- paragraph range: `container_type: "paragraph_range"` and `paragraph_range`

For `replace_clause`, `target_text` must cover the full paragraph, full table-cell paragraph, or full contiguous paragraph range. Use `replace_sentence` for smaller exact text inside one paragraph.

## Failure Handling

The writer logs unresolved actions and continues later actions. Treat these reasons as manual-review items:

- `target_text_not_found`
- `target_text_not_unique`
- `target_text_hash_mismatch`
- `paragraph_index_text_mismatch`
- `replace_clause_requires_full_paragraph_target`
- `paragraph_range_text_not_found`
- `paragraph_range_text_not_unique`
- `paragraph_range_crosses_parent_boundary`
- `paragraph_range_text_mismatch`

Do not rewrite an unresolved action by guessing a nearby paragraph. Re-extract structure or ask the user for a more precise target.

## Validation

Run `validate_docx_redline.py` after every write. Automatic validation checks package structure, revision/comment XML, comment references, authors, and timestamp audit data. WPS acceptance still requires manual opening in WPS.

Display timestamps may be synthetically spread across the two hours before generation to avoid identical visible times. This is an audit-visible display strategy, not a claim of human editing time.
