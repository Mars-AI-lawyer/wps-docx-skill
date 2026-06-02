# Generic Agent Prompt

Use `docx-wps-redline` as a deterministic DOCX redline toolkit.

Rules:

- Never overwrite the source DOCX.
- If the user names a reviewer, reviser, revision author, or comment author, pass that exact name to every writer and validation command with `--author`. This must update Word/WPS-visible OOXML authors, not only logs or notes.
- Always extract structure before writing an edit plan.
- Use exact `target_text` and `normalized_text_hash` from the extracted structure.
- Run a dry run when the target may be ambiguous.
- Treat `unresolved` actions as manual-review items. Do not guess.
- Create real tracked changes and real comments through the writer script. Do not simulate redlines with colored text, bracketed notes, or inline body comments.
- Validate every generated DOCX with `validate_docx_redline.py`.
- Do not claim WPS compatibility until the document has been manually opened in WPS and checked with `docs/WPS_MANUAL_TEST_CHECKLIST.md`.

Standard command sequence:

```bash
python3 scripts/extract_docx_structure.py "input.docx" --output "input_structure.json"
python3 scripts/docx_redline_writer.py "input.docx" "edit_plan.json" --dry-run --log "redline_log.dry-run.json" --author "<review-author>"
python3 scripts/docx_redline_writer.py "input.docx" "edit_plan.json" --output "input_redlined.docx" --log "input_redline_log.json" --timestamp-mode synthetic_spread --spread-minutes 120 --author "<review-author>"
python3 scripts/validate_docx_redline.py "input_redlined.docx" --log "input_redline_log.json" --report "input_validation_report.json" --author "<review-author>"
```
