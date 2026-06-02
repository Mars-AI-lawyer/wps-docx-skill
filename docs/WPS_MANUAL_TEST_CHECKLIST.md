# WPS Manual Test Checklist

Use this checklist after automatic validation passes. Do not claim WPS compatibility until the output DOCX has been opened manually in WPS.

## Open and Repair

- Open the redlined DOCX in WPS.
- Confirm WPS does not show a repair or recovery prompt.
- Confirm the original DOCX remains unchanged.

## Tracked Changes

- Enable or inspect the review/track-changes view.
- Confirm deleted text appears as a real deletion revision.
- Confirm inserted text appears as a real insertion revision.
- Confirm the revision author is exactly the selected author, defaulting to `Reviewer`.
- Confirm the deleted text covers the intended original text.
- Confirm the inserted text covers the intended replacement or insertion.
- Confirm there are no dense word-by-word or character-by-character revisions unless the plan explicitly targeted that exact text.

## Comments

- Open the comment pane or comment balloons.
- Confirm comments are real WPS/Word comments, not inline body text.
- Confirm the comment author is exactly the selected author, defaulting to `Reviewer`.
- Confirm each comment contains the text from the edit plan without an added business-domain prefix.
- Confirm the comment is attached to the intended location.

## Layout and Existing Content

- Confirm paragraph layout is not visibly broken.
- Confirm tables are not displaced.
- Confirm automatic numbering, if present, has no obvious abnormality.
- Confirm headers and footers, if present, are not lost.
- Confirm existing comments and existing revisions remain visible.

## Timestamp and Audit

- Confirm visible revision/comment times are not all identical when multiple actions exist.
- Compare the output with the redline log JSON.
- Confirm the log records the real generation time.
- Confirm the log marks display timestamps as `synthetic_spread`.
- Confirm no delivery note describes synthetic display timestamps as real human editing time.
