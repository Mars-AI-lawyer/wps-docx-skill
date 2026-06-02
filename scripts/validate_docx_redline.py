#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import zipfile
from datetime import datetime
from pathlib import Path

from ooxml_utils import (
    AUTHOR as DEFAULT_AUTHOR,
    COMMENTS_CONTENT_TYPE,
    NS,
    WORD_REL_COMMENTS,
    parse_datetime,
    parse_xml,
    qn,
)


def _collect_text(nodes) -> str:
    return "".join(node.text or "" for node in nodes)


def validate_docx(
    docx_path: Path,
    log_path: Path | None = None,
    expected_author: str = DEFAULT_AUTHOR,
) -> dict:
    checks = []
    log = None

    def add(name: str, passed: bool, detail: str = "") -> None:
        checks.append({"name": name, "passed": bool(passed), "detail": detail})

    try:
        with zipfile.ZipFile(docx_path) as package:
            names = set(package.namelist())
            add("docx_can_unzip", True)
            required = {
                "word/document.xml",
                "word/settings.xml",
                "word/comments.xml",
                "word/_rels/document.xml.rels",
                "[Content_Types].xml",
            }
            for name in sorted(required):
                add(f"part_exists:{name}", name in names)
            document_root = parse_xml(package.read("word/document.xml"))
            settings_root = parse_xml(package.read("word/settings.xml"))
            comments_root = parse_xml(package.read("word/comments.xml"))
            rels_root = parse_xml(package.read("word/_rels/document.xml.rels"))
            content_types_root = parse_xml(package.read("[Content_Types].xml"))
            add("xml_can_parse", True)
    except Exception as exc:
        add("docx_can_unzip_or_parse", False, str(exc))
        return {"valid": False, "checks": checks}

    if log_path is not None:
        try:
            log = json.loads(log_path.read_text(encoding="utf-8"))
        except Exception as exc:
            add("log_can_parse", False, str(exc))

    add(
        "settings_has_track_revisions",
        settings_root.find(qn("w", "trackRevisions")) is not None,
    )
    track_revisions = settings_root.find(qn("w", "trackRevisions"))
    add(
        "settings_track_revisions_enabled",
        track_revisions is not None and track_revisions.get(qn("w", "val")) == "1",
    )
    show_revisions = settings_root.find(qn("w", "showRevisions"))
    add(
        "settings_show_revisions_enabled",
        show_revisions is not None and show_revisions.get(qn("w", "val")) == "1",
    )
    add(
        "settings_revision_view_markup_enabled",
        bool(
            settings_root.xpath(
                "./w:revisionView/w:markup[text()='1']",
                namespaces=NS,
            )
        ),
    )

    deletions = document_root.xpath(".//w:del", namespaces=NS)
    insertions = document_root.xpath(".//w:ins", namespaces=NS)
    generated_revision_ids = None
    generated_comment_ids = None
    needs_deletion_revision = True
    needs_insertion_revision = True
    if log is not None:
        generated_revision_ids = {
            str(revision_id)
            for action in log.get("actions", [])
            for revision_id in action.get("revision_ids", [])
        }
        generated_comment_ids = {
            str(action.get("comment_id"))
            for action in log.get("actions", [])
            if action.get("comment_id") is not None
        }
        generated_events = {
            stamp.get("event")
            for action in log.get("actions", [])
            if action.get("status") in {"applied", "would_apply"}
            for stamp in action.get("display_timestamps", [])
        }
        needs_deletion_revision = "delete_revision" in generated_events
        needs_insertion_revision = "insert_revision" in generated_events
    generated_deletions = (
        [node for node in deletions if node.get(qn("w", "id")) in generated_revision_ids]
        if generated_revision_ids is not None
        else deletions
    )
    generated_insertions = (
        [node for node in insertions if node.get(qn("w", "id")) in generated_revision_ids]
        if generated_revision_ids is not None
        else insertions
    )
    add(
        "document_has_deletion_revision",
        bool(deletions) if needs_deletion_revision else True,
        "not required for this log" if not needs_deletion_revision else "",
    )
    add(
        "document_has_insertion_revision",
        bool(insertions) if needs_insertion_revision else True,
        "not required for this log" if not needs_insertion_revision else "",
    )
    add(
        "generated_deletions_author_matches",
        (
            bool(generated_deletions)
            if needs_deletion_revision
            else True
        )
        and all(
            node.get(qn("w", "author")) == expected_author
            for node in generated_deletions
        ),
        f"expected_author={expected_author}",
    )
    add(
        "generated_insertions_author_matches",
        (
            bool(generated_insertions)
            if needs_insertion_revision
            else True
        )
        and all(
            node.get(qn("w", "author")) == expected_author
            for node in generated_insertions
        ),
        f"expected_author={expected_author}",
    )
    add(
        "deletions_use_del_text",
        (
            bool(document_root.xpath(".//w:del/w:r/w:delText", namespaces=NS))
            if needs_deletion_revision
            else True
        ),
        "not required for this log" if not needs_deletion_revision else "",
    )
    add(
        "insertions_use_text",
        (
            bool(document_root.xpath(".//w:ins/w:r/w:t", namespaces=NS))
            if needs_insertion_revision
            else True
        ),
        "not required for this log" if not needs_insertion_revision else "",
    )

    comments = comments_root.xpath(".//w:comment", namespaces=NS)
    generated_comments = (
        [comment for comment in comments if comment.get(qn("w", "id")) in generated_comment_ids]
        if generated_comment_ids is not None
        else comments
    )
    comment_texts = [
        _collect_text(comment.xpath(".//w:t", namespaces=NS))
        for comment in generated_comments
    ]
    add("comments_exist", bool(comments))
    add(
        "generated_comments_have_text",
        bool(comment_texts) and all(text.strip() for text in comment_texts),
    )
    add(
        "generated_comments_author_matches",
        bool(generated_comments)
        and all(
            comment.get(qn("w", "author")) == expected_author
            for comment in generated_comments
        ),
        f"expected_author={expected_author}",
    )
    add(
        "document_has_comment_references",
        bool(document_root.xpath(".//w:commentReference", namespaces=NS)),
    )
    add(
        "comment_references_have_style",
        bool(document_root.xpath(".//w:commentReference", namespaces=NS))
        and all(
            run.xpath(
                "./w:rPr/w:rStyle[@w:val='CommentReference']",
                namespaces=NS,
            )
            for run in document_root.xpath(".//w:r[w:commentReference]", namespaces=NS)
        ),
    )
    rels = rels_root.xpath(
        "./rel:Relationship[@Type=$type]",
        namespaces=NS,
        type=WORD_REL_COMMENTS,
    )
    add("comments_relationship_exists", bool(rels))
    overrides = content_types_root.xpath(
        "./ct:Override[@PartName='/word/comments.xml' and @ContentType=$content_type]",
        namespaces=NS,
        content_type=COMMENTS_CONTENT_TYPE,
    )
    add("comments_content_type_exists", bool(overrides))

    if log is not None:
        try:
            add(
                "log_timestamp_mode_synthetic_spread",
                log.get("display_timestamp_mode") == "synthetic_spread",
            )
            actual = parse_datetime(log.get("actual_generated_at"))
            display_values = []
            for action in log.get("actions", []):
                for stamp in action.get("display_timestamps", []):
                    display_values.append(parse_datetime(stamp["display_timestamp"]))
            add("log_has_actual_generated_at", log.get("actual_generated_at") is not None)
            add("log_has_display_timestamps", bool(display_values))
            in_range = all(0 <= (actual - stamp).total_seconds() <= 7200 for stamp in display_values)
            add("display_timestamps_within_two_hours", bool(display_values) and in_range)
            unique_seconds = {stamp.isoformat() for stamp in display_values}
            add(
                "display_timestamps_not_all_identical",
                len(unique_seconds) > 1 if len(display_values) > 1 else True,
            )
        except Exception as exc:
            add("log_can_parse", False, str(exc))

    valid = all(check["passed"] for check in checks)
    return {"valid": valid, "checks": checks}


def main() -> int:
    parser = argparse.ArgumentParser(description="Validate DOCX tracked changes and comments.")
    parser.add_argument("docx_path")
    parser.add_argument("--log")
    parser.add_argument("--report")
    parser.add_argument("--author", default=DEFAULT_AUTHOR)
    args = parser.parse_args()

    report = validate_docx(
        Path(args.docx_path),
        Path(args.log) if args.log else None,
        expected_author=args.author,
    )
    payload = json.dumps(report, ensure_ascii=False, indent=2)
    if args.report:
        Path(args.report).write_text(payload + "\n", encoding="utf-8")
    else:
        print(payload)
    return 0 if report["valid"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
