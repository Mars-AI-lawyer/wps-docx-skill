#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import os
import zipfile
from dataclasses import dataclass
from difflib import SequenceMatcher
from pathlib import Path
from typing import Any

from lxml import etree

from ooxml_utils import (
    AUTHOR as DEFAULT_AUTHOR,
    NS,
    append_plain_run,
    direct_run_text,
    ensure_comments_content_type,
    ensure_comments_relationship,
    ensure_track_revisions,
    find_target_paragraph,
    find_target_run_span,
    first_run_properties,
    format_ooxml_utc,
    generate_display_timestamps,
    iter_document_paragraphs,
    iso_beijing,
    make_comment,
    make_comment_reference_run,
    make_revision_run,
    new_comments_root,
    next_comment_id,
    next_revision_id,
    normalize_text,
    normalized_hash,
    parse_datetime,
    parse_xml,
    paragraph_runs,
    paragraph_text,
    qn,
    remove_paragraph_content,
    require_docx_paths,
    serialize_xml,
)


SUPPORTED_ACTIONS = {
    "replace_sentence",
    "replace_words",
    "replace_clause",
    "comment_only",
    "delete_sentence",
    "insert_sentence_after",
}


@dataclass
class PackagePart:
    info: zipfile.ZipInfo
    data: bytes


def load_edit_plan(path: Path) -> dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8"))


def validate_edit_plan(plan: dict[str, Any]) -> list[str]:
    errors = []
    for key in ("document_id", "generated_at", "actions"):
        if key not in plan:
            errors.append(f"missing top-level field: {key}")
    actions = plan.get("actions")
    if not isinstance(actions, list) or not actions:
        errors.append("actions must be a non-empty list")
        return errors
    allowed = {
        "comment_only",
        "replace_sentence",
        "replace_words",
        "replace_clause",
        "insert_sentence_after",
        "delete_sentence",
    }
    for index, action in enumerate(actions):
        prefix = f"actions[{index}]"
        for key in (
            "action_id",
            "action_type",
            "target",
            "comment",
        ):
            if key not in action:
                errors.append(f"{prefix} missing field: {key}")
        action_type = action.get("action_type")
        if action_type not in allowed:
            errors.append(f"{prefix} unsupported action_type in schema: {action_type}")
        if action_type in {"replace_sentence", "replace_words", "replace_clause", "insert_sentence_after"}:
            if not action.get("replacement_text"):
                errors.append(f"{prefix} missing replacement_text")
        if not action.get("comment"):
            errors.append(f"{prefix} comment is required")
        target = action.get("target")
        if not isinstance(target, dict):
            errors.append(f"{prefix}.target must be an object")
            continue
        for key in (
            "container_type",
            "paragraph_index",
            "table_path",
            "target_text",
            "context_before",
            "context_after",
            "normalized_text_hash",
        ):
            if key not in target:
                errors.append(f"{prefix}.target missing field: {key}")
        if not target.get("target_text"):
            errors.append(f"{prefix}.target.target_text is required")
        if target.get("container_type") == "paragraph_range":
            if target.get("paragraph_index") is not None:
                errors.append(f"{prefix}.target.paragraph_index must be null")
            if target.get("table_path") is not None:
                errors.append(f"{prefix}.target.table_path must be null")
            paragraph_range = target.get("paragraph_range")
            if paragraph_range is not None:
                if not isinstance(paragraph_range, dict):
                    errors.append(f"{prefix}.target.paragraph_range must be an object or null")
                    continue
                for key in ("start_paragraph_index", "end_paragraph_index"):
                    if key not in paragraph_range:
                        errors.append(
                            f"{prefix}.target.paragraph_range missing field: {key}"
                        )
    return errors


def read_package(path: Path) -> tuple[list[PackagePart], dict[str, bytes]]:
    package_parts = []
    part_map = {}
    with zipfile.ZipFile(path) as package:
        for info in package.infolist():
            data = package.read(info.filename)
            package_parts.append(PackagePart(info=info, data=data))
            part_map[info.filename] = data
    return package_parts, part_map


def write_package(
    output_path: Path,
    package_parts: list[PackagePart],
    part_map: dict[str, bytes],
    changed_parts: dict[str, bytes],
) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    written = set()
    with zipfile.ZipFile(output_path, "w") as package:
        for part in package_parts:
            data = changed_parts.get(part.info.filename, part.data)
            package.writestr(part.info, data)
            written.add(part.info.filename)
        for name, data in changed_parts.items():
            if name not in written:
                package.writestr(name, data)


def get_or_create_xml(part_map: dict[str, bytes], name: str, default_root: etree._Element):
    if name in part_map:
        return parse_xml(part_map[name]), False
    return default_root, True


def default_rels_root() -> etree._Element:
    return etree.Element(f"{{{NS['rel']}}}Relationships", nsmap={None: NS["rel"]})


def add_comment_to_paragraph(
    paragraph: etree._Element,
    *,
    comment_id: int,
) -> None:
    start = etree.Element(qn("w", "commentRangeStart"))
    start.set(qn("w", "id"), str(comment_id))
    end = etree.Element(qn("w", "commentRangeEnd"))
    end.set(qn("w", "id"), str(comment_id))
    insert_at = 1 if len(paragraph) and paragraph[0].tag == qn("w", "pPr") else 0
    paragraph.insert(insert_at, start)
    paragraph.append(end)
    paragraph.append(make_comment_reference_run(comment_id))


def split_target_text(paragraph: etree._Element, target_text: str) -> tuple[str, str]:
    paragraph_full_text = paragraph_text(paragraph)
    if paragraph_full_text.count(target_text) != 1:
        raise ValueError("target_text must occur exactly once in the target paragraph")
    return paragraph_full_text.split(target_text, 1)


def append_comment_start(paragraph: etree._Element, comment_id: int) -> None:
    comment_start = etree.Element(qn("w", "commentRangeStart"))
    comment_start.set(qn("w", "id"), str(comment_id))
    paragraph.append(comment_start)


def append_comment_end_and_reference(paragraph: etree._Element, comment_id: int) -> None:
    comment_end = etree.Element(qn("w", "commentRangeEnd"))
    comment_end.set(qn("w", "id"), str(comment_id))
    paragraph.append(comment_end)
    paragraph.append(make_comment_reference_run(comment_id))


def make_revision_element(
    tag: str,
    *,
    revision_id: int,
    author: str,
    date,
    text: str,
    deleted: bool,
    run_properties: etree._Element | None,
) -> etree._Element:
    revision = etree.Element(qn("w", tag))
    revision.set(qn("w", "id"), str(revision_id))
    revision.set(qn("w", "author"), author)
    revision.set(qn("w", "date"), format_ooxml_utc(date))
    revision.append(
        make_revision_run(text, deleted=deleted, run_properties=run_properties)
    )
    return revision


def replace_paragraph_sentence(
    paragraph: etree._Element,
    *,
    original_text: str,
    replacement_text: str,
    author: str,
    delete_revision_id: int,
    insert_revision_id: int,
    comment_id: int,
    delete_date,
    insert_date,
) -> None:
    run_properties = first_run_properties(paragraph)
    prefix, suffix = split_target_text(paragraph, original_text)
    remove_paragraph_content(paragraph)
    append_plain_run(paragraph, prefix, run_properties=run_properties)

    append_comment_start(paragraph, comment_id)
    paragraph.append(
        make_revision_element(
            "del",
            revision_id=delete_revision_id,
            author=author,
            date=delete_date,
            text=original_text,
            deleted=True,
            run_properties=run_properties,
        )
    )
    paragraph.append(
        make_revision_element(
            "ins",
            revision_id=insert_revision_id,
            author=author,
            date=insert_date,
            text=replacement_text,
            deleted=False,
            run_properties=run_properties,
        )
    )
    append_comment_end_and_reference(paragraph, comment_id)
    append_plain_run(paragraph, suffix, run_properties=run_properties)


def replace_paragraph_words(
    paragraph: etree._Element,
    *,
    original_text: str,
    replacement_text: str,
    author: str,
    base_revision_id: int,
    max_revision_id: int,
    comment_id: int,
    base_date,
) -> tuple[bool, str, int]:
    """Fine-grained replacement: only changed characters get w:del / w:ins.

    *max_revision_id* is the exclusive upper bound for revision IDs this call
    may use.  Returns ``(applied, granularity, ids_used)`` where *applied* is
    True when the fine-grained path succeeded, *granularity* is ``"word"`` or
    ``"sentence"`` (fallback), and *ids_used* is the number of revision IDs
    consumed.
    """
    # --- locate target_text inside the paragraph's runs ---
    span = find_target_run_span(paragraph, original_text)
    if span is None:
        # Cannot locate precisely → fall back to whole-sentence replace
        return False, "sentence", 0

    start_run_index, start_offset, end_run_index, end_offset = span
    runs = paragraph_runs(paragraph)

    # Collect run properties from the first run for new runs we create
    run_properties = first_run_properties(paragraph)

    # --- detach prefix runs, target runs, suffix runs ---
    prefix_runs = runs[:start_run_index]
    target_runs = runs[start_run_index : end_run_index + 1]
    suffix_runs = runs[end_run_index + 1 :]

    # Extract the actual text from the target run span
    # (may differ from original_text if runs split oddly, but should match)
    target_piece_texts = []
    for i, run in enumerate(target_runs):
        text = direct_run_text(run)
        if i == 0 and i == len(target_runs) - 1:
            # target is within a single run
            target_piece_texts.append(text[start_offset:end_offset])
        elif i == 0:
            target_piece_texts.append(text[start_offset:])
        elif i == len(target_runs) - 1:
            target_piece_texts.append(text[:end_offset])
        else:
            target_piece_texts.append(text)
    actual_target = "".join(target_piece_texts)

    # --- compute character-level diff ---
    matcher = SequenceMatcher(None, actual_target, replacement_text, autojunk=False)
    opcodes = matcher.get_opcodes()

    # Count required revision IDs and check budget
    required_ids = 0
    for tag, i1, i2, j1, j2 in opcodes:
        if tag == "delete":
            required_ids += 1
        elif tag == "insert":
            required_ids += 1
        elif tag == "replace":
            if actual_target[i1:i2]:
                required_ids += 1
            if replacement_text[j1:j2]:
                required_ids += 1
    if required_ids > (max_revision_id - base_revision_id):
        return False, "sentence", 0

    # --- rebuild paragraph ---
    # Remove all children except w:pPr
    remove_paragraph_content(paragraph)

    # Re-attach prefix runs (preserve originals)
    for run in prefix_runs:
        paragraph.append(run)

    # Comment range start (covers entire target region)
    append_comment_start(paragraph, comment_id)

    # Apply diff opcodes
    rev_id = base_revision_id
    for tag, i1, i2, j1, j2 in opcodes:
        if tag == "equal":
            text = actual_target[i1:i2]
            if text:
                append_plain_run(paragraph, text, run_properties=run_properties)
        elif tag == "delete":
            text = actual_target[i1:i2]
            if text:
                paragraph.append(
                    make_revision_element(
                        "del",
                        revision_id=rev_id,
                        author=author,
                        date=base_date,
                        text=text,
                        deleted=True,
                        run_properties=run_properties,
                    )
                )
                rev_id += 1
        elif tag == "insert":
            text = replacement_text[j1:j2]
            if text:
                paragraph.append(
                    make_revision_element(
                        "ins",
                        revision_id=rev_id,
                        author=author,
                        date=base_date,
                        text=text,
                        deleted=False,
                        run_properties=run_properties,
                    )
                )
                rev_id += 1
        elif tag == "replace":
            del_text = actual_target[i1:i2]
            ins_text = replacement_text[j1:j2]
            if del_text:
                paragraph.append(
                    make_revision_element(
                        "del",
                        revision_id=rev_id,
                        author=author,
                        date=base_date,
                        text=del_text,
                        deleted=True,
                        run_properties=run_properties,
                    )
                )
                rev_id += 1
            if ins_text:
                paragraph.append(
                    make_revision_element(
                        "ins",
                        revision_id=rev_id,
                        author=author,
                        date=base_date,
                        text=ins_text,
                        deleted=False,
                        run_properties=run_properties,
                    )
                )
                rev_id += 1

    # Comment range end + reference (covers entire target region)
    append_comment_end_and_reference(paragraph, comment_id)

    # Re-attach suffix runs (preserve originals)
    for run in suffix_runs:
        paragraph.append(run)

    return True, "word", required_ids


def replace_single_paragraph_clause(
    paragraph: etree._Element,
    *,
    original_text: str,
    replacement_text: str,
    author: str,
    delete_revision_id: int,
    insert_revision_id: int,
    comment_id: int,
    delete_date,
    insert_date,
) -> None:
    if normalize_text(paragraph_text(paragraph)) != normalize_text(original_text):
        raise ValueError("replace_clause requires target_text to cover the full paragraph")
    run_properties = first_run_properties(paragraph)
    remove_paragraph_content(paragraph)
    append_comment_start(paragraph, comment_id)
    paragraph.append(
        make_revision_element(
            "del",
            revision_id=delete_revision_id,
            author=author,
            date=delete_date,
            text=original_text,
            deleted=True,
            run_properties=run_properties,
        )
    )
    paragraph.append(
        make_revision_element(
            "ins",
            revision_id=insert_revision_id,
            author=author,
            date=insert_date,
            text=replacement_text,
            deleted=False,
            run_properties=run_properties,
        )
    )
    append_comment_end_and_reference(paragraph, comment_id)


def range_text(paragraphs: list[etree._Element]) -> str:
    return "\n".join(paragraph_text(paragraph) for paragraph in paragraphs)


def find_paragraph_range(
    document_root: etree._Element,
    *,
    paragraph_range: dict | None,
    target_text: str,
    normalized_text_hash: str | None,
) -> tuple[list[etree._Element] | None, str | None]:
    expected_hash = normalized_text_hash or ""
    if (
        expected_hash
        and normalize_text(target_text)
        and expected_hash != normalized_hash(target_text)
    ):
        return None, "target_text_hash_mismatch"

    if not paragraph_range:
        normalized_target = normalize_text(target_text)
        paragraphs = iter_document_paragraphs(document_root)
        matches = []
        for start_index, start_paragraph in enumerate(paragraphs):
            parent = start_paragraph.getparent()
            pieces = []
            for end_index in range(start_index, len(paragraphs)):
                paragraph = paragraphs[end_index]
                if paragraph.getparent() is not parent:
                    break
                pieces.append(paragraph_text(paragraph))
                candidate_text = "\n".join(pieces)
                normalized_candidate = normalize_text(candidate_text)
                if normalized_candidate == normalized_target:
                    matches.append(paragraphs[start_index : end_index + 1])
                if len(normalized_candidate) > len(normalized_target):
                    break
        if len(matches) == 1:
            return matches[0], None
        if not matches:
            return None, "paragraph_range_text_not_found"
        return None, "paragraph_range_text_not_unique"

    start = paragraph_range.get("start_paragraph_index")
    end = paragraph_range.get("end_paragraph_index")
    if start is None or end is None:
        return None, "incomplete_paragraph_range"
    if start < 0 or end < start:
        return None, "invalid_paragraph_range"
    paragraphs = iter_document_paragraphs(document_root)
    if end >= len(paragraphs):
        return None, "paragraph_range_out_of_range"
    selected = paragraphs[start : end + 1]
    if not selected:
        return None, "paragraph_range_empty"
    parent = selected[0].getparent()
    if any(paragraph.getparent() is not parent for paragraph in selected):
        return None, "paragraph_range_crosses_parent_boundary"
    if normalize_text(range_text(selected)) != normalize_text(target_text):
        return None, "paragraph_range_text_mismatch"
    return selected, None


def replace_paragraph_range_clause(
    paragraphs: list[etree._Element],
    *,
    original_text: str,
    replacement_text: str,
    author: str,
    delete_revision_id: int,
    insert_revision_id: int,
    comment_id: int,
    delete_date,
    insert_date,
) -> None:
    if normalize_text(range_text(paragraphs)) != normalize_text(original_text):
        raise ValueError("replace_clause range text does not match target_text")
    for index, paragraph in enumerate(paragraphs):
        run_properties = first_run_properties(paragraph)
        current_text = paragraph_text(paragraph)
        remove_paragraph_content(paragraph)
        if index == 0:
            append_comment_start(paragraph, comment_id)
        if current_text:
            paragraph.append(
                make_revision_element(
                    "del",
                    revision_id=delete_revision_id,
                    author=author,
                    date=delete_date,
                    text=current_text,
                    deleted=True,
                    run_properties=run_properties,
                )
            )
        if index == 0:
            paragraph.append(
                make_revision_element(
                    "ins",
                    revision_id=insert_revision_id,
                    author=author,
                    date=insert_date,
                    text=replacement_text,
                    deleted=False,
                    run_properties=run_properties,
                )
            )
        if index == len(paragraphs) - 1:
            append_comment_end_and_reference(paragraph, comment_id)


def delete_paragraph_sentence(
    paragraph: etree._Element,
    *,
    original_text: str,
    author: str,
    delete_revision_id: int,
    comment_id: int,
    delete_date,
) -> None:
    run_properties = first_run_properties(paragraph)
    prefix, suffix = split_target_text(paragraph, original_text)
    remove_paragraph_content(paragraph)
    append_plain_run(paragraph, prefix, run_properties=run_properties)
    append_comment_start(paragraph, comment_id)
    paragraph.append(
        make_revision_element(
            "del",
            revision_id=delete_revision_id,
            author=author,
            date=delete_date,
            text=original_text,
            deleted=True,
            run_properties=run_properties,
        )
    )
    append_comment_end_and_reference(paragraph, comment_id)
    append_plain_run(paragraph, suffix, run_properties=run_properties)


def insert_sentence_after_target(
    paragraph: etree._Element,
    *,
    target_text: str,
    insertion_text: str,
    author: str,
    insert_revision_id: int,
    comment_id: int,
    insert_date,
) -> None:
    run_properties = first_run_properties(paragraph)
    prefix, suffix = split_target_text(paragraph, target_text)
    remove_paragraph_content(paragraph)
    append_plain_run(paragraph, prefix, run_properties=run_properties)
    append_plain_run(paragraph, target_text, run_properties=run_properties)
    append_comment_start(paragraph, comment_id)
    paragraph.append(
        make_revision_element(
            "ins",
            revision_id=insert_revision_id,
            author=author,
            date=insert_date,
            text=insertion_text,
            deleted=False,
            run_properties=run_properties,
        )
    )
    append_comment_end_and_reference(paragraph, comment_id)
    append_plain_run(paragraph, suffix, run_properties=run_properties)


def event_count(action: dict[str, Any]) -> int:
    if action.get("action_type") in {"replace_sentence", "replace_words", "replace_clause"}:
        return 3
    if action.get("action_type") in {"delete_sentence", "insert_sentence_after"}:
        return 2
    if action.get("action_type") == "comment_only":
        return 1
    return 0


def timestamp_record(event: str, value) -> dict[str, str]:
    return {
        "event": event,
        "display_timestamp": iso_beijing(value),
        "display_timestamp_utc": format_ooxml_utc(value),
    }


def run_redline(
    *,
    input_docx: Path,
    edit_plan_path: Path,
    output_docx: Path | None,
    log_path: Path,
    dry_run: bool = False,
    timestamp_mode: str = "synthetic_spread",
    spread_minutes: int = 120,
    random_seed: int | None = None,
    now: str | None = None,
    author: str = DEFAULT_AUTHOR,
) -> dict[str, Any]:
    input_docx = Path(input_docx)
    output_docx = Path(output_docx) if output_docx else None
    require_docx_paths(input_docx, output_docx)
    if output_docx is None and not dry_run:
        raise ValueError("output_docx is required unless --dry-run is used")

    plan = load_edit_plan(edit_plan_path)
    plan_errors = validate_edit_plan(plan)
    if plan_errors:
        raise ValueError("; ".join(plan_errors))

    package_parts, part_map = read_package(input_docx)
    document_root = parse_xml(part_map["word/document.xml"])
    settings_root, _ = get_or_create_xml(
        part_map,
        "word/settings.xml",
        etree.Element(qn("w", "settings"), nsmap={"w": NS["w"]}),
    )
    comments_root, _ = get_or_create_xml(part_map, "word/comments.xml", new_comments_root())
    rels_root, _ = get_or_create_xml(
        part_map, "word/_rels/document.xml.rels", default_rels_root()
    )
    content_types_root = parse_xml(part_map["[Content_Types].xml"])

    actual_generated_at = parse_datetime(now)
    total_events = sum(event_count(action) for action in plan["actions"])
    timestamps = generate_display_timestamps(
        total_events,
        actual_generated_at,
        spread_minutes=spread_minutes,
        mode=timestamp_mode,
        random_seed=random_seed,
    )
    timestamp_iter = iter(timestamps)

    log = {
        "document_id": plan["document_id"],
        "input_docx": str(input_docx),
        "output_docx": str(output_docx) if output_docx else None,
        "dry_run": dry_run,
        "actual_generated_at": iso_beijing(actual_generated_at),
        "timezone": "Asia/Shanghai",
        "author": author,
        "display_timestamp_mode": timestamp_mode,
        "spread_minutes": spread_minutes,
        "actions": [],
        "unexecuted_actions": [],
    }

    ensure_track_revisions(settings_root)
    ensure_comments_relationship(rels_root)
    ensure_comments_content_type(content_types_root)

    next_rev_id = next_revision_id(document_root)
    next_comm_id = next_comment_id(document_root, comments_root)
    changed = False

    for action in plan["actions"]:
        action_type = action["action_type"]
        target = action["target"]
        action_log = {
            "action_id": action["action_id"],
            "action_type": action_type,
            "target_text": target["target_text"],
            "replacement_text": action.get("replacement_text"),
            "comment": (action.get("comment") or "").strip(),
            "status": None,
            "reason": None,
            "revision_ids": [],
            "comment_id": None,
            "display_timestamps": [],
        }

        if action_type not in SUPPORTED_ACTIONS:
            action_log["status"] = "skipped"
            action_log["reason"] = "unsupported_action_type"
            log["actions"].append(action_log)
            continue

        paragraph = None
        paragraph_range = None
        if target.get("container_type") == "paragraph_range":
            if action_type != "replace_clause":
                action_log["status"] = "unresolved"
                action_log["reason"] = "paragraph_range_only_supported_for_replace_clause"
                log["actions"].append(action_log)
                continue
            paragraph_range, reason = find_paragraph_range(
                document_root,
                paragraph_range=target.get("paragraph_range"),
                target_text=target["target_text"],
                normalized_text_hash=target.get("normalized_text_hash"),
            )
            if paragraph_range is None:
                action_log["status"] = "unresolved"
                action_log["reason"] = reason
                log["actions"].append(action_log)
                continue
        else:
            paragraph, reason = find_target_paragraph(
                document_root,
                container_type=target.get("container_type"),
                paragraph_index=target.get("paragraph_index"),
                table_path=target.get("table_path"),
                target_text=target["target_text"],
                normalized_text_hash=target.get("normalized_text_hash"),
            )
            if paragraph is None:
                action_log["status"] = "unresolved"
                action_log["reason"] = reason
                log["actions"].append(action_log)
                continue
            if (
                action_type == "replace_clause"
                and normalize_text(paragraph_text(paragraph))
                != normalize_text(target["target_text"])
            ):
                action_log["status"] = "unresolved"
                action_log["reason"] = "replace_clause_requires_full_paragraph_target"
                log["actions"].append(action_log)
                continue

        if action_type == "replace_sentence":
            delete_ts = next(timestamp_iter)
            insert_ts = next(timestamp_iter)
            comment_ts = next(timestamp_iter)
            action_log["display_timestamps"].extend(
                [
                    timestamp_record("delete_revision", delete_ts),
                    timestamp_record("insert_revision", insert_ts),
                    timestamp_record("comment", comment_ts),
                ]
            )
            action_log["revision_ids"] = [next_rev_id, next_rev_id + 1]
            action_log["comment_id"] = next_comm_id
            if dry_run:
                action_log["status"] = "would_apply"
            else:
                replace_paragraph_sentence(
                    paragraph,
                    original_text=target["target_text"],
                    replacement_text=action["replacement_text"],
                    author=author,
                    delete_revision_id=next_rev_id,
                    insert_revision_id=next_rev_id + 1,
                    comment_id=next_comm_id,
                    delete_date=delete_ts,
                    insert_date=insert_ts,
                )
                comments_root.append(
                    make_comment(
                        next_comm_id,
                        author,
                        comment_ts,
                        action_log["comment"],
                    )
                )
                action_log["status"] = "applied"
                changed = True
            next_rev_id += 2
            next_comm_id += 1

        elif action_type == "replace_words":
            delete_ts = next(timestamp_iter)
            insert_ts = next(timestamp_iter)
            comment_ts = next(timestamp_iter)
            action_log["display_timestamps"].extend(
                [
                    timestamp_record("delete_revision", delete_ts),
                    timestamp_record("insert_revision", insert_ts),
                    timestamp_record("comment", comment_ts),
                ]
            )
            action_log["revision_ids"] = [next_rev_id, next_rev_id + 1]
            action_log["comment_id"] = next_comm_id
            if dry_run:
                action_log["status"] = "would_apply"
            else:
                # Budget: reserve at most 20 revision IDs for fine-grained diff
                max_rev_budget = 20
                applied, granularity, ids_used = replace_paragraph_words(
                    paragraph,
                    original_text=target["target_text"],
                    replacement_text=action["replacement_text"],
                    author=author,
                    base_revision_id=next_rev_id,
                    max_revision_id=next_rev_id + max_rev_budget,
                    comment_id=next_comm_id,
                    base_date=delete_ts,
                )
                action_log["granularity"] = granularity
                if not applied:
                    # Fallback: fine-grained failed, use sentence-level
                    replace_paragraph_sentence(
                        paragraph,
                        original_text=target["target_text"],
                        replacement_text=action["replacement_text"],
                        author=author,
                        delete_revision_id=next_rev_id,
                        insert_revision_id=next_rev_id + 1,
                        comment_id=next_comm_id,
                        delete_date=delete_ts,
                        insert_date=insert_ts,
                    )
                    ids_used = 2
                action_log["revision_ids"] = list(
                    range(next_rev_id, next_rev_id + ids_used)
                )
                comments_root.append(
                    make_comment(
                        next_comm_id,
                        author,
                        comment_ts,
                        action_log["comment"],
                    )
                )
                action_log["status"] = "applied"
                changed = True
                next_rev_id += ids_used
            next_comm_id += 1

        elif action_type == "replace_clause":
            delete_ts = next(timestamp_iter)
            insert_ts = next(timestamp_iter)
            comment_ts = next(timestamp_iter)
            action_log["display_timestamps"].extend(
                [
                    timestamp_record("delete_revision", delete_ts),
                    timestamp_record("insert_revision", insert_ts),
                    timestamp_record("comment", comment_ts),
                ]
            )
            action_log["revision_ids"] = [next_rev_id, next_rev_id + 1]
            action_log["comment_id"] = next_comm_id
            if dry_run:
                action_log["status"] = "would_apply"
            else:
                if paragraph_range is not None:
                    replace_paragraph_range_clause(
                        paragraph_range,
                        original_text=target["target_text"],
                        replacement_text=action["replacement_text"],
                        author=author,
                        delete_revision_id=next_rev_id,
                        insert_revision_id=next_rev_id + 1,
                        comment_id=next_comm_id,
                        delete_date=delete_ts,
                        insert_date=insert_ts,
                    )
                else:
                    replace_single_paragraph_clause(
                        paragraph,
                        original_text=target["target_text"],
                        replacement_text=action["replacement_text"],
                        author=author,
                        delete_revision_id=next_rev_id,
                        insert_revision_id=next_rev_id + 1,
                        comment_id=next_comm_id,
                        delete_date=delete_ts,
                        insert_date=insert_ts,
                    )
                comments_root.append(
                    make_comment(
                        next_comm_id,
                        author,
                        comment_ts,
                        action_log["comment"],
                    )
                )
                action_log["status"] = "applied"
                changed = True
            next_rev_id += 2
            next_comm_id += 1

        elif action_type == "comment_only":
            comment_ts = next(timestamp_iter)
            action_log["display_timestamps"].append(
                timestamp_record("comment", comment_ts)
            )
            action_log["comment_id"] = next_comm_id
            if dry_run:
                action_log["status"] = "would_apply"
            else:
                add_comment_to_paragraph(paragraph, comment_id=next_comm_id)
                comments_root.append(
                    make_comment(
                        next_comm_id,
                        author,
                        comment_ts,
                        action_log["comment"],
                    )
                )
                action_log["status"] = "applied"
                changed = True
            next_comm_id += 1

        elif action_type == "delete_sentence":
            delete_ts = next(timestamp_iter)
            comment_ts = next(timestamp_iter)
            action_log["display_timestamps"].extend(
                [
                    timestamp_record("delete_revision", delete_ts),
                    timestamp_record("comment", comment_ts),
                ]
            )
            action_log["revision_ids"] = [next_rev_id]
            action_log["comment_id"] = next_comm_id
            if dry_run:
                action_log["status"] = "would_apply"
            else:
                delete_paragraph_sentence(
                    paragraph,
                    original_text=target["target_text"],
                    author=author,
                    delete_revision_id=next_rev_id,
                    comment_id=next_comm_id,
                    delete_date=delete_ts,
                )
                comments_root.append(
                    make_comment(
                        next_comm_id,
                        author,
                        comment_ts,
                        action_log["comment"],
                    )
                )
                action_log["status"] = "applied"
                changed = True
            next_rev_id += 1
            next_comm_id += 1

        elif action_type == "insert_sentence_after":
            insert_ts = next(timestamp_iter)
            comment_ts = next(timestamp_iter)
            action_log["display_timestamps"].extend(
                [
                    timestamp_record("insert_revision", insert_ts),
                    timestamp_record("comment", comment_ts),
                ]
            )
            action_log["revision_ids"] = [next_rev_id]
            action_log["comment_id"] = next_comm_id
            if dry_run:
                action_log["status"] = "would_apply"
            else:
                insert_sentence_after_target(
                    paragraph,
                    target_text=target["target_text"],
                    insertion_text=action["replacement_text"],
                    author=author,
                    insert_revision_id=next_rev_id,
                    comment_id=next_comm_id,
                    insert_date=insert_ts,
                )
                comments_root.append(
                    make_comment(
                        next_comm_id,
                        author,
                        comment_ts,
                        action_log["comment"],
                    )
                )
                action_log["status"] = "applied"
                changed = True
            next_rev_id += 1
            next_comm_id += 1

        log["actions"].append(action_log)

    if not dry_run:
        changed_parts = {
            "word/document.xml": serialize_xml(document_root),
            "word/settings.xml": serialize_xml(settings_root),
            "word/comments.xml": serialize_xml(comments_root),
            "word/_rels/document.xml.rels": serialize_xml(rels_root),
            "[Content_Types].xml": serialize_xml(content_types_root),
        }
        tmp_output = output_docx.with_suffix(output_docx.suffix + ".tmp")
        try:
            write_package(tmp_output, package_parts, part_map, changed_parts)
            os.replace(tmp_output, output_docx)
        except Exception:
            if tmp_output.exists():
                tmp_output.unlink()
            raise

    log["unexecuted_actions"] = [
        {
            "action_id": item["action_id"],
            "status": item["status"],
            "reason": item["reason"],
        }
        for item in log["actions"]
        if item["status"] not in {"applied", "would_apply"}
    ]
    log["summary"] = {
        "applied": sum(1 for item in log["actions"] if item["status"] == "applied"),
        "would_apply": sum(1 for item in log["actions"] if item["status"] == "would_apply"),
        "unresolved": sum(1 for item in log["actions"] if item["status"] == "unresolved"),
        "skipped": sum(1 for item in log["actions"] if item["status"] == "skipped"),
        "not_executed": len(log["unexecuted_actions"]),
        "changed": changed,
    }
    log_path.parent.mkdir(parents=True, exist_ok=True)
    log_path.write_text(json.dumps(log, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return log


def main() -> int:
    parser = argparse.ArgumentParser(description="Write DOCX tracked changes and comments.")
    parser.add_argument("input_docx")
    parser.add_argument("edit_plan")
    parser.add_argument("--output")
    parser.add_argument("--log", required=True)
    parser.add_argument("--dry-run", action="store_true")
    parser.add_argument("--timestamp-mode", default="synthetic_spread")
    parser.add_argument("--spread-minutes", type=int, default=120)
    parser.add_argument("--random-seed", type=int)
    parser.add_argument("--now")
    parser.add_argument("--author", default=DEFAULT_AUTHOR)
    args = parser.parse_args()

    run_redline(
        input_docx=Path(args.input_docx),
        edit_plan_path=Path(args.edit_plan),
        output_docx=Path(args.output) if args.output else None,
        log_path=Path(args.log),
        dry_run=args.dry_run,
        timestamp_mode=args.timestamp_mode,
        spread_minutes=args.spread_minutes,
        random_seed=args.random_seed,
        now=args.now,
        author=args.author,
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
