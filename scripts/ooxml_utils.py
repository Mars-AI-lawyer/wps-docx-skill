from __future__ import annotations

import copy
import hashlib
import re
from datetime import datetime, timedelta, timezone
from pathlib import Path
from random import Random
from typing import Iterable
from zoneinfo import ZoneInfo

from lxml import etree


NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
}

WORD_REL_COMMENTS = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
)
COMMENTS_CONTENT_TYPE = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"
)
XML_NS = "http://www.w3.org/XML/1998/namespace"
BEIJING = ZoneInfo("Asia/Shanghai")
AUTHOR = "Reviewer"


def qn(prefix: str, tag: str) -> str:
    return f"{{{NS[prefix]}}}{tag}"


def parse_xml(data: bytes) -> etree._Element:
    parser = etree.XMLParser(remove_blank_text=False, resolve_entities=False)
    return etree.fromstring(data, parser=parser)


def serialize_xml(root: etree._Element) -> bytes:
    return etree.tostring(
        root,
        xml_declaration=True,
        encoding="UTF-8",
        standalone=None,
    )


def normalize_text(text: str) -> str:
    return re.sub(r"\s+", "", text or "")


def normalized_hash(text: str) -> str:
    return hashlib.sha256(normalize_text(text).encode("utf-8")).hexdigest()


def paragraph_text(paragraph: etree._Element) -> str:
    pieces = []
    for node in paragraph.xpath(".//w:t | .//w:delText", namespaces=NS):
        pieces.append(node.text or "")
    return "".join(pieces)


def paragraph_runs(paragraph: etree._Element) -> list[etree._Element]:
    return list(paragraph.xpath("./w:r", namespaces=NS))


def direct_run_text(run: etree._Element) -> str:
    return "".join(t.text or "" for t in run.xpath("./w:t", namespaces=NS))


def iter_document_paragraphs(document_root: etree._Element) -> list[etree._Element]:
    return list(document_root.xpath(".//w:body//w:p", namespaces=NS))


def table_cell_paragraph(
    document_root: etree._Element,
    table_path: dict | None,
) -> tuple[etree._Element | None, str | None]:
    if not table_path:
        return None, "missing_table_path"
    tables = document_root.xpath(".//w:tbl", namespaces=NS)
    table_index = table_path.get("table_index")
    row_index = table_path.get("row_index")
    cell_index = table_path.get("cell_index")
    paragraph_index = table_path.get("paragraph_index_in_cell", 0)
    if None in (table_index, row_index, cell_index, paragraph_index):
        return None, "incomplete_table_path"
    if table_index < 0 or table_index >= len(tables):
        return None, "table_index_out_of_range"
    rows = tables[table_index].xpath("./w:tr", namespaces=NS)
    if row_index < 0 or row_index >= len(rows):
        return None, "row_index_out_of_range"
    cells = rows[row_index].xpath("./w:tc", namespaces=NS)
    if cell_index < 0 or cell_index >= len(cells):
        return None, "cell_index_out_of_range"
    paragraphs = cells[cell_index].xpath("./w:p", namespaces=NS)
    if paragraph_index < 0 or paragraph_index >= len(paragraphs):
        return None, "cell_paragraph_index_out_of_range"
    return paragraphs[paragraph_index], None


def has_exactly_one_target(paragraph: etree._Element, target_text: str) -> bool:
    return paragraph_text(paragraph).count(target_text) == 1


def find_unique_paragraph(
    document_root: etree._Element,
    *,
    paragraph_index: int | None,
    target_text: str,
    normalized_text_hash: str | None = None,
) -> tuple[etree._Element | None, str | None]:
    paragraphs = iter_document_paragraphs(document_root)
    normalized_target = normalize_text(target_text)
    expected_hash = normalized_text_hash or normalized_hash(target_text)

    def is_match(paragraph: etree._Element) -> bool:
        text = paragraph_text(paragraph)
        if normalize_text(text) != normalized_target:
            return False
        if expected_hash and normalized_hash(text) != expected_hash:
            return False
        return True

    if paragraph_index is not None:
        if paragraph_index < 0 or paragraph_index >= len(paragraphs):
            return None, "paragraph_index_out_of_range"
        paragraph = paragraphs[paragraph_index]
        if is_match(paragraph):
            return paragraph, None
        return None, "paragraph_index_text_mismatch"

    matches = [paragraph for paragraph in paragraphs if is_match(paragraph)]
    if len(matches) == 1:
        return matches[0], None
    if not matches:
        return None, "target_text_not_found"
    return None, "target_text_not_unique"


def find_target_paragraph(
    document_root: etree._Element,
    *,
    container_type: str,
    paragraph_index: int | None,
    table_path: dict | None,
    target_text: str,
    normalized_text_hash: str | None = None,
) -> tuple[etree._Element | None, str | None]:
    expected_hash = normalized_text_hash or normalized_hash(target_text)
    if expected_hash and normalized_hash(target_text) != expected_hash:
        return None, "target_text_hash_mismatch"

    if container_type == "table_cell":
        paragraph, reason = table_cell_paragraph(document_root, table_path)
        if paragraph is None:
            return None, reason
        if has_exactly_one_target(paragraph, target_text):
            return paragraph, None
        count = paragraph_text(paragraph).count(target_text)
        return None, "target_text_not_found" if count == 0 else "target_text_not_unique"

    if container_type != "paragraph":
        return None, "unsupported_container_type"

    paragraphs = iter_document_paragraphs(document_root)
    if paragraph_index is not None:
        if paragraph_index < 0 or paragraph_index >= len(paragraphs):
            return None, "paragraph_index_out_of_range"
        paragraph = paragraphs[paragraph_index]
        if has_exactly_one_target(paragraph, target_text):
            return paragraph, None
        count = paragraph_text(paragraph).count(target_text)
        return None, "target_text_not_found" if count == 0 else "target_text_not_unique"

    matches = [
        paragraph
        for paragraph in paragraphs
        if has_exactly_one_target(paragraph, target_text)
    ]
    if len(matches) == 1:
        return matches[0], None
    if not matches:
        return None, "target_text_not_found"
    return None, "target_text_not_unique"


def first_run_properties(paragraph: etree._Element) -> etree._Element | None:
    run = paragraph.find(qn("w", "r"))
    if run is None:
        return None
    run_properties = run.find(qn("w", "rPr"))
    if run_properties is None:
        return None
    return copy.deepcopy(run_properties)


def remove_paragraph_content(paragraph: etree._Element) -> None:
    for child in list(paragraph):
        if child.tag != qn("w", "pPr"):
            paragraph.remove(child)


def text_element(tag: str, text: str) -> etree._Element:
    element = etree.Element(qn("w", tag))
    if text.startswith(" ") or text.endswith(" "):
        element.set(f"{{{XML_NS}}}space", "preserve")
    element.text = text
    return element


def make_revision_run(text: str, *, deleted: bool, run_properties: etree._Element | None):
    run = etree.Element(qn("w", "r"))
    if run_properties is not None:
        run.append(copy.deepcopy(run_properties))
    run.append(text_element("delText" if deleted else "t", text))
    return run


def append_plain_run(
    paragraph: etree._Element,
    text: str,
    *,
    run_properties: etree._Element | None,
) -> None:
    if not text:
        return
    run = etree.Element(qn("w", "r"))
    if run_properties is not None:
        run.append(copy.deepcopy(run_properties))
    run.append(text_element("t", text))
    paragraph.append(run)


def max_w_id(elements: Iterable[etree._Element]) -> int:
    max_id = -1
    for element in elements:
        value = element.get(qn("w", "id"))
        if value is None:
            continue
        try:
            max_id = max(max_id, int(value))
        except ValueError:
            continue
    return max_id


def next_revision_id(document_root: etree._Element) -> int:
    nodes = document_root.xpath(".//w:ins | .//w:del", namespaces=NS)
    return max_w_id(nodes) + 1


def next_comment_id(document_root: etree._Element, comments_root: etree._Element) -> int:
    nodes = list(comments_root.xpath(".//w:comment", namespaces=NS))
    nodes.extend(document_root.xpath(".//w:commentRangeStart", namespaces=NS))
    nodes.extend(document_root.xpath(".//w:commentRangeEnd", namespaces=NS))
    nodes.extend(document_root.xpath(".//w:commentReference", namespaces=NS))
    return max_w_id(nodes) + 1


def ensure_track_revisions(settings_root: etree._Element) -> bool:
    changed = False

    track_revisions = settings_root.find(qn("w", "trackRevisions"))
    if track_revisions is None:
        track_revisions = etree.Element(qn("w", "trackRevisions"))
        settings_root.append(track_revisions)
        changed = True
    if track_revisions.get(qn("w", "val")) != "1":
        track_revisions.set(qn("w", "val"), "1")
        changed = True

    show_revisions = settings_root.find(qn("w", "showRevisions"))
    if show_revisions is None:
        show_revisions = etree.Element(qn("w", "showRevisions"))
        settings_root.append(show_revisions)
        changed = True
    if show_revisions.get(qn("w", "val")) != "1":
        show_revisions.set(qn("w", "val"), "1")
        changed = True

    revision_view = settings_root.find(qn("w", "revisionView"))
    if revision_view is None:
        revision_view = etree.Element(qn("w", "revisionView"))
        settings_root.append(revision_view)
        changed = True
    markup = revision_view.find(qn("w", "markup"))
    if markup is None:
        markup = etree.Element(qn("w", "markup"))
        revision_view.append(markup)
        changed = True
    if markup.text != "1":
        markup.text = "1"
        changed = True

    return changed


def ensure_comments_relationship(rels_root: etree._Element) -> bool:
    existing = rels_root.xpath(
        "./rel:Relationship[@Type=$type]",
        namespaces=NS,
        type=WORD_REL_COMMENTS,
    )
    if existing:
        return False
    ids = []
    for rel in rels_root.xpath("./rel:Relationship", namespaces=NS):
        rel_id = rel.get("Id", "")
        if rel_id.startswith("rId") and rel_id[3:].isdigit():
            ids.append(int(rel_id[3:]))
    next_id = max(ids, default=0) + 1
    relationship = etree.Element(f"{{{NS['rel']}}}Relationship")
    relationship.set("Id", f"rId{next_id}")
    relationship.set("Type", WORD_REL_COMMENTS)
    relationship.set("Target", "comments.xml")
    rels_root.append(relationship)
    return True


def ensure_comments_content_type(content_types_root: etree._Element) -> bool:
    existing = content_types_root.xpath(
        "./ct:Override[@PartName='/word/comments.xml']",
        namespaces=NS,
    )
    if existing:
        return False
    override = etree.Element(f"{{{NS['ct']}}}Override")
    override.set("PartName", "/word/comments.xml")
    override.set("ContentType", COMMENTS_CONTENT_TYPE)
    content_types_root.append(override)
    return True


def new_comments_root() -> etree._Element:
    return etree.Element(qn("w", "comments"), nsmap={"w": NS["w"]})


def normalize_comment_text(comment: str) -> str:
    return (comment or "").strip()


def make_comment(comment_id: int, author: str, date_utc: datetime, text: str):
    comment = etree.Element(qn("w", "comment"))
    comment.set(qn("w", "id"), str(comment_id))
    comment.set(qn("w", "author"), author)
    comment.set(qn("w", "date"), format_ooxml_utc(date_utc))
    paragraph = etree.SubElement(comment, qn("w", "p"))
    run = etree.SubElement(paragraph, qn("w", "r"))
    run.append(text_element("t", normalize_comment_text(text)))
    return comment


def make_comment_reference_run(comment_id: int) -> etree._Element:
    run = etree.Element(qn("w", "r"))
    run_properties = etree.SubElement(run, qn("w", "rPr"))
    run_style = etree.SubElement(run_properties, qn("w", "rStyle"))
    run_style.set(qn("w", "val"), "CommentReference")
    reference = etree.SubElement(run, qn("w", "commentReference"))
    reference.set(qn("w", "id"), str(comment_id))
    return run


def parse_datetime(value: str | None) -> datetime:
    if not value:
        return datetime.now(tz=BEIJING)
    cleaned = value.strip()
    if cleaned.endswith("Z"):
        cleaned = f"{cleaned[:-1]}+00:00"
    parsed = datetime.fromisoformat(cleaned)
    if parsed.tzinfo is None:
        parsed = parsed.replace(tzinfo=BEIJING)
    return parsed.astimezone(BEIJING)


def generate_display_timestamps(
    count: int,
    now_beijing: datetime,
    *,
    spread_minutes: int = 120,
    mode: str = "synthetic_spread",
    random_seed: int | None = None,
) -> list[datetime]:
    if mode != "synthetic_spread":
        raise ValueError(f"Unsupported timestamp mode: {mode}")
    if count <= 0:
        return []
    now = now_beijing.astimezone(BEIJING)
    start = now - timedelta(minutes=spread_minutes)
    total_seconds = max(1, int((now - start).total_seconds()))
    rng = Random(random_seed)
    offsets = sorted(rng.sample(range(total_seconds + 1), k=min(count, total_seconds + 1)))
    while len(offsets) < count:
        offsets.append(offsets[-1])
    return [start + timedelta(seconds=offset) for offset in offsets[:count]]


def format_ooxml_utc(value: datetime) -> str:
    return value.astimezone(timezone.utc).replace(microsecond=0).isoformat().replace(
        "+00:00", "Z"
    )


def iso_beijing(value: datetime) -> str:
    return value.astimezone(BEIJING).replace(microsecond=0).isoformat()


def require_docx_paths(input_path: Path, output_path: Path | None) -> None:
    if output_path is None:
        return
    if input_path.resolve() == output_path.resolve():
        raise ValueError("output path must not overwrite the source DOCX")
