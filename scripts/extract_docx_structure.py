#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import zipfile
from pathlib import Path

from lxml import etree

from ooxml_utils import (
    NS,
    direct_run_text,
    iter_document_paragraphs,
    normalize_text,
    normalized_hash,
    paragraph_runs,
    paragraph_text,
    parse_xml,
)


def extract_docx_structure(docx_path: Path) -> dict:
    with zipfile.ZipFile(docx_path) as package:
        document_root = parse_xml(package.read("word/document.xml"))

    paragraphs = []
    for paragraph_index, paragraph in enumerate(iter_document_paragraphs(document_root)):
        text = paragraph_text(paragraph)
        runs = [
            {"run_index": index, "text": direct_run_text(run)}
            for index, run in enumerate(paragraph_runs(paragraph))
        ]
        paragraphs.append(
            {
                "paragraph_index": paragraph_index,
                "container_type": "paragraph",
                "table_path": None,
                "text": text,
                "normalized_text": normalize_text(text),
                "text_hash": normalized_hash(text),
                "normalized_text_hash": normalized_hash(text),
                "runs": runs,
            }
        )

    tables = []
    for table_index, table in enumerate(document_root.xpath(".//w:tbl", namespaces=NS)):
        rows = []
        for row_index, row in enumerate(table.xpath("./w:tr", namespaces=NS)):
            cells = []
            for cell_index, cell in enumerate(row.xpath("./w:tc", namespaces=NS)):
                cell_paragraphs = []
                for paragraph_index, paragraph in enumerate(
                    cell.xpath("./w:p", namespaces=NS)
                ):
                    text = paragraph_text(paragraph)
                    cell_paragraphs.append(
                        {
                            "paragraph_index_in_cell": paragraph_index,
                            "container_type": "table_cell",
                            "table_path": {
                                "table_index": table_index,
                                "row_index": row_index,
                                "cell_index": cell_index,
                                "paragraph_index_in_cell": paragraph_index,
                            },
                            "text": text,
                            "normalized_text": normalize_text(text),
                            "text_hash": normalized_hash(text),
                            "normalized_text_hash": normalized_hash(text),
                            "runs": [
                                {"run_index": index, "text": direct_run_text(run)}
                                for index, run in enumerate(paragraph_runs(paragraph))
                            ],
                        }
                    )
                cells.append({"cell_index": cell_index, "paragraphs": cell_paragraphs})
            rows.append({"row_index": row_index, "cells": cells})
        tables.append({"table_index": table_index, "rows": rows})

    return {"paragraphs": paragraphs, "tables": tables}


def main() -> int:
    parser = argparse.ArgumentParser(description="Extract minimal DOCX structure.")
    parser.add_argument("docx_path")
    parser.add_argument("--output")
    args = parser.parse_args()

    result = extract_docx_structure(Path(args.docx_path))
    payload = json.dumps(result, ensure_ascii=False, indent=2)
    if args.output:
        Path(args.output).write_text(payload + "\n", encoding="utf-8")
    else:
        print(payload)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
