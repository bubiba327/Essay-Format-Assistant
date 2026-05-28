#!/usr/bin/env python3
"""Contract checks for evaluation report serialization."""

from __future__ import annotations

import csv
import json
from pathlib import Path
from tempfile import TemporaryDirectory

from PIL import Image

import report_thesis_eval as reports


def assert_true(value, message: str) -> None:
    if not value:
        raise AssertionError(message)


def main() -> int:
    results = [
        {
            "case_id": "table-format1-sample2",
            "description": "Regression case",
            "status": "pass",
            "qa_level": "fast",
            "sample": "/tmp/format.doc",
            "target": "/tmp/target.doc",
            "output": "/tmp/output.docx",
            "learned_roles": ["table_header", "table_body_value"],
            "missing_roles": [],
            "lint_codes": ["mixed_page_number_examples"],
            "unexpected_lint_codes": [],
            "visual_check": "skipped-fast",
            "blank_like_pages": [],
            "edge_overflow_pages": [],
            "selected_review_pages": [],
            "contact_sheet": None,
            "structural_status": "pass",
            "structural_issues": [],
            "structural_report": {
                "metrics": {
                    "table_count": 1,
                    "toc_entry_count": 6,
                    "heading_count": 4,
                    "figure_caption_count": 2,
                    "image_count": 2,
                    "equation_count": 1,
                    "protected_paragraph_count": 3,
                    "page_number_count": 1,
                },
                "checks": {
                    "captioned_tables": {
                        "captioned_table_count": 1,
                        "checked_cell_count": 5,
                    }
                }
            },
            "warnings": [],
            "failures": [],
            "performance": {
                "stage1_seconds": 8.25,
                "stage2_seconds": 13.5,
                "total_seconds": 21.75,
                "pdf_conversion_passes": 2,
                "png_raster_passes": 1,
                "toc_refreshed_entries": 57,
            },
        }
    ]
    with TemporaryDirectory() as temp_dir:
        paths = reports.write_reports(Path(temp_dir), results)
        payload = json.loads(paths["json"].read_text(encoding="utf-8"))
        markdown = paths["markdown"].read_text(encoding="utf-8")
        with paths["csv"].open(newline="", encoding="utf-8") as handle:
            rows = list(csv.DictReader(handle))
        assert_true(payload["status"] == "pass", "aggregate JSON passes")
        assert_true("table-format1-sample2" in markdown, "Markdown names the case")
        assert_true("mixed_page_number_examples" in markdown, "Markdown exposes accepted warnings")
        assert_true("skipped-fast" in markdown, "Markdown states visual rendering was skipped")
        assert_true(
            "- Edge-overflow pages: `not checked`" in markdown,
            "Fast report does not imply rendered edge checks",
        )
        assert_true("- Structural QA: `pass` (0 issue(s))" in markdown, "Markdown states structural QA")
        assert_true("TOC entries: 6" in markdown, "Markdown gives structural metric details")
        assert_true("Total runtime: `21.75s`" in markdown, "Markdown gives combined runtime")
        assert_true("PDF/PNG passes: `2/1`" in markdown, "Markdown gives render pass counts")
        assert_true("TOC refreshed entries: `57`" in markdown, "Markdown gives refreshed TOC count")
        assert_true(rows[0]["status"] == "pass", "CSV contains case status")
        assert_true(rows[0]["structural_status"] == "pass", "CSV exposes structural status")
        assert_true(rows[0]["total_seconds"] == "21.75", "CSV exposes total runtime")

    failing_result = dict(results[0])
    failing_result.update(
        {
            "case_id": "table-format1-sample2-unformatted",
            "status": "fail",
            "structural_status": "fail",
            "structural_issues": [
                {
                    "code": "captioned_table_cell_format_mismatch",
                    "table_index": 1,
                    "row": 1,
                    "column": 2,
                    "text": "count",
                    "role": "table_header",
                    "differences": [
                        {
                            "field": "size_pt",
                            "actual_display": "12 磅",
                            "expected_display": "9 磅",
                        }
                    ],
                }
            ],
            "visual_check": "rendered",
            "blank_like_pages": ["4"],
            "edge_overflow_pages": [],
            "selected_review_pages": [1, 4],
            "contact_sheet": "/tmp/contact-sheet.png",
            "warnings": ["Blank-like rendered pages require review: 4"],
        }
    )
    with TemporaryDirectory() as temp_dir:
        case_sheet = Path(temp_dir) / "case-contact-sheet.png"
        Image.new("RGB", (20, 20), "white").save(case_sheet)
        failing_result["contact_sheet"] = str(case_sheet)
        paths = reports.write_reports(Path(temp_dir), [failing_result])
        markdown = paths["markdown"].read_text(encoding="utf-8")
        with paths["csv"].open(newline="", encoding="utf-8") as handle:
            rows = list(csv.DictReader(handle))
        assert_true("- Structural QA: `fail` (1 issue(s))" in markdown, "Failure count is visible")
        assert_true("table 1, row 1, column 2" in markdown, "Failure location is visible")
        assert_true("size_pt" in markdown, "Mismatch field is visible")
        assert_true(str(case_sheet) in markdown, "Rendered contact sheet is linked")
        assert_true(paths["contact_sheet"].exists(), "Run-level contact sheet is generated")
        assert_true("Run contact sheet:" in markdown, "Run-level contact sheet is reported")
        assert_true("Manual review pages: `1, 4`" in markdown, "Selected pages are visible")
        assert_true("Blank-like rendered pages require review: 4" in markdown, "Warnings are visible")
        assert_true(rows[0]["structural_issue_count"] == "1", "CSV exposes mismatch count")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
