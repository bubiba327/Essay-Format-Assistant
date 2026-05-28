#!/usr/bin/env python3
"""Contract checks for stable evaluation baselines and drift reporting."""

from __future__ import annotations

from pathlib import Path
from tempfile import TemporaryDirectory

import evaluate_thesis_skill as evaluator
import report_thesis_eval as reports


def assert_equal(actual, expected, message: str) -> None:
    if actual != expected:
        raise AssertionError(f"{message}: expected {expected!r}, got {actual!r}")


def assert_true(value, message: str) -> None:
    if not value:
        raise AssertionError(message)


def fixture_result() -> dict:
    return {
        "case_id": "stable-case",
        "description": "Stable baseline case",
        "status": "pass",
        "qa_level": "fast",
        "sample": "/tmp/sample.docx",
        "target": "/tmp/target.docx",
        "output": "/tmp/output.docx",
        "learned_roles": ["body", "table_header", "equation"],
        "missing_roles": [],
        "lint_codes": ["known-warning"],
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
                "table_count": 2,
                "table_cell_count": 12,
                "toc_entry_count": 5,
                "heading_count": 3,
                "figure_caption_count": 4,
                "image_count": 4,
                "equation_count": 1,
                "protected_paragraph_count": 2,
                "protected_table_count": 1,
                "page_number_count": 1,
            },
            "checks": {"captioned_tables": {"captioned_table_count": 2, "checked_cell_count": 12}},
        },
        "warnings": [],
        "failures": [],
    }


def main() -> int:
    result = fixture_result()
    summary = evaluator.baseline_summary(result)
    assert_equal(summary["role_count"], 3, "role count is stable")
    assert_equal(summary["warning_count"], 1, "warning count is stable")
    assert_equal(summary["table_count"], 2, "table count is stable")
    assert_equal(summary["image_count"], 4, "image count is stable")
    assert_equal(summary["equation_count"], 1, "equation count is stable")

    with TemporaryDirectory() as temp_dir:
        baseline_path = Path(temp_dir) / "expected-qa" / "stable-case.json"
        evaluator.update_baseline(baseline_path, summary)
        stored = evaluator.read_json(baseline_path)
        assert_equal(evaluator.compare_baseline(summary, stored), [], "stored baseline matches")
        drifted = dict(summary)
        drifted["image_count"] = 5
        differences = evaluator.compare_baseline(drifted, stored)
        assert_equal(differences[0]["field"], "image_count", "image count drift is diagnosed")

        changed = fixture_result()
        changed["status"] = "fail"
        changed["baseline_status"] = "changed"
        changed["baseline_differences"] = differences
        paths = reports.write_reports(Path(temp_dir) / "report", [changed])
        markdown = paths["markdown"].read_text(encoding="utf-8")
        assert_true("- Baseline: `changed`" in markdown, "report exposes baseline status")
        assert_true("image_count: expected `4`, actual `5`" in markdown, "report explains drift")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
