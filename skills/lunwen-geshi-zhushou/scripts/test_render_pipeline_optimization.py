#!/usr/bin/env python3
"""Contract checks for TOC pagination probing before PNG visual QA."""

from __future__ import annotations

from pathlib import Path
from tempfile import TemporaryDirectory

import thesis_format_from_sample as fmt


def assert_equal(actual, expected, message: str) -> None:
    if actual != expected:
        raise AssertionError(f"{message}: expected {expected!r}, got {actual!r}")


def exercise_pipeline(updated_count: int) -> tuple[list[str], dict, dict]:
    events: list[str] = []
    original = {
        name: getattr(fmt, name, None)
        for name in (
            "render_pdf_only",
            "render_pdf_pages_for_visual_qa",
            "render_for_visual_qa",
            "refresh_toc_page_results_from_pdf",
            "check_toc_page_results_against_pdf",
            "write_toc_page_validation_report",
            "validate_docx_output",
        )
    }
    with TemporaryDirectory() as temp_dir:
        root = Path(temp_dir)
        pdf = root / "probe.pdf"
        pdf.write_bytes(b"%PDF")

        def render_pdf_only(*_args, **_kwargs):
            events.append("pdf-probe")
            return {"pdf": str(pdf)}

        def render_pdf_pages(*_args, **_kwargs):
            events.append("png-from-probe")
            return {"pdf": str(pdf), "page_count": 1, "render_dir": str(root)}

        def render_final(*_args, **_kwargs):
            events.append("full-final-render")
            return {"pdf": str(pdf), "page_count": 1, "render_dir": str(root)}

        def refresh(*_args, **_kwargs):
            events.append("refresh")
            return {"updated_count": updated_count, "entry_count": 58, "changes": []}

        def validate_toc(*_args, **_kwargs):
            events.append("validate-toc")
            return {"status": "pass", "entry_count": 58, "mismatches": []}

        def write_validation(*_args, **_kwargs):
            events.append("write-validation")

        def validate_docx(*_args, **_kwargs):
            events.append("validate-docx")

        fmt.render_pdf_only = render_pdf_only
        fmt.render_pdf_pages_for_visual_qa = render_pdf_pages
        fmt.render_for_visual_qa = render_final
        fmt.refresh_toc_page_results_from_pdf = refresh
        fmt.check_toc_page_results_against_pdf = validate_toc
        fmt.write_toc_page_validation_report = write_validation
        fmt.validate_docx_output = validate_docx
        tracker = fmt.PerformanceTracker("apply", "review")
        output_report, toc_report = fmt.render_output_with_toc_validation(
            root / "output.docx",
            root / "visual-qa",
            Path("/tmp/soffice"),
            fmt.qa_policy("review", 1800, 2600),
            "review",
            expected_comments=0,
            comment_mode="none",
            tracker=tracker,
        )
        metrics = dict(tracker.metrics)
    for name, value in original.items():
        if value is None:
            delattr(fmt, name)
        else:
            setattr(fmt, name, value)
    return events, metrics, toc_report


def main() -> int:
    refreshed_events, refreshed_metrics, refreshed_report = exercise_pipeline(57)
    assert_equal(
        refreshed_events,
        ["pdf-probe", "refresh", "validate-docx", "full-final-render", "validate-toc", "write-validation"],
        "refreshed TOC performs only one final PNG QA pass",
    )
    assert_equal(refreshed_metrics["pdf_conversion_passes"], 2, "refreshed PDF pass count")
    assert_equal(refreshed_metrics["png_raster_passes"], 1, "refreshed PNG pass count")
    assert_equal(refreshed_report["refreshed_count"], 57, "refreshed TOC metrics")

    unchanged_events, unchanged_metrics, unchanged_report = exercise_pipeline(0)
    assert_equal(
        unchanged_events,
        ["pdf-probe", "refresh", "png-from-probe", "validate-toc", "write-validation"],
        "unchanged TOC reuses the probe PDF for PNG QA",
    )
    assert_equal(unchanged_metrics["pdf_conversion_passes"], 1, "unchanged PDF pass count")
    assert_equal(unchanged_metrics["png_raster_passes"], 1, "unchanged PNG pass count")
    assert_equal(unchanged_report["refreshed_count"], 0, "unchanged TOC metrics")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
