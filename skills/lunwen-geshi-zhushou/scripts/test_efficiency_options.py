#!/usr/bin/env python3
"""Small regression checks for faster QA and quiet/profile-reuse options."""

from __future__ import annotations

import json
import sys
from pathlib import Path
from tempfile import TemporaryDirectory

import thesis_format_from_sample as fmt


def assert_equal(actual, expected, message: str) -> None:
    if actual != expected:
        raise AssertionError(f"{message}: expected {expected!r}, got {actual!r}")


def assert_true(value, message: str) -> None:
    if not value:
        raise AssertionError(message)


def parse_with(argv: list[str]):
    old_argv = sys.argv
    try:
        sys.argv = ["thesis_format_from_sample.py", *argv]
        return fmt.parse_args()
    finally:
        sys.argv = old_argv


def main() -> int:
    strict = fmt.qa_policy("strict", 1800, 2600)
    review = fmt.qa_policy("review", 1800, 2600)
    fast = fmt.qa_policy("fast", 1800, 2600)

    assert_true(strict["render_enabled"], "strict QA renders pages")
    assert_true(strict["edge_overflow_is_error"], "strict QA keeps edge overflow as hard failure")
    assert_equal(strict["render_width"], 1800, "strict width")
    assert_equal(strict["render_height"], 2600, "strict height")
    assert_true(review["render_enabled"], "review QA renders pages")
    assert_true(not review["edge_overflow_is_error"], "review QA reports edge overflow as risk")
    assert_true(review["render_width"] < strict["render_width"], "review uses lighter width")
    assert_true(fast["render_enabled"] is False, "fast QA skips visual rendering")

    pages = [{"page_index": index, "blank_like": index == 5, "edge_overflow": index == 7} for index in range(1, 22)]
    selected = fmt.select_review_page_indices(pages, max_pages=8)
    assert_equal(selected[:3], [1, 2, 3], "review selection starts with first pages")
    assert_true(5 in selected, "review selection includes blank-like risk page")
    assert_true(7 in selected, "review selection includes edge-overflow risk page")
    assert_true(21 in selected, "review selection includes final page")
    assert_true(len(selected) <= 8, "review selection respects max pages")

    args = parse_with([
        "--format-table",
        "confirmed.xlsx",
        "--target",
        "target.docx",
        "--reuse-profile",
        "profile.json",
        "--apply-engine",
        "direct",
        "--qa-level",
        "review",
        "--quiet",
    ])
    assert_equal(args.qa_level, "review", "qa-level parser")
    assert_true(args.quiet, "quiet parser")
    assert_equal(args.reuse_profile, Path("profile.json"), "reuse-profile parser")
    assert_equal(args.apply_engine, "direct", "apply-engine parser")
    assert_true(args.shared_sample_cache_dir is not None, "cross-run sample cache defaults on")

    cache_disabled_args = parse_with([
        "--sample",
        "sample.docx",
        "--analyze-only",
        "--no-shared-sample-cache",
    ])
    assert_true(cache_disabled_args.no_shared_sample_cache, "cross-run sample cache can be disabled")

    with TemporaryDirectory() as tmpdir:
        tracker = fmt.PerformanceTracker("apply", "review")
        with tracker.measure("output_pagination_pdf"):
            pass
        tracker.increment("pdf_conversion_passes")
        tracker.increment("png_raster_passes")
        tracker.set_metric("toc_refreshed_entries", 57)
        report_path = tracker.write(Path(tmpdir))
        payload = json.loads(report_path.read_text(encoding="utf-8"))
        assert_equal(report_path.name, "performance-apply.json", "performance artifact name")
        assert_equal(payload["operation"], "apply", "performance operation")
        assert_equal(payload["qa_level"], "review", "performance QA level")
        assert_true(payload["elapsed_seconds"] >= 0, "performance elapsed time")
        assert_true(payload["phases"]["output_pagination_pdf"] >= 0, "performance phase time")
        assert_equal(payload["metrics"]["pdf_conversion_passes"], 1, "performance PDF pass count")
        assert_equal(payload["metrics"]["png_raster_passes"], 1, "performance PNG pass count")
        assert_equal(payload["metrics"]["toc_refreshed_entries"], 57, "performance TOC count")

    with TemporaryDirectory() as tmpdir:
        root = Path(tmpdir)
        sample = root / "sample.docx"
        sample_docx = root / "sample.normalized.docx"
        profile_json = root / "sample.format-profile.json"
        render_dir = root / "visual-qa" / "sample_sample"
        shared_root = root / "shared-cache"
        sample.write_text("sample-input", encoding="utf-8")
        sample_docx.write_text("normalized-input", encoding="utf-8")
        profile_json.write_text(json.dumps({"roles": {"body": {}}, "enabled_roles": ["body"]}), encoding="utf-8")
        render_dir.mkdir(parents=True)
        (render_dir / "visual-report.json").write_text(
            json.dumps({"render_dir": str(render_dir), "page_count": 3}),
            encoding="utf-8",
        )
        fmt.publish_shared_sample_analysis(
            sample,
            shared_root,
            sample_docx,
            profile_json,
            render_dir,
            1200,
            1700,
            False,
            "review",
        )
        cached = fmt.load_shared_sample_analysis(sample, shared_root, 1200, 1700, False, "review")
        assert_true(cached is not None, "published cross-run sample cache can be loaded")
        assert_equal(cached["profile"]["enabled_roles"], ["body"], "shared cache preserves profile")
        assert_equal(cached["sample_report"]["page_count"], 3, "shared cache preserves visual evidence")
        assert_true(
            str(shared_root.resolve()) in cached["sample_report"]["render_dir"],
            "shared visual report points to reusable cached evidence",
        )
        assert_true(not list(shared_root.rglob("*.xlsx")), "shared cache never stores editable confirmation table")

    cleanup_args = parse_with([
        "--format-table",
        "confirmed.xlsx",
        "--target",
        "target.docx",
        "--cleanup-after-delivery",
    ])
    assert_true(cleanup_args.cleanup_after_delivery, "cleanup is explicitly selectable after delivery")

    with TemporaryDirectory() as tmpdir:
        root = Path(tmpdir)
        original = root / "target.doc"
        output = root / "target_formatted.docx"
        qa_dir = root / "visual-qa"
        sample_qa_dir = root / "sample-visual-qa"
        analysis = root / "analysis"
        table = root / "confirmed.xlsx"
        normalized = analysis / "target.normalized.docx"
        converted = analysis / "target.converted.docx"
        profile = analysis / "sample.format-profile.json"
        lint = analysis / "profile-lint-report.txt"
        original.write_text("original", encoding="utf-8")
        output.write_text("deliverable", encoding="utf-8")
        qa_dir.mkdir()
        (qa_dir / "visual-report.txt").write_text("qa", encoding="utf-8")
        sample_qa_dir.mkdir()
        (sample_qa_dir / "sample-page-1.png").write_text("generated", encoding="utf-8")
        analysis.mkdir()
        for path in (table, normalized, converted, profile, lint):
            path.write_text("generated", encoding="utf-8")
        removed = fmt.cleanup_after_successful_delivery(
            candidates=[table, normalized, converted, profile, lint, sample_qa_dir],
            protected=[original, output, qa_dir],
        )
        assert_true(table.resolve() in removed, "confirmed format table is removed after successful final handoff")
        assert_true(all(not path.exists() for path in (table, normalized, converted, profile, lint)), "generated working files are removed")
        assert_true(not sample_qa_dir.exists(), "generated sample visual QA is removed after final output QA exists")
        assert_true(original.exists(), "source thesis remains untouched by cleanup")
        assert_true(output.exists(), "final DOCX remains after cleanup")
        assert_true((qa_dir / "visual-report.txt").exists(), "final QA evidence remains after cleanup")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
