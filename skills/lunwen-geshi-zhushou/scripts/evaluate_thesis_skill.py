#!/usr/bin/env python3
"""Run repeatable real-document evaluations for the thesis formatting skill."""

from __future__ import annotations

import argparse
import json
import subprocess
import sys
from datetime import datetime
from pathlib import Path
from typing import Any

from report_thesis_eval import write_reports
from thesis_qa_checks import run_structural_checks


SCRIPT_DIR = Path(__file__).resolve().parent
SKILL_DIR = SCRIPT_DIR.parent
DEFAULT_MANIFEST = SKILL_DIR / "fixtures" / "manifest.json"
DEFAULT_BASELINE_DIR = SKILL_DIR / "fixtures" / "baselines" / "expected-qa"
RUNNER = SCRIPT_DIR / "run_thesis_format_from_sample.sh"


def load_cases(manifest_path: Path) -> dict[str, dict[str, Any]]:
    data = json.loads(manifest_path.read_text(encoding="utf-8"))
    if data.get("schema_version") != 1 or not isinstance(data.get("cases"), dict):
        raise ValueError(f"Unsupported evaluation manifest: {manifest_path}")
    return data["cases"]


def resolve_fixture_path(raw_path: str, manifest_path: Path) -> Path:
    path = Path(raw_path).expanduser()
    if not path.is_absolute():
        path = manifest_path.parent / path
    return path.resolve()


def resolve_case_paths(raw_case: dict[str, Any], manifest_path: Path) -> dict[str, Any]:
    case = dict(raw_case)
    case["sample"] = resolve_fixture_path(str(raw_case["sample"]), manifest_path)
    case["target"] = resolve_fixture_path(str(raw_case["target"]), manifest_path)
    return case


def case_layout(run_dir: Path, case_id: str, sample: Path, target: Path) -> dict[str, Path]:
    case_dir = run_dir / case_id
    analysis_dir = case_dir / "analysis"
    visual_qa_dir = case_dir / "visual-qa"
    output = case_dir / f"{target.stem}_评测输出.docx"
    return {
        "case_dir": case_dir,
        "analysis_dir": analysis_dir,
        "visual_qa_dir": visual_qa_dir,
        "format_table": case_dir / "格式确认表.xlsx",
        "profile_json": analysis_dir / f"{sample.stem}.format-profile.json",
        "lint_json": analysis_dir / "profile-lint-report.json",
        "analyze_performance_json": analysis_dir / "performance-analyze.json",
        "apply_performance_json": analysis_dir / "performance-apply.json",
        "source_docx": analysis_dir / f"{target.stem}.normalized.docx",
        "output": output,
        "output_risk_json": visual_qa_dir / f"{output.stem}_output" / "visual-risk-report.json",
        "output_visual_report_json": visual_qa_dir / f"{output.stem}_output" / "visual-report.json",
        "structural_report_json": case_dir / "structural-qa-report.json",
    }


def build_formatter_commands(
    case: dict[str, Any],
    layout: dict[str, Path],
    qa_level: str,
    shared_sample_cache_dir: Path | None = None,
    no_shared_sample_cache: bool = False,
) -> tuple[list[str], list[str]]:
    common = [
        "--analysis-dir",
        str(layout["analysis_dir"]),
        "--visual-qa-dir",
        str(layout["visual_qa_dir"]),
        "--qa-level",
        qa_level,
        "--quiet",
    ]
    if shared_sample_cache_dir is not None:
        common.extend(["--shared-sample-cache-dir", str(shared_sample_cache_dir)])
    if no_shared_sample_cache:
        common.append("--no-shared-sample-cache")
    analyze = [
        str(RUNNER),
        "--sample",
        str(case["sample"]),
        "--export-format-table",
        str(layout["format_table"]),
        *common,
        "--analyze-only",
    ]
    apply = [
        str(RUNNER),
        "--format-table",
        str(layout["format_table"]),
        "--reuse-profile",
        str(layout["profile_json"]),
        "--target",
        str(case["target"]),
        "--output",
        str(layout["output"]),
        *common,
        "--apply-engine",
        str(case.get("apply_engine", "direct")),
        "--comment-mode",
        str(case.get("comment_mode", "role")),
    ]
    return analyze, apply


def read_json(path: Path) -> dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8"))


def read_performance(layout: dict[str, Path]) -> dict[str, Any]:
    analyze = read_json(layout["analyze_performance_json"]) if layout["analyze_performance_json"].exists() else {}
    apply = read_json(layout["apply_performance_json"]) if layout["apply_performance_json"].exists() else {}
    analyze_metrics = analyze.get("metrics") or {}
    apply_metrics = apply.get("metrics") or {}
    stage1_seconds = float(analyze.get("elapsed_seconds", 0.0))
    stage2_seconds = float(apply.get("elapsed_seconds", 0.0))
    return {
        "stage1_seconds": stage1_seconds,
        "stage2_seconds": stage2_seconds,
        "total_seconds": round(stage1_seconds + stage2_seconds, 4),
        "pdf_conversion_passes": int(analyze_metrics.get("pdf_conversion_passes", 0))
        + int(apply_metrics.get("pdf_conversion_passes", 0)),
        "png_raster_passes": int(analyze_metrics.get("png_raster_passes", 0))
        + int(apply_metrics.get("png_raster_passes", 0)),
        "toc_refreshed_entries": int(apply_metrics.get("toc_refreshed_entries", 0)),
        "stage1_report": str(layout["analyze_performance_json"]) if analyze else None,
        "stage2_report": str(layout["apply_performance_json"]) if apply else None,
    }


def baseline_summary(result: dict[str, Any]) -> dict[str, Any]:
    metrics = (result.get("structural_report") or {}).get("metrics") or {}
    lint_codes = sorted(result.get("lint_codes", []))
    runtime_warnings = list(result.get("warnings", []))
    return {
        "schema_version": 1,
        "case_id": result["case_id"],
        "role_count": len(result.get("learned_roles", [])),
        "roles": sorted(result.get("learned_roles", [])),
        "warning_count": len(lint_codes) + len(runtime_warnings),
        "lint_codes": lint_codes,
        "table_count": int(metrics.get("table_count", 0)),
        "table_cell_count": int(metrics.get("table_cell_count", 0)),
        "toc_entry_count": int(metrics.get("toc_entry_count", 0)),
        "heading_count": int(metrics.get("heading_count", 0)),
        "figure_caption_count": int(metrics.get("figure_caption_count", 0)),
        "image_count": int(metrics.get("image_count", 0)),
        "equation_count": int(metrics.get("equation_count", 0)),
        "protected_paragraph_count": int(metrics.get("protected_paragraph_count", 0)),
        "protected_table_count": int(metrics.get("protected_table_count", 0)),
        "page_number_count": int(metrics.get("page_number_count", 0)),
    }


def update_baseline(path: Path, summary: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(summary, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def compare_baseline(actual: dict[str, Any], expected: dict[str, Any]) -> list[dict[str, Any]]:
    differences: list[dict[str, Any]] = []
    for field, expected_value in expected.items():
        if field == "schema_version":
            continue
        actual_value = actual.get(field)
        if actual_value == expected_value:
            continue
        differences.append({"field": field, "expected": expected_value, "actual": actual_value})
    return differences


def apply_baseline(
    result: dict[str, Any],
    baseline_dir: Path,
    update: bool = False,
) -> dict[str, Any]:
    path = baseline_dir / f"{result['case_id']}.json"
    summary = baseline_summary(result)
    result["baseline_path"] = str(path)
    result["baseline_summary"] = summary
    result["baseline_differences"] = []
    if update:
        if result["status"] == "pass":
            update_baseline(path, summary)
            result["baseline_status"] = "updated"
        else:
            result["baseline_status"] = "skipped-failed"
        return result
    if not path.exists():
        result["baseline_status"] = "missing"
        result.setdefault("warnings", []).append(f"No stored baseline: {path}")
        return result
    differences = compare_baseline(summary, read_json(path))
    result["baseline_differences"] = differences
    if differences:
        result["baseline_status"] = "changed"
        detail = "; ".join(
            f"{difference['field']} expected {difference['expected']!r}, got {difference['actual']!r}"
            for difference in differences
        )
        result.setdefault("failures", []).append("Baseline differs: " + detail)
        result["status"] = "fail"
    else:
        result["baseline_status"] = "matched"
    return result


def structural_issue_failure(issue: dict[str, Any]) -> str:
    if issue.get("code") == "captioned_table_cell_format_mismatch":
        detail = "; ".join(
            f"{difference['field']} {difference['actual_display']} != {difference['expected_display']}"
            for difference in issue.get("differences", [])
        )
        return (
            "Structural QA failed at "
            f"table {issue['table_index']}, row {issue['row']}, column {issue['column']}: {detail}"
        )
    text = f" at `{issue['text']}`" if issue.get("text") else ""
    return f"Structural QA failed: {issue.get('message') or issue.get('code', 'unknown issue')}{text}"


def assess_case(
    case_id: str,
    case: dict[str, Any],
    layout: dict[str, Path],
    qa_level: str,
    stages: list[subprocess.CompletedProcess[str]],
) -> dict[str, Any]:
    expected = case.get("expected", {})
    failures: list[str] = []
    warnings: list[str] = []
    for name, completed in zip(("analyze", "apply"), stages):
        if completed.returncode != 0:
            failures.append(f"{name} command exited {completed.returncode}: {completed.stderr.strip()[-400:]}")

    required_artifacts = [
        layout["profile_json"],
        layout["lint_json"],
        layout["output"],
        layout["output_risk_json"],
    ]
    for artifact in required_artifacts:
        if not artifact.exists():
            failures.append(f"Missing artifact: {artifact}")

    profile = read_json(layout["profile_json"]) if layout["profile_json"].exists() else {}
    learned_roles = list(profile.get("enabled_roles", []))
    required_roles = list(expected.get("required_roles", []))
    missing_roles = sorted(set(required_roles) - set(learned_roles))
    if missing_roles:
        failures.append("Missing required learned roles: " + ", ".join(missing_roles))
    expected_language = expected.get("language")
    language = (profile.get("document_language") or {}).get("dominant")
    if expected_language and language != expected_language:
        failures.append(f"Language profile mismatch: expected {expected_language}, got {language}")

    lint_payload = read_json(layout["lint_json"]) if layout["lint_json"].exists() else {"issues": []}
    lint_codes = [issue["code"] for issue in lint_payload.get("issues", [])]
    allowed_codes = set(expected.get("allowed_lint_codes", []))
    unexpected_lint_codes = sorted(set(lint_codes) - allowed_codes)
    if unexpected_lint_codes:
        failures.append("Unexpected profile lint warnings: " + ", ".join(unexpected_lint_codes))

    risk = read_json(layout["output_risk_json"]) if layout["output_risk_json"].exists() else {}
    visual_check = "skipped-fast" if qa_level == "fast" else "rendered"
    blank_like_pages = risk.get("blank_like_pages", [])
    edge_overflow_pages = risk.get("edge_overflow_pages", [])
    selected_review_pages = risk.get("selected_review_pages", [])
    contact_sheet = risk.get("contact_sheet")
    if visual_check == "rendered" and blank_like_pages:
        warnings.append(
            "Blank-like rendered pages require review: " + ", ".join(str(value) for value in blank_like_pages)
        )
    if edge_overflow_pages and qa_level == "strict":
        failures.append(
            "Output has edge-overflow risk pages: " + ", ".join(str(value) for value in edge_overflow_pages)
        )
    elif edge_overflow_pages and visual_check == "rendered":
        warnings.append(
            "Edge-overflow risk pages require strict verification: "
            + ", ".join(str(value) for value in edge_overflow_pages)
        )

    structural_report: dict[str, Any] = {"status": "not-run", "checks": {}, "issues": []}
    if layout["output"].exists() and layout["profile_json"].exists():
        try:
            output_visual_report = (
                read_json(layout["output_visual_report_json"])
                if layout["output_visual_report_json"].exists()
                else {}
            )
            rendered_pdf = output_visual_report.get("pdf")
            structural_report = run_structural_checks(
                layout["output"],
                profile,
                source_path=layout["source_docx"] if layout["source_docx"].exists() else None,
                rendered_pdf_path=Path(rendered_pdf) if rendered_pdf else None,
            )
        except Exception as exc:
            structural_report = {
                "status": "fail",
                "output": str(layout["output"]),
                "checks": {},
                "issues": [
                    {
                        "code": "structural_qa_execution_error",
                        "severity": "error",
                        "message": str(exc),
                    }
                ],
            }
        layout["structural_report_json"].write_text(
            json.dumps(structural_report, ensure_ascii=False, indent=2) + "\n",
            encoding="utf-8",
        )
    structural_issues = list(structural_report.get("issues", []))
    for issue in structural_issues:
        failures.append(structural_issue_failure(issue))
    performance = read_performance(layout)

    return {
        "case_id": case_id,
        "description": case.get("description", ""),
        "status": "pass" if not failures else "fail",
        "qa_level": qa_level,
        "sample": str(case["sample"]),
        "target": str(case["target"]),
        "output": str(layout["output"]),
        "learned_roles": learned_roles,
        "missing_roles": missing_roles,
        "lint_codes": lint_codes,
        "unexpected_lint_codes": unexpected_lint_codes,
        "visual_check": visual_check,
        "blank_like_pages": blank_like_pages,
        "edge_overflow_pages": edge_overflow_pages,
        "selected_review_pages": selected_review_pages,
        "contact_sheet": contact_sheet,
        "structural_status": structural_report.get("status", "not-run"),
        "structural_issues": structural_issues,
        "structural_report": structural_report,
        "performance": performance,
        "warnings": warnings,
        "failures": failures,
    }


def run_case(
    case_id: str,
    case: dict[str, Any],
    run_dir: Path,
    qa_level: str,
    shared_sample_cache_dir: Path | None = None,
    no_shared_sample_cache: bool = False,
) -> dict[str, Any]:
    layout = case_layout(run_dir, case_id, case["sample"], case["target"])
    layout["case_dir"].mkdir(parents=True, exist_ok=True)
    missing_inputs = [path for path in (case["sample"], case["target"]) if not path.exists()]
    if missing_inputs:
        return {
            "case_id": case_id,
            "description": case.get("description", ""),
            "status": "fail",
            "qa_level": qa_level,
            "sample": str(case["sample"]),
            "target": str(case["target"]),
            "output": str(layout["output"]),
            "learned_roles": [],
            "missing_roles": list(case.get("expected", {}).get("required_roles", [])),
            "lint_codes": [],
            "unexpected_lint_codes": [],
            "visual_check": "not-run",
            "blank_like_pages": [],
            "edge_overflow_pages": [],
            "selected_review_pages": [],
            "contact_sheet": None,
            "structural_status": "not-run",
            "structural_issues": [],
            "structural_report": {"status": "not-run", "checks": {}, "issues": []},
            "performance": {},
            "warnings": [],
            "failures": ["Missing fixture input: " + ", ".join(str(path) for path in missing_inputs)],
        }
    commands = build_formatter_commands(
        case,
        layout,
        qa_level,
        shared_sample_cache_dir=shared_sample_cache_dir,
        no_shared_sample_cache=no_shared_sample_cache,
    )
    stages: list[subprocess.CompletedProcess[str]] = []
    for command in commands:
        completed = subprocess.run(command, text=True, capture_output=True, check=False)
        stages.append(completed)
        if completed.returncode != 0:
            break
    return assess_case(case_id, case, layout, qa_level, stages)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--manifest", type=Path, default=DEFAULT_MANIFEST, help="Evaluation fixture manifest.")
    parser.add_argument("--case", default="all", help="Case id to run, or 'all'.")
    parser.add_argument("--qa-level", choices=("fast", "review", "strict"), default="fast")
    parser.add_argument("--output-dir", type=Path, help="Directory for this evaluation run.")
    parser.add_argument(
        "--shared-sample-cache-dir",
        type=Path,
        help="Isolated cross-run sample-cache directory passed through to formatter stages.",
    )
    parser.add_argument(
        "--no-shared-sample-cache",
        action="store_true",
        help="Disable cross-run sample-cache reuse and publication during evaluation.",
    )
    parser.add_argument(
        "--baseline-dir",
        type=Path,
        default=DEFAULT_BASELINE_DIR,
        help="Directory containing stable per-case expected QA summaries.",
    )
    parser.add_argument(
        "--update-baseline",
        action="store_true",
        help="Write stable summaries for passing cases instead of comparing existing baselines.",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    manifest = args.manifest.resolve()
    cases = load_cases(manifest)
    if args.case == "all":
        selected = list(cases.items())
    elif args.case in cases:
        selected = [(args.case, cases[args.case])]
    else:
        print(f"Unknown evaluation case: {args.case}", file=sys.stderr)
        return 2
    run_dir = (
        args.output_dir or SKILL_DIR / "evaluation-runs" / datetime.now().strftime("%Y%m%d-%H%M%S")
    ).resolve()
    results = []
    for case_id, raw_case in selected:
        result = run_case(
            case_id,
            resolve_case_paths(raw_case, manifest),
            run_dir,
            args.qa_level,
            shared_sample_cache_dir=args.shared_sample_cache_dir.resolve()
            if args.shared_sample_cache_dir
            else None,
            no_shared_sample_cache=args.no_shared_sample_cache,
        )
        results.append(apply_baseline(result, args.baseline_dir.resolve(), update=args.update_baseline))
    paths = write_reports(run_dir, results)
    print(f"Evaluation report: {paths['markdown']}")
    print(f"Evaluation JSON: {paths['json']}")
    print(f"Case summary: {paths['csv']}")
    return 0 if all(result["status"] == "pass" for result in results) else 1


if __name__ == "__main__":
    raise SystemExit(main())
