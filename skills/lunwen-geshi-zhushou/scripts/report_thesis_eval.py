#!/usr/bin/env python3
"""Write concise human- and machine-readable thesis evaluation reports."""

from __future__ import annotations

import csv
import json
from pathlib import Path
from typing import Any


def create_run_contact_sheet(output_dir: Path, results: list[dict[str, Any]]) -> Path | None:
    inputs = [
        (result["case_id"], Path(result["contact_sheet"]))
        for result in results
        if result.get("contact_sheet") and Path(result["contact_sheet"]).exists()
    ]
    if not inputs:
        return None
    try:
        from PIL import Image, ImageDraw
    except ImportError:
        return None
    max_width = 960
    gutter = 20
    label_height = 28
    images: list[tuple[str, Any]] = []
    for case_id, path in inputs:
        with Image.open(path) as image:
            rendered = image.convert("RGB")
            if rendered.width > max_width:
                ratio = max_width / rendered.width
                rendered = rendered.resize((max_width, max(1, int(rendered.height * ratio))))
            images.append((case_id, rendered.copy()))
    sheet_width = max(image.width for _, image in images) + gutter * 2
    sheet_height = gutter + sum(label_height + image.height + gutter for _, image in images)
    sheet = Image.new("RGB", (sheet_width, sheet_height), "white")
    draw = ImageDraw.Draw(sheet)
    y = gutter
    for case_id, image in images:
        draw.text((gutter, y), case_id, fill="black")
        y += label_height
        sheet.paste(image, (gutter, y))
        y += image.height + gutter
    path = output_dir / "contact-sheet.png"
    sheet.save(path)
    return path


def write_reports(output_dir: Path, results: list[dict[str, Any]]) -> dict[str, Path]:
    output_dir.mkdir(parents=True, exist_ok=True)
    overall_status = "pass" if results and all(item["status"] == "pass" for item in results) else "fail"
    json_path = output_dir / "evaluation-report.json"
    markdown_path = output_dir / "evaluation-report.md"
    csv_path = output_dir / "case-summary.csv"
    contact_sheet_path = create_run_contact_sheet(output_dir, results)

    json_path.write_text(
        json.dumps(
            {
                "status": overall_status,
                "contact_sheet": str(contact_sheet_path) if contact_sheet_path else None,
                "cases": results,
            },
            ensure_ascii=False,
            indent=2,
        )
        + "\n",
        encoding="utf-8",
    )
    with csv_path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(
            handle,
            fieldnames=[
                "case_id",
                "status",
                "qa_level",
                "missing_roles",
                "lint_codes",
                "unexpected_lint_codes",
                "structural_status",
                "structural_issue_count",
                "baseline_status",
                "baseline_change_count",
                "visual_check",
                "blank_like_pages",
                "edge_overflow_pages",
                "stage1_seconds",
                "stage2_seconds",
                "total_seconds",
                "pdf_conversion_passes",
                "png_raster_passes",
                "toc_refreshed_entries",
            ],
        )
        writer.writeheader()
        for result in results:
            writer.writerow(
                {
                    "case_id": result["case_id"],
                    "status": result["status"],
                    "qa_level": result["qa_level"],
                    "missing_roles": ", ".join(result.get("missing_roles", [])),
                    "lint_codes": ", ".join(result.get("lint_codes", [])),
                    "unexpected_lint_codes": ", ".join(result.get("unexpected_lint_codes", [])),
                    "structural_status": result.get("structural_status", "not-run"),
                    "structural_issue_count": len(result.get("structural_issues", [])),
                    "baseline_status": result.get("baseline_status", "not-checked"),
                    "baseline_change_count": len(result.get("baseline_differences", [])),
                    "visual_check": result.get("visual_check", ""),
                    "blank_like_pages": ", ".join(
                        str(value) for value in result.get("blank_like_pages", [])
                    ),
                    "edge_overflow_pages": ", ".join(
                        str(value) for value in result.get("edge_overflow_pages", [])
                    ),
                    "stage1_seconds": (result.get("performance") or {}).get("stage1_seconds", ""),
                    "stage2_seconds": (result.get("performance") or {}).get("stage2_seconds", ""),
                    "total_seconds": (result.get("performance") or {}).get("total_seconds", ""),
                    "pdf_conversion_passes": (result.get("performance") or {}).get("pdf_conversion_passes", ""),
                    "png_raster_passes": (result.get("performance") or {}).get("png_raster_passes", ""),
                    "toc_refreshed_entries": (result.get("performance") or {}).get("toc_refreshed_entries", ""),
                }
            )

    lines = ["# Thesis Skill Evaluation Report", "", f"Overall status: **{overall_status.upper()}**"]
    if contact_sheet_path:
        lines.append(f"Run contact sheet: `{contact_sheet_path}`")
    lines.append("")
    for result in results:
        visual_check = result.get("visual_check", "unknown")
        structural_status = result.get("structural_status", "not-run")
        structural_issues = result.get("structural_issues", [])
        performance = result.get("performance") or {}
        metrics = (result.get("structural_report") or {}).get("metrics") or {}
        captioned_table_check = (
            ((result.get("structural_report") or {}).get("checks") or {}).get("captioned_tables") or {}
        )
        if visual_check == "skipped-fast":
            blank_like_summary = "not checked"
            edge_overflow_summary = "not checked"
        else:
            blank_like_summary = ", ".join(str(value) for value in result.get("blank_like_pages", [])) or "none"
            edge_overflow_summary = ", ".join(
                str(value) for value in result.get("edge_overflow_pages", [])
            ) or "none"
        lines.extend(
            [
                f"## {result['case_id']}: {result['status'].upper()}",
                "",
                result.get("description", ""),
                "",
                f"- Sample: `{result['sample']}`",
                f"- Target: `{result['target']}`",
                f"- Output: `{result['output']}`",
                f"- QA level: `{result['qa_level']}`",
                f"- Learned roles checked: `{', '.join(result.get('learned_roles', [])) or 'none'}`",
                f"- Missing required roles: `{', '.join(result.get('missing_roles', [])) or 'none'}`",
                f"- Lint warnings: `{', '.join(result.get('lint_codes', [])) or 'none'}`",
                f"- Unexpected lint warnings: `{', '.join(result.get('unexpected_lint_codes', [])) or 'none'}`",
                f"- Structural QA: `{structural_status}` ({len(structural_issues)} issue(s))",
                (
                    "- Captioned tables checked: "
                    f"`{captioned_table_check.get('captioned_table_count', 0)} table(s), "
                    f"{captioned_table_check.get('checked_cell_count', 0)} cell(s)`"
                ),
                (
                    "- Structural counts: "
                    f"`TOC entries: {metrics.get('toc_entry_count', 0)}, "
                    f"headings: {metrics.get('heading_count', 0)}, "
                    f"figure captions/images: {metrics.get('figure_caption_count', 0)}/{metrics.get('image_count', 0)}, "
                    f"equations: {metrics.get('equation_count', 0)}, "
                    f"protected paragraphs: {metrics.get('protected_paragraph_count', 0)}, "
                    f"page numbers: {metrics.get('page_number_count', 0)}`"
                ),
                f"- Baseline: `{result.get('baseline_status', 'not-checked')}`",
                (
                    "- Runtime: "
                    f"Stage 1 `{performance.get('stage1_seconds', 0)}s`, "
                    f"Stage 2 `{performance.get('stage2_seconds', 0)}s`, "
                    f"Total runtime: `{performance.get('total_seconds', 0)}s`"
                ),
                (
                    "- Render work: "
                    f"PDF/PNG passes: `{performance.get('pdf_conversion_passes', 0)}/"
                    f"{performance.get('png_raster_passes', 0)}`; "
                    f"TOC refreshed entries: `{performance.get('toc_refreshed_entries', 0)}`"
                ),
                f"- Visual QA execution: `{visual_check}`",
                f"- Blank-like pages: `{blank_like_summary}`",
                f"- Edge-overflow pages: `{edge_overflow_summary}`",
                (
                    "- Manual review pages: `"
                    f"{', '.join(str(value) for value in result.get('selected_review_pages', [])) or 'none'}`"
                ),
                f"- Contact sheet: `{result.get('contact_sheet') or 'not generated'}`",
            ]
        )
        for difference in result.get("baseline_differences", []):
            lines.append(
                "- Baseline change: "
                f"{difference['field']}: expected `{difference['expected']}`, actual `{difference['actual']}`"
            )
        for issue in structural_issues:
            if issue.get("code") == "captioned_table_cell_format_mismatch":
                detail = "; ".join(
                    f"{difference['field']} {difference['actual_display']} != {difference['expected_display']}"
                    for difference in issue.get("differences", [])
                )
                lines.append(
                    "- Structural mismatch: "
                    f"table {issue['table_index']}, row {issue['row']}, column {issue['column']}, "
                    f"`{issue.get('text', '')}` (`{issue.get('role', '')}`): {detail}"
                )
            else:
                lines.append(
                    f"- Structural issue: `{issue.get('code', 'unknown')}` "
                    f"{issue.get('message', '')}".rstrip()
                )
        for warning in result.get("warnings", []):
            lines.append(f"- Warning: {warning}")
        for failure in result.get("failures", []):
            if not failure.startswith("Structural QA failed"):
                lines.append(f"- Failure: {failure}")
        lines.append("")
    markdown_path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    paths = {"json": json_path, "markdown": markdown_path, "csv": csv_path}
    if contact_sheet_path:
        paths["contact_sheet"] = contact_sheet_path
    return paths
