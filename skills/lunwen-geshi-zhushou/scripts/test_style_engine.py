#!/usr/bin/env python3
"""Small regression checks for style-first paragraph formatting."""

from __future__ import annotations

from docx import Document
from docx.oxml.ns import qn

import thesis_format_from_sample as fmt
import thesis_style_engine as style_engine


def assert_equal(actual, expected, message: str) -> None:
    if actual != expected:
        raise AssertionError(f"{message}: expected {expected!r}, got {actual!r}")


def assert_true(value, message: str) -> None:
    if not value:
        raise AssertionError(message)


def assert_close(actual: float | None, expected: float, message: str) -> None:
    if actual is None or abs(actual - expected) > 0.01:
        raise AssertionError(f"{message}: expected {expected!r}, got {actual!r}")


def main() -> int:
    profile = fmt.default_profile()
    body = profile["roles"]["body"]
    body.update(
        {
            "font": "Times New Roman",
            "east_asia_font": "宋体",
            "font_color_hex": "#112233",
            "size_pt": 12,
            "bold": False,
            "italic": False,
            "align": "center",
            "line_spacing": 1.25,
            "before_pt": 6,
            "after_pt": 3,
            "first_line_pt": 21,
            "left_pt": 7,
            "right_pt": 9,
        }
    )

    doc = Document()
    styles = style_engine.ensure_role_styles(doc, profile, {"body"})
    assert_true("body" in styles, "body role style is created")
    assert_equal(styles["body"].name, "ThesisBody", "body style name")
    assert_equal(styles["body"].font.name, "Times New Roman", "latin font")
    assert_equal(str(styles["body"].font.color.rgb), "112233", "font color")
    assert_close(fmt.pt(styles["body"].font.size), 12.0, "font size")
    assert_equal(styles["body"].font.bold, False, "bold toggle")
    assert_equal(styles["body"].font.italic, False, "italic toggle")

    r_fonts = styles["body"].element.rPr.rFonts
    assert_equal(r_fonts.get(qn("w:eastAsia")), "宋体", "east Asian font")
    assert_equal(fmt.align_name(styles["body"].paragraph_format.alignment), "center", "style alignment")
    assert_close(float(styles["body"].paragraph_format.line_spacing), 1.25, "style line spacing")
    assert_close(fmt.pt(styles["body"].paragraph_format.space_before), 6.0, "before spacing")
    assert_close(fmt.pt(styles["body"].paragraph_format.space_after), 3.0, "after spacing")
    assert_close(fmt.pt(styles["body"].paragraph_format.first_line_indent), 21.0, "first-line indent")
    assert_close(fmt.pt(styles["body"].paragraph_format.left_indent), 7.0, "left indent")
    assert_close(fmt.pt(styles["body"].paragraph_format.right_indent), 9.0, "right indent")

    paragraph = doc.add_paragraph("This is body text.")
    applied = style_engine.apply_role_style(paragraph, "body", doc, profile, {"body"})
    assert_true(applied, "body style was applied")
    assert_equal(paragraph.style.name, "ThesisBody", "paragraph uses thesis body style")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
