#!/usr/bin/env python3
"""Style-first paragraph formatting helpers for thesis_format_from_sample."""

from __future__ import annotations

import re
from typing import Any

from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor


ROLE_STYLE_NAMES: dict[str, str] = {
    "front_heading": "ThesisFrontHeading",
    "abstract_heading": "ThesisAbstractHeading",
    "contents_heading": "ThesisContentsHeading",
    "toc1": "ThesisTOC1",
    "toc2": "ThesisTOC2",
    "toc3": "ThesisTOC3",
    "chapter": "ThesisChapter",
    "second": "ThesisSection",
    "third": "ThesisSubsection",
    "body": "ThesisBody",
    "keywords": "ThesisKeywords",
    "figure_caption": "ThesisFigureCaption",
    "figure_note": "ThesisFigureNote",
    "table_caption": "ThesisTableCaption",
    "table_note": "ThesisTableNote",
    "table_footnote": "ThesisTableFootnote",
    "reference_heading": "ThesisReferenceHeading",
    "reference_entry": "ThesisReferenceEntry",
    "ack_heading": "ThesisAcknowledgementsHeading",
    "ack_body": "ThesisAcknowledgementsBody",
}

STYLEABLE_ROLE_SET = set(ROLE_STYLE_NAMES)

ALIGNMENTS = {
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
    "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
}

TAB_LEADERS = {
    "none": WD_TAB_LEADER.SPACES,
    "spaces": WD_TAB_LEADER.SPACES,
    "dots": WD_TAB_LEADER.DOTS,
    "hyphens": WD_TAB_LEADER.DASHES,
    "underscore": WD_TAB_LEADER.HEAVY,
}


def style_name_for_role(role: str) -> str | None:
    """Return the generated paragraph style name for a role."""
    return ROLE_STYLE_NAMES.get(role)


def styleable_roles(enabled_roles: set[str] | list[str] | tuple[str, ...]) -> set[str]:
    """Return enabled roles that are safe to route through paragraph styles."""
    return set(enabled_roles).intersection(STYLEABLE_ROLE_SET)


def _get_or_add_child(parent, tag_name: str):
    child = parent.find(qn(tag_name))
    if child is None:
        child = OxmlElement(tag_name)
        parent.append(child)
    return child


def _set_east_asia_font(style, latin: str, east_asia: str) -> None:
    style.font.name = latin
    r_pr = style.element.get_or_add_rPr()
    r_fonts = r_pr.rFonts
    if r_fonts is None:
        r_fonts = OxmlElement("w:rFonts")
        r_pr.append(r_fonts)
    for attr, value in {
        "w:ascii": latin,
        "w:hAnsi": latin,
        "w:cs": latin,
        "w:eastAsia": east_asia,
    }.items():
        r_fonts.set(qn(attr), value)


def _set_font_color(style, color_hex: str | None) -> None:
    value = (color_hex or "").strip().upper()
    if not value:
        return
    if not value.startswith("#") and re.fullmatch(r"[0-9A-F]{6}", value):
        value = f"#{value}"
    if re.fullmatch(r"#[0-9A-F]{6}", value):
        style.font.color.rgb = RGBColor.from_string(value[1:])


def _to_alignment(name: str | None):
    return ALIGNMENTS.get((name or "justify").strip().lower(), WD_ALIGN_PARAGRAPH.JUSTIFY)


def _to_tab_leader(name: str | None):
    return TAB_LEADERS.get((name or "dots").strip().lower(), WD_TAB_LEADER.DOTS)


def _set_font_properties(style, fmt: dict[str, Any], profile: dict[str, Any]) -> None:
    latin = fmt.get("font") or profile.get("latin_font") or "Times New Roman"
    east_asia = fmt.get("east_asia_font") or profile.get("east_asia_font") or "宋体"
    _set_east_asia_font(style, latin, east_asia)
    _set_font_color(style, fmt.get("font_color_hex") or "#000000")
    if fmt.get("size_pt") is not None:
        style.font.size = Pt(float(fmt["size_pt"]))
    for key, attr in (
        ("bold", "bold"),
        ("italic", "italic"),
        ("all_caps", "all_caps"),
        ("small_caps", "small_caps"),
    ):
        if fmt.get(key) is not None:
            setattr(style.font, attr, bool(fmt[key]))


def _set_paragraph_properties(style, fmt: dict[str, Any]) -> None:
    paragraph_format = style.paragraph_format
    paragraph_format.alignment = _to_alignment(fmt.get("align"))
    paragraph_format.line_spacing = float(fmt.get("line_spacing") or 1.5)
    paragraph_format.space_before = Pt(float(fmt.get("before_pt") or 0))
    paragraph_format.space_after = Pt(float(fmt.get("after_pt") or 0))
    hanging_pt = float(fmt.get("hanging_pt") or 0)
    first_line_pt = float(fmt.get("first_line_pt") or 0)
    paragraph_format.first_line_indent = Pt(-hanging_pt) if hanging_pt > 0 else Pt(first_line_pt)
    paragraph_format.left_indent = Pt(float(fmt.get("left_pt") or 0))
    paragraph_format.right_indent = Pt(float(fmt.get("right_pt") or 0))
    if fmt.get("toc_page_number_right_aligned") and fmt.get("toc_right_tab_pt"):
        paragraph_format.tab_stops.clear_all()
        paragraph_format.tab_stops.add_tab_stop(
            Pt(float(fmt["toc_right_tab_pt"])),
            alignment=WD_TAB_ALIGNMENT.RIGHT,
            leader=_to_tab_leader(fmt.get("toc_leader", "dots")),
        )


def _ensure_paragraph_style(doc, name: str):
    try:
        return doc.styles[name]
    except KeyError:
        style = doc.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
        try:
            style.base_style = doc.styles["Normal"]
        except KeyError:
            pass
        return style


def ensure_role_styles(doc, profile: dict[str, Any], enabled_roles: set[str] | list[str]) -> dict[str, Any]:
    """Create or update custom paragraph styles for enabled thesis roles."""
    roles = profile.get("roles", {})
    configured: dict[str, Any] = {}
    for role in sorted(styleable_roles(enabled_roles)):
        fmt = roles.get(role)
        style_name = style_name_for_role(role)
        if not fmt or not style_name:
            continue
        style = _ensure_paragraph_style(doc, style_name)
        _set_font_properties(style, fmt, profile)
        _set_paragraph_properties(style, fmt)
        configured[role] = style
    return configured


def apply_role_style(paragraph, role: str, doc, profile: dict[str, Any], enabled_roles: set[str] | list[str]) -> bool:
    """Apply a generated thesis style to a paragraph when the role supports it."""
    if role not in styleable_roles(enabled_roles):
        return False
    style_name = style_name_for_role(role)
    if not style_name:
        return False
    try:
        paragraph.style = doc.styles[style_name]
        return True
    except KeyError:
        styles = ensure_role_styles(doc, profile, enabled_roles)
        if role not in styles:
            return False
        paragraph.style = styles[role]
        return True
