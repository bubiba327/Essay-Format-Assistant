#!/usr/bin/env python3
"""Regression checks for TOC content-control validation."""

from __future__ import annotations

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

import thesis_format_from_sample as fmt


def make_toc_paragraph(style_id: str, text: str):
    paragraph = OxmlElement("w:p")
    p_pr = OxmlElement("w:pPr")
    p_style = OxmlElement("w:pStyle")
    p_style.set(qn("w:val"), style_id)
    p_pr.append(p_style)
    paragraph.append(p_pr)
    run = OxmlElement("w:r")
    text_element = OxmlElement("w:t")
    text_element.text = text
    run.append(text_element)
    paragraph.append(run)
    return paragraph


def make_doc_with_toc(style_id: str, text: str) -> Document:
    doc = Document()
    sdt = OxmlElement("w:sdt")
    sdt_pr = OxmlElement("w:sdtPr")
    gallery = OxmlElement("w:docPartGallery")
    gallery.set(qn("w:val"), "Table of Contents")
    sdt_pr.append(gallery)
    content = OxmlElement("w:sdtContent")
    content.append(make_toc_paragraph(style_id, text))
    sdt.append(sdt_pr)
    sdt.append(content)
    doc._element.body.append(sdt)
    return doc


def make_doc_with_toc_entries(entries: list[tuple[str, str]]) -> Document:
    doc = Document()
    sdt = OxmlElement("w:sdt")
    sdt_pr = OxmlElement("w:sdtPr")
    gallery = OxmlElement("w:docPartGallery")
    gallery.set(qn("w:val"), "Table of Contents")
    sdt_pr.append(gallery)
    content = OxmlElement("w:sdtContent")
    for style_id, text in entries:
        content.append(make_toc_paragraph(style_id, text))
    sdt.append(sdt_pr)
    sdt.append(content)
    doc._element.body.append(sdt)
    return doc


def expect_runtime_error(fn, message: str) -> None:
    try:
        fn()
    except RuntimeError:
        return
    raise AssertionError(message)


def main() -> int:
    doc = make_doc_with_toc("TOC2", "1.1.1 Research Background6")
    fmt.validate_toc_content_controls(doc, {"toc1", "toc2"})

    doc = make_doc_with_toc("TOC2", "1.1.1 Research Background6")
    expect_runtime_error(
        lambda: fmt.validate_toc_content_controls(doc, {"toc1", "toc2", "toc3"}),
        "enabled toc3 entries should still be validated",
    )

    doc = make_doc_with_toc_entries(
        [
            ("TOC1", "Abstract2"),
            ("TOC1", "Chapter 1 Introduction6"),
            ("TOC1", "References42"),
        ]
    )
    rendered_pages = [
        "Cover",
        "Declaration",
        "Abstract",
        "CONTENTS Abstract 2 Chapter 1 Introduction 6 References 42",
        "Chapter 1 Introduction",
        "Body",
        "References",
    ]
    stale_report = fmt.check_toc_page_results_against_rendered_text(doc, rendered_pages)
    assert stale_report["status"] == "fail", "stale displayed TOC pages must fail validation"
    assert len(stale_report["mismatches"]) == 3, "every stale TOC entry must be reported"
    refreshed = fmt.refresh_toc_page_results_from_rendered_text(doc, rendered_pages)
    assert refreshed["updated_count"] == 3, "stale TOC display results are rewritten"
    refreshed_report = fmt.check_toc_page_results_against_rendered_text(doc, rendered_pages)
    assert refreshed_report["status"] == "pass", "rewritten TOC results match rendered headings"
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
