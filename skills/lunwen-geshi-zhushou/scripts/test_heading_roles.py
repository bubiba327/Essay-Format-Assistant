#!/usr/bin/env python3
"""Regression checks for heading role recognition."""

from __future__ import annotations

import thesis_format_from_sample as fmt


def assert_equal(actual, expected, message: str) -> None:
    if actual != expected:
        raise AssertionError(f"{message}: expected {expected!r}, got {actual!r}")


def main() -> int:
    long_chapter = "Chapter 5 Empirical Analysis of the Impact of Population Aging on Export Technological Sophistication"
    assert_equal(fmt.role_for_target(long_chapter, "", False, False), "chapter", "long explicit English chapter")
    assert_equal(fmt.role_for_target("Chapter 3 Analysis of the Current Situation of Population Aging", "", False, False), "chapter", "explicit English chapter")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
