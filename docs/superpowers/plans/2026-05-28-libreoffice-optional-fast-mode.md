# LibreOffice Optional Fast Mode Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Allow users without LibreOffice to try the `.docx` fast workflow while preserving LibreOffice as the required dependency for legacy conversion and final rendered QA.

**Architecture:** Keep `review` and `strict` unchanged as rendered QA modes that require LibreOffice, Swift, and the PDF renderer. Add a startup decision that only permits missing LibreOffice when `--qa-level fast` is used and all supplied input files are already `.docx`.

**Tech Stack:** Python 3.12, python-docx, openpyxl, LibreOffice `soffice`, Swift/PDFKit renderer on macOS.

---

### Task 1: Dependency Policy

**Files:**
- Modify: `/Users/hezhengyu/Desktop/论文格式skill研究/lunwen-geshi-zhushou-github/skills/lunwen-geshi-zhushou/scripts/thesis_format_from_sample.py`
- Test: `/Users/hezhengyu/Desktop/论文格式skill研究/lunwen-geshi-zhushou-github/skills/lunwen-geshi-zhushou/scripts/test_efficiency_options.py`

- [x] Add helpers that decide whether a run requires LibreOffice based on QA level and input suffixes.
- [x] Make `fast` + `.docx` allowed without `soffice`.
- [x] Keep `.doc`, `.rtf`, `.odt`, `review`, and `strict` blocked when `soffice` is missing.
- [x] Record fast reports with `"soffice": null` when LibreOffice is unavailable.

### Task 2: Skill And README Guidance

**Files:**
- Modify: `/Users/hezhengyu/Desktop/论文格式skill研究/lunwen-geshi-zhushou-github/README.md`
- Modify: `/Users/hezhengyu/Desktop/论文格式skill研究/lunwen-geshi-zhushou-github/skills/lunwen-geshi-zhushou/SKILL.md`
- Modify: `/Users/hezhengyu/Desktop/论文格式skill研究/lunwen-geshi-zhushou-github/.codex-plugin/plugin.json`

- [x] Document three modes: `fast` trial, `review`, and `strict` final delivery.
- [x] Warn that `fast` skips visual QA and is not final delivery.
- [x] Explain that legacy `.doc` users should install LibreOffice or convert to `.docx` first.

### Task 3: Verification And Publish

**Files:**
- Verify all modified files.

- [x] Run JSON validation for `.codex-plugin/plugin.json`.
- [x] Run the script test suite.
- [ ] Commit the change.
- [ ] Push to `origin/main` using the local HTTP proxy if needed.
