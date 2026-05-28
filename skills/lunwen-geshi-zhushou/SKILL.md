---
name: lunwen-geshi-zhushou
description: Use when Codex needs to infer formatting from a thesis/sample DOC or DOCX, produce an editable XLSX confirmation table, and later apply the user-confirmed table to another thesis, dissertation, paper, or academic Word document.
---

# 论文格式助手

## Core Rule

Never edit the user's original thesis in place. First analyze only the format sample, then produce an editable `.xlsx` format confirmation table. Wait for the user to review/edit/confirm that table. Only after confirmation should the user provide the target thesis; then apply the confirmed table to a named copy. Add Word comments beside changed paragraphs/tables that state the original format and the adjusted format.

Use the Documents plugin workflow for DOC/DOCX work. Always load workspace dependencies first and run the bundled Codex Python directly through `scripts/run_thesis_format_from_sample.sh` (or the exact bundled Python path from `load_workspace_dependencies`). Do not probe system Python first and do not fall back through multiple local runtimes.

LibreOffice is the quality gate dependency, not a casual optional extra. Use `/Applications/LibreOffice.app/Contents/MacOS/soffice` on macOS when present, or `SOFFICE` if the user has configured a different path. `review` and `strict` runs require LibreOffice because the script renders PDF/page PNGs, validates TOC displayed pages, and writes visual QA artifacts. Legacy `.doc`/`.rtf`/`.odt` inputs also require LibreOffice for conversion. If LibreOffice is unavailable, you may use `--qa-level fast` only for `.docx`-only trial/smoke runs; clearly tell the user this skips rendering and is not final delivery. For a final handoff, install LibreOffice or run on a machine with LibreOffice and rerun `--qa-level strict`. On macOS, the script uses Swift/PDFKit to render pages, avoiding Poppler.

## Workflow

1. Locate the sample format file. Do not ask for the target thesis yet.
2. Pass legacy `.doc`/`.rtf`/`.odt` files directly to the script. Do not manually pre-convert with `textutil` or LibreOffice. The script creates analysis copies, uses LibreOffice for the main conversion, and only uses `textutil` internally as a narrow fallback to recover missing TOC/table-caption roles from old `.doc` files.
3. Extract the sample format and generate an editable `.xlsx` confirmation table:
   - page size/margins if available
   - header distance, footer distance, odd/even header-footer setting, first-page-different setting, page-number location, numbering format, and restart/continue section numbering
   - abstract/contents headings
   - TOC level 1/2/3
   - chapter, section, subsection headings
   - body paragraphs
   - figure captions, table captions, notes, equations, equation numbers, plain-text fallback equations, references, acknowledgements
   - table title/caption, table header/column-label text, table body variable-name column, table body definition/description column, table body numeric/value cells, table footnotes/subnotes, and table overall settings
   - image layout profile for automatic application
   - document language profile based on Chinese character and Latin-letter counts; use this as diagnostic context, not as two isolated Chinese/English pipelines
   - cache repeated sample analysis with `run-manifest.json` when the sample file, script version, render size, and relevant options have not changed; on a cache hit, reuse the saved `format-profile.json` and existing sample visual QA report instead of re-converting/re-rendering the sample
   - reuse unchanged sample assets across separate work runs through the default shared cache at `~/.codex/cache/lunwen-geshi-zhushou/sample-analysis`; this cache stores the learned profile, normalized sample, and sample visual-QA evidence only, never the user's editable confirmation table. Use `--no-shared-sample-cache` when diagnosing extraction behavior or when the user specifically does not want reusable template assets retained.
   - use `--qa-level fast` for code/debug smoke checks, `--qa-level review` for normal working runs with lighter rendered contact-sheet QA, and `--qa-level strict` for final delivery
4. The `.xlsx` table must include only roles actually observed in the sample. If the sample has no TOC, do not include TOC rows. If the sample has no tables, do not include table text rows.
5. Every text row in the `.xlsx` confirmation table must expose the complete editable format set: western font, Chinese font, font color name, font color value, size, bold, italic, all caps, small caps, alignment, line spacing, before/after spacing, heading blank lines before/after when observed, first-line indent, left indent, hanging indent, right indent, and role-specific notes.
6. Do not use technical wording like `继承` in the user-facing table. Use clear values such as `未特别设置` for bold/italic, `0` for no spacing or no indent, `自动（通常黑色）` plus `#000000` for automatic black text.
7. TOC rows need extra editable fields: page number right aligned, dot leader, and right tab position. If a converted legacy `.doc` hides tab stops but TOC entries contain tabs/page references, infer a right-aligned dotted tab at the usable page width and expose it in the table for user correction.
   - Do not import false italic from legacy TOC field-code conversion when the rendered TOC is upright. Prefer explicit all-caps/small-caps fields over changing the visible text.
   - During application, if the paper title page or `Chapter 1` starts immediately after TOC entries without a page break, insert a page break so the body does not begin on the last contents page.
   - Recognize Chinese contents headings such as `正文目录`, and flat Chinese/legacy TOC entries that use tabs, dot leaders, or long dash leaders before the page number. Do not classify those TOC entries as body headings.
   - After formatting changes pagination, first generate an output PDF pagination probe to locate each TOC target heading and refresh stale displayed page results in the copied DOCX. Generate PNG visual-QA pages only for the corrected final DOCX; if no TOC value changes, reuse the probe PDF for that final PNG pass. Fail final delivery if any displayed TOC page still differs from the final rendered heading page.
8. Table rows need both text and table-layout coverage:
   - include `表格标题`, `表头（列标题）`, `表格正文-变量名列`, `表格正文-定义/说明列`, `表格正文-数值列`, `表尾标注`, and `表格整体设置` when observed
   - table text rows use the same full font/color/spacing/indent fields as normal text
   - do not expose `表格文字` as a normal user-facing row; it is only a backward-compatible fallback for old confirmation tables
   - `表格整体设置` exposes table horizontal position, text wrapping, layout mode, total width, and cell margins
   - recognize Chinese table captions such as `表1 标题` and `表3-1 标题` as table captions, while excluding narrative references such as `表1显示...`
   - when the sample contains captioned tables, learn table layout/text formats primarily from those captioned body tables instead of cover/template tables
   - normalize legacy `.doc` table XML that reports visually centered thesis tables as left-aligned; default table overall alignment to centered, prefer near-full percentage width when the sample shows full-width tables, clear leftover table indentation, and do not keep target fixed row heights unless explicitly learned from the sample
   - if legacy `.doc` conversion flattens a table into paragraphs after a `Table x-x` caption, infer table header/body roles from those paragraphs and expose them for correction using the same fine-grained table-role names where possible
   - classify repeated table subnote lines such as `t statistics in parentheses`, `Note: ...`, and `* p < 0.1, ** p < 0.05, *** p < 0.01` as table footnotes, not body text
   - heading distribution matters: learn and apply blank paragraphs around heading roles. For common thesis defaults, chapter headings may have two blank lines before, while section/subsection headings normally have no blank paragraph before or after.
9. Image formatting is automatically inferred and applied, but not exposed in the `.xlsx` table. Apply safe image layout properties such as image paragraph alignment, common image width, and inline embedding when the sample image is inline. Preserve complex wrapping when direct conversion would be unsafe. Figure/table names and notes around images are text roles and must be exposed in the `.xlsx`.
   - recognize Chinese figure captions such as `图1 标题` and `图2-1 标题` as figure captions, while excluding narrative references such as `图2显示...`
   - when the sample contains images near figure captions, learn image layout/size primarily from caption-nearby figures instead of decorative or template images
   - If legacy `.doc` conversion puts a figure image and its `Figure x-x ...` caption text in the same paragraph with the caption text before the image, split it into an image paragraph followed by a caption paragraph before applying image and caption formatting.
   - If legacy `.doc` conversion produces an adjacent `Figure x-x ...` caption paragraph immediately before an image paragraph, move the caption paragraph below the image before rendering and visual QA.
   - If legacy `.doc` conversion wraps an image-only figure in a single-cell table, flatten that image-only table into a normal inline image paragraph before caption ordering, sizing, centering, and visual QA.
   - Do not treat legacy `w:cs` font strings such as `宋体;SimSun` as the Western font for figure captions. If the sample lacks an explicit Latin font for a figure/table caption, keep `Times New Roman` for Western text and use `宋体` for Chinese text.
10. Equation formatting must protect math content before applying paragraph formatting:
   - recognize true Word/WPS equations stored as `m:oMath` or `m:oMathPara`, legacy `EQ`/`MERGEFORMAT` equation fields, standalone equation numbers such as `(1)`, `（1）`, and `(3-1)`, and plain-text fallback equations such as `share_ikt=X_ikt/X_it (2)` or `edu_it = Σ (...) (6)`
   - expose observed equation roles as `公式`, `公式编号`, and `普通文本公式` in the confirmation table
   - for true Word/WPS equation objects, change only paragraph-level layout such as alignment, spacing, line spacing, indentation, and tab stops; do not rewrite internal math XML
   - for plain-text fallback equations and equation numbers, remove inherited body first-line indentation and apply the learned equation paragraph layout
   - do not click-drive WPS or automatically convert every plain-text equation into an editable formula object
11. Header, footer, and page-number formatting must be inferred from the sample and exposed in the confirmation table:
   - include `页眉整体设置`, `页眉文字`, `页脚整体设置`, `页脚文字`, and `页码` rows when observed
   - expose header/footer distance, `首页不同`, `奇偶页不同`, `保留目标页码`, page-number location, page-number format, start value, and section numbering mode on the `页面设置` sheet
   - default behavior is to preserve the target thesis header/footer text and target page numbers; do not restyle, renumber, insert, normalize, or replace page-number fields unless the confirmed table explicitly sets `保留目标页码` to `否`
   - use the thesis default recommended route: front matter uses Roman numbering where the sample does, body sections use decimal numbering, and body numbering restarts at 1 when the sample has a restarted body section
12. Every profile/table generation writes `profile-lint-report.txt` and `profile-lint-report.json` beside the profile JSON. Read the lint report before asking the user to confirm the table. Treat warnings such as suspicious caption Latin fonts, non-centered table overall settings, conflicting table widths, missing table-header rows, mixed Roman/decimal page-number examples, or target page-number overwrite risk as issues to discuss or fix before applying formatting.
13. Ask the user to edit the `.xlsx` if needed and send back the confirmed table together with the target thesis. The confirmed table is the source of truth, even when it differs from the automatic sample inference.
14. Pass the target Word file directly to the script, including legacy `.doc`; it creates a target DOCX copy with a clear suffix, for example `_格式调整试改.docx`. When a previous `format-profile.json` is available, pass it with `--reuse-profile` together with the confirmed table so Stage 2 does not re-analyze the sample.
   - Stage 2 uses a style-first paragraph engine by default: it creates Word/WPS paragraph styles such as `ThesisBody`, `ThesisChapter`, and `ThesisTableCaption`, applies them by detected role, then keeps the direct-format fallback for compatibility with older target documents.
   - Use `--apply-engine direct` only for debugging or compatibility checks when you need the legacy direct-format-only route.
15. Before applying formatting, automatically detect and protect target-only front matter such as school cover pages, originality/诚信声明, copyright authorization, confidentiality choices, and handwritten signature/date pages. Use content keywords, section/page-break structure, and the academic start point (`摘 要`, `Abstract`, `正文目录`, or equivalent) together; never hard-code a fixed number of pages to skip. Protected front matter paragraphs, tables, images, sections, headers, footers, and page-number settings must be left unchanged.
16. Apply formatting to the copy only. Keep content, chapter order, figures, tables, and cover text intact unless the user asks otherwise.
17. Add Word comments for format changes:
   - default to `--comment-mode role`, which comments the first changed paragraph/table for each role such as body, headings, captions, references, and table text
   - use `--comment-mode all` when the user explicitly wants every changed paragraph/table cell annotated
   - use `--comment-mode none` only when the user asks for a clean copy without comments
   - comments must say `original format -> adjusted format` for changed properties such as font, font color, size, bold, italic, alignment, line spacing, paragraph spacing, indentation, and TOC tab/leader settings
18. Validate with the LibreOffice visual gate:
   - For `review` and `strict`, the script must find `soffice` before it starts.
   - For `fast`, `.docx`-only runs may continue without `soffice`; treat this as a trial/smoke run and not as a deliverable final QA pass.
   - Legacy `.doc`/`.rtf`/`.odt` inputs require `soffice` even in `fast`; if the user cannot install LibreOffice, ask them to save the file as `.docx` in Word/WPS first.
   - For sample analysis in `strict` or `review`, render the converted/read sample into `page-*.png` images and a PDF under `visual-qa/`.
   - For formatted output, validate the DOCX container, re-open with `python-docx`, confirm comments when comments were generated, then follow the selected QA level. In `review`/`strict`, use the PDF-only pagination probe before final PNG generation so a stale-TOC intermediate version is never rasterized merely to be discarded.
   - `fast`: skip LibreOffice rendering and write a visual report that says rendering was skipped; use only for code/debug smoke checks or no-LibreOffice `.docx` trials, not final delivery.
   - `review`: render at lighter dimensions, write `contact-sheet.png`, `visual-risk-report.txt/json`, and inspect the selected review/risk pages instead of opening every page.
   - `strict`: render full high-resolution pages and keep side-edge ink detection as a hard failure.
   - Open the rendered PNGs/contact sheet and inspect layout, TOC, tables, figures, captions, comments, spacing, overlap, blank pages, and page breaks according to the QA level.
   - Run structural QA for loose heading numbering such as `5.2. 3 ...`; if such a heading keeps body first-line indentation or wrong blank-line distribution, fail the run and fix role recognition/application before delivery.
   - Treat a TOC displayed-page mismatch as a delivery error. In `review`/`strict`, the script must write `toc-page-validation.txt/json` after rendering; a formatted DOCX cannot pass final delivery while a TOC entry points to a stale page.
   - Treat automatic side-edge ink detection as a hard failure: if a rendered body page has text/table/image ink touching the left or right page edge, rerun after fixing table width, table text wrapping, image layout, or paragraph formatting.
   - If visual QA finds any problem, do not deliver the DOCX. Identify whether the cause is the confirmed format table, role recognition, formatting application, comments, table/image handling, or rendering setup; fix the smallest responsible part; rerun the formatting step; render again; inspect again.
   - Repeat the edit -> render -> inspect loop until the rendered PNGs pass visual QA.
   - Do not deliver a formatted DOCX until the final rendered PNGs have been visually checked and accepted. If rendering fails, fix rendering or formatting first.
19. For a completed final handoff, remove generated working artifacts only after `--qa-level strict` has passed:
   - pass `--cleanup-after-delivery` when the user no longer needs to edit/reuse the confirmation table
   - this removes the confirmed `.xlsx`, generated analysis/conversion working files, and manifest-identified sample visual-QA working directory used before final delivery
   - never remove the user's original sample or target file, the final output DOCX, or its final visual-QA directory/reports
   - omit the flag when the user wants to apply the same confirmed table to additional theses

## Script

Stage 1: analyze the sample and create the editable confirmation table:

```bash
<skill-dir>/scripts/run_thesis_format_from_sample.sh \
  --sample 格式范文.doc \
  --analysis-dir .thesis_format_analysis \
  --export-format-table 格式确认表.xlsx \
  --visual-qa-dir .thesis_format_analysis/visual-qa \
  --qa-level review \
  --quiet \
  --analyze-only
```

Stage 2: after the user confirms/edits the `.xlsx`, apply that table to the target thesis:

```bash
<skill-dir>/scripts/run_thesis_format_from_sample.sh \
  --format-table 格式确认表.xlsx \
  --reuse-profile .thesis_format_analysis/格式范文.format-profile.json \
  --target 毕业论文修改.doc \
  --output 毕业论文修改_格式调整试改.docx \
  --analysis-dir .thesis_format_analysis \
  --visual-qa-dir .thesis_format_analysis/visual-qa \
  --qa-level strict \
  --quiet \
  --apply-engine styles \
  --comment-mode role \
  --cleanup-after-delivery
```

The script writes a JSON format profile, an editable `.xlsx` table, and mandatory visual QA artifacts. The table has `格式确认表` and `页面设置` sheets. Users can edit fonts, font color name/value, sizes, bold/italic/all-caps/small-caps, alignment, line spacing, paragraph spacing, heading blank-line counts, first-line/left/hanging/right indents, TOC right tabs/dot leaders, table position/text wrapping/layout/width/cell margins, page margins, header/footer distance, first-page/odd-even header-footer settings, whether target page numbers are preserved, page-number location/format/start/section restart mode, and `是否应用`. When applying, `--format-table` overrides automatic inference and controls which roles are formatted.

Stage 1 also writes `run-manifest.json` for cache validation, `profile-lint-report.txt/json` for high-risk profile warnings, and `performance-analyze.json` for elapsed-phase/render-pass measurements. A cache hit is allowed only when the sample hash, script hash, QA level, render dimensions, and relevant options match; otherwise rerun conversion/analysis/rendering. The lint report is part of the handoff and should be mentioned when it contains warnings.

Efficiency rule: do not repeat Stage 1 visual rendering during Stage 2 when a confirmed format table is already loaded. Prefer `--reuse-profile` with the previous `format-profile.json`; allow the shared sample cache to avoid re-converting/re-rendering a previously verified unchanged template in a later work run; analyze only what is needed for non-table fields such as image layout; and render the formatted output only. Stage 2 writes `performance-apply.json`; read this short summary before verbose page-level JSON when reporting runtime or investigating slowness. Use `--quiet` by default for routine runs; read the written reports for details instead of flooding the chat with role summaries.

Chinese/English rule: keep one unified role-recognition pipeline. The script records whether the sample is Chinese-dominant, English-dominant, or mixed, then applies both English and Chinese role heuristics where they match. This supports mixed theses with Chinese body text, English abstracts, English table heads, or bilingual captions.

Front-matter protection rule: during Stage 2, detect school cover/declaration/authorization material as protected front matter before formatting. The detector should protect by structure and content, not by fixed page count. When protection is active, skip formatting protected paragraphs/tables/images and skip section-level page/header/footer/page-number changes for protected sections; report the number of skipped paragraphs and sections in the run output.

The script preserves target page size by default and skips the first table because thesis cover forms are often tables. It accepts `.doc` or `.docx` targets directly. It adds format-change comments by default with `--comment-mode role`; use `--comment-mode all` for exhaustive paragraph/table annotations, or `--comment-mode none` for no comments. Use `--use-sample-page-size` or `--format-cover-table` only when the user asks or inspection shows they are needed.

Do not add manual post-processing for common caption variants. The script recognizes title captions such as `Figure 4-1. Title`, `Figure 3 - 3Title`, `Table 5-1: Title`, `Table 5 -6．Title`, `图2-1 标题`, `图1 标题`, `表3-1 标题`, and `表1 标题`, while leaving narrative references like `Figure 3-4 illustrates...`, `Table 5-1 presents...`, `图2显示...`, or `表1说明...` as body text.

Visual QA artifacts are written under `--visual-qa-dir` or `<analysis-dir>/visual-qa` by default. Rendered `review`/`strict` directories contain `page-*.png`, a PDF, `visual-report.json`, `visual-report.txt`, `visual-risk-report.txt/json`, and usually `contact-sheet.png`. Treat reports and contact sheets as inspection aids, not as substitutes for checking suspicious pages. A final delivery run should use `--qa-level strict`; a `fast` run is only a smoke check.

Final artifact cleanup is opt-in at the CLI and standard for a one-document finished handoff: pass `--cleanup-after-delivery` only with `--qa-level strict`. It runs after final TOC validation and visual QA succeed, deleting the used confirmation table, generated analysis/conversion work files, and manifest-identified per-run sample QA render directory while retaining original inputs, the final DOCX, final output QA evidence, and reusable shared sample-cache assets. When the user asks not to retain reusable template assets, run with `--no-shared-sample-cache` and remove any cache entry created in an earlier retained run only with explicit user approval.

## Role Heuristics

Read `references/role-mapping.md` when a document has mixed direct formatting, unreliable Word styles, or captions/explanatory paragraphs that are easy to confuse.

Default mapping:
- `Chapter 1 ...` / `第一章 ...` -> Heading 1
- `1.1 ...` -> Heading 2
- `1.1.1 ...` -> Heading 3
- `一、...` / `1 ...` -> Heading 1 when the text is title-like, not a long narrative paragraph
- `（一）...` -> Heading 2 when the text is title-like
- `1.理论背景` -> Heading 3 when the text is title-like
- `Figure 1-1 Title...`, `Figure 1-1: Title...`, `Figure 1 - 1Title...`, `Table 5-1: Title...`, `图2-1 标题`, `表1 标题` -> caption
- `Figure 3-1 shows...` or `Table 5-1 reports...` -> body paragraph, not caption
- `图2显示...` or `表1说明...` -> body paragraph, not caption
- true Word/WPS `m:oMath` / `m:oMathPara` paragraphs -> `公式`, preserving internal math XML
- standalone `(1)`, `（1）`, `(3-1)` -> `公式编号`
- `share_ikt=X_ikt/X_it (2)` or `edu_it = Σ (P_j,it × W_j) / P_it (6)` -> `普通文本公式`
- text after `References` / `参考文献` -> reference entries
- text after `Acknowledgements` / `致谢` -> acknowledgement body

## Reporting Back

Keep the final response short:
- link the copied output DOCX
- when only stage 1 is complete, link the editable `.xlsx` confirmation table and ask the user to send back the confirmed table plus target thesis
- summarize the main formatting changes
- state whether comments were added and what comment mode was used
- state the LibreOffice visual QA directory, whether the rendered PNGs were inspected, and whether any visual-QA correction loop was needed
- when runtime matters, report the `performance-analyze.json`/`performance-apply.json` total seconds and PDF/PNG pass counts rather than copying detailed logs
- state whether delivery cleanup removed the used confirmation table and generated work files, or whether they were retained for reuse
- state any limits, especially legacy `.doc` conversion uncertainty
