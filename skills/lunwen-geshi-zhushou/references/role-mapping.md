# Thesis Role Mapping

Use this mapping when the sample or target document uses direct formatting instead of reliable Word styles.

| Role | Detection | Formatting Intent |
| --- | --- | --- |
| Front heading | `摘  要`, `摘要`, `Abstract`, `ABSTRACT` | Centered, bold, sample heading size. |
| Contents heading | `CONTENTS`, `目录`, `目  录` | Centered, bold, usually larger than abstract heading. |
| TOC level 1 | Word `toc 1`, `CHAPTER`, `第...章`, `References`, `Acknowledgements` entries | No left indent, larger/bolder, paragraph spacing around major entries. |
| TOC level 2 | Word `toc 2`, `1.1 ...` entries | Slight left indent, usually bold. |
| TOC level 3 | Word `toc 3`, `1.1.1 ...` entries | Deeper indent; confirm italic/all-caps/small-caps from the rendered sample instead of assuming field-code styling is visible. |
| Paper title/meta | English paper title, `Class: ...`, `Student ID...`, `Name: ...` | Front matter/title block formatting; often appears between abstracts and Chapter 1. |
| Chapter heading | `Chapter 1 ...` or `第一章 ...` | Primary body heading; apply Word `Heading 1`. |
| Section heading | `1.1 ...` | Secondary body heading; apply Word `Heading 2`. |
| Subsection heading | `1.1.1 ...`, including loose spacing like `5.2. 3 ...` | Tertiary body heading; apply Word `Heading 3`, not body indentation. |
| Body paragraph | Long prose outside TOC/references/acknowledgements | Justified, first-line indent, thesis body font. |
| Figure caption | `Figure 1-1 Title...` | Centered, smaller, bold. Do not treat explanatory paragraphs like `Figure 3-1 shows...` as captions. |
| Table caption | `Table 5-1 Title...` | Centered, smaller, bold. |
| Figure note / table note | `Source:...`, `Note:...`, `注：...` near a figure or table | Usually smaller than body; keep as editable text rows. |
| Equation | True Word/WPS `m:oMath` / `m:oMathPara` paragraphs, or legacy `EQ` / `MERGEFORMAT` equation fields | Preserve internal math XML; apply only paragraph-level layout such as alignment, spacing, indentation, and tab stops. |
| Equation number | Standalone `(1)`, `（1）`, `(3-1)` | Separate from body text; usually no first-line indent and often right aligned. |
| Plain-text fallback equation | `share_ikt=X_ikt/X_it (2)`, `edu_it = Σ (P_j,it × W_j) / P_it (6)` | Treat as an equation paragraph, not body prose; remove inherited body first-line indent. |
| Table footnote/subnote | `t statistics in parentheses`, `Note: ...`, `* p < 0.1, ** p < 0.05, *** p < 0.01` under a table | Separate from table body and body prose; expose as `表下注释` in the confirmation table. |
| Table overall | A Word table object, or flattened paragraphs after a `Table x-x` caption in converted legacy `.doc` | Table alignment, text wrapping, layout mode, total width, and cell margins. |
| Table header text | First row of a Word table, or leading labels after a flattened table caption | Complete text formatting, usually centered/bold. |
| Table body text | Remaining Word table rows, or body-like paragraphs after a flattened table caption | Complete text formatting; may differ from headers. |
| References | After `References` or `参考文献` | Usually smaller than body, justified, no first-line indent unless sample shows otherwise. |
| Acknowledgements | After `Acknowledgements` or `致谢` | Heading centered; prose like body; signature/date right aligned. |

Important: treat converted legacy `.doc` samples as approximate. Cross-check with rendered or previewed pages before making final claims about exact page layout.

The editable `.xlsx` confirmation table should include only roles observed in the sample. The user-confirmed table is authoritative: if a row is edited, use the edited values; if `是否应用` is `否`, do not format that role in the target thesis.

Confirmation-table rows must expose complete typography and paragraph controls for every role: western font, Chinese font, font color name/value, size, bold, italic, all caps, small caps, alignment, line spacing, before/after spacing, first-line indent, left indent, hanging indent, and right indent. Use human-readable values: `未特别设置` for unset bold/italic/caps effects, `0` for no spacing or indent, and `自动（通常黑色）` plus `#000000` for automatic black text.

For TOC roles, also include page-number-right-aligned, dot-leader, and right-tab-position fields. If a converted legacy `.doc` has TOC entries with tabs or page references but no exposed tab stops, infer a right-aligned dotted tab at the usable page width and let the user correct it in the `.xlsx`.

For table roles, include `表题`, `表头文字`, `表体文字`, `表下注释`, `表注`, and `表格整体设置` when observed. `表格整体设置` must preserve the sample's overall table width and text wrapping. A table with Word `tblpPr` is floating/around text; a normal thesis table should usually be `不环绕` so surrounding paragraphs break before and after the table instead of flowing beside it. If legacy `.doc` conversion flattens a table into paragraphs, infer the first label-like paragraphs after the table caption as headers and later paragraphs as body text. Image layout settings are inferred and applied automatically; figure/table captions and notes remain editable text roles in the `.xlsx`.

Keyword paragraphs may need split-run formatting: the `关键词:` / `Key words:` label and the keyword content can use different font properties even when Word stores them in one run. Split the label from the content before applying role formats.

Heading blank paragraphs are part of format, not content. Learn explicit blank paragraph counts around `chapter`, `second`, and `third` headings. Remove extra blank paragraphs around section/subsection headings when the sample has none; preserve or insert the sample count around chapter headings, commonly two blank lines before a new chapter.

After applying the profile, run structural heading QA. Loose heading numbers such as `5.2. 3 Explained variable` must not keep body first-line indentation. If indentation or blank-line counts differ from the target heading role, the run has not passed.

Visual QA must catch layout overflow. In addition to manual PNG inspection, treat side-edge ink in the rendered body area as a failed visual gate because it often indicates a table, image, or paragraph has exceeded the printable area.

When adding format-change comments, attach comments to the first changed text run in the relevant paragraph/table cell. In `role` mode, add one representative comment per role; in `all` mode, add one comment to each changed paragraph/table cell until the safety cap is reached.
