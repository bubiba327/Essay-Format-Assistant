# 论文格式助手 Codex Plugin

This repository packages the `lunwen-geshi-zhushou` Codex skill as a GitHub-ready Codex plugin.

The skill analyzes a thesis format sample, exports an editable XLSX confirmation table, then applies the confirmed table to a copied Word thesis with comments and tiered LibreOffice visual QA.

## 最简单用法

选中这个插件，或明确告诉 Codex “使用论文格式助手这个 skill”。然后直接把格式范文发给 Codex，并说明“这是我的格式范文”。发送后，Codex 会分析范文并反馈一个格式确认表；你确认之后，再把需要修改格式的论文发给 Codex，它会创建一个副本，并在副本中修改你的论文格式。

## Contents

- `.codex-plugin/plugin.json` - Codex plugin manifest
- `skills/lunwen-geshi-zhushou/SKILL.md` - skill instructions
- `skills/lunwen-geshi-zhushou/scripts/` - formatter, QA, and evaluation scripts
- `skills/lunwen-geshi-zhushou/references/` - role-mapping reference
- `skills/lunwen-geshi-zhushou/fixtures/` - public fixture manifest template and baseline summaries

Generated evaluation runs, shared caches, local fixture documents, and Python bytecode are intentionally excluded from Git.

## Requirements

- Python 3.12+
- Python packages from `requirements.txt`
- Recommended for final delivery: macOS with LibreOffice installed at `/Applications/LibreOffice.app/Contents/MacOS/soffice`, or `SOFFICE` set to another `soffice` executable path

Install Python dependencies:

```bash
python3 -m pip install -r requirements.txt
```

## LibreOffice Modes

- `--qa-level fast`: trial mode. Works without LibreOffice only when all inputs are already `.docx`. It skips PDF/PNG rendering and is not suitable for final delivery.
- `--qa-level review`: normal working mode. Requires LibreOffice and renders lighter contact-sheet visual QA.
- `--qa-level strict`: final delivery mode. Requires LibreOffice and runs the full visual QA and TOC page validation gate.

Legacy `.doc`, `.rtf`, and `.odt` files require LibreOffice for conversion. If you do not have LibreOffice, first save those files as `.docx` in Word/WPS, then run `fast` as a trial.

## Validate

Run script-level checks:

```bash
./scripts/test_all.sh
```

Run real-document evaluation after adding local fixture documents:

```bash
mkdir -p skills/lunwen-geshi-zhushou/fixtures/local
# Add your private documents here:
# - format1.doc
# - format2.doc
# - format3.docx
# - sample1.docx
# - sample2.doc

python3 skills/lunwen-geshi-zhushou/scripts/evaluate_thesis_skill.py \
  --case all \
  --qa-level review \
  --output-dir /tmp/lunwen-review
```

The local fixture documents are ignored by Git so private thesis samples are not published.

## Basic Usage

Stage 1, analyze a sample and create a confirmation table:

```bash
skills/lunwen-geshi-zhushou/scripts/run_thesis_format_from_sample.sh \
  --sample path/to/format-sample.docx \
  --analysis-dir .thesis_format_analysis \
  --export-format-table 格式确认表.xlsx \
  --qa-level review \
  --quiet \
  --analyze-only
```

Stage 2, apply the confirmed table to a target thesis copy:

```bash
skills/lunwen-geshi-zhushou/scripts/run_thesis_format_from_sample.sh \
  --format-table 格式确认表.xlsx \
  --reuse-profile .thesis_format_analysis/format-sample.format-profile.json \
  --target path/to/thesis.docx \
  --output thesis_格式调整试改.docx \
  --analysis-dir .thesis_format_analysis \
  --qa-level strict \
  --quiet \
  --apply-engine styles \
  --comment-mode role
```

## Publishing To GitHub

After creating a GitHub repository, add it as the remote and push:

```bash
git remote add origin git@github.com:YOUR_ACCOUNT/lunwen-geshi-zhushou.git
git push -u origin main
```

If you prefer HTTPS:

```bash
git remote add origin https://github.com/YOUR_ACCOUNT/lunwen-geshi-zhushou.git
git push -u origin main
```

## License

MIT. See `LICENSE`.
