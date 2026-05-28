#!/usr/bin/env bash
set -euo pipefail

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
SCRIPT_DIR="$ROOT/skills/lunwen-geshi-zhushou/scripts"
PYTHON="${PYTHON:-python3}"

cd "$SCRIPT_DIR"
"$PYTHON" -m py_compile thesis_format_from_sample.py evaluate_thesis_skill.py report_thesis_eval.py thesis_qa_checks.py thesis_style_engine.py

for test_file in test_*.py; do
  "$PYTHON" "$test_file"
done
