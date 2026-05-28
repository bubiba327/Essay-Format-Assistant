#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PYTHON="${CODEX_BUNDLED_PYTHON:-$HOME/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3}"

if [[ ! -x "$PYTHON" ]]; then
  echo "Codex bundled Python not found: $PYTHON" >&2
  echo "Run load_workspace_dependencies or set CODEX_BUNDLED_PYTHON to the bundled python3 path." >&2
  exit 127
fi

exec "$PYTHON" "$SCRIPT_DIR/thesis_format_from_sample.py" "$@"
