#!/usr/bin/env bash
# MyFinance - one-command setup (macOS / Linux / Git Bash on Windows)
# Installs the Python packages, then runs the first-time wizard:
#   ./setup.sh                                  (interactive)
#   ./setup.sh --yes --name "You" --passcode x --income 10000 --no-repo
set -e
cd "$(dirname "$0")"
echo "=== MyFinance setup ==="

# Find a REAL Python 3 - rejects the Windows Store "python" stub, which
# prints "Python was not found" and exits non-zero.
PY=""
for c in python3 python "py -3"; do
  if $c -c "import sys" >/dev/null 2>&1; then PY=$c; break; fi
done
if [ -z "$PY" ]; then
  echo "[FAIL] Python 3 not found."
  echo "       Install: brew install python   (or https://www.python.org/downloads/)"
  exit 1
fi
echo "Python found: $($PY --version 2>&1 | head -n 1)"

echo "Installing packages: openpyxl python-docx plotly"
if ! $PY -m pip install --quiet openpyxl python-docx plotly; then
  # PEP 668 (Homebrew/system Pythons block direct installs) - retry for this user
  if ! $PY -m pip install --quiet --user --break-system-packages openpyxl python-docx plotly; then
    # last resort: isolated virtualenv
    echo "pip blocked -> creating .venv"
    if $PY -m venv .venv; then
      PY=".venv/bin/python"
      if ! $PY -m pip install --quiet openpyxl python-docx plotly; then
        echo "[FAIL] pip install failed inside .venv"
        exit 1
      fi
    else
      echo "[FAIL] pip blocked and .venv could not be created."
      echo "       Fix Python: https://www.python.org/downloads/ then re-run ./setup.sh"
      exit 1
    fi
  fi
fi

echo "Running first-time setup (config, sample SMS, build, verify)..."
$PY init_project.py "$@"

echo ""
echo "Setup done."
echo "Next step: install OpenCode from https://opencode.ai then run: opencode"
echo "OpenCode loads AGENTS.md + the myfinance skill automatically,"
echo "and system files are read-only for the agent (permission rules)."
