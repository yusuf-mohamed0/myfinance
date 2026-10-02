@echo off
setlocal EnableExtensions
rem MyFinance - one-command setup (Windows)
rem   setup.bat
rem   setup.bat --yes --name "You" --passcode x --income 10000 --no-repo
cd /d "%~dp0"
echo === MyFinance setup ===

set "PY="
py -3 -c "import sys" >nul 2>nul && set "PY=py -3"
if not defined PY python -c "import sys" >nul 2>nul && set "PY=python"

if not defined PY (
  echo [FAIL] Python 3 not found.
  echo        Install:  winget install Python.Python.3.13
  echo        or:       https://www.python.org/downloads/
  exit /b 1
)
echo Python found: %PY%

echo Installing packages: openpyxl python-docx plotly
%PY% -m pip install --quiet openpyxl python-docx plotly
if errorlevel 1 (
  echo [FAIL] pip install failed - check your Python/pip and re-run setup.bat
  exit /b 1
)

echo Running first-time setup (config, sample SMS, build, verify)...
%PY% init_project.py %*
set "RC=%ERRORLEVEL%"
if not "%RC%"=="0" exit /b %RC%

echo.
echo Setup done.
echo Next step: install OpenCode from https://opencode.ai then run: opencode
echo OpenCode loads AGENTS.md + the myfinance skill automatically,
echo and system files are read-only for the agent (permission rules).
exit /b 0
