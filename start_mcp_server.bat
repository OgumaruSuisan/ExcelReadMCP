@echo off
setlocal

set "SCRIPT_DIR=%~dp0"
set "VENV_PYTHON=%SCRIPT_DIR%.venv\Scripts\python.exe"

rem ensure package root is discoverable by Python
set "PYTHONPATH=%SCRIPT_DIR%;%PYTHONPATH%"

if exist "%VENV_PYTHON%" (
    "%VENV_PYTHON%" -m excel_read_mcp.server
) else (
    python -m excel_read_mcp.server
)
