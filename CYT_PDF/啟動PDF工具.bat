@echo off
cd /d "%~dp0"

echo =========================================
echo      Starting CYT PDF Tool...
echo =========================================
echo.

set VENV_PATH=""

if exist "..\.venv\Scripts\python.exe" (
    set VENV_PATH="..\.venv\Scripts\python.exe"
) else if exist ".venv\Scripts\python.exe" (
    set VENV_PATH=".venv\Scripts\python.exe"
)

if not "%VENV_PATH%" == "" (
    echo [INFO] Virtual environment detected: %VENV_PATH%
    %VENV_PATH% app.py
) else (
    echo [INFO] No venv found, trying "py -3 app.py"...
    py -3 app.py
)

if %errorlevel% neq 0 (
    echo.
    echo =========================================
    echo [ERROR] Failed to start. 
    echo Current directory files:
    dir /b
    echo =========================================
    pause
)
