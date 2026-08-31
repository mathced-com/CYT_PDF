@echo off
cd /d "%~dp0"

echo ===============================================
echo      Installing required packages...
echo ===============================================
echo.

if exist "..\.venv\Scripts\python.exe" (
    echo [INFO] Environment found. Installing...
    "..\.venv\Scripts\python.exe" -m pip install customtkinter pypdf pymupdf Pillow pyinstaller pdf2docx
) else (
    echo [ERROR] Virtual environment not found at ..\.venv
)

echo.
echo ===============================================
echo      Installation Process Ended.
echo ===============================================
pause
