@echo off
cd /d "%~dp0"

echo Starting YouTube Downloader...
py -3 main.py

if %errorlevel% neq 0 (
    echo.
    echo =========================================
    echo [ERROR] Failed to start! Please screenshot this window.
    echo =========================================
    pause
)
