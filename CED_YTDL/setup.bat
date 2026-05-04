@echo off
echo ==============================================
echo      YouTube Downloader - Environment Setup
echo ==============================================
echo.

echo [1/3] Updating Python pip...
py -3 -m pip install --upgrade pip
echo.

echo [2/3] Installing/Updating packages...
py -3 -m pip install -U yt-dlp Pillow
echo.

echo [3/3] Checking and installing FFmpeg...
py -3 install_ffmpeg.py
echo.

echo ==============================================
echo      All installations completed!
echo      You can close this window now.
echo ==============================================
pause
