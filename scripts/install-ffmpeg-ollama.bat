@echo off
REM Requires winget (App Installer). Re-open the terminal after install if PATH is not updated.

echo Installing ffmpeg (Gyan.FFmpeg)...
winget install -e --id Gyan.FFmpeg --accept-package-agreements --accept-source-agreements

echo Installing Ollama...
winget install -e --id Ollama.Ollama --accept-package-agreements --accept-source-agreements

echo Done.
