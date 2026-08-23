@echo off
REM Requires winget (App Installer). Re-open the terminal after install if PATH is not updated.

echo Installing ffmpeg (Gyan.FFmpeg)...
winget install -e --id Gyan.FFmpeg --accept-package-agreements --accept-source-agreements

echo Installing llama.cpp (llama-server)...
winget install -e --id ggml.llamacpp --accept-package-agreements --accept-source-agreements

echo Done.
