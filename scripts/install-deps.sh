#!/usr/bin/env sh
# macOS: Homebrew. Linux: apt-based (Debian/Ubuntu) only; elsewhere install
# ffmpeg yourself and put a llama.cpp release's llama-server on PATH.
set -e

case "$(uname -s)" in
Darwin)
  brew install ffmpeg llama.cpp
  ;;
Linux)
  sudo apt-get update
  sudo apt-get install -y ffmpeg
  echo "llama.cpp has no apt package: download a release from"
  echo "https://github.com/ggml-org/llama.cpp/releases and put llama-server on PATH."
  ;;
*)
  echo "Unsupported OS: need Darwin (macOS) or Linux (apt)." >&2
  exit 1
  ;;
esac
