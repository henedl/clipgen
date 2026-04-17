#!/usr/bin/env sh
# macOS: Homebrew. Linux: apt-based (Debian/Ubuntu) only; elsewhere install ffmpeg and Ollama yourself.
set -e

case "$(uname -s)" in
Darwin)
  brew install ffmpeg ollama
  ;;
Linux)
  sudo apt-get update
  sudo apt-get install -y ffmpeg
  curl -fsSL https://ollama.com/install.sh | sh
  ;;
*)
  echo "Unsupported OS: need Darwin (macOS) or Linux (apt)." >&2
  exit 1
  ;;
esac
