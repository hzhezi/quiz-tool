#!/bin/bash
# Install dependencies (PyMuPDF for PDF image extraction) — run once before first use
cd "$(dirname "$0")"
echo "Checking Python3…"
command -v python3 >/dev/null 2>&1 || { echo "Python3 not found. Install it first (https://www.python.org or brew install python)"; read -n 1 -s -r -p "Press any key to exit"; exit 1; }

if [ -x ".venv/bin/python" ]; then
    echo "Dependencies already installed."
    echo "Double-click start.command to launch."
    read -n 1 -s -r -p "Press any key to exit"; exit 0
fi

echo "Creating venv and installing PyMuPDF (PDF image support)…"
if command -v uv >/dev/null 2>&1; then
    uv venv .venv && uv pip install --python .venv/bin/python pymupdf
else
    python3 -m venv .venv && .venv/bin/pip install --upgrade pip && .venv/bin/pip install pymupdf
fi

# 公式(WMF)转图工具：题目中的分数/公式显示需要
if ! command -v wmf2svg >/dev/null 2>&1 || ! command -v rsvg-convert >/dev/null 2>&1; then
    echo "Installing formula-image tools (libwmf + librsvg)…"
    if command -v brew >/dev/null 2>&1; then
        brew install libwmf librsvg
    else
        echo "⚠️ 未找到 Homebrew，公式图可能无法显示（题目文字不受影响）"
    fi
fi

if [ -x ".venv/bin/python" ]; then
    echo "Done. Double-click start.command to launch."
else
    echo "Install failed, please check your network and retry."
fi
read -n 1 -s -r -p "Press any key to exit"
