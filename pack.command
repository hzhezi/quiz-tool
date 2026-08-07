#!/bin/bash
# Package the tool into a clean shareable zip (no venv / temp files / personal records)
cd "$(dirname "$0")"
STAMP=$(date +%Y%m%d_%H%M)
PKG="/tmp/quiz-tool_$STAMP"
ZIP="/tmp/quiz-tool_$STAMP.zip"

rm -rf "$PKG" "$ZIP"
mkdir -p "$PKG/data"

# Core code
cp app.py extract.py start.command install.command README.md "$PKG/" 2>/dev/null
cp -R static "$PKG/static" 2>/dev/null
# Question bank
cp data/questions.json "$PKG/data/" 2>/dev/null
cp -R data/images "$PKG/data/images" 2>/dev/null

# Excluded: .venv / tmp_docx / quiz.db(learning records) / imported.json(user imports)

cd /tmp
zip -rq "$ZIP" "quiz-tool_$STAMP"
rm -rf "$PKG"

echo "Done: $ZIP"
echo "Send this zip to classmates. Unzip, double-click install.command (PDF image support), then start.command to launch."
echo "Note: importing .doc files requires Microsoft Word installed; docx / pdf / txt do not."
open -R "$ZIP"
