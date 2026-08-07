#!/bin/bash
# Start the quiz tool (double-click to run)
cd "$(dirname "$0")"
# Kill any leftover server process to avoid port conflict
kill $(lsof -ti :8000 2>/dev/null) 2>/dev/null
sleep 1
echo "Starting quiz tool…"
echo "URL: http://localhost:8000"
(sleep 1; open http://localhost:8000) &
# Prefer venv python (has PyMuPDF for PDF images)
if [ -x ".venv/bin/python" ]; then
    exec .venv/bin/python app.py
else
    exec python3 app.py
fi
