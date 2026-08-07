# Quiz Tool (行测刷题工具)

Extract questions from Word/PDF question banks, practice locally with wrong-question book and statistics. All data stays on your machine.

## Requirements

| Item | Requirement |
|---|---|
| OS | **macOS** (startup script is .command) |
| Python | Python 3 (via Homebrew or python.org) |
| Importing `.doc` | Requires **Microsoft Word** installed |
| PDF image support | Run `install.command` once (installs PyMuPDF) |

> docx / pdf / txt imports do NOT need Word; only legacy `.doc` format does.

## First Run

1. Unzip the package
2. Double-click **`install.command`** (one-time dependency install)
3. Double-click **`start.command`** — browser opens http://localhost:8000

## Features

- **Practice**: answer question-by-question, auto-grading + explanation; answered questions locked, wrong ones can be redone
- **Bank**: browse/search by paper, filter unanswered / wrong / correct
- **Wrong book**: redo recently-wrong questions, auto-removed when correct
- **To-do**: questions without answers, set answers in one place
- **Import**: "📥 Import" button — doc / docx / pdf / txt; upload question + answer files, or a single document containing both
- **Stats**: progress & accuracy per paper; jump to practice / delete paper
- **Persistence**: answers, records, wrong book stored in local SQLite — survive restarts

## Data Files

- `data/questions.json` — built-in question bank (12 papers + imported)
- `data/imported.json` — questions you imported
- `data/quiz.db` — your answer records, wrong book, edited answers
- `data/images/` — question images

To fully reset (clear records & your answer edits): click "重置记录" in the top bar.

## Re-extract built-in bank (optional, usually not needed)

```bash
python3 extract.py    # or .venv/bin/python extract.py
```
Re-parses the `.doc` papers in the parent directory (requires Word).
