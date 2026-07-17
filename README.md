# Citation Renumberer

A small **local web app** to renumber bracketed citations in a Microsoft Word
(`.docx`) document and keep the reference list in sync. Runs entirely on your
machine, nothing is uploaded anywhere.

## Features
- Order-of-appearance numbering (IEEE/Vancouver): first citation becomes [1], next [2], etc.
- Manual mapping mode for custom old -> new pairs.
- Handles single and grouped citations: [1], [52], [1, 2, 3], [1,2], [12-15].
- Collision-safe two-pass placeholder method (52 -> 1 and 1 -> 3 never clash).
- Only numbers inside citation brackets change; years, DOIs, page numbers are untouched.
- Reference list is renumbered and re-sorted to match.
- Preserves run-level formatting (font, bold, italic, color).
- Live preview of the old -> new plan before download.

## Requirements
- Python 3.9+
- Flask and python-docx (see requirements.txt)

## Installation
    python -m pip install -r requirements.txt

## Usage
    python app.py

Then open http://127.0.0.1:5000, choose your .docx, preview the plan, and
download the renumbered file. Press Ctrl+C to stop the server.

## Limitations
- Reference list must be in the same file under a "References"/"Bibliography" heading.
- Assumes one paragraph per reference entry.
- Citations in footnotes/endnotes or hyperlinks are not modified.
- Always keep your original; the tool outputs a new copy.

## License
MIT
