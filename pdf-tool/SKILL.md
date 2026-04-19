---
name: pdf-tool
description: >
  Command-line utility for common PDF operations: merge all PDFs in a directory
  into one file (with optional PPTX conversion and title pages), and extract
  specific pages from a PDF into a new file. Use when the user needs to combine,
  split, or reorganize PDF documents.
version: 1.0.0
author: Zijian Guo
license: MIT
platforms: [linux, macos]
prerequisites:
  commands: [python3, uv]
metadata:
  hermes:
    tags: [pdf, documents, merge, extract, convert, pptx, office]
---

# pdf-tool — PDF merge and extract utility

Merge and extract PDF files from the command line using a single Python script.

## When to Use

- User wants to merge all PDFs (and optionally PPTX files) in a directory into one file.
- User wants to extract specific pages from a PDF into a new file.
- User wants title pages inserted before each document when merging.

## Setup

### Install dependencies

```bash
uv pip install pypdf reportlab
```

For PPTX conversion, LibreOffice must also be installed:
- macOS: `brew install --cask libreoffice`
- Linux: `apt install libreoffice` or `dnf install libreoffice`

### Script location

After installing this skill, the helper script is at:

```
~/.hermes/skills/pdf-tool/scripts/pdf_tool.py
```

Set a short alias for convenience:

```bash
PDF="~/.hermes/skills/pdf-tool/scripts/pdf_tool.py"
```

## Quick Reference

| Action | Command |
|---|---|
| Merge all PDFs in current directory | `python3 $PDF merge` |
| Merge PDFs in a specific directory | `python3 $PDF merge -d ./slides -o output.pdf` |
| Merge with PPTX files included | `python3 $PDF merge -d ./slides --include-pptx` |
| Merge with title pages | `python3 $PDF merge -d ./slides --prepend-titles` |
| Extract pages 1, 3, and 5–8 | `python3 $PDF extract report.pdf -p 1,3,5-8` |
| Extract page 2 to a specific path | `python3 $PDF extract report.pdf -p 2 -o page2.pdf` |

## Procedure

### Merging

1. Confirm the source directory and desired output path.
2. Run `merge` with `-d` and `-o`. Add `--include-pptx` if PPTX files should be included, `--prepend-titles` for title pages.
3. Verify the output file exists and report its path.

### Extracting

1. Confirm the source PDF path and the page selection.
2. Run `extract` with the input path, `-p` for pages, and `-o` for output.
3. Report the number of pages extracted and the output path.

## Options Reference

### `merge`

| Option | Default | Description |
|---|---|---|
| `-d`, `--directory` | `.` | Directory to scan for PDF files (non-recursive) |
| `-o`, `--output` | `merged.pdf` | Output file path |
| `--include-pptx` | — | Also convert and include PPTX files (requires LibreOffice) |
| `--prepend-titles` | — | Insert a title page with the source filename before each document |
| `--soffice-path` | — | Explicit path to the LibreOffice `soffice` binary |

### `extract`

| Option | Default | Description |
|---|---|---|
| `input` | (required) | Path to the source PDF |
| `-p`, `--pages` | (required) | Page selection: comma-separated numbers and ranges, e.g. `1,3-5,8` |
| `-o`, `--output` | `extracted.pdf` | Output file path (placed next to source file if relative) |

## Pitfalls

- **PPTX conversion fails**: LibreOffice must be installed and on PATH, or pass `--soffice-path` explicitly.
- **Page numbers are 1-based**: `1` is the first page, not `0`.
- **Output file excluded from merge inputs**: if the output path already exists inside the source directory, it is automatically excluded to prevent reading the output as an input.
- **Range direction**: `start` must be less than or equal to `end` in a range spec (e.g. `5-3` is invalid).
