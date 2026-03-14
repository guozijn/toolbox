# toolbox
A set of useful toolkits and scripts

---

## pdf_tool

A command-line utility for common PDF operations.

### Requirements

```
pip install pypdf reportlab
```

For PPTX conversion, [LibreOffice](https://www.libreoffice.org/) must be installed.

### Usage

```
python pdf_tool.py <command> [options]
```

### Commands

#### `merge`

Merge all PDF files in a directory into a single PDF, sorted in natural order (e.g. `Lec 2` before `Lec 10`).

```
python pdf_tool.py merge [options]
```

| Option | Default | Description |
|---|---|---|
| `-d`, `--directory` | `.` | Directory to scan for PDF files (non-recursive) |
| `-o`, `--output` | `merged.pdf` | Output file path |
| `--include-pptx` | — | Also convert and include PPTX files (requires LibreOffice) |
| `--prepend-titles` | — | Insert a title page with the source filename before each document |
| `--soffice-path` | — | Explicit path to the LibreOffice `soffice` binary |

**Examples**

```bash
# Merge all PDFs in the current directory
python pdf_tool.py merge

# Merge PDFs in a specific folder, save to a custom path
python pdf_tool.py merge -d ./slides -o ./slides/all.pdf

# Include PPTX files and add a title page before each document
python pdf_tool.py merge -d ./slides --include-pptx --prepend-titles
```

---

#### `extract`

Extract specific pages from a PDF into a new file.

```
python pdf_tool.py extract <input> -p <pages> [options]
```

| Argument | Description |
|---|---|
| `input` | Path to the source PDF |
| `-p`, `--pages` | Page selection (1-based): comma-separated numbers and ranges, e.g. `1,3-5,8` |
| `-o`, `--output` | Output file path (default: `extracted.pdf` next to the source file) |

**Examples**

```bash
# Extract pages 1, 3, and 5 through 8
python pdf_tool.py extract report.pdf -p 1,3,5-8

# Extract page 2 and save to a specific path
python pdf_tool.py extract report.pdf -p 2 -o page2.pdf
```
