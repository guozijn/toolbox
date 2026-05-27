---
name: csv-preview
description: >
  Preview CSV or TSV files in a compact table with delimiter sniffing. Use when
  inspecting tabular text data before transforming or importing it.
version: 1.0.0
author: Zijian Guo
license: MIT
platforms: [linux, macos]
prerequisites:
  commands: [python3]
metadata:
  hermes:
    tags: [agent-tools, csv, tsv, data, preview]
---

# csv-preview - Tabular file preview

Print the header and first rows of CSV or TSV data.

## When to Use

- User asks what columns a CSV has.
- You need to inspect rows before writing a transform.
- You need quick delimiter-aware output.

## Script

```
~/.hermes/skills/csv-preview/scripts/csv_preview.py
```

## Options

| Option | Default | Description |
|---|---|---|
| `input` | required | CSV/TSV file path |
| `--rows` | `5` | Number of data rows to print |
| `--delimiter` | auto | Explicit delimiter |

## Pitfalls

- Very wide tables are truncated to 40 characters per cell for readability.
