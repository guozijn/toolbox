---
name: text-search
description: >
  Search project files for text or regex patterns, preferring ripgrep when
  available and falling back to a Python implementation. Use for fast codebase
  and document text discovery.
version: 1.0.0
author: Zijian Guo
license: MIT
platforms: [linux, macos]
prerequisites:
  commands: [python3]
metadata:
  hermes:
    tags: [agent-tools, search, grep, ripgrep, codebase]
---

# text-search - Project text search

Search files with stable line-numbered output. The script uses `rg` when
available and falls back to Python otherwise.

## When to Use

- User asks where a symbol, phrase, TODO, route, or setting appears.
- You need evidence before editing code.
- You need context lines around matches.

## Script

```
~/.hermes/skills/text-search/scripts/text_search.py
```

## Options

| Option | Default | Description |
|---|---|---|
| `pattern` | required | Text or regex pattern |
| `path` | `.` | File or directory to search |
| `--context` | `0` | Context lines before and after matches |
| `--fixed` | off | Treat pattern as literal text |
| `--ignore-case` | off | Case-insensitive search |
| `--glob` | repeatable | Include/exclude glob passed to `rg` when available |
| `--max-count` | `100` | Stop after this many printed lines in fallback mode |

## Pitfalls

- The fallback skips binary-looking files and common dependency directories.
- `--glob` is only honored by the `rg` path.
