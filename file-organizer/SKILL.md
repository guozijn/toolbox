---
name: file-organizer
description: >
  Plan and optionally apply local file organization tasks: group files by
  extension, date, or media type, and find duplicate files by hash. Use when a
  user asks to clean up Downloads, Desktop, project artifacts, or duplicate files.
version: 1.0.0
author: Zijian Guo
license: MIT
platforms: [linux, macos]
prerequisites:
  commands: [python3]
metadata:
  hermes:
    tags: [computer, files, cleanup, organize, duplicates, dry-run]
---

# file-organizer - Local file cleanup planning

Use this skill for safe file cleanup on a user's computer.

## Workflow

1. Start with a dry-run plan.
2. Review moves or duplicates with the user for high-value folders.
3. Apply only when the user explicitly wants changes.
4. Keep destination folders predictable and reversible.

## Script

```
~/.hermes/skills/file-organizer/scripts/file_organizer.py
```

## Pitfalls

- File moves can break project references. Avoid organizing source repositories
  unless the user explicitly asks.
- Duplicate detection by hash is accurate for content equality, but deletion is
  intentionally not implemented.
