---
name: archive-manager
description: >
  Safely list, create, and extract zip/tar archives with path traversal checks.
  Use when packaging artifacts, unpacking downloaded files, or inspecting archive
  contents on a user's computer.
version: 1.0.0
author: Zijian Guo
license: MIT
platforms: [linux, macos]
prerequisites:
  commands: [python3]
metadata:
  hermes:
    tags: [computer, archive, zip, tar, packaging, extraction]
---

# archive-manager - Safe archive operations

Use this skill for local zip and tar operations where predictable behavior and
safe extraction matter.

## Workflow

1. List archive contents before extracting unknown files.
2. Extract into a dedicated directory.
3. Reject members that would escape the destination path.
4. Create archives from explicit files or folders.

## Script

```
~/.hermes/skills/archive-manager/scripts/archive_manager.py
```

## Pitfalls

- Archives can contain absolute paths or `..` traversal entries. This tool blocks
  those during extraction.
- File permissions in tar archives may be platform-specific.
