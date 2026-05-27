---
name: file-hash
description: >
  Compute stable checksums for one or more files using hashlib algorithms.
  Use when the user needs reproducible content verification or artifact IDs.
version: 1.0.0
author: Zijian Guo
license: MIT
platforms: [linux, macos]
prerequisites:
  commands: [python3]
metadata:
  hermes:
    tags: [agent-tools, checksum, sha256, hash, files]
---

# file-hash - File checksums

Compute file checksums with Python's standard library.

## When to Use

- User asks for a file hash or checksum.
- You need to verify generated artifacts.
- You need a stable identifier for file contents.

## Script

```
~/.hermes/skills/file-hash/scripts/file_hash.py
```

## Options

| Option | Default | Description |
|---|---|---|
| `paths` | required | One or more files to hash |
| `--algorithm` | `sha256` | Any algorithm supported by `hashlib` |

## Pitfalls

- Directories are not supported; pass specific files.
