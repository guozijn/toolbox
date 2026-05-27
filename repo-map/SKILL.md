---
name: repo-map
description: >
  Map a repository or directory with a compact file listing, extension summary,
  and total size. Use before modifying unfamiliar projects or when the user asks
  for a quick project overview.
version: 1.0.0
author: Zijian Guo
license: MIT
platforms: [linux, macos]
prerequisites:
  commands: [python3]
metadata:
  hermes:
    tags: [agent-tools, repository, files, inspection]
---

# repo-map - Repository overview

Generate a compact map of a project without third-party dependencies.

## When to Use

- User asks what is in a repository or folder.
- You need orientation before making code changes.
- You need a machine-readable file inventory for downstream processing.

## Script

```
~/.hermes/skills/repo-map/scripts/repo_map.py
```

## Options

| Option | Default | Description |
|---|---|---|
| `path` | `.` | Directory to inspect |
| `--depth` | `3` | Maximum directory depth |
| `--max-files` | `200` | Maximum files to print |
| `--include-hidden` | off | Include hidden files and directories |
| `--json` | off | Emit machine-readable JSON |

## Pitfalls

- Common dependency and build directories are skipped by default.
- Hidden files are skipped unless `--include-hidden` is passed.
