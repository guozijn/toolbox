---
name: clipboard-tool
description: >
  Read from and write to the system clipboard on macOS or Linux when clipboard
  commands are available. Use for transferring short text snippets between agent
  work and the user's desktop applications.
version: 1.0.0
author: Zijian Guo
license: MIT
platforms: [linux, macos]
prerequisites:
  commands: [python3]
metadata:
  hermes:
    tags: [computer, clipboard, desktop, copy, paste]
---

# clipboard-tool - System clipboard bridge

Use this skill when the user explicitly wants text copied to or read from the
local clipboard.

## Workflow

1. Use clipboard reads only when the user asks or it is clearly part of the task.
2. Avoid putting secrets on the clipboard unless explicitly requested.
3. Prefer file artifacts for large content; use clipboard for short snippets.

## Script

```
~/.hermes/skills/clipboard-tool/scripts/clipboard_tool.py
```

## Pitfalls

- Clipboard contents may be visible to other local applications.
- Linux requires `wl-copy`/`wl-paste`, `xclip`, or `xsel`.
