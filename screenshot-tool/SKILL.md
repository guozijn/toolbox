---
name: screenshot-tool
description: >
  Capture desktop screenshots on macOS or Linux using available system commands.
  Use when the user asks to inspect visible UI state, preserve evidence, or debug
  desktop/browser rendering outside a controlled Playwright session.
version: 1.0.0
author: Zijian Guo
license: MIT
platforms: [linux, macos]
prerequisites:
  commands: [python3]
metadata:
  hermes:
    tags: [computer, screenshot, desktop, ui, evidence]
---

# screenshot-tool - Desktop screenshot capture

Use this skill when a real desktop screenshot is needed.

## Workflow

1. Capture to an explicit file path.
2. Use interactive/area capture only when the user can select the region.
3. Inspect the image after capture when the task depends on visual details.
4. Avoid capturing sensitive screens unless the user asks.

## Script

```
~/.hermes/skills/screenshot-tool/scripts/screenshot_tool.py
```

## Pitfalls

- macOS may require Screen Recording permission for terminal applications.
- Headless Linux sessions may not have a screenshot backend.
