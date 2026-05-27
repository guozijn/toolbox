---
name: system-inventory
description: >
  Inspect a user's computer environment: OS, architecture, Python/runtime paths,
  common developer commands, disk space, selected environment variables, and
  network identity. Use before installing tools, debugging setup problems, or
  deciding which local commands are available.
version: 1.0.0
author: Zijian Guo
license: MIT
platforms: [linux, macos]
prerequisites:
  commands: [python3]
metadata:
  hermes:
    tags: [computer, system, environment, diagnostics, setup]
---

# system-inventory - Local environment inspection

Use this skill to understand the machine before running setup, build, browser,
or deployment workflows.

## Workflow

1. Collect OS, Python, shell, current directory, disk, and command availability.
2. Redact sensitive environment values; report presence rather than contents.
3. Use JSON output when another script or agent will consume the result.
4. Avoid changing system state during inventory.

## Script

```
~/.hermes/skills/system-inventory/scripts/system_inventory.py
```

## Checks

| Check | Purpose |
|---|---|
| OS and architecture | Pick platform-specific instructions |
| Command availability | Detect missing tools before work starts |
| Disk space | Avoid failed downloads/builds |
| Python executable/version | Choose compatible scripts |
| Sensitive env presence | Know whether API-backed tools can run without leaking values |
