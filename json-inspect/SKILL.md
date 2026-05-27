---
name: json-inspect
description: >
  Validate, summarize, pretty-print, and extract simple dot paths from JSON
  files or stdin. Use for API responses, config files, and structured data.
version: 1.0.0
author: Zijian Guo
license: MIT
platforms: [linux, macos]
prerequisites:
  commands: [python3]
metadata:
  hermes:
    tags: [agent-tools, json, data, inspect]
---

# json-inspect - JSON inspection

Inspect JSON data without third-party dependencies.

## When to Use

- User asks whether JSON is valid.
- You need to pretty-print a JSON file.
- You need to extract a simple nested value.

## Script

```
~/.hermes/skills/json-inspect/scripts/json_inspect.py
```

## Options

| Option | Default | Description |
|---|---|---|
| `input` | required | JSON file path, or `-` for stdin |
| `--pretty` | off | Pretty-print the loaded JSON |
| `--path` | none | Dot path to extract, e.g. `items.0.name` |

## Pitfalls

- Dot paths do not support object keys that contain literal dots.
