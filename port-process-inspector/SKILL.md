---
name: port-process-inspector
description: >
  Inspect local processes and network ports to debug dev servers, port conflicts,
  hung processes, and common localhost issues. Use before killing processes or
  changing ports; this skill is read-only by default.
version: 1.0.0
author: Zijian Guo
license: MIT
platforms: [linux, macos]
prerequisites:
  commands: [python3]
metadata:
  hermes:
    tags: [computer, ports, processes, localhost, dev-server, diagnostics]
---

# port-process-inspector - Local process and port diagnostics

Use this skill when a server will not start, a port is already in use, or a
localhost app is not reachable.

## Workflow

1. List listening ports and owning processes.
2. Inspect a specific port before deciding whether to reuse another port.
3. Report process IDs and commands; do not kill anything unless the user asks.
4. Prefer read-only inspection for shared development machines.

## Script

```
~/.hermes/skills/port-process-inspector/scripts/port_process_inspector.py
```

## Pitfalls

- Some process details may require elevated permissions.
- `lsof` is preferred when available; fallback output varies by platform.
