---
name: mcp-toolsmith
description: >
  Design, lint, and review Model Context Protocol (MCP) tools, resources,
  prompts, and server configurations. Use when creating MCP servers, exposing
  APIs to agents, reviewing tool schemas, or checking tool metadata for safety
  and usability before publishing.
version: 1.0.0
author: Zijian Guo
license: MIT
platforms: [linux, macos]
prerequisites:
  commands: [python3]
metadata:
  hermes:
    tags: [agent, mcp, tools, schemas, interoperability, tool-design]
---

# mcp-toolsmith - MCP tool design and linting

Use this skill for agent-tool interoperability work, especially when converting
local scripts, APIs, or internal systems into agent-callable MCP tools.

## Research Signal

MCP has become a public integration layer for agent tools. MCP servers expose
tools, resources, and prompts; the market gap is no longer only "can an agent
call a tool", but whether the tool schema is clear, least-privilege, auditable,
and resistant to malicious metadata.

## Workflow

1. Model read-only data as resources where possible; reserve tools for actions.
2. Give every tool a specific name, narrow description, typed input schema, and
   explicit side-effect notes.
3. Avoid tool descriptions that contain hidden instructions to the model.
4. Add validation and authorization outside the model.
5. Lint schemas/configs before publishing or installing them.

## Script

```
~/.hermes/skills/mcp-toolsmith/scripts/mcp_schema_lint.py
```

## Checked Items

| Check | Purpose |
|---|---|
| Tool name format | Keeps tools discoverable and stable |
| Description quality | Avoids vague or instruction-like metadata |
| Input schema | Requires object schemas with typed properties |
| Side-effect flags | Highlights write/delete/send/execute tools |
| Poisoning phrases | Detects metadata that tries to steer the model |

## Pitfalls

- A valid schema can still be unsafe. Review runtime authorization separately.
- Tool descriptions are part of model context; treat them as executable influence.
