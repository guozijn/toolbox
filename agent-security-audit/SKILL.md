---
name: agent-security-audit
description: >
  Audit agent repositories, MCP configurations, tool schemas, prompts, and logs
  for common security risks: secret exposure, prompt injection payloads, tool
  poisoning, command execution, broad filesystem access, and unsafe auto-approval.
version: 1.0.0
author: Zijian Guo
license: MIT
platforms: [linux, macos]
prerequisites:
  commands: [python3]
metadata:
  hermes:
    tags: [agent, security, mcp, prompt-injection, tool-poisoning, secrets]
---

# agent-security-audit - Agent tool-surface review

Use this skill before enabling new tools, installing MCP servers, approving
agent memory, or exposing agents to untrusted content.

## Research Signal

Public MCP and LLM security guidance highlights tool poisoning, prompt injection,
command execution, excessive privileges, secret exposure, and context
over-sharing as practical agent failure modes.

## Workflow

1. Scan the repository, MCP config, tool schemas, and prompt files.
2. Review high-severity findings before enabling auto-approval or write tools.
3. Separate untrusted retrieval/browser content from privileged tools.
4. Prefer least-privilege tool scopes and short-lived credentials.
5. Re-scan when tools, prompts, or dependencies change.

## Script

```
~/.hermes/skills/agent-security-audit/scripts/agent_security_audit.py
```

## Finding Types

| Type | Examples |
|---|---|
| `secret_like_value` | API keys, tokens, passwords with assigned values |
| `prompt_injection_phrase` | "ignore previous instructions", "reveal system prompt" |
| `command_execution` | `subprocess`, `os.system`, `child_process.exec` |
| `risky_mcp_config` | MCP server commands, broad env passing, remote tools |
| `broad_file_access` | file read/write/delete tools with broad scope |

## Pitfalls

- Static scanning is incomplete. Treat clean output as a starting point, not a
  guarantee.
- Do not paste sensitive findings into shared prompts or tickets.
