---
name: agent-trace-inspector
description: >
  Inspect agent traces, tool-call logs, and OpenTelemetry/OpenInference-style
  span exports. Use when debugging why an agent skipped a tool, called the wrong
  tool, hit errors, exceeded latency budgets, or produced an unsupported answer.
version: 1.0.0
author: Zijian Guo
license: MIT
platforms: [linux, macos]
prerequisites:
  commands: [python3]
metadata:
  hermes:
    tags: [agent, observability, traces, opentelemetry, openinference, tool-calls]
---

# agent-trace-inspector - Trace and tool-call inspection

Use this skill when an agent outcome is wrong and the important question is
which intermediate step failed.

## Research Signal

OpenTelemetry now defines GenAI agent span conventions, including agent
operations and tool execution spans. OpenInference similarly standardizes traces
for LLM application execution. The practical gap is that teams often have logs
but lack a fast local way to inspect operation mix, failed spans, tool calls, and
latency hotspots.

## Workflow

1. Export traces or JSONL logs from the agent runtime.
2. Inspect operation counts, error spans, tool-call names, and latency totals.
3. Check whether retrieval, tool execution, and synthesis were all represented.
4. Compare traces before and after a prompt/tool/runtime change.

## Script

```
~/.hermes/skills/agent-trace-inspector/scripts/trace_inspect.py
```

## Supported Inputs

- JSON arrays of spans or runs.
- JSONL files containing one span/run per line.
- Nested objects containing `spans`, `children`, `events`, or `attributes`.

## Pitfalls

- Trace schemas differ across vendors. Treat this script as a local normalizer,
  not as a replacement for production observability.
- Sensitive prompts and tool arguments may be present in trace exports.
