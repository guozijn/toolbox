---
name: agent-workflow-graph
description: >
  Design and validate graph/state-machine agent workflows with explicit starts,
  terminal states, guarded cycles, human approval points, and side-effecting tool
  boundaries. Use when building LangGraph-style, multi-agent, or durable agent
  workflows.
version: 1.0.0
author: Zijian Guo
license: MIT
platforms: [linux, macos]
prerequisites:
  commands: [python3]
metadata:
  hermes:
    tags: [agent, workflow, graph, state-machine, orchestration, multi-agent]
---

# agent-workflow-graph - Agent workflow validation

Use this skill when an agent should be modeled as a controlled workflow instead
of an unconstrained loop.

## Research Signal

Public agent frameworks are moving toward graph/state-machine orchestration for
durable, inspectable workflows. The gap in many projects is not lack of an LLM,
but missing control-flow invariants: unreachable nodes, unbounded loops, missing
terminal states, and side-effecting tools without approval boundaries.

## Workflow

1. Describe the workflow as nodes, edges, start, and terminals.
2. Mark side-effecting tool nodes and human approval nodes explicitly.
3. Check graph reachability and terminal reachability.
4. Add guards or iteration limits to cycles.
5. Keep the graph spec near the code that implements it.

## Script

```
~/.hermes/skills/agent-workflow-graph/scripts/workflow_graph_check.py
```

## Spec Shape

```json
{
  "start": "plan",
  "terminal": ["done"],
  "nodes": [{"id": "plan"}, {"id": "send_email", "side_effect": true, "requires_approval": true}, {"id": "done"}],
  "edges": [{"from": "plan", "to": "send_email"}, {"from": "send_email", "to": "done"}]
}
```

## Pitfalls

- Graph validation does not prove the node implementation is correct.
- Approval nodes should show raw tool arguments, not only the agent summary.
