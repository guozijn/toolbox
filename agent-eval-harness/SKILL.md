---
name: agent-eval-harness
description: >
  Build and run lightweight regression evaluations for agents, tools, prompts,
  and workflows. Use when comparing agent versions, catching behavior drift,
  measuring pass rates over repeated runs, or turning production failures into
  replayable test cases.
version: 1.0.0
author: Zijian Guo
license: MIT
platforms: [linux, macos]
prerequisites:
  commands: [python3]
metadata:
  hermes:
    tags: [agent, evals, regression, pass-rate, testing, reliability]
---

# agent-eval-harness - Agent regression evaluation

Use this skill when an agent needs repeatable evidence that a prompt, tool, or
workflow still behaves correctly after changes.

## Research Signal

Public agent platforms increasingly treat evals as unit tests for agents:
curated datasets catch regressions during development, while online evals detect
production drift. Agent benchmarks such as tau-bench also emphasize repeated
trials because agent behavior can be inconsistent across runs.

## Workflow

1. Define cases as JSONL. Keep each case narrow and tied to one observable user
   outcome.
2. Run each case against the candidate command or harness.
3. Capture stdout, stderr, exit status, latency, and match results.
4. Summarize pass rate and inspect failures before changing prompts or code.
5. Add any production failure as a new case before fixing it.

## Script

```
~/.hermes/skills/agent-eval-harness/scripts/agent_eval.py
```

## Case Format

Each JSONL line is one case:

```json
{"id":"case-001","input":"question or task","expected_contains":["done"],"expected_regex":["order #[0-9]+"],"expected_exit_code":0}
```

Supported checks:

| Field | Meaning |
|---|---|
| `id` | Stable case identifier |
| `input` | Input passed through `AGENT_EVAL_INPUT` or stdin |
| `expected_contains` | Strings that must appear in stdout |
| `expected_regex` | Regex patterns that must match stdout |
| `forbidden_contains` | Strings that must not appear in stdout |
| `expected_exit_code` | Required process exit code |

## Guidance

- Prefer deterministic assertions over LLM-as-judge for small behavior checks.
- Run stochastic agents multiple times and report pass rates, not one lucky run.
- Keep raw outputs as artifacts; summaries alone are not enough for debugging.
