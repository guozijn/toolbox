---
name: context-provenance
description: >
  Build source-grounded context packs with file hashes, line ranges, chunk text,
  and simple retrieval selection. Use when preparing reliable context for agents,
  RAG experiments, codebase analysis, or long-document tasks where provenance
  and source traceability matter.
version: 1.0.0
author: Zijian Guo
license: MIT
platforms: [linux, macos]
prerequisites:
  commands: [python3]
metadata:
  hermes:
    tags: [agent, context-engineering, rag, provenance, citations, memory]
---

# context-provenance - Source-grounded context packs

Use this skill when an agent needs context that can be traced back to exact
files, line spans, and content hashes.

## Research Signal

Long-context models still show position sensitivity, and RAG systems often fail
because retrieved context is stale, weakly attributed, or hard to audit. A small
provenance pack gives agents enough structure to cite and verify sources instead
of relying on copied blobs.

## Workflow

1. Pack relevant files into chunks with line ranges and hashes.
2. Select chunks by query before putting them into the model context.
3. Preserve chunk IDs in the final answer or downstream trace.
4. Rebuild packs when source files change.

## Script

```
~/.hermes/skills/context-provenance/scripts/context_pack.py
```

## Output Fields

| Field | Meaning |
|---|---|
| `id` | Stable chunk identifier |
| `path` | Source file path |
| `start_line`, `end_line` | Inclusive source line range |
| `sha256` | Hash of the chunk text |
| `token_estimate` | Approximate token count |
| `text` | Chunk content |

## Guidance

- Prefer smaller chunks for code and policy text; use larger chunks for prose.
- Keep provenance metadata with any summary derived from a pack.
- Do not mix private corpora unless access control is explicit.
