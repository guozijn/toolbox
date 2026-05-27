---
name: browser-agent
description: >
  Build, run, and review browser-using AI agents with Playwright, Browser Use,
  MCP browser tools, or hosted browser runtimes. Use for web navigation agents,
  form-filling, UI task automation, screenshots, DOM/vision checks, auth/session
  handling, and browser-task reliability reviews.
version: 1.0.0
author: Zijian Guo
license: MIT
platforms: [linux, macos]
prerequisites:
  commands: [python3]
metadata:
  hermes:
    tags: [agent, browser, playwright, browser-use, web-automation, ui]
---

# browser-agent - Browser automation agents

Use this skill when an agent needs to operate a real website rather than only
fetch static HTML.

## Research Signal

Browser agents are a fast-growing public category. Browser Use reports real-world
browser task benchmarks and exposes custom tools, hosted/cloud execution, MCP,
and authentication patterns. The production gap is reliability: screenshots,
DOM state, session handling, CAPTCHAs, consent screens, and side effects must be
handled deliberately.

## Workflow

1. Classify the task: read-only browse, authenticated browse, form fill, purchase
   or destructive action.
2. Use Playwright for deterministic UI tests; use browser-agent frameworks when
   high-level planning across pages is needed.
3. Capture screenshots, current URL, visible text, console errors, and network
   failures at each checkpoint.
4. Require explicit approval before submitting forms, purchasing, sending
   messages, or changing account data.
5. Convert repeated browser failures into eval cases.

## Script

```
~/.hermes/skills/browser-agent/scripts/browser_env_check.py
```

## Readiness Checks

| Check | Why it matters |
|---|---|
| Python packages | Detects Browser Use and Playwright availability |
| Node commands | Detects Playwright CLI/npm availability |
| Browser binaries | Indicates whether a local browser can run |
| Environment variables | Warns about API keys needed by hosted runtimes |

## Pitfalls

- Static HTTP fetch is not a substitute for JS-heavy pages.
- CAPTCHA and bot-detection handling may require hosted browser infrastructure.
- Never let a browser agent auto-submit high-impact actions without raw-action
  approval.
