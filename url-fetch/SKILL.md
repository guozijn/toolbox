---
name: url-fetch
description: >
  Fetch HTTP or HTTPS URLs, preview response metadata and text, or save the body
  to disk. Use for lightweight web/API retrieval when JavaScript execution is
  not required.
version: 1.0.0
author: Zijian Guo
license: MIT
platforms: [linux, macos]
prerequisites:
  commands: [python3]
metadata:
  hermes:
    tags: [agent-tools, http, web, api, fetch]
---

# url-fetch - HTTP fetch utility

Fetch a URL with Python standard-library networking.

## When to Use

- User asks to retrieve a raw URL or API endpoint.
- You need to save a remote file.
- You need a quick text preview and status code.

## Script

```
~/.hermes/skills/url-fetch/scripts/url_fetch.py
```

## Options

| Option | Default | Description |
|---|---|---|
| `url` | required | HTTP or HTTPS URL |
| `-o`, `--output` | none | Save response body to a file |
| `--max-chars` | `4000` | Maximum characters to print when not saving |
| `--timeout` | `20` | Request timeout in seconds |
| `--header` | repeatable | Extra header as `Name: value` |

## Pitfalls

- This does not execute JavaScript. Use browser automation for JS-heavy pages.
- Secrets passed in headers may be visible in shell history.
