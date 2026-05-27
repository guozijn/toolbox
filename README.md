# toolbox

A collection of command-line scripts and Hermes skills.

Each tool lives in its own skill directory and can be used directly or installed
into a Hermes agent via `hermes skills install`.

## Tool Index

| Tool | Category | Purpose |
|---|---|---|
| `pdf-tool` | Documents | Merge PDFs, extract pages, and optionally convert PPTX during merges |
| `social-matrix` | Publishing | Cross-post messages to X, LinkedIn, Facebook Page, Reddit, and Buffer |
| `repo-map` | Core agent utility | Produce compact repository or directory overviews |
| `text-search` | Core agent utility | Search project text with ripgrep-first behavior |
| `url-fetch` | Core agent utility | Fetch HTTP(S) responses for preview or saving |
| `json-inspect` | Core agent utility | Validate, summarize, pretty-print, and query JSON |
| `csv-preview` | Core agent utility | Preview CSV/TSV headers and rows |
| `file-hash` | Core agent utility | Compute reproducible file checksums |
| `mcp-toolsmith` | Advanced agent | Design and lint MCP tool schemas and metadata |
| `agent-eval-harness` | Advanced agent | Run repeatable agent regression evals and pass-rate checks |
| `agent-trace-inspector` | Advanced agent | Inspect traces, spans, tool calls, errors, and latency |
| `context-provenance` | Advanced agent | Build source-grounded context packs with hashes and line ranges |
| `agent-security-audit` | Advanced agent | Scan prompts, MCP configs, tools, and logs for agent security risks |
| `agent-workflow-graph` | Advanced agent | Validate graph/state-machine workflows and approval boundaries |
| `browser-agent` | Advanced agent | Build and review browser-using agents and browser runtime readiness |
| `system-inventory` | Computer operation | Inspect OS, runtimes, commands, disk space, and env presence |
| `port-process-inspector` | Computer operation | Inspect local listening ports and owning processes |
| `archive-manager` | Computer operation | Safely list, create, and extract zip/tar archives |
| `file-organizer` | Computer operation | Plan file cleanup, grouping, and duplicate detection |
| `clipboard-tool` | Computer operation | Read from or write to the system clipboard |
| `screenshot-tool` | Computer operation | Capture desktop screenshots with available OS commands |

---

## Tools

### pdf-tool

Merge all PDFs in a directory into one file, or extract specific pages from a PDF.

**Install via Hermes:**
```bash
hermes skills install guozijn/toolbox/pdf-tool --force
```
> `--force` is required to bypass a false positive in hermes's security scanner (it flags subprocess calls as a supply-chain risk).

**Use directly:**
```bash
pip install pypdf reportlab        # or: uv pip install pypdf reportlab
python3 pdf-tool/scripts/pdf_tool.py merge -d ./slides -o all.pdf
python3 pdf-tool/scripts/pdf_tool.py extract report.pdf -p 1,3,5-8 -o out.pdf
```

| Command | Description |
|---|---|
| `merge` | Merge all PDFs in a directory (optionally include PPTX, add title pages) |
| `extract` | Extract specific pages from a PDF into a new file |

**merge options:**

| Option | Default | Description |
|---|---|---|
| `-d`, `--directory` | `.` | Directory to scan for PDF files |
| `-o`, `--output` | `merged.pdf` | Output file path |
| `--include-pptx` | — | Also convert and include PPTX files (requires LibreOffice) |
| `--prepend-titles` | — | Insert a title page with the source filename before each document |
| `--soffice-path` | — | Explicit path to the LibreOffice `soffice` binary |

**extract options:**

| Option | Default | Description |
|---|---|---|
| `input` | (required) | Path to the source PDF |
| `-p`, `--pages` | (required) | Page selection: e.g. `1,3-5,8` |
| `-o`, `--output` | `extracted.pdf` | Output file path |

**Examples:**
```bash
# Merge all PDFs in current directory
python3 pdf-tool/scripts/pdf_tool.py merge

# Merge PDFs and PPTX files with title pages
python3 pdf-tool/scripts/pdf_tool.py merge -d ./slides --include-pptx --prepend-titles

# Extract pages 1, 3, and 5 through 8
python3 pdf-tool/scripts/pdf_tool.py extract report.pdf -p 1,3,5-8

# Extract page 2 to a specific path
python3 pdf-tool/scripts/pdf_tool.py extract report.pdf -p 2 -o page2.pdf
```

---

### social-matrix

Post messages to X (Twitter), LinkedIn, Facebook Page, Reddit, and Buffer
simultaneously or selectively from a single command.

**Install via Hermes:**
```bash
hermes skills install guozijn/toolbox/social-matrix --force
```
> `--force` is required to bypass a false positive in hermes's security scanner (it flags reading API keys from env vars as exfiltration).

**Use directly:**
```bash
pip install tweepy requests praw python-dotenv    # or: uv pip install ...
cp social-matrix/.env.example .env   # fill in your API credentials
python3 social-matrix/scripts/social_post.py "Your message" --platforms all \
    --reddit-subreddit python --reddit-title "Hello"
```

**Setup:** Copy `.env.example` to `.env` and fill in credentials for each platform you use.

| Platform | Required credentials |
|---|---|
| X (Twitter) | `X_API_KEY`, `X_API_SECRET`, `X_ACCESS_TOKEN`, `X_ACCESS_TOKEN_SECRET` |
| LinkedIn | `LINKEDIN_ACCESS_TOKEN` |
| Facebook Page | `FACEBOOK_PAGE_ID`, `FACEBOOK_PAGE_ACCESS_TOKEN` |
| Reddit | `REDDIT_CLIENT_ID`, `REDDIT_CLIENT_SECRET`, `REDDIT_USERNAME`, `REDDIT_PASSWORD` |
| Buffer | `BUFFER_ACCESS_TOKEN` (optional: `BUFFER_CHANNEL_IDS`) |

**Options:**

| Option | Default | Description |
|---|---|---|
| `--platforms` | `all` | Space-separated: `x`, `linkedin`, `facebook`, `reddit`, `buffer`, `all` |
| `--image` | — | Path to image file to attach |
| `--reddit-subreddit` | — | Subreddit name (required when Reddit is targeted) |
| `--reddit-title` | — | Post title for Reddit (required when Reddit is targeted) |
| `--buffer-channel-ids` | — | Buffer channel IDs (default: all connected channels) |
| `--buffer-facebook-post-type` | `post` | Facebook post type via Buffer: `post`, `story`, `reel` |
| `--dry-run` | — | Preview without sending |
| `--output` | `text` | Output format: `text` or `json` |

**Examples:**
```bash
# Post to all platforms
python3 social-matrix/scripts/social_post.py "Hello world!" --platforms all \
    --reddit-subreddit python --reddit-title "Hello"

# Post to X and LinkedIn only
python3 social-matrix/scripts/social_post.py "New post" --platforms x linkedin

# Post via Buffer with a Facebook story
python3 social-matrix/scripts/social_post.py "Check this" --platforms buffer \
    --buffer-facebook-post-type story

# Dry run
python3 social-matrix/scripts/social_post.py "Test" --platforms all --dry-run \
    --reddit-subreddit test --reddit-title "Test"
```

---

### Core agent tools

Small standard-library utilities for common AI-agent workflows. Each tool is a
separate skill so agents can install only what they need.

| Tool | Purpose | Install via Hermes |
|---|---|---|
| `repo-map` | Compact repository or directory overview | `hermes skills install guozijn/toolbox/repo-map --force` |
| `text-search` | Project text search with `rg` fallback behavior | `hermes skills install guozijn/toolbox/text-search --force` |
| `url-fetch` | HTTP(S) fetch and response preview/save | `hermes skills install guozijn/toolbox/url-fetch --force` |
| `json-inspect` | Validate, summarize, pretty-print, or query JSON | `hermes skills install guozijn/toolbox/json-inspect --force` |
| `csv-preview` | Preview CSV/TSV headers and rows | `hermes skills install guozijn/toolbox/csv-preview --force` |
| `file-hash` | Compute reproducible file checksums | `hermes skills install guozijn/toolbox/file-hash --force` |

---

### Advanced agent skills

Research-informed skills for building, evaluating, observing, and securing
modern AI agents.

| Tool | Purpose | Install via Hermes |
|---|---|---|
| `mcp-toolsmith` | Design and lint MCP tool schemas and server metadata | `hermes skills install guozijn/toolbox/mcp-toolsmith --force` |
| `agent-eval-harness` | Run repeatable agent regression evals and pass-rate checks | `hermes skills install guozijn/toolbox/agent-eval-harness --force` |
| `agent-trace-inspector` | Inspect agent traces, spans, tool calls, errors, and latency | `hermes skills install guozijn/toolbox/agent-trace-inspector --force` |
| `context-provenance` | Build source-grounded context packs with hashes and line ranges | `hermes skills install guozijn/toolbox/context-provenance --force` |
| `agent-security-audit` | Scan agent prompts, MCP configs, tools, and logs for common risks | `hermes skills install guozijn/toolbox/agent-security-audit --force` |
| `agent-workflow-graph` | Validate graph/state-machine workflows for reachability and approvals | `hermes skills install guozijn/toolbox/agent-workflow-graph --force` |
| `browser-agent` | Build and review browser-using agents and browser runtime readiness | `hermes skills install guozijn/toolbox/browser-agent --force` |

---

### Computer operation tools

Common local-computer skills for setup diagnostics, desktop workflow, and safe
file operations.

| Tool | Purpose | Install via Hermes |
|---|---|---|
| `system-inventory` | Inspect OS, runtimes, commands, disk space, and env presence | `hermes skills install guozijn/toolbox/system-inventory --force` |
| `port-process-inspector` | Inspect local listening ports and owning processes | `hermes skills install guozijn/toolbox/port-process-inspector --force` |
| `archive-manager` | Safely list, create, and extract zip/tar archives | `hermes skills install guozijn/toolbox/archive-manager --force` |
| `file-organizer` | Plan file cleanup, grouping, and duplicate detection with dry-run defaults | `hermes skills install guozijn/toolbox/file-organizer --force` |
| `clipboard-tool` | Read from or write to the system clipboard | `hermes skills install guozijn/toolbox/clipboard-tool --force` |
| `screenshot-tool` | Capture desktop screenshots with available OS commands | `hermes skills install guozijn/toolbox/screenshot-tool --force` |
