# toolbox

A collection of command-line scripts and Hermes skills.

Each tool lives in its own skill directory and can be used directly or installed
into a Hermes agent via `hermes skills install`.

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
