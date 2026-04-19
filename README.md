# toolbox
A set of useful toolkits and scripts

---

## social_post

A command-line tool for posting messages to a social media matrix — X (Twitter), LinkedIn, Facebook Page, Reddit, and Buffer — simultaneously or selectively.

### Requirements

```
pip install tweepy requests praw python-dotenv
```

### Setup

Copy `.env.example` to `.env` and fill in your API credentials for each platform you intend to use.

| Platform | Where to get credentials |
|---|---|
| X (Twitter) | https://developer.twitter.com/en/portal/dashboard |
| LinkedIn | https://www.linkedin.com/developers/apps — scopes: `w_member_social`, `r_liteprofile` |
| Facebook Page | https://developers.facebook.com/apps — Page Access Token with `pages_manage_posts` scope |
| Reddit | https://www.reddit.com/prefs/apps — script-type app |
| Buffer | https://buffer.com/developers/apps — Personal Access Token or OAuth2 app |

### Usage

```
python social_post.py "Your message" [options]
```

| Option | Default | Description |
|---|---|---|
| `--platforms` | `all` | Space-separated list: `x`, `linkedin`, `facebook`, `reddit`, `buffer`, `all` |
| `--image` | — | Path to an image file to attach |
| `--reddit-subreddit` | — | Subreddit name (required when Reddit is targeted) |
| `--reddit-title` | — | Post title for Reddit (required when Reddit is targeted) |
| `--buffer-profile-ids` | — | Space-separated Buffer profile IDs (default: all connected profiles) |
| `--dry-run` | — | Preview what would be posted without sending |
| `--output` | `text` | Output format: `text` or `json` |

**Examples**

```bash
# Post to all platforms
python social_post.py "Hello world!" --platforms all \
    --reddit-subreddit python --reddit-title "Hello from social_post"

# Post only to X and LinkedIn
python social_post.py "New blog post is live" --platforms x linkedin

# Post with an image to Facebook and LinkedIn
python social_post.py "Check this photo" --platforms facebook linkedin --image ./photo.jpg

# Dry run to preview before sending
python social_post.py "Test" --platforms all --dry-run \
    --reddit-subreddit test --reddit-title "Test post"

# Get machine-readable JSON output
python social_post.py "Hello" --platforms x linkedin --output json
```

---

## pdf_tool

A command-line utility for common PDF operations.

### Requirements

```
pip install pypdf reportlab
```

For PPTX conversion, [LibreOffice](https://www.libreoffice.org/) must be installed.

### Usage

```
python pdf_tool.py <command> [options]
```

### Commands

#### `merge`

Merge all PDF files in a directory into a single PDF, sorted in natural order (e.g. `Lec 2` before `Lec 10`).

```
python pdf_tool.py merge [options]
```

| Option | Default | Description |
|---|---|---|
| `-d`, `--directory` | `.` | Directory to scan for PDF files (non-recursive) |
| `-o`, `--output` | `merged.pdf` | Output file path |
| `--include-pptx` | — | Also convert and include PPTX files (requires LibreOffice) |
| `--prepend-titles` | — | Insert a title page with the source filename before each document |
| `--soffice-path` | — | Explicit path to the LibreOffice `soffice` binary |

**Examples**

```bash
# Merge all PDFs in the current directory
python pdf_tool.py merge

# Merge PDFs in a specific folder, save to a custom path
python pdf_tool.py merge -d ./slides -o ./slides/all.pdf

# Include PPTX files and add a title page before each document
python pdf_tool.py merge -d ./slides --include-pptx --prepend-titles
```

---

#### `extract`

Extract specific pages from a PDF into a new file.

```
python pdf_tool.py extract <input> -p <pages> [options]
```

| Argument | Description |
|---|---|
| `input` | Path to the source PDF |
| `-p`, `--pages` | Page selection (1-based): comma-separated numbers and ranges, e.g. `1,3-5,8` |
| `-o`, `--output` | Output file path (default: `extracted.pdf` next to the source file) |

**Examples**

```bash
# Extract pages 1, 3, and 5 through 8
python pdf_tool.py extract report.pdf -p 1,3,5-8

# Extract page 2 and save to a specific path
python pdf_tool.py extract report.pdf -p 2 -o page2.pdf
```
