---
name: social-matrix
description: >
  Post messages to a social media matrix — X (Twitter), LinkedIn, Facebook Page,
  Reddit, and Buffer — simultaneously or selectively from a single command.
  Use when the user wants to publish content to one or more social platforms at once,
  cross-post announcements, or automate social media publishing workflows.
version: 1.0.0
author: Zijian Guo
license: MIT
platforms: [linux, macos]
required_environment_variables:
  - name: BUFFER_ACCESS_TOKEN
    prompt: Buffer API key
    help: Generate at https://publish.buffer.com/settings/api
    required_for: Buffer posting
  - name: X_API_KEY
    prompt: X (Twitter) API key
    help: Create an app at https://developer.twitter.com/en/portal/dashboard
    required_for: X posting
  - name: X_API_SECRET
    prompt: X (Twitter) API secret
    help: From the same app at https://developer.twitter.com/en/portal/dashboard
    required_for: X posting
  - name: X_ACCESS_TOKEN
    prompt: X (Twitter) access token
    help: Generate under your app's Keys and Tokens tab (requires Read and Write permissions)
    required_for: X posting
  - name: X_ACCESS_TOKEN_SECRET
    prompt: X (Twitter) access token secret
    help: Generated alongside X_ACCESS_TOKEN
    required_for: X posting
  - name: LINKEDIN_ACCESS_TOKEN
    prompt: LinkedIn OAuth2 access token
    help: Create an app at https://www.linkedin.com/developers/apps with w_member_social scope
    required_for: LinkedIn posting
  - name: FACEBOOK_PAGE_ID
    prompt: Facebook Page ID
    help: Found in your Page settings or via the Graph API Explorer
    required_for: Facebook posting
  - name: FACEBOOK_PAGE_ACCESS_TOKEN
    prompt: Facebook Page access token
    help: Generate at https://developers.facebook.com/apps with pages_manage_posts permission
    required_for: Facebook posting
  - name: REDDIT_CLIENT_ID
    prompt: Reddit app client ID
    help: Create a script-type app at https://www.reddit.com/prefs/apps
    required_for: Reddit posting
  - name: REDDIT_CLIENT_SECRET
    prompt: Reddit app client secret
    help: From the same app at https://www.reddit.com/prefs/apps
    required_for: Reddit posting
  - name: REDDIT_USERNAME
    prompt: Reddit account username
    required_for: Reddit posting
  - name: REDDIT_PASSWORD
    prompt: Reddit account password
    required_for: Reddit posting
prerequisites:
  commands: [python3, uv]
metadata:
  hermes:
    tags: [social-media, twitter, x, linkedin, facebook, reddit, buffer, publishing, cross-posting]
    related_skills: [xitter]
---

# social-post — Multi-platform social media publishing

Publish a message to X (Twitter), LinkedIn, Facebook Page, Reddit, and Buffer
simultaneously or selectively using a single Python script.

## When to Use

- User wants to cross-post an announcement, blog post, or update to multiple platforms at once.
- User wants to publish to a specific subset of platforms.
- User wants to post with an attached image.
- User wants to push content through Buffer to all connected accounts.

Do NOT use this skill for reading timelines, searching posts, or interacting with
existing content — use the `xitter` skill for X-specific read operations.

## Setup

### Install dependencies

```bash
uv pip install --python ~/.hermes/hermes-agent/venv/bin/python tweepy praw
```

`requests` and `python-dotenv` are already present in the Hermes venv.

### Script location

After installing this skill, the helper script is at:

```
~/.hermes/skills/social-matrix/scripts/social_post.py
```

Set a short alias for convenience in this session:

```bash
SOCIAL="$(python3 -c "import os; print(os.path.expanduser('~/.hermes/skills/social-matrix/scripts/social_post.py'))")"
PYTHON="$HOME/.hermes/hermes-agent/venv/bin/python"
```

### Credentials

Each platform only needs its own credentials. Unset platforms are simply skipped.

| Platform | Required env vars |
|---|---|
| X (Twitter) | `X_API_KEY`, `X_API_SECRET`, `X_ACCESS_TOKEN`, `X_ACCESS_TOKEN_SECRET` |
| LinkedIn | `LINKEDIN_ACCESS_TOKEN` |
| Facebook Page | `FACEBOOK_PAGE_ID`, `FACEBOOK_PAGE_ACCESS_TOKEN` |
| Reddit | `REDDIT_CLIENT_ID`, `REDDIT_CLIENT_SECRET`, `REDDIT_USERNAME`, `REDDIT_PASSWORD` |
| Buffer | `BUFFER_ACCESS_TOKEN` (optional: `BUFFER_CHANNEL_IDS`) |

Store credentials in `~/.hermes/.env` — they are loaded automatically.

## Quick Reference

| Action | Command |
|---|---|
| Post to all platforms | `$PYTHON $SOCIAL "text" --platforms all --reddit-subreddit NAME --reddit-title "Title"` |
| Post to specific platforms | `$PYTHON $SOCIAL "text" --platforms x linkedin` |
| Post via Buffer only | `$PYTHON $SOCIAL "text" --platforms buffer` |
| Post with image | `$PYTHON $SOCIAL "text" --platforms all --image /path/to/photo.jpg ...` |
| Dry run preview | `$PYTHON $SOCIAL "text" --platforms all --dry-run ...` |
| JSON output | `$PYTHON $SOCIAL "text" --platforms x --output json` |

## Procedure

1. Confirm which platforms the user wants to post to.
2. Verify that credentials for those platforms are present in `~/.hermes/.env`.
3. For Reddit, confirm the subreddit name and post title.
4. Run with `--dry-run` first to preview what will be sent.
5. Run without `--dry-run` to publish.
6. Report back post IDs and URLs from the output.

## Platform Notes

### X (Twitter)

Requires a developer app with **Read and Write** permissions. Regenerate the access
token and secret after enabling write permissions or posts will fail with a 403.
X API free tier has strict rate limits — avoid posting identical content twice in
quick succession.

### LinkedIn

Requires an OAuth2 access token with `w_member_social` scope. Access tokens expire;
if posting fails with a 401, a new token must be generated via the OAuth2 flow.

### Facebook Page

Requires a Page Access Token (not a User token) with `pages_manage_posts` permission.
Text-only short posts render with large font on Facebook — this is Facebook's design,
not a bug. Attach an image or write longer text to get normal font rendering.

### Reddit

Uses username + password authentication (script-type app). Reddit blocks posting
identical content twice in a short window — vary the text if reposting.

### Buffer

Uses the Buffer GraphQL API (`https://api.buffer.com`) with a personal API key.
Get the key at https://publish.buffer.com/settings/api.

By default the script posts to **all channels** connected to the Buffer account.
To restrict to specific channels, set `BUFFER_CHANNEL_IDS` as a comma-separated
list of channel IDs, or pass `--buffer-channel-ids ID1 ID2`.

Facebook channels connected via Buffer require a post type. The default is `post`.
Override with `--buffer-facebook-post-type story` or `--buffer-facebook-post-type reel`.

## Options Reference

| Option | Default | Description |
|---|---|---|
| `--platforms` | `all` | Space-separated: `x`, `linkedin`, `facebook`, `reddit`, `buffer`, `all` |
| `--image` | — | Path to image file to attach |
| `--reddit-subreddit` | — | Subreddit name (required when Reddit is targeted) |
| `--reddit-title` | — | Post title for Reddit (required when Reddit is targeted) |
| `--buffer-channel-ids` | — | Buffer channel IDs (overrides `BUFFER_CHANNEL_IDS` env var) |
| `--buffer-facebook-post-type` | `post` | Facebook post type via Buffer: `post`, `story`, `reel` |
| `--dry-run` | — | Preview without sending |
| `--output` | `text` | Output format: `text` or `json` |

## Verification

- `[OK]` lines appear for each platform targeted.
- For JSON output, check `"success": true` per platform entry.
- Each successful post includes an `id` or `url` field that can be used to verify the post exists.
- If a platform returns `[FAIL]`, check the error message — most failures are credential
  or rate-limit issues, not script bugs.

## Pitfalls

- **X 403 oauth1-permissions**: Regenerate the access token after enabling Read and Write permissions.
- **LinkedIn 401**: Access token has expired — generate a new one via OAuth2.
- **Reddit duplicate**: Reddit rejects identical posts posted close together — vary the text.
- **Buffer Facebook missing type**: Always passes `metadata.facebook.type` — defaults to `post`. If the channel requires a different type, use `--buffer-facebook-post-type`.
- **Buffer OIDC token rejected**: The token must be the API key from https://publish.buffer.com/settings/api, not an OAuth login token.
