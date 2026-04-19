#!/usr/bin/env python3
"""
social_post.py - Social media matrix posting tool.

Supports posting to X (Twitter), LinkedIn, Facebook Page, Reddit, and Buffer
simultaneously or selectively via a single command.

Usage:
    python social_post.py "Your message here" [options]

Setup:
    Copy .env.example to .env and fill in your API credentials.
"""

import argparse
import os
import sys
import json
from pathlib import Path

try:
    from dotenv import load_dotenv
except ImportError:
    print("Missing dependency: pip install python-dotenv", file=sys.stderr)
    sys.exit(1)

load_dotenv()


# ---------------------------------------------------------------------------
# Platform clients
# ---------------------------------------------------------------------------


def post_to_x(text: str, image_path: str | None = None) -> dict:
    """Post to X (Twitter) using the v2 API via tweepy."""
    try:
        import tweepy
    except ImportError:
        return {"platform": "X", "success": False, "error": "tweepy not installed: pip install tweepy"}

    api_key = os.getenv("X_API_KEY")
    api_secret = os.getenv("X_API_SECRET")
    access_token = os.getenv("X_ACCESS_TOKEN")
    access_token_secret = os.getenv("X_ACCESS_TOKEN_SECRET")

    missing = [k for k, v in {
        "X_API_KEY": api_key,
        "X_API_SECRET": api_secret,
        "X_ACCESS_TOKEN": access_token,
        "X_ACCESS_TOKEN_SECRET": access_token_secret,
    }.items() if not v]
    if missing:
        return {"platform": "X", "success": False, "error": f"Missing env vars: {', '.join(missing)}"}

    client = tweepy.Client(
        consumer_key=api_key,
        consumer_secret=api_secret,
        access_token=access_token,
        access_token_secret=access_token_secret,
    )

    media_ids = None
    if image_path:
        auth = tweepy.OAuth1UserHandler(api_key, api_secret, access_token, access_token_secret)
        api_v1 = tweepy.API(auth)
        media = api_v1.media_upload(image_path)
        media_ids = [media.media_id]

    kwargs = {"text": text}
    if media_ids:
        kwargs["media_ids"] = media_ids

    response = client.create_tweet(**kwargs)
    tweet_id = response.data["id"]
    return {
        "platform": "X",
        "success": True,
        "id": tweet_id,
        "url": f"https://x.com/i/web/status/{tweet_id}",
    }


def post_to_linkedin(text: str, image_path: str | None = None) -> dict:
    """Post to LinkedIn as a personal share using the v2 API."""
    try:
        import requests
    except ImportError:
        return {"platform": "LinkedIn", "success": False, "error": "requests not installed: pip install requests"}

    access_token = os.getenv("LINKEDIN_ACCESS_TOKEN")
    if not access_token:
        return {"platform": "LinkedIn", "success": False, "error": "Missing env var: LINKEDIN_ACCESS_TOKEN"}

    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json",
        "X-Restli-Protocol-Version": "2.0.0",
    }

    # Resolve the authenticated member's URN
    me_resp = requests.get("https://api.linkedin.com/v2/userinfo", headers=headers)
    if not me_resp.ok:
        return {"platform": "LinkedIn", "success": False, "error": f"Failed to fetch profile: {me_resp.text}"}
    profile = me_resp.json()
    person_urn = f"urn:li:person:{profile['sub']}"

    media_asset = None
    if image_path:
        # Step 1: register upload
        register_body = {
            "registerUploadRequest": {
                "recipes": ["urn:li:digitalmediaRecipe:feedshare-image"],
                "owner": person_urn,
                "serviceRelationships": [
                    {"relationshipType": "OWNER", "identifier": "urn:li:userGeneratedContent"}
                ],
            }
        }
        reg_resp = requests.post(
            "https://api.linkedin.com/v2/assets?action=registerUpload",
            headers=headers,
            json=register_body,
        )
        if not reg_resp.ok:
            return {"platform": "LinkedIn", "success": False, "error": f"Image register failed: {reg_resp.text}"}
        reg_data = reg_resp.json()
        upload_url = reg_data["value"]["uploadMechanism"][
            "com.linkedin.digitalmedia.uploading.MediaUploadHttpRequest"
        ]["uploadUrl"]
        media_asset = reg_data["value"]["asset"]

        # Step 2: upload binary
        with open(image_path, "rb") as fh:
            upload_resp = requests.put(upload_url, data=fh, headers={"Authorization": f"Bearer {access_token}"})
        if not upload_resp.ok:
            return {"platform": "LinkedIn", "success": False, "error": f"Image upload failed: {upload_resp.text}"}

    post_body: dict = {
        "author": person_urn,
        "lifecycleState": "PUBLISHED",
        "specificContent": {
            "com.linkedin.ugc.ShareContent": {
                "shareCommentary": {"text": text},
                "shareMediaCategory": "IMAGE" if media_asset else "NONE",
            }
        },
        "visibility": {"com.linkedin.ugc.MemberNetworkVisibility": "PUBLIC"},
    }

    if media_asset:
        post_body["specificContent"]["com.linkedin.ugc.ShareContent"]["media"] = [
            {"status": "READY", "media": media_asset}
        ]

    post_resp = requests.post("https://api.linkedin.com/v2/ugcPosts", headers=headers, json=post_body)
    if not post_resp.ok:
        return {"platform": "LinkedIn", "success": False, "error": f"Post failed: {post_resp.text}"}

    post_id = post_resp.headers.get("x-restli-id", "")
    return {"platform": "LinkedIn", "success": True, "id": post_id}


def post_to_facebook(text: str, image_path: str | None = None) -> dict:
    """Post to a Facebook Page using the Graph API."""
    try:
        import requests
    except ImportError:
        return {"platform": "Facebook", "success": False, "error": "requests not installed: pip install requests"}

    page_id = os.getenv("FACEBOOK_PAGE_ID")
    page_access_token = os.getenv("FACEBOOK_PAGE_ACCESS_TOKEN")

    missing = [k for k, v in {
        "FACEBOOK_PAGE_ID": page_id,
        "FACEBOOK_PAGE_ACCESS_TOKEN": page_access_token,
    }.items() if not v]
    if missing:
        return {"platform": "Facebook", "success": False, "error": f"Missing env vars: {', '.join(missing)}"}

    base_url = f"https://graph.facebook.com/v19.0/{page_id}"

    if image_path:
        with open(image_path, "rb") as fh:
            resp = requests.post(
                f"{base_url}/photos",
                data={"message": text, "access_token": page_access_token},
                files={"source": fh},
            )
    else:
        resp = requests.post(
            f"{base_url}/feed",
            data={"message": text, "access_token": page_access_token},
        )

    if not resp.ok:
        return {"platform": "Facebook", "success": False, "error": f"Post failed: {resp.text}"}

    data = resp.json()
    post_id = data.get("post_id") or data.get("id", "")
    return {"platform": "Facebook", "success": True, "id": post_id}


def post_to_buffer(
    text: str,
    image_path: str | None = None,
    channel_ids: list[str] | None = None,
    facebook_post_type: str = "post",
) -> dict:
    """Post to one or more connected channels via the Buffer GraphQL API."""
    try:
        import requests
    except ImportError:
        return {"platform": "Buffer", "success": False, "error": "requests not installed: pip install requests"}

    access_token = os.getenv("BUFFER_ACCESS_TOKEN")
    if not access_token:
        return {"platform": "Buffer", "success": False, "error": "Missing env var: BUFFER_ACCESS_TOKEN"}

    endpoint = "https://api.buffer.com"
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json",
    }

    def gql(query: str, variables: dict | None = None) -> dict:
        payload: dict = {"query": query}
        if variables:
            payload["variables"] = variables
        resp = requests.post(endpoint, headers=headers, json=payload)
        resp.raise_for_status()
        return resp.json()

    # Always fetch all channels to get service info for metadata decisions
    account_data = gql("{ account { organizations { id } } }")
    errors = account_data.get("errors")
    if errors:
        return {"platform": "Buffer", "success": False, "error": f"Failed to fetch account: {errors}"}
    orgs = account_data["data"]["account"]["organizations"]
    if not orgs:
        return {"platform": "Buffer", "success": False, "error": "No organizations found in Buffer account"}
    org_id = orgs[0]["id"]

    channels_data = gql(
        "query GetChannels($input: ChannelsInput!) { channels(input: $input) { id name service } }",
        {"input": {"organizationId": org_id}},
    )
    errors = channels_data.get("errors")
    if errors:
        return {"platform": "Buffer", "success": False, "error": f"Failed to fetch channels: {errors}"}
    all_channels = channels_data["data"]["channels"]
    if not all_channels:
        return {"platform": "Buffer", "success": False, "error": "No connected channels found in Buffer account"}

    # Resolve which channels to post to: CLI arg > env var > all
    if not channel_ids:
        env_ids = os.getenv("BUFFER_CHANNEL_IDS", "")
        channel_ids = [c.strip() for c in env_ids.split(",") if c.strip()]

    channel_map = {c["id"]: c["service"] for c in all_channels}

    if channel_ids:
        # Filter to only known channels; warn about unknown IDs
        unknown = [cid for cid in channel_ids if cid not in channel_map]
        if unknown:
            return {"platform": "Buffer", "success": False, "error": f"Unknown channel IDs: {', '.join(unknown)}"}
        target_channels = [(cid, channel_map[cid]) for cid in channel_ids]
    else:
        target_channels = [(c["id"], c["service"]) for c in all_channels]

    # Create a post for each channel
    create_mutation = """
    mutation CreatePost($input: CreatePostInput!) {
      createPost(input: $input) {
        ... on PostActionSuccess {
          post { id text dueAt }
        }
        ... on MutationError {
          message
        }
      }
    }
    """

    results_per_channel = []
    errors_per_channel = []

    for cid, service in target_channels:
        post_input: dict = {
            "channelId": cid,
            "text": text,
            "schedulingType": "automatic",
            "mode": "shareNow",
        }
        if service == "facebook":
            post_input["metadata"] = {"facebook": {"type": facebook_post_type}}

        variables: dict = {"input": post_input}
        data = gql(create_mutation, variables)
        gql_errors = data.get("errors")
        if gql_errors:
            errors_per_channel.append(f"{cid}({service}): {gql_errors}")
            continue
        result = data["data"]["createPost"]
        if "message" in result:
            errors_per_channel.append(f"{cid}({service}): {result['message']}")
        else:
            post = result.get("post", {})
            results_per_channel.append({
                "channel_id": cid,
                "service": service,
                "post_id": post.get("id", ""),
                "due_at": post.get("dueAt", ""),
            })

    if errors_per_channel and not results_per_channel:
        return {"platform": "Buffer", "success": False, "error": "; ".join(errors_per_channel)}

    return {
        "platform": "Buffer",
        "success": True,
        "posts": results_per_channel,
        "channel_count": len(results_per_channel),
        **({"partial_errors": errors_per_channel} if errors_per_channel else {}),
    }


def post_to_reddit(text: str, subreddit_name: str, title: str, image_path: str | None = None) -> dict:
    """Post to a Reddit subreddit (text or image post)."""
    try:
        import praw
    except ImportError:
        return {"platform": "Reddit", "success": False, "error": "praw not installed: pip install praw"}

    client_id = os.getenv("REDDIT_CLIENT_ID")
    client_secret = os.getenv("REDDIT_CLIENT_SECRET")
    username = os.getenv("REDDIT_USERNAME")
    password = os.getenv("REDDIT_PASSWORD")
    user_agent = os.getenv("REDDIT_USER_AGENT", "social_post/1.0")

    missing = [k for k, v in {
        "REDDIT_CLIENT_ID": client_id,
        "REDDIT_CLIENT_SECRET": client_secret,
        "REDDIT_USERNAME": username,
        "REDDIT_PASSWORD": password,
    }.items() if not v]
    if missing:
        return {"platform": "Reddit", "success": False, "error": f"Missing env vars: {', '.join(missing)}"}

    reddit = praw.Reddit(
        client_id=client_id,
        client_secret=client_secret,
        username=username,
        password=password,
        user_agent=user_agent,
    )

    subreddit = reddit.subreddit(subreddit_name)

    if image_path:
        submission = subreddit.submit_image(title=title, image_path=image_path)
    else:
        submission = subreddit.submit(title=title, selftext=text)

    return {
        "platform": "Reddit",
        "success": True,
        "id": submission.id,
        "url": submission.url,
    }


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------


PLATFORM_CHOICES = ["x", "linkedin", "facebook", "reddit", "buffer", "all"]


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Post a message to one or more social media platforms.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Post to all platforms
  python social_post.py "Hello world!" --platforms all --reddit-subreddit python --reddit-title "Hello"

  # Post only to X and LinkedIn
  python social_post.py "New blog post" --platforms x linkedin

  # Post with an image
  python social_post.py "Check this out" --platforms all --image ./photo.jpg \\
      --reddit-subreddit python --reddit-title "Check this out"

  # Post via Buffer to specific profiles
  python social_post.py "New post" --platforms buffer --buffer-profile-ids abc123 def456

  # Dry run (preview without posting)
  python social_post.py "Test message" --platforms all --dry-run \\
      --reddit-subreddit test --reddit-title "Test"
""",
    )

    parser.add_argument("message", help="The text content to post")
    parser.add_argument(
        "--platforms",
        nargs="+",
        choices=PLATFORM_CHOICES,
        default=["all"],
        metavar="PLATFORM",
        help=f"Platforms to post to. Choices: {', '.join(PLATFORM_CHOICES)} (default: all)",
    )
    parser.add_argument("--image", metavar="PATH", help="Optional image file to attach to the post")
    parser.add_argument("--reddit-subreddit", metavar="NAME", help="Subreddit name (required when posting to Reddit)")
    parser.add_argument("--reddit-title", metavar="TITLE", help="Post title for Reddit (required when posting to Reddit)")
    parser.add_argument(
        "--buffer-channel-ids",
        nargs="+",
        metavar="ID",
        help="Buffer channel IDs to post to (overrides BUFFER_CHANNEL_IDS env var; defaults to all connected channels)",
    )
    parser.add_argument(
        "--buffer-facebook-post-type",
        choices=["post", "story", "reel"],
        default="post",
        help="Post type for Facebook channels connected via Buffer (default: post)",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Preview what would be posted without actually sending anything",
    )
    parser.add_argument(
        "--output",
        choices=["text", "json"],
        default="text",
        help="Output format (default: text)",
    )
    return parser


def resolve_platforms(platforms: list[str]) -> list[str]:
    if "all" in platforms:
        return ["x", "linkedin", "facebook", "reddit", "buffer"]
    return list(dict.fromkeys(platforms))  # deduplicate while preserving order


def print_result(result: dict, fmt: str) -> None:
    if fmt == "json":
        print(json.dumps(result))
        return
    platform = result["platform"]
    if result["success"]:
        url = result.get("url", "")
        post_id = result.get("id", "")
        detail = url or (f"id={post_id}" if post_id else "")
        print(f"[OK]    {platform}" + (f" — {detail}" if detail else ""))
    else:
        print(f"[FAIL]  {platform} — {result['error']}")


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()

    platforms = resolve_platforms(args.platforms)
    message: str = args.message
    image: str | None = args.image

    if image and not Path(image).is_file():
        print(f"Error: image file not found: {image}", file=sys.stderr)
        return 1

    needs_reddit = "reddit" in platforms
    if needs_reddit and (not args.reddit_subreddit or not args.reddit_title):
        print("Error: --reddit-subreddit and --reddit-title are required when posting to Reddit.", file=sys.stderr)
        return 1

    if args.dry_run:
        print("Dry run — no posts will be sent.\n")
        print(f"Message : {message}")
        print(f"Image   : {image or '(none)'}")
        print(f"Platforms: {', '.join(platforms)}")
        if needs_reddit:
            print(f"Reddit  : r/{args.reddit_subreddit} — \"{args.reddit_title}\"")
        if "buffer" in platforms:
            buf_ids = args.buffer_channel_ids or os.getenv("BUFFER_CHANNEL_IDS") or "(all connected channels)"
            print(f"Buffer  : channels={buf_ids}")
        return 0

    results: list[dict] = []

    for platform in platforms:
        if platform == "x":
            results.append(post_to_x(message, image))
        elif platform == "linkedin":
            results.append(post_to_linkedin(message, image))
        elif platform == "facebook":
            results.append(post_to_facebook(message, image))
        elif platform == "reddit":
            results.append(
                post_to_reddit(message, args.reddit_subreddit, args.reddit_title, image)
            )
        elif platform == "buffer":
            results.append(post_to_buffer(message, image, args.buffer_channel_ids, args.buffer_facebook_post_type))

    exit_code = 0
    if args.output == "json":
        print(json.dumps(results, indent=2))
    else:
        for result in results:
            print_result(result, "text")
        failures = [r for r in results if not r["success"]]
        if failures:
            exit_code = 1

    return exit_code


if __name__ == "__main__":
    sys.exit(main())
