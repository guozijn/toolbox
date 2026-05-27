#!/usr/bin/env python3
from __future__ import annotations

import argparse
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen


def parse_headers(values: list[str] | None) -> dict[str, str]:
    headers = {"User-Agent": "url-fetch/1.0"}
    for value in values or []:
        if ":" not in value:
            raise ValueError(f"header must be 'Name: value': {value}")
        name, header_value = value.split(":", 1)
        headers[name.strip()] = header_value.strip()
    return headers


def main() -> int:
    parser = argparse.ArgumentParser(description="Fetch an HTTP(S) URL.")
    parser.add_argument("url")
    parser.add_argument("-o", "--output")
    parser.add_argument("--max-chars", type=int, default=4000)
    parser.add_argument("--timeout", type=float, default=20)
    parser.add_argument("--header", action="append")
    args = parser.parse_args()

    try:
        request = Request(args.url, headers=parse_headers(args.header))
        with urlopen(request, timeout=args.timeout) as response:
            body = response.read()
            status = getattr(response, "status", 200)
            content_type = response.headers.get("content-type", "")
    except HTTPError as exc:
        body = exc.read()
        status = exc.code
        content_type = exc.headers.get("content-type", "")
    except (URLError, ValueError) as exc:
        print(f"error: {exc}")
        return 1

    if args.output:
        output = Path(args.output).expanduser()
        output.parent.mkdir(parents=True, exist_ok=True)
        output.write_bytes(body)
        print(f"status={status} content_type={content_type} bytes={len(body)} output={output}")
        return 0 if 200 <= status < 400 else 1

    text = body.decode("utf-8", errors="replace")
    print(f"status={status} content_type={content_type} bytes={len(body)}")
    print("")
    print(text[:args.max_chars])
    if len(text) > args.max_chars:
        print(f"\n... truncated after {args.max_chars} characters")
    return 0 if 200 <= status < 400 else 1


if __name__ == "__main__":
    raise SystemExit(main())
