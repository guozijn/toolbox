#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any


def load_json(path_value: str) -> Any:
    if path_value == "-":
        return json.load(sys.stdin)
    with Path(path_value).expanduser().open("r", encoding="utf-8") as handle:
        return json.load(handle)


def extract_path(data: Any, path: str) -> Any:
    current = data
    for part in path.split("."):
        if isinstance(current, list):
            current = current[int(part)]
        elif isinstance(current, dict):
            current = current[part]
        else:
            raise KeyError(f"cannot descend into {type(current).__name__} at {part}")
    return current


def main() -> int:
    parser = argparse.ArgumentParser(description="Validate, summarize, or query JSON.")
    parser.add_argument("input")
    parser.add_argument("--pretty", action="store_true")
    parser.add_argument("--path")
    args = parser.parse_args()

    try:
        data = load_json(args.input)
        if args.path:
            data = extract_path(data, args.path)
    except Exception as exc:
        print(f"error: {exc}")
        return 1

    if args.pretty or args.path:
        print(json.dumps(data, indent=2, sort_keys=True, ensure_ascii=False))
    elif isinstance(data, dict):
        print(f"valid JSON object with {len(data)} keys")
        print("keys:", ", ".join(map(str, list(data.keys())[:20])))
    elif isinstance(data, list):
        print(f"valid JSON array with {len(data)} items")
    else:
        print(f"valid JSON {type(data).__name__}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
