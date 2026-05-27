#!/usr/bin/env python3
from __future__ import annotations

import argparse
import hashlib
from pathlib import Path


def main() -> int:
    parser = argparse.ArgumentParser(description="Hash one or more files.")
    parser.add_argument("paths", nargs="+")
    parser.add_argument("--algorithm", default="sha256")
    args = parser.parse_args()

    algorithm = args.algorithm.lower()
    if algorithm not in hashlib.algorithms_available:
        print(f"error: unsupported hash algorithm: {args.algorithm}")
        return 1

    for path_value in args.paths:
        path = Path(path_value).expanduser()
        if not path.is_file():
            print(f"error: file not found: {path}")
            return 1
        digest = hashlib.new(algorithm)
        with path.open("rb") as handle:
            for chunk in iter(lambda: handle.read(1024 * 1024), b""):
                digest.update(chunk)
        print(f"{digest.hexdigest()}  {path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
