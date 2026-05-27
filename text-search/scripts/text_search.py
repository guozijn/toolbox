#!/usr/bin/env python3
from __future__ import annotations

import argparse
import os
import re
import shutil
import subprocess
from pathlib import Path
from typing import Iterable


SKIP_DIRS = {".git", "__pycache__", ".venv", "venv", "node_modules", "dist", "build", "target"}


def iter_files(root: Path) -> Iterable[Path]:
    if root.is_file():
        yield root
        return
    for dirpath, dirnames, filenames in os.walk(root):
        dirnames[:] = sorted(d for d in dirnames if d not in SKIP_DIRS and not d.startswith("."))
        for filename in sorted(filenames):
            if not filename.startswith("."):
                yield Path(dirpath) / filename


def looks_binary(path: Path) -> bool:
    try:
        return b"\0" in path.read_bytes()[:2048]
    except OSError:
        return True


def run_rg(args: argparse.Namespace, path: Path) -> int | None:
    rg = shutil.which("rg")
    if not rg:
        return None
    command = [rg, "--line-number", "--heading", "--color", "never"]
    if args.context:
        command.extend(["--context", str(args.context)])
    if args.fixed:
        command.append("--fixed-strings")
    if args.ignore_case:
        command.append("--ignore-case")
    for glob in args.glob or []:
        command.extend(["--glob", glob])
    command.extend([args.pattern, str(path)])
    completed = subprocess.run(command)
    return 0 if completed.returncode in {0, 1} else completed.returncode


def main() -> int:
    parser = argparse.ArgumentParser(description="Search text, using rg when available.")
    parser.add_argument("pattern")
    parser.add_argument("path", nargs="?", default=".")
    parser.add_argument("--context", type=int, default=0)
    parser.add_argument("--fixed", action="store_true")
    parser.add_argument("--ignore-case", action="store_true")
    parser.add_argument("--glob", action="append")
    parser.add_argument("--max-count", type=int, default=100)
    args = parser.parse_args()

    path = Path(args.path).expanduser()
    if not path.exists():
        print(f"error: path not found: {path}")
        return 1

    rg_code = run_rg(args, path)
    if rg_code is not None:
        return rg_code

    flags = re.IGNORECASE if args.ignore_case else 0
    pattern = re.escape(args.pattern) if args.fixed else args.pattern
    regex = re.compile(pattern, flags)
    printed = 0

    for file_path in iter_files(path):
        if printed >= args.max_count or looks_binary(file_path):
            continue
        try:
            lines = file_path.read_text(errors="replace").splitlines()
        except OSError:
            continue
        matches = [i for i, line in enumerate(lines) if regex.search(line)]
        if not matches:
            continue
        print(file_path)
        emitted: set[int] = set()
        for match in matches:
            for index in range(max(0, match - args.context), min(len(lines), match + args.context + 1)):
                if index in emitted:
                    continue
                emitted.add(index)
                marker = ":" if index == match else "-"
                print(f"{index + 1}{marker}{lines[index]}")
                printed += 1
                if printed >= args.max_count:
                    break
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
