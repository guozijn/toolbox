#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import mimetypes
import os
from pathlib import Path
from typing import Any, Iterable


SKIP_DIRS = {
    ".git", ".hg", ".svn", "__pycache__", ".pytest_cache", ".mypy_cache",
    ".ruff_cache", ".venv", "venv", "node_modules", "dist", "build", "target",
}


def human_size(size: int) -> str:
    units = ["B", "KB", "MB", "GB", "TB"]
    value = float(size)
    for unit in units:
        if value < 1024 or unit == units[-1]:
            return f"{int(value)} {unit}" if unit == "B" else f"{value:.1f} {unit}"
        value /= 1024
    return f"{size} B"


def iter_files(root: Path, include_hidden: bool, max_depth: int) -> Iterable[Path]:
    root = root.resolve()
    for dirpath, dirnames, filenames in os.walk(root):
        current = Path(dirpath)
        rel = current.relative_to(root)
        depth = 0 if rel == Path(".") else len(rel.parts)
        dirnames[:] = sorted(
            d for d in dirnames
            if d not in SKIP_DIRS and (include_hidden or not d.startswith("."))
        )
        if depth >= max_depth:
            dirnames[:] = []
        for filename in sorted(filenames):
            if not include_hidden and filename.startswith("."):
                continue
            yield current / filename


def main() -> int:
    parser = argparse.ArgumentParser(description="Print a compact repository map.")
    parser.add_argument("path", nargs="?", default=".")
    parser.add_argument("--depth", type=int, default=3)
    parser.add_argument("--max-files", type=int, default=200)
    parser.add_argument("--include-hidden", action="store_true")
    parser.add_argument("--json", action="store_true")
    args = parser.parse_args()

    root = Path(args.path).expanduser().resolve()
    if not root.is_dir():
        print(f"error: directory not found: {root}")
        return 1

    files: list[dict[str, Any]] = []
    ext_counts: dict[str, int] = {}
    total_bytes = 0
    for path in iter_files(root, args.include_hidden, args.depth):
        try:
            stat = path.stat()
        except OSError:
            continue
        rel = path.relative_to(root).as_posix()
        ext = path.suffix.lower() or "[no extension]"
        ext_counts[ext] = ext_counts.get(ext, 0) + 1
        total_bytes += stat.st_size
        if len(files) < args.max_files:
            files.append({
                "path": rel,
                "size": stat.st_size,
                "type": mimetypes.guess_type(path.name)[0] or "unknown",
            })

    result = {
        "root": str(root),
        "files_shown": len(files),
        "total_files_seen": sum(ext_counts.values()),
        "total_size": total_bytes,
        "extensions": dict(sorted(ext_counts.items(), key=lambda item: (-item[1], item[0]))),
        "files": files,
    }
    if args.json:
        print(json.dumps(result, indent=2, sort_keys=True))
        return 0

    print(f"Root: {root}")
    print(f"Files seen: {result['total_files_seen']} ({human_size(total_bytes)})")
    if ext_counts:
        print("Top extensions: " + ", ".join(
            f"{ext}: {count}" for ext, count in list(result["extensions"].items())[:10]
        ))
    print("")
    for item in files:
        print(f"{item['path']}\t{human_size(item['size'])}")
    if result["total_files_seen"] > len(files):
        print(f"... {result['total_files_seen'] - len(files)} more files not shown")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
