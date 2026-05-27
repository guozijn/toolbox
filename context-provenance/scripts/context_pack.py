#!/usr/bin/env python3
from __future__ import annotations

import argparse
import fnmatch
import hashlib
import json
import os
import re
from pathlib import Path
from typing import Iterable


SKIP_DIRS = {".git", "__pycache__", ".venv", "venv", "node_modules", "dist", "build", "target"}


def looks_binary(path: Path) -> bool:
    try:
        with path.open("rb") as handle:
            return b"\0" in handle.read(2048)
    except OSError:
        return True


def iter_paths(paths: list[Path], includes: list[str], excludes: list[str]) -> Iterable[Path]:
    for root in paths:
        if root.is_file():
            candidates = [root]
        else:
            candidates = []
            for dirpath, dirnames, filenames in os.walk(root):
                dirnames[:] = [d for d in dirnames if d not in SKIP_DIRS and not d.startswith(".")]
                for filename in filenames:
                    candidates.append(Path(dirpath) / filename)
        for path in candidates:
            rel = path.as_posix()
            if includes and not any(fnmatch.fnmatch(rel, pattern) for pattern in includes):
                continue
            if any(fnmatch.fnmatch(rel, pattern) for pattern in excludes):
                continue
            if path.is_file() and not looks_binary(path):
                yield path


def chunk_file(path: Path, max_chars: int, overlap_lines: int) -> Iterable[dict]:
    text = path.read_text(encoding="utf-8", errors="replace")
    lines = text.splitlines()
    index = 0
    chunk_no = 1
    while index < len(lines):
        current: list[str] = []
        start = index
        size = 0
        while index < len(lines) and (not current or size + len(lines[index]) + 1 <= max_chars):
            current.append(lines[index])
            size += len(lines[index]) + 1
            index += 1
        chunk_text = "\n".join(current)
        digest = hashlib.sha256(chunk_text.encode("utf-8")).hexdigest()
        yield {
            "id": f"{path.as_posix()}:{start + 1}-{index}",
            "path": path.as_posix(),
            "chunk": chunk_no,
            "start_line": start + 1,
            "end_line": index,
            "sha256": digest,
            "token_estimate": max(1, len(chunk_text) // 4),
            "text": chunk_text,
        }
        chunk_no += 1
        if overlap_lines and index < len(lines):
            index = max(index - overlap_lines, start + 1)


def command_pack(args: argparse.Namespace) -> int:
    paths = [Path(value).expanduser() for value in args.paths]
    chunks = []
    for path in iter_paths(paths, args.include or [], args.exclude or []):
        try:
            chunks.extend(chunk_file(path, args.max_chars, args.overlap_lines))
        except OSError as exc:
            print(f"warning: skipped {path}: {exc}")

    output = Path(args.output).expanduser()
    output.parent.mkdir(parents=True, exist_ok=True)
    with output.open("w", encoding="utf-8") as handle:
        for chunk in chunks:
            handle.write(json.dumps(chunk, ensure_ascii=False) + "\n")
    print(f"chunks={len(chunks)} output={output}")
    return 0


def score(query_terms: set[str], chunk: dict) -> int:
    text = chunk.get("text", "").lower()
    return sum(len(re.findall(re.escape(term), text)) for term in query_terms)


def command_select(args: argparse.Namespace) -> int:
    query_terms = {term.lower() for term in re.findall(r"[\w.-]+", args.query) if len(term) > 1}
    chunks = []
    with Path(args.pack).expanduser().open("r", encoding="utf-8") as handle:
        for line in handle:
            if line.strip():
                chunk = json.loads(line)
                chunk["_score"] = score(query_terms, chunk)
                chunks.append(chunk)
    selected = [chunk for chunk in sorted(chunks, key=lambda c: (-c["_score"], c["id"])) if chunk["_score"] > 0]
    for chunk in selected[:args.limit]:
        print(json.dumps(chunk, ensure_ascii=False))
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(description="Build and query provenance-preserving context packs.")
    subparsers = parser.add_subparsers(dest="command", required=True)

    pack = subparsers.add_parser("pack")
    pack.add_argument("paths", nargs="+")
    pack.add_argument("-o", "--output", required=True)
    pack.add_argument("--include", action="append")
    pack.add_argument("--exclude", action="append")
    pack.add_argument("--max-chars", type=int, default=4000)
    pack.add_argument("--overlap-lines", type=int, default=0)
    pack.set_defaults(func=command_pack)

    select = subparsers.add_parser("select")
    select.add_argument("pack")
    select.add_argument("--query", required=True)
    select.add_argument("--limit", type=int, default=5)
    select.set_defaults(func=command_select)

    args = parser.parse_args()
    return args.func(args)


if __name__ == "__main__":
    raise SystemExit(main())
