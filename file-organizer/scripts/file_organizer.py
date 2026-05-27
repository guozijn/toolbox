#!/usr/bin/env python3
from __future__ import annotations

import argparse
import hashlib
import json
import shutil
from collections import defaultdict
from datetime import datetime
from pathlib import Path


MEDIA_TYPES = {
    "images": {".jpg", ".jpeg", ".png", ".gif", ".webp", ".heic", ".svg"},
    "documents": {".pdf", ".doc", ".docx", ".ppt", ".pptx", ".xls", ".xlsx", ".txt", ".md"},
    "archives": {".zip", ".tar", ".gz", ".tgz", ".rar", ".7z"},
    "audio": {".mp3", ".wav", ".m4a", ".flac"},
    "video": {".mp4", ".mov", ".mkv", ".avi"},
    "code": {".py", ".js", ".ts", ".tsx", ".jsx", ".go", ".rs", ".java", ".c", ".cpp"},
}


def category(path: Path, mode: str) -> str:
    if mode == "extension":
        return (path.suffix.lower().lstrip(".") or "no-extension")
    if mode == "date":
        return datetime.fromtimestamp(path.stat().st_mtime).strftime("%Y-%m")
    if mode == "type":
        suffix = path.suffix.lower()
        for name, suffixes in MEDIA_TYPES.items():
            if suffix in suffixes:
                return name
        return "other"
    raise ValueError(f"unknown mode: {mode}")


def plan_moves(source: Path, destination: Path, mode: str) -> list[dict]:
    moves = []
    for path in sorted(source.iterdir()):
        if not path.is_file():
            continue
        target_dir = destination / category(path, mode)
        target = target_dir / path.name
        counter = 1
        while target.exists() and target.resolve() != path.resolve():
            target = target_dir / f"{path.stem}-{counter}{path.suffix}"
            counter += 1
        if target.resolve() != path.resolve():
            moves.append({"from": path.as_posix(), "to": target.as_posix()})
    return moves


def command_plan(args: argparse.Namespace) -> int:
    source = Path(args.source).expanduser().resolve()
    destination = Path(args.destination).expanduser().resolve() if args.destination else source
    moves = plan_moves(source, destination, args.mode)
    if args.json:
        print(json.dumps({"moves": moves}, indent=2))
    else:
        for move in moves:
            print(f"{move['from']} -> {move['to']}")
        print(f"moves={len(moves)}")
    if args.apply:
        for move in moves:
            target = Path(move["to"])
            target.parent.mkdir(parents=True, exist_ok=True)
            shutil.move(move["from"], move["to"])
        print("applied=true")
    return 0


def file_hash(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def command_duplicates(args: argparse.Namespace) -> int:
    root = Path(args.path).expanduser()
    groups: dict[tuple[int, str], list[str]] = defaultdict(list)
    for path in root.rglob("*") if args.recursive else root.iterdir():
        if path.is_file():
            groups[(path.stat().st_size, file_hash(path))].append(path.as_posix())
    duplicates = [paths for paths in groups.values() if len(paths) > 1]
    if args.json:
        print(json.dumps({"duplicates": duplicates}, indent=2))
    else:
        for paths in duplicates:
            print("duplicate group:")
            for path in paths:
                print(f"  {path}")
        print(f"groups={len(duplicates)}")
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(description="Plan safe local file organization.")
    subparsers = parser.add_subparsers(dest="command", required=True)
    plan = subparsers.add_parser("plan")
    plan.add_argument("source")
    plan.add_argument("--destination")
    plan.add_argument("--mode", choices=["extension", "date", "type"], default="type")
    plan.add_argument("--apply", action="store_true")
    plan.add_argument("--json", action="store_true")
    plan.set_defaults(func=command_plan)
    duplicates = subparsers.add_parser("duplicates")
    duplicates.add_argument("path")
    duplicates.add_argument("--recursive", action="store_true")
    duplicates.add_argument("--json", action="store_true")
    duplicates.set_defaults(func=command_duplicates)
    args = parser.parse_args()
    try:
        return args.func(args)
    except Exception as exc:
        print(f"error: {exc}")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
