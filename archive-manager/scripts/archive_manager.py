#!/usr/bin/env python3
from __future__ import annotations

import argparse
import tarfile
import zipfile
from pathlib import Path


def is_zip(path: Path) -> bool:
    return path.suffix.lower() == ".zip"


def safe_target(base: Path, name: str) -> Path:
    target = (base / name).resolve()
    if not str(target).startswith(str(base.resolve())):
        raise ValueError(f"unsafe archive member path: {name}")
    return target


def command_list(args: argparse.Namespace) -> int:
    archive = Path(args.archive).expanduser()
    if is_zip(archive):
        with zipfile.ZipFile(archive) as zf:
            for info in zf.infolist():
                print(f"{info.file_size}\t{info.filename}")
    else:
        with tarfile.open(archive) as tf:
            for member in tf.getmembers():
                print(f"{member.size}\t{member.name}")
    return 0


def command_extract(args: argparse.Namespace) -> int:
    archive = Path(args.archive).expanduser()
    destination = Path(args.destination).expanduser().resolve()
    destination.mkdir(parents=True, exist_ok=True)
    if is_zip(archive):
        with zipfile.ZipFile(archive) as zf:
            for member in zf.infolist():
                safe_target(destination, member.filename)
            zf.extractall(destination)
    else:
        with tarfile.open(archive) as tf:
            for member in tf.getmembers():
                safe_target(destination, member.name)
            tf.extractall(destination)
    print(destination)
    return 0


def command_create(args: argparse.Namespace) -> int:
    output = Path(args.output).expanduser()
    output.parent.mkdir(parents=True, exist_ok=True)
    paths = [Path(value).expanduser() for value in args.paths]
    if is_zip(output):
        with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED) as zf:
            for path in paths:
                if path.is_dir():
                    for child in path.rglob("*"):
                        if child.is_file():
                            zf.write(child, child.as_posix())
                else:
                    zf.write(path, path.name)
    else:
        mode = "w:gz" if output.suffixes[-2:] in [[".tar", ".gz"], [".tgz"]] or output.suffix == ".tgz" else "w"
        with tarfile.open(output, mode) as tf:
            for path in paths:
                tf.add(path, arcname=path.name)
    print(output)
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(description="List, create, and safely extract archives.")
    subparsers = parser.add_subparsers(dest="command", required=True)
    list_parser = subparsers.add_parser("list")
    list_parser.add_argument("archive")
    list_parser.set_defaults(func=command_list)
    extract_parser = subparsers.add_parser("extract")
    extract_parser.add_argument("archive")
    extract_parser.add_argument("destination")
    extract_parser.set_defaults(func=command_extract)
    create_parser = subparsers.add_parser("create")
    create_parser.add_argument("output")
    create_parser.add_argument("paths", nargs="+")
    create_parser.set_defaults(func=command_create)
    args = parser.parse_args()
    try:
        return args.func(args)
    except Exception as exc:
        print(f"error: {exc}")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
