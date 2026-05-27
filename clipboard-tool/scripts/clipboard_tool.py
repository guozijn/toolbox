#!/usr/bin/env python3
from __future__ import annotations

import argparse
import shutil
import subprocess
import sys


def copy_command() -> list[str] | None:
    for command in (["pbcopy"], ["wl-copy"], ["xclip", "-selection", "clipboard"], ["xsel", "--clipboard", "--input"]):
        if shutil.which(command[0]):
            return command
    return None


def paste_command() -> list[str] | None:
    for command in (["pbpaste"], ["wl-paste"], ["xclip", "-selection", "clipboard", "-o"], ["xsel", "--clipboard", "--output"]):
        if shutil.which(command[0]):
            return command
    return None


def main() -> int:
    parser = argparse.ArgumentParser(description="Read or write the system clipboard.")
    subparsers = parser.add_subparsers(dest="command", required=True)
    copy_parser = subparsers.add_parser("copy")
    copy_parser.add_argument("text", nargs="?")
    subparsers.add_parser("paste")
    args = parser.parse_args()

    if args.command == "copy":
        command = copy_command()
        if not command:
            print("error: no clipboard copy command found")
            return 1
        text = args.text if args.text is not None else sys.stdin.read()
        subprocess.run(command, input=text, text=True, check=True)
        print(f"copied={len(text)} chars")
        return 0

    command = paste_command()
    if not command:
        print("error: no clipboard paste command found")
        return 1
    completed = subprocess.run(command, text=True, capture_output=True, check=True)
    print(completed.stdout, end="")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
