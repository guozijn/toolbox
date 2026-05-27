#!/usr/bin/env python3
from __future__ import annotations

import argparse
import shutil
import subprocess
from datetime import datetime
from pathlib import Path


def default_output() -> Path:
    return Path.cwd() / f"screenshot-{datetime.now().strftime('%Y%m%d-%H%M%S')}.png"


def command_for(output: Path, interactive: bool) -> list[str] | None:
    if shutil.which("screencapture"):
        command = ["screencapture"]
        if interactive:
            command.append("-i")
        command.append(str(output))
        return command
    if shutil.which("gnome-screenshot"):
        command = ["gnome-screenshot", "-f", str(output)]
        if interactive:
            command.insert(1, "-a")
        return command
    if shutil.which("import"):
        return ["import", str(output)] if interactive else ["import", "-window", "root", str(output)]
    return None


def main() -> int:
    parser = argparse.ArgumentParser(description="Capture a desktop screenshot.")
    parser.add_argument("-o", "--output")
    parser.add_argument("--interactive", action="store_true")
    args = parser.parse_args()

    output = Path(args.output).expanduser() if args.output else default_output()
    output.parent.mkdir(parents=True, exist_ok=True)
    command = command_for(output, args.interactive)
    if not command:
        print("error: no supported screenshot command found")
        return 1
    completed = subprocess.run(command, text=True, capture_output=True)
    if completed.returncode != 0:
        print(completed.stderr.strip() or "error: screenshot command failed")
        return completed.returncode
    print(output)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
