#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import os
import platform
import shutil
import socket
import sys
from pathlib import Path


DEFAULT_COMMANDS = [
    "git", "python3", "node", "npm", "uv", "pip", "rg", "jq", "curl",
    "brew", "docker", "gh", "code", "playwright", "npx",
]
SENSITIVE_ENV = [
    "OPENAI_API_KEY", "ANTHROPIC_API_KEY", "GITHUB_TOKEN", "GH_TOKEN",
    "AWS_ACCESS_KEY_ID", "GOOGLE_API_KEY", "BROWSER_USE_API_KEY",
]


def disk_info(path: Path) -> dict:
    usage = shutil.disk_usage(path)
    return {"total": usage.total, "used": usage.used, "free": usage.free}


def main() -> int:
    parser = argparse.ArgumentParser(description="Inspect local computer environment.")
    parser.add_argument("--command", action="append", help="Additional command to check.")
    parser.add_argument("--json", action="store_true")
    args = parser.parse_args()

    commands = sorted(set(DEFAULT_COMMANDS + (args.command or [])))
    result = {
        "platform": {
            "system": platform.system(),
            "release": platform.release(),
            "version": platform.version(),
            "machine": platform.machine(),
            "processor": platform.processor(),
        },
        "python": {
            "version": sys.version.split()[0],
            "executable": sys.executable,
        },
        "cwd": str(Path.cwd()),
        "home": str(Path.home()),
        "shell": os.getenv("SHELL"),
        "hostname": socket.gethostname(),
        "disk": disk_info(Path.cwd()),
        "commands": {command: shutil.which(command) for command in commands},
        "env_present": {name: bool(os.getenv(name)) for name in SENSITIVE_ENV},
    }

    if args.json:
        print(json.dumps(result, indent=2, sort_keys=True))
        return 0

    print(f"os={result['platform']['system']} {result['platform']['release']} {result['platform']['machine']}")
    print(f"python={result['python']['version']} executable={result['python']['executable']}")
    print(f"cwd={result['cwd']}")
    print(f"disk_free={result['disk']['free'] / (1024 ** 3):.1f}GB")
    print("commands:")
    for command, path in result["commands"].items():
        print(f"  {command}: {path or 'not found'}")
    print("env present:")
    for name, present in result["env_present"].items():
        print(f"  {name}: {'set' if present else 'unset'}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
