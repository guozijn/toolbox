#!/usr/bin/env python3
from __future__ import annotations

import argparse
import importlib.util
import json
import os
import shutil
from pathlib import Path


PY_PACKAGES = ["playwright", "browser_use"]
COMMANDS = ["node", "npm", "npx", "playwright", "google-chrome", "chromium", "chromium-browser"]
ENV_VARS = ["OPENAI_API_KEY", "ANTHROPIC_API_KEY", "BROWSER_USE_API_KEY"]


def check() -> dict:
    packages = {name: importlib.util.find_spec(name) is not None for name in PY_PACKAGES}
    commands = {name: shutil.which(name) for name in COMMANDS}
    env = {name: bool(os.getenv(name)) for name in ENV_VARS}
    chrome_paths = [
        Path("/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"),
        Path("/Applications/Chromium.app/Contents/MacOS/Chromium"),
    ]
    browsers = {path.as_posix(): path.exists() for path in chrome_paths}
    return {"python_packages": packages, "commands": commands, "env": env, "browser_paths": browsers}


def main() -> int:
    parser = argparse.ArgumentParser(description="Check browser-agent environment readiness.")
    parser.add_argument("--json", action="store_true")
    args = parser.parse_args()

    result = check()
    if args.json:
        print(json.dumps(result, indent=2, sort_keys=True))
        return 0

    print("python packages:")
    for name, present in result["python_packages"].items():
        print(f"  {name}: {'yes' if present else 'no'}")
    print("commands:")
    for name, path in result["commands"].items():
        print(f"  {name}: {path or 'not found'}")
    print("env:")
    for name, present in result["env"].items():
        print(f"  {name}: {'set' if present else 'unset'}")
    print("browser paths:")
    for path, present in result["browser_paths"].items():
        print(f"  {path}: {'yes' if present else 'no'}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
