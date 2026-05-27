#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import shutil
import subprocess


def run(command: list[str]) -> tuple[int, str, str]:
    completed = subprocess.run(command, text=True, capture_output=True)
    return completed.returncode, completed.stdout, completed.stderr


def lsof_ports(port: int | None = None) -> list[dict]:
    command = ["lsof", "-nP", "-iTCP", "-sTCP:LISTEN"]
    if port is not None:
        command = ["lsof", "-nP", f"-iTCP:{port}", "-sTCP:LISTEN"]
    code, stdout, stderr = run(command)
    if code not in {0, 1}:
        raise RuntimeError(stderr.strip() or "lsof failed")
    rows = []
    for line in stdout.splitlines()[1:]:
        parts = line.split()
        if len(parts) < 9:
            continue
        name = parts[0]
        pid = parts[1]
        user = parts[2]
        address = parts[-2] if parts[-1].startswith("(") else parts[-1]
        rows.append({"command": name, "pid": pid, "user": user, "address": address})
    return rows


def fallback_ports() -> list[dict]:
    for command in (["ss", "-ltnp"], ["netstat", "-anv"]):
        if shutil.which(command[0]):
            code, stdout, stderr = run(command)
            if code == 0:
                return [{"raw": line} for line in stdout.splitlines() if "LISTEN" in line]
            raise RuntimeError(stderr.strip() or f"{command[0]} failed")
    raise RuntimeError("no supported port inspection command found")


def main() -> int:
    parser = argparse.ArgumentParser(description="Inspect local listening ports and processes.")
    parser.add_argument("--port", type=int)
    parser.add_argument("--json", action="store_true")
    args = parser.parse_args()

    try:
        if shutil.which("lsof"):
            rows = lsof_ports(args.port)
        else:
            rows = fallback_ports()
    except Exception as exc:
        print(f"error: {exc}")
        return 1

    if args.json:
        print(json.dumps({"listeners": rows}, indent=2, sort_keys=True))
        return 0
    for row in rows:
        if "raw" in row:
            print(row["raw"])
        else:
            print(f"{row['address']} pid={row['pid']} user={row['user']} command={row['command']}")
    print(f"listeners={len(rows)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
