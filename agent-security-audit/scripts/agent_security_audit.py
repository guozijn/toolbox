#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import os
import re
from pathlib import Path
from typing import Iterable


SKIP_DIRS = {".git", "__pycache__", ".venv", "venv", "node_modules", "dist", "build", "target"}
TEXT_EXTS = {".json", ".jsonl", ".yaml", ".yml", ".toml", ".md", ".txt", ".py", ".js", ".ts", ".tsx", ".jsx", ".sh"}

PATTERNS = [
    ("HIGH", "secret_like_value", re.compile(r"(?i)(api[_-]?key|secret|token|password)\s*[:=]\s*['\"]?[A-Za-z0-9_./+=-]{12,}")),
    ("HIGH", "prompt_injection_phrase", re.compile(r"(?i)(ignore (all )?(previous|prior) instructions|reveal (the )?system prompt|developer message|exfiltrate|send.*secret)")),
    ("HIGH", "command_execution", re.compile(r"\b(os\.system|subprocess\.(run|Popen|call)|child_process\.(exec|spawn)|eval\(|exec\()")),
    ("MEDIUM", "curl_pipe_shell", re.compile(r"(?i)curl .*[\|] *(sh|bash)|wget .*[\|] *(sh|bash)")),
    ("MEDIUM", "broad_file_access", re.compile(r"(?i)(read_file|write_file|delete_file|filesystem|file system|rm -rf|chmod 777)")),
    ("MEDIUM", "risky_mcp_config", re.compile(r"(?i)(mcpServers|autoapprove|auto_approve|alwaysAllow|command\"?\s*:|args\"?\s*:)")),
    ("LOW", "persistent_memory", re.compile(r"(?i)(memory|remember this|store forever|persistent context)")),
]


def iter_files(root: Path) -> Iterable[Path]:
    if root.is_file():
        yield root
        return
    for dirpath, dirnames, filenames in os.walk(root):
        dirnames[:] = [d for d in dirnames if d not in SKIP_DIRS and not d.startswith(".")]
        for filename in filenames:
            path = Path(dirpath) / filename
            if path.suffix.lower() in TEXT_EXTS:
                yield path


def scan_file(path: Path) -> list[dict]:
    findings = []
    try:
        lines = path.read_text(encoding="utf-8", errors="replace").splitlines()
    except OSError:
        return findings
    for line_no, line in enumerate(lines, 1):
        for severity, kind, pattern in PATTERNS:
            if pattern.search(line):
                snippet = line.strip()
                if kind == "secret_like_value":
                    snippet = re.sub(r"([:=]\s*['\"]?)[^'\"\s]+", r"\1[REDACTED]", snippet)
                findings.append({
                    "severity": severity,
                    "type": kind,
                    "path": path.as_posix(),
                    "line": line_no,
                    "snippet": snippet[:240],
                })
    return findings


def main() -> int:
    parser = argparse.ArgumentParser(description="Scan agent tool surfaces for common security risks.")
    parser.add_argument("paths", nargs="+")
    parser.add_argument("--json", action="store_true")
    parser.add_argument("--fail-on", choices=["LOW", "MEDIUM", "HIGH"], default="HIGH")
    args = parser.parse_args()

    severity_rank = {"LOW": 1, "MEDIUM": 2, "HIGH": 3}
    findings = []
    for value in args.paths:
        for path in iter_files(Path(value).expanduser()):
            findings.extend(scan_file(path))

    findings.sort(key=lambda item: (-severity_rank[item["severity"]], item["path"], item["line"]))
    if args.json:
        print(json.dumps({"findings": findings}, indent=2, sort_keys=True))
    else:
        for item in findings:
            print(f"{item['severity']} {item['type']} {item['path']}:{item['line']} {item['snippet']}")
        print(f"findings={len(findings)}")

    threshold = severity_rank[args.fail_on]
    return 1 if any(severity_rank[item["severity"]] >= threshold for item in findings) else 0


if __name__ == "__main__":
    raise SystemExit(main())
