#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import re
from pathlib import Path
from typing import Any


NAME_RE = re.compile(r"^[a-zA-Z][a-zA-Z0-9_.-]{1,63}$")
POISON_RE = re.compile(r"(?i)(ignore previous|system prompt|developer message|secret|exfiltrate|do not tell|hidden instruction)")
SIDE_EFFECT_RE = re.compile(r"(?i)\b(write|delete|send|create|update|execute|run|transfer|purchase|email)\b")


def load_json(path: Path) -> Any:
    with path.open("r", encoding="utf-8") as handle:
        return json.load(handle)


def find_tools(data: Any) -> list[dict[str, Any]]:
    if isinstance(data, dict):
        if isinstance(data.get("tools"), list):
            return [tool for tool in data["tools"] if isinstance(tool, dict)]
        if {"name", "description"} <= set(data):
            return [data]
        tools = []
        for value in data.values():
            tools.extend(find_tools(value))
        return tools
    if isinstance(data, list):
        tools = []
        for item in data:
            tools.extend(find_tools(item))
        return tools
    return []


def lint_tool(tool: dict[str, Any], index: int) -> list[dict[str, Any]]:
    findings = []
    name = str(tool.get("name", ""))
    description = str(tool.get("description", ""))
    schema = tool.get("inputSchema") or tool.get("input_schema") or tool.get("parameters")

    def add(severity: str, message: str) -> None:
        findings.append({"severity": severity, "tool": name or f"#{index}", "message": message})

    if not name:
        add("HIGH", "missing tool name")
    elif not NAME_RE.match(name):
        add("MEDIUM", "tool name should be stable ASCII identifier-like text")

    if len(description.strip()) < 20:
        add("MEDIUM", "description is too short to guide reliable tool use")
    if POISON_RE.search(description):
        add("HIGH", "description contains prompt-injection or secret-seeking language")
    if SIDE_EFFECT_RE.search(name + " " + description):
        if not any(key in tool for key in ("sideEffects", "side_effects", "requiresApproval", "requires_approval")):
            add("MEDIUM", "side-effecting tool should declare side effects or approval requirements")

    if not isinstance(schema, dict):
        add("HIGH", "missing input schema object")
    else:
        schema_type = schema.get("type")
        properties = schema.get("properties")
        if schema_type and schema_type != "object":
            add("MEDIUM", "input schema should usually be an object")
        if properties is not None and not isinstance(properties, dict):
            add("HIGH", "schema properties must be an object")
        if isinstance(properties, dict):
            for prop_name, prop_schema in properties.items():
                if not isinstance(prop_schema, dict) or not prop_schema.get("type"):
                    add("LOW", f"property {prop_name!r} should declare a type")
    return findings


def main() -> int:
    parser = argparse.ArgumentParser(description="Lint MCP-style tool schemas.")
    parser.add_argument("paths", nargs="+")
    parser.add_argument("--json", action="store_true")
    args = parser.parse_args()

    all_findings = []
    for value in args.paths:
        path = Path(value).expanduser()
        try:
            tools = find_tools(load_json(path))
        except Exception as exc:
            all_findings.append({"severity": "HIGH", "tool": path.as_posix(), "message": f"could not parse JSON: {exc}"})
            continue
        if not tools:
            all_findings.append({"severity": "MEDIUM", "tool": path.as_posix(), "message": "no tools found"})
        for index, tool in enumerate(tools, 1):
            for finding in lint_tool(tool, index):
                finding["path"] = path.as_posix()
                all_findings.append(finding)

    if args.json:
        print(json.dumps({"findings": all_findings}, indent=2, sort_keys=True))
    else:
        for finding in all_findings:
            print(f"{finding['severity']} {finding.get('path', '')} {finding['tool']}: {finding['message']}")
        print(f"findings={len(all_findings)}")
    return 1 if any(item["severity"] == "HIGH" for item in all_findings) else 0


if __name__ == "__main__":
    raise SystemExit(main())
