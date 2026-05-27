#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
from collections import defaultdict, deque
from pathlib import Path
from typing import Any


def normalize(data: dict[str, Any]) -> tuple[dict[str, dict[str, Any]], list[dict[str, Any]], str, set[str]]:
    raw_nodes = data.get("nodes", [])
    if isinstance(raw_nodes, dict):
        nodes = {str(node_id): (value if isinstance(value, dict) else {"id": node_id}) for node_id, value in raw_nodes.items()}
    else:
        nodes = {str(node.get("id")): node for node in raw_nodes if isinstance(node, dict) and node.get("id")}

    edges = data.get("edges", [])
    if isinstance(edges, dict):
        expanded = []
        for source, targets in edges.items():
            for target in targets:
                expanded.append({"from": source, "to": target})
        edges = expanded

    start = str(data.get("start") or data.get("entrypoint") or "")
    terminal_value = data.get("terminal") or data.get("terminals") or data.get("end") or []
    if isinstance(terminal_value, str):
        terminals = {terminal_value}
    else:
        terminals = {str(value) for value in terminal_value}
    return nodes, [edge for edge in edges if isinstance(edge, dict)], start, terminals


def reachable_from(start: str, graph: dict[str, list[str]]) -> set[str]:
    seen = set()
    queue = deque([start])
    while queue:
        node = queue.popleft()
        if node in seen:
            continue
        seen.add(node)
        queue.extend(graph.get(node, []))
    return seen


def can_reach_terminal(node: str, reverse: dict[str, list[str]], terminals: set[str]) -> bool:
    reachable = set()
    queue = deque(terminals)
    while queue:
        current = queue.popleft()
        if current in reachable:
            continue
        reachable.add(current)
        queue.extend(reverse.get(current, []))
    return node in reachable


def main() -> int:
    parser = argparse.ArgumentParser(description="Validate an agent workflow graph JSON spec.")
    parser.add_argument("spec")
    parser.add_argument("--json", action="store_true")
    args = parser.parse_args()

    try:
        data = json.loads(Path(args.spec).expanduser().read_text(encoding="utf-8"))
    except Exception as exc:
        print(f"error: {exc}")
        return 1

    nodes, edges, start, terminals = normalize(data)
    findings = []

    def add(severity: str, message: str) -> None:
        findings.append({"severity": severity, "message": message})

    if not start:
        add("HIGH", "missing start node")
    elif start not in nodes:
        add("HIGH", f"start node not declared: {start}")
    if not terminals:
        add("HIGH", "missing terminal nodes")
    for terminal in terminals:
        if terminal not in nodes:
            add("HIGH", f"terminal node not declared: {terminal}")

    graph: dict[str, list[str]] = defaultdict(list)
    reverse: dict[str, list[str]] = defaultdict(list)
    for edge in edges:
        source = str(edge.get("from", ""))
        target = str(edge.get("to", ""))
        if source not in nodes:
            add("HIGH", f"edge source not declared: {source}")
        if target not in nodes:
            add("HIGH", f"edge target not declared: {target}")
        graph[source].append(target)
        reverse[target].append(source)
        if source == target and not (edge.get("condition") or edge.get("max_iterations")):
            add("MEDIUM", f"self-cycle lacks condition or max_iterations: {source}")

    if start in nodes:
        reachable = reachable_from(start, graph)
        for node_id in sorted(set(nodes) - reachable):
            add("MEDIUM", f"unreachable node: {node_id}")
        for node_id in sorted(reachable - terminals):
            if not graph.get(node_id):
                add("HIGH", f"non-terminal dead end: {node_id}")
            if terminals and not can_reach_terminal(node_id, reverse, terminals):
                add("HIGH", f"node cannot reach terminal: {node_id}")

    for node_id, node in nodes.items():
        if node.get("side_effect") and not node.get("requires_approval"):
            add("MEDIUM", f"side-effecting node lacks approval marker: {node_id}")

    if args.json:
        print(json.dumps({"findings": findings}, indent=2, sort_keys=True))
    else:
        for finding in findings:
            print(f"{finding['severity']} {finding['message']}")
        print(f"findings={len(findings)} nodes={len(nodes)} edges={len(edges)}")
    return 1 if any(finding["severity"] == "HIGH" for finding in findings) else 0


if __name__ == "__main__":
    raise SystemExit(main())
