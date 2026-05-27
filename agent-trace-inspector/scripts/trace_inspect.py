#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
from collections import Counter
from pathlib import Path
from typing import Any, Iterable


SPAN_KEYS = ("spans", "children", "events")


def load_records(path: Path) -> list[Any]:
    text = path.read_text(encoding="utf-8")
    stripped = text.strip()
    if not stripped:
        return []
    if stripped[0] in "[{":
        return [json.loads(stripped)]
    records = []
    for line_no, line in enumerate(text.splitlines(), 1):
        line = line.strip()
        if not line:
            continue
        try:
            records.append(json.loads(line))
        except json.JSONDecodeError as exc:
            raise ValueError(f"{path}:{line_no}: invalid JSON: {exc}") from exc
    return records


def walk(value: Any) -> Iterable[dict[str, Any]]:
    if isinstance(value, list):
        for item in value:
            yield from walk(item)
    elif isinstance(value, dict):
        if looks_like_span(value):
            yield value
        for key in SPAN_KEYS:
            if key in value:
                yield from walk(value[key])


def looks_like_span(value: dict[str, Any]) -> bool:
    keys = set(value)
    return bool(keys & {"name", "span_name", "span_id", "trace_id", "attributes", "status", "duration_ms"})


def attrs(span: dict[str, Any]) -> dict[str, Any]:
    attributes = span.get("attributes")
    return attributes if isinstance(attributes, dict) else {}


def operation(span: dict[str, Any]) -> str:
    attributes = attrs(span)
    for key in ("gen_ai.operation.name", "openinference.span.kind", "operation", "kind", "type"):
        if attributes.get(key):
            return str(attributes[key])
        if span.get(key):
            return str(span[key])
    name = str(span.get("name") or span.get("span_name") or "unknown").lower()
    if "tool" in name:
        return "execute_tool"
    if "retriev" in name:
        return "retrieval"
    if "chat" in name or "llm" in name:
        return "chat"
    return "unknown"


def tool_name(span: dict[str, Any]) -> str | None:
    attributes = attrs(span)
    for key in ("gen_ai.tool.name", "tool.name", "tool_name", "name"):
        value = attributes.get(key) or span.get(key)
        if value and operation(span) in {"execute_tool", "tool"}:
            return str(value)
    return None


def duration_ms(span: dict[str, Any]) -> float:
    for key in ("duration_ms", "latency_ms", "elapsed_ms"):
        if span.get(key) is not None:
            try:
                return float(span[key])
            except (TypeError, ValueError):
                return 0.0
    return 0.0


def is_error(span: dict[str, Any]) -> bool:
    status = span.get("status")
    attributes = attrs(span)
    if isinstance(status, dict):
        status = status.get("code") or status.get("status_code")
    status_text = str(status or attributes.get("error.type") or "").lower()
    return any(token in status_text for token in ("error", "fail", "exception", "timeout"))


def summarize(paths: list[Path]) -> dict[str, Any]:
    spans: list[dict[str, Any]] = []
    for path in paths:
        for record in load_records(path):
            spans.extend(walk(record))

    operations = Counter(operation(span) for span in spans)
    tools = Counter(name for span in spans for name in [tool_name(span)] if name)
    errors = [span for span in spans if is_error(span)]
    durations = Counter()
    for span in spans:
        durations[operation(span)] += duration_ms(span)

    missing_ids = sum(1 for span in spans if not (span.get("span_id") or span.get("id")))
    return {
        "span_count": len(spans),
        "operation_counts": dict(operations),
        "tool_counts": dict(tools),
        "error_count": len(errors),
        "missing_span_ids": missing_ids,
        "duration_ms_by_operation": dict(durations),
        "errors": [
            {
                "name": span.get("name") or span.get("span_name"),
                "operation": operation(span),
                "status": span.get("status") or attrs(span).get("error.type"),
            }
            for span in errors[:20]
        ],
    }


def main() -> int:
    parser = argparse.ArgumentParser(description="Inspect agent trace exports.")
    parser.add_argument("paths", nargs="+")
    parser.add_argument("--json", action="store_true")
    args = parser.parse_args()

    try:
        result = summarize([Path(path).expanduser() for path in args.paths])
    except Exception as exc:
        print(f"error: {exc}")
        return 1

    if args.json:
        print(json.dumps(result, indent=2, sort_keys=True))
        return 0

    print(f"spans={result['span_count']} errors={result['error_count']} missing_span_ids={result['missing_span_ids']}")
    print("operations:")
    for name, count in sorted(result["operation_counts"].items()):
        duration = result["duration_ms_by_operation"].get(name, 0)
        print(f"  {name}: {count} spans, {duration:.1f}ms")
    if result["tool_counts"]:
        print("tools:")
        for name, count in sorted(result["tool_counts"].items()):
            print(f"  {name}: {count}")
    if result["errors"]:
        print("errors:")
        for error in result["errors"]:
            print(f"  {error['operation']} {error['name']}: {error['status']}")
    return 0 if result["error_count"] == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())
