#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import os
import re
import subprocess
import sys
import time
from pathlib import Path
from typing import Any


def load_jsonl(path: Path) -> list[dict[str, Any]]:
    cases = []
    with path.open("r", encoding="utf-8") as handle:
        for line_no, line in enumerate(handle, 1):
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            try:
                case = json.loads(line)
            except json.JSONDecodeError as exc:
                raise ValueError(f"{path}:{line_no}: invalid JSON: {exc}") from exc
            if "id" not in case:
                case["id"] = f"line-{line_no}"
            cases.append(case)
    return cases


def evaluate_output(case: dict[str, Any], stdout: str, exit_code: int) -> tuple[bool, list[str]]:
    failures: list[str] = []
    expected_exit = case.get("expected_exit_code")
    if expected_exit is not None and exit_code != int(expected_exit):
        failures.append(f"exit_code expected {expected_exit}, got {exit_code}")

    for expected in case.get("expected_contains", []):
        if expected not in stdout:
            failures.append(f"missing text: {expected!r}")

    for forbidden in case.get("forbidden_contains", []):
        if forbidden in stdout:
            failures.append(f"forbidden text present: {forbidden!r}")

    for pattern in case.get("expected_regex", []):
        if not re.search(pattern, stdout, flags=re.MULTILINE):
            failures.append(f"regex did not match: {pattern!r}")

    return not failures, failures


def run_case(case: dict[str, Any], command: list[str], timeout: float, use_stdin: bool) -> dict[str, Any]:
    env = os.environ.copy()
    env["AGENT_EVAL_CASE_ID"] = str(case["id"])
    env["AGENT_EVAL_INPUT"] = str(case.get("input", ""))

    started = time.monotonic()
    try:
        completed = subprocess.run(
            command,
            input=str(case.get("input", "")) if use_stdin else None,
            text=True,
            capture_output=True,
            timeout=timeout,
            env=env,
        )
        timed_out = False
        exit_code = completed.returncode
        stdout = completed.stdout
        stderr = completed.stderr
    except subprocess.TimeoutExpired as exc:
        timed_out = True
        exit_code = 124
        stdout = exc.stdout or ""
        stderr = exc.stderr or f"timed out after {timeout}s"

    elapsed_ms = round((time.monotonic() - started) * 1000, 2)
    passed, failures = evaluate_output(case, stdout, exit_code)
    if timed_out:
        passed = False
        failures.append("timeout")

    return {
        "id": case["id"],
        "passed": passed,
        "failures": failures,
        "exit_code": exit_code,
        "elapsed_ms": elapsed_ms,
        "stdout": stdout,
        "stderr": stderr,
    }


def command_run(args: argparse.Namespace) -> int:
    cases = load_jsonl(Path(args.suite).expanduser())
    if not cases:
        print("error: no cases found", file=sys.stderr)
        return 1
    results: list[dict[str, Any]] = []
    for repeat in range(args.repeat):
        for case in cases:
            result = run_case(case, args.command, args.timeout, args.stdin)
            result["repeat"] = repeat + 1
            results.append(result)
            status = "PASS" if result["passed"] else "FAIL"
            print(f"{status} {result['id']} repeat={repeat + 1} {result['elapsed_ms']}ms")
            for failure in result["failures"]:
                print(f"  - {failure}")

    if args.output:
        output = Path(args.output).expanduser()
        output.parent.mkdir(parents=True, exist_ok=True)
        with output.open("w", encoding="utf-8") as handle:
            for result in results:
                handle.write(json.dumps(result, ensure_ascii=False) + "\n")

    passed = sum(1 for result in results if result["passed"])
    total = len(results)
    print(f"\npass_rate={passed}/{total} ({passed / total:.1%})")
    return 0 if passed == total else 1


def command_summarize(args: argparse.Namespace) -> int:
    results = load_jsonl(Path(args.results).expanduser())
    total = len(results)
    passed = sum(1 for result in results if result.get("passed"))
    by_case: dict[str, list[dict[str, Any]]] = {}
    for result in results:
        by_case.setdefault(str(result.get("id", "unknown")), []).append(result)

    print(f"total={total} passed={passed} failed={total - passed} pass_rate={passed / total:.1%}" if total else "total=0")
    for case_id, case_results in sorted(by_case.items()):
        case_passed = sum(1 for result in case_results if result.get("passed"))
        print(f"{case_id}: {case_passed}/{len(case_results)}")
        for result in case_results:
            if not result.get("passed"):
                for failure in result.get("failures", []):
                    print(f"  - {failure}")
                break
    return 0 if passed == total else 1


def command_init(args: argparse.Namespace) -> int:
    output = Path(args.output).expanduser()
    if output.exists() and not args.force:
        print(f"error: file exists: {output}")
        return 1
    output.parent.mkdir(parents=True, exist_ok=True)
    sample = {
        "id": "case-001",
        "input": "Say done and include order #123.",
        "expected_contains": ["done"],
        "expected_regex": ["order #[0-9]+"],
        "forbidden_contains": ["ERROR"],
        "expected_exit_code": 0,
    }
    output.write_text(json.dumps(sample) + "\n", encoding="utf-8")
    print(output)
    return 0


def parse_run(argv: list[str]) -> argparse.Namespace:
    if "--" not in argv:
        raise ValueError("run requires '--' before the command to evaluate")
    separator = argv.index("--")
    option_argv = argv[:separator]
    command = argv[separator + 1:]
    if not command:
        raise ValueError("missing command after '--'")

    parser = argparse.ArgumentParser(prog="agent_eval.py run")
    parser.add_argument("suite")
    parser.add_argument("--repeat", type=int, default=1)
    parser.add_argument("--timeout", type=float, default=60)
    parser.add_argument("--stdin", action="store_true", help="Send case input to command stdin.")
    parser.add_argument("-o", "--output")
    args = parser.parse_args(option_argv)
    args.command = command
    return args


def main(argv: list[str] | None = None) -> int:
    argv = list(sys.argv[1:] if argv is None else argv)
    if argv[:1] == ["run"]:
        try:
            args = parse_run(argv[1:])
        except ValueError as exc:
            print(f"error: {exc}", file=sys.stderr)
            return 1
        return command_run(args)

    parser = argparse.ArgumentParser(description="Run lightweight regression evals for agents.")
    subparsers = parser.add_subparsers(dest="command_name", required=True)

    summarize_parser = subparsers.add_parser("summarize")
    summarize_parser.add_argument("results")
    summarize_parser.set_defaults(func=command_summarize)

    init_parser = subparsers.add_parser("init")
    init_parser.add_argument("output")
    init_parser.add_argument("--force", action="store_true")
    init_parser.set_defaults(func=command_init)

    args = parser.parse_args(argv)
    return args.func(args)


if __name__ == "__main__":
    raise SystemExit(main())
