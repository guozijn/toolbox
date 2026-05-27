#!/usr/bin/env python3
from __future__ import annotations

import argparse
import csv
from pathlib import Path


def main() -> int:
    parser = argparse.ArgumentParser(description="Preview CSV or TSV rows.")
    parser.add_argument("input")
    parser.add_argument("--rows", type=int, default=5)
    parser.add_argument("--delimiter")
    args = parser.parse_args()

    path = Path(args.input).expanduser()
    if not path.exists():
        print(f"error: input file not found: {path}")
        return 1

    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        sample = handle.read(4096)
        handle.seek(0)
        delimiter = args.delimiter
        if delimiter is None:
            try:
                delimiter = csv.Sniffer().sniff(sample).delimiter
            except csv.Error:
                delimiter = "\t" if path.suffix.lower() == ".tsv" else ","
        reader = csv.reader(handle, delimiter=delimiter)
        rows = []
        for index, row in enumerate(reader):
            rows.append(row)
            if index >= args.rows:
                break

    if not rows:
        print("empty file")
        return 0

    widths = [0] * max(len(row) for row in rows)
    for row in rows:
        for index, value in enumerate(row):
            widths[index] = max(widths[index], min(len(value), 40))

    for row in rows:
        cells = []
        for index, value in enumerate(row):
            display = value if len(value) <= 40 else value[:37] + "..."
            cells.append(display.ljust(widths[index]))
        print(" | ".join(cells).rstrip())
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
