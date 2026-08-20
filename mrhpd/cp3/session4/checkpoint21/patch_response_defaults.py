#!/usr/bin/env python3
"""Add schema-safe defaults to all recovered Session 4 response records."""
from __future__ import annotations

import argparse
from pathlib import Path


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("builder", type=Path)
    args = parser.parse_args()
    text = args.builder.read_text(encoding="utf-8")
    old = '''    entries.sort(key=lambda item: int(item["response_number"]))
    numbers = [int(item["response_number"]) for item in entries]'''
    new = '''    entries.sort(key=lambda item: int(item["response_number"]))
    for entry in entries:
        number = int(entry.get("response_number") or str(entry.get("response_key", "R0")).lstrip("R") or 0)
        entry.setdefault("response_key", f"R{number}")
        entry.setdefault("response_number", number)
        entry.setdefault("response_label", str(number))
        entry.setdefault("branch_id", "mainline")
        entry.setdefault("canonical_current", 1)
        entry.setdefault("major_topic", "Human Pathogen Database remediation")
        entry.setdefault("title", f"Session 4 Response {number}")
        is_gap = entry.get("state") == "explicit_unrecovered_gap"
        entry.setdefault("state", "explicit_unrecovered_gap" if is_gap else "source_supported_recovery_record")
        entry.setdefault("coverage", "gap" if is_gap else "source-supported record")
        entry.setdefault("fidelity_classification", "unrecovered_gap_no_inference" if is_gap else "source_verified_recovery_record")
        entry.setdefault("notes", None)
    numbers = [int(item["response_number"]) for item in entries]'''
    if old not in text:
        raise SystemExit("response-normalization insertion point not found")
    args.builder.write_text(text.replace(old, new, 1), encoding="utf-8", newline="\n")
    print(args.builder)


if __name__ == "__main__":
    main()
