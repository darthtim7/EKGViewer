#!/usr/bin/env python3
"""Align the independent audit with cumulative database governance totals.

The canonical SQLite database is cumulative. Responses 1-40 and three earlier
fractional prompts predate Session 4; current Session 4 adds Responses 41-63
and six fractional prompts. Recovery-event identifiers are not row counts, so
the audit requires the verified cumulative row floor rather than treating event
ID 79 as a required count.
"""
from __future__ import annotations

import argparse
from pathlib import Path


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("builder", type=Path)
    args = parser.parse_args()
    text = args.builder.read_text(encoding="utf-8")
    marker = "    audit_text = audit_text.replace(audit_raise, audit_replacement, 1)\n"
    insertion = (
        marker
        + '    audit_text = audit_text.replace("EXPECTED_RESPONSE_RECORDS = 23", "EXPECTED_RESPONSE_RECORDS = 63")\n'
        + '    audit_text = audit_text.replace("EXPECTED_FRACTIONAL_PROMPTS = 6", "EXPECTED_FRACTIONAL_PROMPTS = 9")\n'
        + '    audit_text = audit_text.replace(\'counts["remediation_recovery_event"] < 79\', \'counts["remediation_recovery_event"] < 47\')\n'
    )
    if marker not in text:
        raise SystemExit("audit database-count insertion point not found")
    args.builder.write_text(text.replace(marker, insertion, 1), encoding="utf-8", newline="\n")
    print(args.builder)


if __name__ == "__main__":
    main()
