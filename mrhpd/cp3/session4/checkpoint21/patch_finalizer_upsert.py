#!/usr/bin/env python3
"""Correct the recovered finalizer's overly broad *_id exclusion rule."""
from __future__ import annotations

import argparse
from pathlib import Path


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("builder", type=Path)
    args = parser.parse_args()
    text = args.builder.read_text(encoding="utf-8")
    old = '''    repair_wrapped_python(finalizer)
    subprocess.run([sys.executable, "-m", "py_compile", str(finalizer)], check=True)'''
    new = '''    repair_wrapped_python(finalizer)
    repaired_finalizer = finalizer.read_text(encoding="utf-8")
    buggy_filter = '    values = {k: v for k, v in row.items() if k in columns and not k.endswith("_id")}'
    fixed_filter = (
        '    primary_keys = {info[1] for info in con.execute(f\\'PRAGMA table_info("{table}")\\') if info[5]}\\n'
        '    values = {k: v for k, v in row.items() if k in columns and k not in primary_keys}'
    )
    if buggy_filter not in repaired_finalizer:
        raise RuntimeError("Recovered finalizer upsert filter was not found")
    finalizer.write_text(repaired_finalizer.replace(buggy_filter, fixed_filter, 1), encoding="utf-8", newline="\\n")
    subprocess.run([sys.executable, "-m", "py_compile", str(finalizer)], check=True)'''
    if old not in text:
        raise SystemExit("finalizer post-repair insertion point not found")
    args.builder.write_text(text.replace(old, new, 1), encoding="utf-8", newline="\n")
    print(args.builder)


if __name__ == "__main__":
    main()
