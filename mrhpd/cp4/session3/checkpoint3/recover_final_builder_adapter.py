#!/usr/bin/env python3
"""Recover small mechanical defects in the final builder adapter.

The adapter and generated builder are disposable execution sources. This
utility never changes accepted clinical artifacts or frozen publication files.
"""
from __future__ import annotations

import hashlib
from pathlib import Path

PATH = Path("mrhpd/cp4/session3/checkpoint3/prepare_session3_complete_builder.py")


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


text = PATH.read_text(encoding="utf-8")
original = text
replacements = [
    ("section3_checkpoint3_governance.py", "session3_checkpoint3_governance.py", "governance filename marker"),
]
applied: list[str] = []
for old, new, label in replacements:
    if old in text:
        text = text.replace(old, new)
        applied.append(label)
    elif new not in text:
        raise SystemExit(f"Recovery target missing: {label}")
if text != original:
    PATH.write_text(text, encoding="utf-8")
print({"status":"passed","path":PATH.as_posix(),"applied":applied,"sha256":sha256_file(PATH)})
