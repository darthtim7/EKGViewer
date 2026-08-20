#!/usr/bin/env python3
from __future__ import annotations

from pathlib import Path

TARGET = Path(__file__).with_name("build_section5_checkpoint1.py")
OLD = 'str(package_root / "TOOLS" / "apply_checkpoint_recovery.py")'
NEW = 'str((package_root / "TOOLS" / "apply_checkpoint_recovery.py").resolve())'

text = TARGET.read_text(encoding="utf-8")
count = text.count(OLD)
if count != 1:
    raise SystemExit({"expected_occurrences": 1, "observed": count, "target": str(TARGET)})
TARGET.write_text(text.replace(OLD, NEW, 1), encoding="utf-8")
print({"status": "passed", "target": str(TARGET), "replacement_count": count})
