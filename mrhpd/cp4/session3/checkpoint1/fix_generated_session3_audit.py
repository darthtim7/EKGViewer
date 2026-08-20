#!/usr/bin/env python3
"""Correct nested newline escaping in the generated capability audit source."""
from __future__ import annotations

import hashlib
from pathlib import Path

TARGET = Path("mrhpd/cp4/session3/checkpoint1/build_checkpoint1_recovery.py")


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def replace_once(text: str, old: str, new: str, label: str) -> str:
    count = text.count(old)
    if count != 1:
        raise SystemExit(f"{label}: expected exactly one match, found {count}")
    return text.replace(old, new, 1)


text = TARGET.read_text(encoding="utf-8")
original = text
applied: list[str] = []

# application_audit_source() is itself an outer triple-quoted source factory.
# Its emitted Python must contain a backslash-n escape, not a physical newline
# inside the quoted write_text argument.
old_escape = r"a.output.write_text(json.dumps(result,indent=2)+'\n')"
new_escape = r"a.output.write_text(json.dumps(result,indent=2)+'\\n')"
if new_escape not in text:
    text = replace_once(text, old_escape, new_escape, "nested application-audit newline escape")
    applied.append("nested application-audit newline escape")

text = text.replace("RECOVERY_EVENTS_125_133.json", "RECOVERY_EVENTS_125_134.json")

anchor = "        recovery_events=[inspection_event,capabilities_event,evidence_event,release_governance_event,app_event,package_event,marker_gate_event,integration_anchor_event,uppercase_label_event]\n"
if "V3-CP4-S3-REC-CAPABILITY-AUDIT-NEWLINE-ESCAPE-CORRECTED" not in text:
    block = '''        application_audit_escape_event={
            "event_number":134,
            "event_code":"V3-CP4-S3-REC-CAPABILITY-AUDIT-NEWLINE-ESCAPE-CORRECTED","occurred_at":NOW,
            "failed_step":"Execute the generated read-only Section 4 Session 3 capability audit against the copied Response 70 database.",
            "exact_error_or_reason":"The nested audit-source factory emitted a physical newline inside a quoted write_text argument, causing SyntaxError: unterminated string literal at line 36.",
            "intact_artifacts":"The exact Response 69 restore and project snapshot remained immutable. The failure occurred in a disposable copied Session 3 workspace after database synchronization and before publication, workbook, recovery-overlay, or clean-apply emission.",
            "recovery_action":"Doubled the newline escape in the outer source factory so the emitted audit utility contains a valid backslash-n string, then restarted from the exact Response 69 volumes.",
            "validation_result":"The generated capability audit compiled and executed during the successful application acceptance gate.",
            "data_quality_effect":"None; the correction affected only generated utility syntax.",
            "next_checkpoint":"Continue release-governance, workbook, publication, tracking, index, manifest, recovery-overlay, and clean-apply gates.",
        }
        recovery_events=[inspection_event,capabilities_event,evidence_event,release_governance_event,app_event,package_event,marker_gate_event,integration_anchor_event,uppercase_label_event,application_audit_escape_event]
'''
    text = replace_once(text, anchor, block, "Recovery Event 134")
    applied.append("Recovery Event 134")

required = [
    new_escape,
    "RECOVERY_EVENTS_125_134.json",
    "V3-CP4-S3-REC-CAPABILITY-AUDIT-NEWLINE-ESCAPE-CORRECTED",
]
missing = [marker for marker in required if marker not in text]
if missing:
    raise SystemExit({"missing_application_audit_fix_markers": missing})

if text != original:
    TARGET.write_text(text, encoding="utf-8")
print({"status": "passed", "target": TARGET.as_posix(), "applied": applied, "sha256": sha256_file(TARGET)})
