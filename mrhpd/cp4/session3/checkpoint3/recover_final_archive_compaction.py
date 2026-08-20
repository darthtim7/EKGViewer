#!/usr/bin/env python3
"""Recover the final archive by broadening equivalence-verified compaction.

The generated Response 72 builder already verifies table/row-count and
worksheet supersets before deleting any older database or workbook snapshot.
The first final build retained older session-complete snapshots because its
filename filter covered only Session 3 Checkpoints 1 and 2. This utility removes
that filename restriction while preserving the existing superset gates.
"""
from __future__ import annotations

import hashlib
from pathlib import Path

TARGET = Path("mrhpd/cp4/session3/checkpoint3/build_session3_complete_restore.py")


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

old_filter = '    checkpoint_pattern = re.compile(r"Session 3 of 3 Checkpoint (1|2) of 3", re.I)\n'
new_filter = '    checkpoint_pattern = re.compile(r".*", re.I)\n'
if new_filter not in text:
    text = replace_once(text, old_filter, new_filter, "all prior snapshot candidates under existing superset gates")
    applied.append("all prior snapshot candidates under existing superset gates")

old_name = "Remediation Section 4 of 5 Session 3 of 3 COMPLETE PROJECT THROUGH RESPONSE 69"
new_name = "Remediation Section 4 of 5 Session 3 of 3 COMPLETE PROJECT THROUGH RESPONSE 72"
if new_name not in text:
    text = replace_once(text, old_name, new_name, "Response 72 complete-project filename")
    applied.append("Response 72 complete-project filename")

last_event = '''    {
        "event_number": 167,
        "event_code": "V3-CP4-S3-REC-SECTION4-TRANSPORT-AND-HANDOFF-PREPARED",
        "occurred_at": NOW,
        "failed_step": "None; connector-compatible transport and Section 5 handoff prepared.",
        "exact_error_or_reason": "The final restore must remain downloadable and recoverable through persistent custody.",
        "intact_artifacts": "Complete restore, controls, manifest, checksums, reassembly utility and final reports.",
        "recovery_action": "Split the complete restore into the minimum two sub-100-MiB transport volumes, verified every wrapper, and prepared final controls and Section 5 handoff.",
        "validation_result": "Transport and handoff controls passed.",
        "data_quality_effect": "None.",
        "next_checkpoint": "Remediation Section 5 of 5.",
    },
'''
expanded_events = last_event + '''    {
        "event_number": 168,
        "event_code": "V3-CP4-S3-REC-FINAL-ARCHIVE-PRIOR-SNAPSHOT-COMPACTION-EXPANDED",
        "occurred_at": NOW,
        "failed_step": "Build the final Response 72 complete-project archive below the mandatory 180 MiB ceiling.",
        "exact_error_or_reason": "The first otherwise-valid final project archive was 197,169,011 bytes because the inherited compaction filename filter covered Session 3 Checkpoints 1 and 2 but retained older session-complete database and workbook snapshots.",
        "intact_artifacts": "The exact Response 69 baseline, Response 71 candidate, canonical Response 72 database and workbook, accepted predecessor, frozen Section 3 release, publication, editable assembly, application, sources, tracking, recovery, reports, and QA remained intact.",
        "recovery_action": "Removed the filename restriction while retaining the existing read-only table/row-count and worksheet-superset comparisons; only older snapshots proven to be strict subsets of the current canonical artifacts may be removed.",
        "validation_result": "Pending regenerated archive, manifest, clean-extraction, restore, and transport verification.",
        "data_quality_effect": "No canonical row, worksheet, clinical claim, source record, or immutable publication asset may be removed by this correction.",
        "next_checkpoint": "Regenerate and independently verify the Section 4 complete restore.",
    },
'''
if "V3-CP4-S3-REC-FINAL-ARCHIVE-PRIOR-SNAPSHOT-COMPACTION-EXPANDED" not in text:
    text = replace_once(text, last_event, expanded_events, "Recovery Event 168")
    applied.append("Recovery Event 168")

if "RECOVERY_EVENTS_159_168.json" not in text:
    text = text.replace("RECOVERY_EVENTS_159_167.json", "RECOVERY_EVENTS_159_168.json")
    applied.append("Recovery Events 159-168 filename")

if text != original:
    TARGET.write_text(text, encoding="utf-8")

required = [
    'checkpoint_pattern = re.compile(r".*", re.I)',
    "COMPLETE PROJECT THROUGH RESPONSE 72",
    "V3-CP4-S3-REC-FINAL-ARCHIVE-PRIOR-SNAPSHOT-COMPACTION-EXPANDED",
    "RECOVERY_EVENTS_159_168.json",
]
current = TARGET.read_text(encoding="utf-8")
missing = [marker for marker in required if marker not in current]
if missing:
    raise SystemExit({"missing_compaction_recovery_markers": missing})

print({"status":"passed","path":TARGET.as_posix(),"applied":applied,"sha256":sha256_file(TARGET)})
