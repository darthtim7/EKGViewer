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


def insert_once(text: str, anchor: str, payload: str, label: str) -> str:
    count = text.count(anchor)
    if count != 1:
        raise SystemExit(f"{label}: expected exactly one anchor, found {count}")
    return text.replace(anchor, payload + anchor, 1)


text = PATH.read_text(encoding="utf-8")
original = text
applied: list[str] = []

replacements = [
    ("section3_checkpoint3_governance.py", "session3_checkpoint3_governance.py", "governance filename marker"),
]
for old, new, label in replacements:
    if old in text:
        text = text.replace(old, new)
        applied.append(label)
    elif new not in text:
        raise SystemExit(f"Recovery target missing: {label}")

archive_marker = "# MRHPD Session 3 final archive compaction recovery"
archive_anchor = "# Canonical output names and state keys.\n"
archive_block = '''# MRHPD Session 3 final archive compaction recovery
# The inherited compaction routine retains read-only table/row-count and
# worksheet-superset proofs. Broaden only its filename candidate filter so
# older session-complete snapshots can be considered under those same gates.
text = text.replace(
    'checkpoint_pattern = re.compile(r"Session 3 of 3 Checkpoint (1|2) of 3", re.I)',
    'checkpoint_pattern = re.compile(r".*", re.I)',
)
text = text.replace(
    "Remediation Section 4 of 5 Session 3 of 3 COMPLETE PROJECT THROUGH RESPONSE 69",
    "Remediation Section 4 of 5 Session 3 of 3 COMPLETE PROJECT THROUGH RESPONSE 72",
)
if "V3-CP4-S3-REC-FINAL-ARCHIVE-PRIOR-SNAPSHOT-COMPACTION-EXPANDED" not in text:
    _archive_event = r"""    {
        "event_number": 168,
        "event_code": "V3-CP4-S3-REC-FINAL-ARCHIVE-PRIOR-SNAPSHOT-COMPACTION-EXPANDED",
        "occurred_at": NOW,
        "failed_step": "Build the final Response 72 complete-project archive below the mandatory 180 MiB ceiling.",
        "exact_error_or_reason": "The first otherwise-valid final project archive was 197,169,011 bytes because the inherited filename filter covered Session 3 Checkpoints 1 and 2 but retained older session-complete database and workbook snapshots.",
        "intact_artifacts": "The exact Response 69 baseline, Response 71 candidate, canonical Response 72 database and workbook, accepted predecessor, frozen Section 3 release, publication, editable assembly, application, sources, tracking, recovery, reports, and QA remained intact.",
        "recovery_action": "Removed the filename restriction while retaining the existing read-only table/row-count and worksheet-superset comparisons; only older snapshots proven to be subsets of the current canonical artifacts may be removed.",
        "validation_result": "Regenerated archive, manifest, clean-extraction, restore, and transport verification required.",
        "data_quality_effect": "No canonical row, worksheet, clinical claim, source record, or immutable publication asset may be removed by this correction.",
        "next_checkpoint": "Regenerate and independently verify the Section 4 complete restore.",
    },
"""
    _archive_event_anchor = "]\\n\\nNET_PROMPT = ("
    if _archive_event_anchor not in text:
        raise SystemExit("Response 72 Recovery Event 168 insertion anchor missing")
    text = text.replace(_archive_event_anchor, _archive_event + _archive_event_anchor, 1)
text = text.replace("RECOVERY_EVENTS_159_167.json", "RECOVERY_EVENTS_159_168.json")

'''
if archive_marker not in text:
    text = insert_once(text, archive_anchor, archive_block, "final archive compaction recovery")
    applied.append("final archive compaction recovery")

required_old = '    "RECOVERY_EVENTS_159_167.json",\n'
required_new = '    "RECOVERY_EVENTS_159_168.json",\n    "V3-CP4-S3-REC-FINAL-ARCHIVE-PRIOR-SNAPSHOT-COMPACTION-EXPANDED",\n'
if required_old in text:
    text = text.replace(required_old, required_new, 1)
    applied.append("Recovery Events 159-168 required markers")
elif required_new not in text:
    raise SystemExit("Recovery Events 159-168 required-marker target missing")

if text != original:
    PATH.write_text(text, encoding="utf-8")
print({"status":"passed","path":PATH.as_posix(),"applied":applied,"sha256":sha256_file(PATH)})
