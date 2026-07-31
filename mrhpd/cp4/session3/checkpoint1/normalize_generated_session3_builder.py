#!/usr/bin/env python3
"""Normalize remaining generated Session 3 labels and recovery history."""
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

# Some archive names use a spaced all-uppercase RESPONSE token. Token-aware and
# title-case advancement does not cover that spelling.
for old, new in [("RESPONSE 66", "RESPONSE 69"), ("RESPONSE 67", "RESPONSE 70")]:
    if old in text:
        text = text.replace(old, new)
        applied.append(f"{old} -> {new}")

# Harden the reassembly invocation against future relative-working-directory
# changes without altering the exact input or output identities.
old_run = "result = subprocess.run([sys.executable, str(utility)], cwd=staging, text=True, capture_output=True, timeout=1200)"
new_run = "result = subprocess.run([sys.executable, str(utility.resolve())], cwd=staging, text=True, capture_output=True, timeout=1200)"
if old_run in text:
    text = replace_once(text, old_run, new_run, "absolute reassembly utility path")
    applied.append("absolute reassembly utility path")

text = text.replace("RECOVERY_EVENTS_125_130.json", "RECOVERY_EVENTS_125_133.json")

anchor = "        recovery_events=[inspection_event,capabilities_event,evidence_event,release_governance_event,app_event,package_event]\n"
if "V3-CP4-S3-REC-UPPERCASE-BASELINE-ARCHIVE-LABEL-CORRECTED" not in text:
    block = '''        marker_gate_event={
            "event_number":131,
            "event_code":"V3-CP4-S3-REC-GENERATED-MARKER-GATE-CORRECTED","occurred_at":NOW,
            "failed_step":"Generate the Session 3 Checkpoint 1 builder from the governed Session 2 source.",
            "exact_error_or_reason":"A validation marker required one literal recovery-name phrase that is assembled dynamically at runtime and therefore was not present in the generated source text.",
            "intact_artifacts":"Both exact Response 69 transport volumes, verification package, accepted predecessor, frozen Section 3 release, publication, editable assembly, database, workbook, and application source remained intact.",
            "recovery_action":"Removed only the invalid literal-source marker while retaining compile, response, session, checkpoint, baseline, restore-argument, and prohibited-marker gates.",
            "validation_result":"The generated Response 70 builder passed the remaining semantic marker checks.",
            "data_quality_effect":"None.",
            "next_checkpoint":"Integrate final-session release-governance controls.",
        }
        integration_anchor_event={
            "event_number":132,
            "event_code":"V3-CP4-S3-REC-GENERATED-INTEGRATION-ANCHORS-HARDENED","occurred_at":NOW,
            "failed_step":"Patch the generated clean-apply utility with release-governance and release-readiness checks.",
            "exact_error_or_reason":"A whole-block source substitution expected byte-identical formatting across generated versions and found no exact match despite the intended semantic anchors being present.",
            "intact_artifacts":"The generated builder had not executed; all baseline and project artifacts remained unchanged.",
            "recovery_action":"Replaced the brittle whole-block substitution with two small semantic-anchor insertions: one after the capability-registry gate and one after the existing application-audit gate.",
            "validation_result":"The integration patch applied all fourteen governed modifications and the generated builder compiled.",
            "data_quality_effect":"None.",
            "next_checkpoint":"Execute the Response 70 build from the exact Response 69 restore.",
        }
        uppercase_label_event={
            "event_number":133,
            "event_code":"V3-CP4-S3-REC-UPPERCASE-BASELINE-ARCHIVE-LABEL-CORRECTED","occurred_at":NOW,
            "failed_step":"Locate the exact Response 69 complete project snapshot inside the reconstructed complete restore.",
            "exact_error_or_reason":"One inherited all-uppercase spaced selector still searched for COMPLETE PROJECT THROUGH RESPONSE 66 after all title-case, compact-uppercase, lowercase, and identifier forms had advanced to Response 69.",
            "intact_artifacts":"The exact Response 69 restore reconstructed and verified successfully; failure occurred before the copied project tree was mutated.",
            "recovery_action":"Normalized the remaining spaced-uppercase response label to RESPONSE 69 and hardened the reassembly utility path resolution.",
            "validation_result":"The project-snapshot selector now targets the exact 182,223,875-byte Response 69 project archive with governed SHA-256 aa9d563357c375fd45b1e90e0e7f3465ebb855f1d4f997f817f46389ceb2b438.",
            "data_quality_effect":"None.",
            "next_checkpoint":"Continue the full database, workbook, application, publication, tracking, index, manifest, recovery, and clean-apply gates.",
        }
        recovery_events=[inspection_event,capabilities_event,evidence_event,release_governance_event,app_event,package_event,marker_gate_event,integration_anchor_event,uppercase_label_event]
'''
    text = replace_once(text, anchor, block, "Recovery Events 131-133")
    applied.append("Recovery Events 131-133")

required = [
    "*COMPLETE PROJECT THROUGH RESPONSE 69*.zip",
    "RECOVERY_EVENTS_125_133.json",
    "V3-CP4-S3-REC-GENERATED-MARKER-GATE-CORRECTED",
    "V3-CP4-S3-REC-GENERATED-INTEGRATION-ANCHORS-HARDENED",
    "V3-CP4-S3-REC-UPPERCASE-BASELINE-ARCHIVE-LABEL-CORRECTED",
]
missing = [marker for marker in required if marker not in text]
if missing:
    raise SystemExit({"missing_normalization_markers": missing})

prohibited = ["*COMPLETE PROJECT THROUGH RESPONSE 66*.zip", "RECOVERY_EVENTS_125_130.json"]
remaining = [marker for marker in prohibited if marker in text]
if remaining:
    raise SystemExit({"remaining_obsolete_markers": remaining})

if text != original:
    TARGET.write_text(text, encoding="utf-8")
print({"status": "passed", "target": TARGET.as_posix(), "applied": applied, "sha256": sha256_file(TARGET)})
