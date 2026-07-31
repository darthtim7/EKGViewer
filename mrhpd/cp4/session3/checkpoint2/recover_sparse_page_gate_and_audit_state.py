#!/usr/bin/env python3
"""Recover Checkpoint 2 sparse-page classification and audit-state sequencing.

The corrections apply only to disposable execution-source copies. No accepted
or frozen clinical artifact is edited in place.
"""
from __future__ import annotations

import hashlib
from pathlib import Path

GOVERNANCE = Path("mrhpd/cp4/session3/checkpoint2/session3_checkpoint2_governance.py")
BUILDER = Path("mrhpd/cp4/session3/checkpoint2/build_checkpoint2_recovery.py")


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


text = GOVERNANCE.read_text(encoding="utf-8")
original = text
applied: list[str] = []

old_page_gate = '''            status = "passed" if text.strip() and rect.width > 500 and rect.height > 700 and nonwhite_ratio > 0.0005 else "failed"
            notes = "searchable_rendered_page" if status == "passed" else "possible_blank_or_geometry_anomaly"
'''
new_page_gate = '''            render_valid = bool(samples) and pix.width > 100 and pix.height > 100
            status = "passed" if text.strip() and rect.width > 500 and rect.height > 700 and render_valid else "failed"
            if status != "passed":
                notes = "possible_blank_or_geometry_anomaly"
            elif nonwhite_ratio > 0.0005:
                notes = "searchable_rendered_page"
            else:
                notes = "searchable_sparse_or_low_ink_page"
'''
if new_page_gate not in text:
    text = replace_once(text, old_page_gate, new_page_gate, "searchable sparse-page classification")
    applied.append("searchable sparse-page classification")

old_state_sequence = '''    workbook_qa = augment_workbook(workbook, base_workbook_qa, source_rows, page_rows, graphics_rows, drift_rows, risks, summaries)
    application_audit = run_application_audit(audit, audit_output, db, workbook, publication, application)
    database_qa = update_checkpoint_state(db, workbook_qa["status"], application_audit["status"], generated_at)
'''
new_state_sequence = '''    workbook_qa = augment_workbook(workbook, base_workbook_qa, source_rows, page_rows, graphics_rows, drift_rows, risks, summaries)
    update_checkpoint_state(db, workbook_qa["status"], "audit_pending", generated_at)
    application_audit = run_application_audit(audit, audit_output, db, workbook, publication, application)
    database_qa = update_checkpoint_state(db, workbook_qa["status"], application_audit["status"], generated_at)
'''
if new_state_sequence not in text:
    text = replace_once(text, old_state_sequence, new_state_sequence, "release-candidate audit-state sequencing")
    applied.append("release-candidate audit-state sequencing")

last_event = '''        ("V3-CP4-S3-REC-CHECKPOINT2-RELEASE-CANDIDATE-PREPARED", "Prepare the Section 4 release candidate for independent Checkpoint 3 verification.", "Checkpoint 2 may resolve and freeze the candidate but may not self-declare final release.", "Closed or explicitly deferred each controlled risk, rebuilt reports/QA/index/manifest/recovery surfaces, and handed a cleanly applicable candidate to Checkpoint 3."),
'''
recovery_events = '''        ("V3-CP4-S3-REC-CHECKPOINT2-RELEASE-CANDIDATE-PREPARED", "Prepare the Section 4 release candidate for independent Checkpoint 3 verification.", "Checkpoint 2 may resolve and freeze the candidate but may not self-declare final release.", "Closed or explicitly deferred each controlled risk, rebuilt reports/QA/index/manifest/recovery surfaces, and handed a cleanly applicable candidate to Checkpoint 3."),
        ("V3-CP4-S3-REC-PUBLICATION-SPARSE-PAGE-CLASSIFICATION-CORRECTED", "Run the 537-page publication render audit.", "The initial disposable audit confirmed 537 searchable pages but incorrectly failed 39 intentionally sparse pages because a low-ink heuristic was treated as a blank-page gate.", "Separated searchable text and valid render geometry from the descriptive ink-density metric; sparse pages now remain explicitly labeled while genuine empty or malformed pages still fail."),
        ("V3-CP4-S3-REC-RELEASE-CANDIDATE-AUDIT-STATE-SEQUENCING-CORRECTED", "Run the independent read-only release-candidate audit.", "Code review identified that the audit expected checkpoint_complete before the state transition occurred.", "Marked the prepared candidate checkpoint_complete with audit_pending status before the read-only audit, then finalized application_status only after that audit passed."),
'''
if "V3-CP4-S3-REC-PUBLICATION-SPARSE-PAGE-CLASSIFICATION-CORRECTED" not in text:
    text = replace_once(text, last_event, recovery_events, "Recovery Events 155-156")
    applied.append("Recovery Events 155-156")

if text != original:
    GOVERNANCE.write_text(text, encoding="utf-8")

builder_text = BUILDER.read_text(encoding="utf-8")
builder_original = builder_text
for old_name in ("RECOVERY_EVENTS_139_154.json", "RECOVERY_EVENTS_139_155.json"):
    builder_text = builder_text.replace(old_name, "RECOVERY_EVENTS_139_156.json")
if "RECOVERY_EVENTS_139_156.json" not in builder_text:
    raise SystemExit("Checkpoint 2 recovery-event filename was not updated")
if builder_text != builder_original:
    BUILDER.write_text(builder_text, encoding="utf-8")
    applied.append("recovery-event filename 139-156")

required = [
    "searchable_sparse_or_low_ink_page",
    'update_checkpoint_state(db, workbook_qa["status"], "audit_pending", generated_at)',
    "V3-CP4-S3-REC-PUBLICATION-SPARSE-PAGE-CLASSIFICATION-CORRECTED",
    "V3-CP4-S3-REC-RELEASE-CANDIDATE-AUDIT-STATE-SEQUENCING-CORRECTED",
    "RECOVERY_EVENTS_139_156.json",
]
combined = GOVERNANCE.read_text(encoding="utf-8") + "\n" + BUILDER.read_text(encoding="utf-8")
missing = [marker for marker in required if marker not in combined]
if missing:
    raise SystemExit({"missing_recovery_markers": missing})

print({
    "status": "passed",
    "applied": applied,
    "governance_sha256": sha256_file(GOVERNANCE),
    "builder_sha256": sha256_file(BUILDER),
})
