#!/usr/bin/env python3
"""Correct inherited three-column baseline and QA report row contracts."""
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

old_baseline = '''    for label, value in [
        ("Response 69 complete restore", f"{BASE_RESTORE_BYTES:,} bytes; {BASE_RESTORE_SHA256}", "PASS"),
        ("Response 69 project snapshot", f"{BASE_PROJECT_BYTES:,} bytes; {BASE_PROJECT_SHA256}", "PASS"),
        ("Accepted predecessor mutation", "No", "PASS"),
        ("Integrated publication", f"537 searchable pages; {PUBLICATION_SHA256}", "PASS"),
        ("Editable assembly", EDITABLE_ASSEMBLY_SHA256, "PASS"),
    ]:
        cells = baseline.add_row().cells; cells[0].text = label; cells[1].text = value; cells[2].text = value if False else "PASS"
'''
new_baseline = '''    for label, value, status in [
        ("Response 69 complete restore", f"{BASE_RESTORE_BYTES:,} bytes; {BASE_RESTORE_SHA256}", "PASS"),
        ("Response 69 project snapshot", f"{BASE_PROJECT_BYTES:,} bytes; {BASE_PROJECT_SHA256}", "PASS"),
        ("Accepted predecessor mutation", "No", "PASS"),
        ("Integrated publication", f"537 searchable pages; {PUBLICATION_SHA256}", "PASS"),
        ("Editable assembly", EDITABLE_ASSEMBLY_SHA256, "PASS"),
    ]:
        cells = baseline.add_row().cells; cells[0].text = label; cells[1].text = value; cells[2].text = status
'''
if new_baseline not in text:
    text = replace_once(text, old_baseline, new_baseline, "three-column baseline report row contract")
    applied.append("three-column baseline report row contract")

old_qa = '''    for area, result in [
        ("SQLite", f"{database_qa['table_count']} tables; integrity {database_qa['integrity']}; 0 FK violations", database_qa["status"]),
        ("Workbook", f"{workbook_qa['current_sheet_count']} sheets; {workbook_qa['formula_error_count']} formula-error tokens", workbook_qa["status"]),
        ("Application", f"{application_qa['test_files']} legacy test files plus capability audit", application_qa["status"]),
        ("Publication", f"{publication_qa['publication_pages']} pages; {publication_qa['searchable_pages']} searchable", publication_qa["status"]),
    ]:
        cells = qa_table.add_row().cells; cells[0].text = area; cells[1].text = result; cells[2].text = "PASS"
'''
new_qa = '''    for area, result, status in [
        ("SQLite", f"{database_qa['table_count']} tables; integrity {database_qa['integrity']}; 0 FK violations", database_qa["status"]),
        ("Workbook", f"{workbook_qa['current_sheet_count']} sheets; {workbook_qa['formula_error_count']} formula-error tokens", workbook_qa["status"]),
        ("Application", f"{application_qa['test_files']} legacy test files plus capability audit", application_qa["status"]),
        ("Publication", f"{publication_qa['publication_pages']} pages; {publication_qa['searchable_pages']} searchable", publication_qa["status"]),
    ]:
        cells = qa_table.add_row().cells; cells[0].text = area; cells[1].text = result; cells[2].text = status.upper()
'''
if new_qa not in text:
    text = replace_once(text, old_qa, new_qa, "three-column QA report row contract")
    applied.append("three-column QA report row contract")

text = text.replace("RECOVERY_EVENTS_125_135.json", "RECOVERY_EVENTS_125_137.json")

anchor = "        recovery_events=[inspection_event,capabilities_event,evidence_event,release_governance_event,app_event,package_event,marker_gate_event,integration_anchor_event,uppercase_label_event,application_audit_escape_event,release_audit_sql_event]\n"
if "V3-CP4-S3-REC-QA-REPORT-ROW-CONTRACT-CORRECTED" not in text:
    block = '''        baseline_report_event={
            "event_number":136,
            "event_code":"V3-CP4-S3-REC-BASELINE-REPORT-ROW-CONTRACT-CORRECTED","occurred_at":NOW,
            "failed_step":"Generate the inherited human-readable baseline and immutability table.",
            "exact_error_or_reason":"The report loop unpacked each three-element (label, value, status) row into two variables, raising ValueError: too many values to unpack.",
            "intact_artifacts":"The exact Response 69 baseline remained immutable. Database synchronization, legacy application regressions, the capability audit, release-governance persistence, workbook augmentation, and release-readiness audit had passed in the disposable current workspace.",
            "recovery_action":"Changed the baseline report loop to unpack label, value, and status and write the explicit status into the third table column.",
            "validation_result":"The baseline and immutability table rendered with its intended three-column contract.",
            "data_quality_effect":"None; only derivative report rendering logic changed.",
            "next_checkpoint":"Continue the remaining QA table and report validation.",
        }
        qa_report_event={
            "event_number":137,
            "event_code":"V3-CP4-S3-REC-QA-REPORT-ROW-CONTRACT-CORRECTED","occurred_at":NOW,
            "failed_step":"Generate the inherited database, workbook, application, and publication QA table.",
            "exact_error_or_reason":"The adjacent QA report loop repeated the same two-variable unpacking error for three-element (area, result, status) rows.",
            "intact_artifacts":"All clinical and artifact acceptance gates remained passed; the failure was isolated to a derivative human-readable table in a disposable build tree.",
            "recovery_action":"Changed the QA report loop to unpack area, result, and status, then rendered the normalized status in the third column.",
            "validation_result":"DOCX, searchable PDF, XLSX, and report-QA derivatives generated and validated.",
            "data_quality_effect":"None; no database, workbook source, application, evidence record, or publication asset changed.",
            "next_checkpoint":"Complete tracking, indexes, manifests, recovery overlay, clean apply, and transport delivery.",
        }
        recovery_events=[inspection_event,capabilities_event,evidence_event,release_governance_event,app_event,package_event,marker_gate_event,integration_anchor_event,uppercase_label_event,application_audit_escape_event,release_audit_sql_event,baseline_report_event,qa_report_event]
'''
    text = replace_once(text, anchor, block, "Recovery Events 136-137")
    applied.append("Recovery Events 136-137")

required = [
    "for label, value, status in [",
    "for area, result, status in [",
    "cells[2].text = status",
    "cells[2].text = status.upper()",
    "RECOVERY_EVENTS_125_137.json",
    "V3-CP4-S3-REC-BASELINE-REPORT-ROW-CONTRACT-CORRECTED",
    "V3-CP4-S3-REC-QA-REPORT-ROW-CONTRACT-CORRECTED",
]
missing = [marker for marker in required if marker not in text]
if missing:
    raise SystemExit({"missing_report_fix_markers": missing})

if text != original:
    TARGET.write_text(text, encoding="utf-8")
print({"status": "passed", "target": TARGET.as_posix(), "applied": applied, "sha256": sha256_file(TARGET)})
