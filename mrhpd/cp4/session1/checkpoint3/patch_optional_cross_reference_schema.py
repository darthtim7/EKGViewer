#!/usr/bin/env python3
"""Patch optional cross-reference current-state column handling.

The accepted database's publication_cross_reference table contains the governed
current rows but does not expose an is_current column. The Session 1 builder must
inspect the schema before selecting a current-state predicate rather than assume a
column that is absent.
"""
from pathlib import Path

path = Path(__file__).with_name("build_session1_complete_restore.py")
text = path.read_text(encoding="utf-8")
old = '''        cross_refs = con.execute("SELECT COUNT(*) FROM publication_cross_reference WHERE COALESCE(is_current,1)=1").fetchone()[0]
'''
new = '''        cross_ref_columns = table_columns(con, "publication_cross_reference")
        cross_ref_sql = (
            "SELECT COUNT(*) FROM publication_cross_reference WHERE COALESCE(is_current,1)=1"
            if "is_current" in cross_ref_columns
            else "SELECT COUNT(*) FROM publication_cross_reference"
        )
        cross_refs = con.execute(cross_ref_sql).fetchone()[0]
'''
if old not in text:
    raise SystemExit("database cross-reference schema patch anchor not found")
text = text.replace(old, new, 1)
old2 = '''    con = sqlite3.connect(db)
    try:
        direct_checks = {
'''
new2 = '''    con = sqlite3.connect(db)
    try:
        cross_ref_columns = table_columns(con, "publication_cross_reference")
        cross_ref_sql = (
            "SELECT COUNT(*) FROM publication_cross_reference WHERE COALESCE(is_current,1)=1"
            if "is_current" in cross_ref_columns
            else "SELECT COUNT(*) FROM publication_cross_reference"
        )
        current_cross_ref_count = con.execute(cross_ref_sql).fetchone()[0]
        direct_checks = {
'''
if old2 not in text:
    raise SystemExit("application direct-check schema patch anchor not found")
text = text.replace(old2, new2, 1)
old3 = '''            "cross_references": con.execute("SELECT COUNT(*) FROM publication_cross_reference WHERE COALESCE(is_current,1)=1").fetchone()[0] == 12,
'''
new3 = '''            "cross_references": current_cross_ref_count == 12,
'''
if old3 not in text:
    raise SystemExit("application cross-reference count patch anchor not found")
text = text.replace(old3, new3, 1)
event = '''RECOVERY_EVENTS.append({
    "event_number": 99,
    "event_code": "V3-CP4-S1-REC-OPTIONAL-CROSS-REFERENCE-CURRENT-FLAG-HANDLED",
    "occurred_at": NOW,
    "failed_step": "Count current publication cross-references using an assumed is_current column.",
    "exact_error_or_reason": "sqlite3.OperationalError: no such column: is_current",
    "intact_artifacts": "The reconstructed Checkpoint 2 state, current database copy, workbook, application, publication, accepted predecessor and frozen Section 3 release remained intact.",
    "recovery_action": "Made the query schema-aware: use the current-state predicate when the column exists, otherwise count the governed current cross-reference table directly.",
    "validation_result": "The final database and application gates still require exactly twelve current cross-references.",
    "data_quality_effect": "None; no record was changed to resolve the schema mismatch.",
    "next_checkpoint": "Resume final Session 1 synchronization and complete restore generation.",
})

'''
anchor = "\nNET_PROMPT = (\n"
if anchor not in text:
    raise SystemExit("recovery-event insertion anchor not found")
text = text.replace(anchor, "\n" + event + "NET_PROMPT = (\n", 1)
path.write_text(text, encoding="utf-8")
print({"optional_cross_reference_schema_patch": "applied", "builder": str(path)})
