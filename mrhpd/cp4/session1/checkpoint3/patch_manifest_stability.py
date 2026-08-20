#!/usr/bin/env python3
"""Prevent a late QA rewrite from invalidating the final project manifest."""
from pathlib import Path

path = Path(__file__).with_name("build_session1_complete_restore.py")
text = path.read_text(encoding="utf-8")
old = '''        # Rebuild indexes once more so final QA/session records are included.
        index_qa=build_indexes_and_manifest(project); session_qa["indexes"]=index_qa
        json_write(qa_dir/"SESSION_1_COMPLETE_QA.json",session_qa)
        project_archive,project_archive_qa=build_project_archive(project,dist)
'''
new = '''        # Rebuild indexes once more so the already-frozen final QA/session record
        # is included. Do not rewrite that record after the manifest is frozen;
        # doing so would invalidate its recorded hash.
        index_qa=build_indexes_and_manifest(project)
        session_qa["final_indexes_after_qa_inclusion"] = index_qa
        project_archive,project_archive_qa=build_project_archive(project,dist)
'''
if old not in text:
    raise SystemExit("manifest stability patch anchor not found")
text = text.replace(old, new, 1)
event = '''RECOVERY_EVENTS.append({
    "event_number": 100,
    "event_code": "V3-CP4-S1-REC-LATE-QA-REWRITE-MANIFEST-MISMATCH-CORRECTED",
    "occurred_at": NOW,
    "failed_step": "Verify the clean-extracted project against the final Session 1 project manifest.",
    "exact_error_or_reason": "SESSION_1_COMPLETE_QA.json was rewritten after its hash had been frozen in the project manifest.",
    "intact_artifacts": "The copied project tree, synchronized database, workbook, application, publication, tracking, recovery records and all source artifacts remained intact.",
    "recovery_action": "Freeze the final QA/session record before the final index and manifest pass, include it in that pass, and prohibit any subsequent rewrite before clean verification.",
    "validation_result": "Clean verification requires zero project-manifest mismatches.",
    "data_quality_effect": "None; only build ordering was corrected.",
    "next_checkpoint": "Resume complete project archive and self-contained restore generation.",
})

'''
anchor = "\nNET_PROMPT = (\n"
if anchor not in text:
    raise SystemExit("recovery-event insertion anchor not found")
text = text.replace(anchor, "\n" + event + "NET_PROMPT = (\n", 1)
path.write_text(text, encoding="utf-8")
print({"manifest_stability_patch": "applied", "builder": str(path)})
