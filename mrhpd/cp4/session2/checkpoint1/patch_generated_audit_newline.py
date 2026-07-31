#!/usr/bin/env python3
"""Patch recoverable Checkpoint 1 builder defects exposed by disposable remote runs.

The first disposable build exposed a generated-source newline escape defect.
The second and third progressed farther and exposed two three-column report
rows being unpacked into two variables. This patch applies all corrections and
records each recovery in the copied project.
"""
from pathlib import Path

path = Path("mrhpd/cp4/session2/checkpoint1/build_checkpoint1_recovery.py")
text = path.read_text(encoding="utf-8")

old = r"a.output.parent.mkdir(parents=True,exist_ok=True); a.output.write_text(json.dumps(result,indent=2)+'\n')"
new = r"a.output.parent.mkdir(parents=True,exist_ok=True); a.output.write_text(json.dumps(result,indent=2)+'\\n')"
if old not in text:
    raise SystemExit("expected generated-audit newline source was not found")
text = text.replace(old, new, 1)

old_loop = "    for label, value in [\n"
new_loop = "    for label, value, status in [\n"
if old_loop not in text:
    raise SystemExit("expected three-column baseline report loop was not found")
text = text.replace(old_loop, new_loop, 1)
old_status = 'cells[0].text = label; cells[1].text = value; cells[2].text = value if False else "PASS"'
new_status = 'cells[0].text = label; cells[1].text = value; cells[2].text = status'
if old_status not in text:
    raise SystemExit("expected baseline report status assignment was not found")
text = text.replace(old_status, new_status, 1)

old_qa_loop = "    for area, result in [\n"
new_qa_loop = "    for area, result, status in [\n"
if old_qa_loop not in text:
    raise SystemExit("expected three-column QA report loop was not found")
text = text.replace(old_qa_loop, new_qa_loop, 1)
old_qa_status = 'cells[0].text = area; cells[1].text = result; cells[2].text = "PASS"'
new_qa_status = 'cells[0].text = area; cells[1].text = result; cells[2].text = status.upper()'
if old_qa_status not in text:
    raise SystemExit("expected QA report status assignment was not found")
text = text.replace(old_qa_status, new_qa_status, 1)

old_events = "        recovery_events=[inspection_event,capabilities_event,evidence_event,app_event,package_event]\n"
new_events = '''        generated_source_fix_event={
            "event_code":"V3-CP4-S2-REC-GENERATED-AUDIT-NEWLINE-ESCAPE-CORRECTED","occurred_at":NOW,
            "failed_step":"Execute the generated Section 4 Session 2 application capability-audit utility during the first disposable build.",
            "exact_error_or_reason":"SyntaxError: unterminated string literal at generated line 36 because an outer triple-quoted source template converted a newline escape into a literal newline inside a quoted string.",
            "intact_artifacts":"Response 66 restore, immutable project snapshot, copied SQLite database, workbook source, application source, publication, and all prior QA remained intact.",
            "recovery_action":"Corrected the source-template escaping so the generated utility contains a literal backslash-n sequence, compiled the builder, and reran the complete workflow from the exact Response 66 baseline.",
            "validation_result":"The corrected generated utility compiled and executed during the subsequent full rerun.",
            "data_quality_effect":"None; the failed operation occurred only in a disposable copied working tree before package emission.",
            "next_checkpoint":"Complete Checkpoint 1 recovery packaging and clean-apply verification.",
        }
        report_tuple_fix_event={
            "event_code":"V3-CP4-S2-REC-REPORT-BASELINE-TUPLE-UNPACK-CORRECTED","occurred_at":NOW,
            "failed_step":"Generate the editable Checkpoint 1 baseline-and-immutability report table during the second disposable build.",
            "exact_error_or_reason":"ValueError: too many values to unpack (expected 2) because each baseline report row contained label, value, and status while the loop accepted only label and value.",
            "intact_artifacts":"The exact Response 66 restore, immutable project snapshot, copied database, synchronized application, synchronized workbook, publication, editable assembly, and prior QA remained intact.",
            "recovery_action":"Changed the baseline report loop to unpack label, value, and status, used the supplied status value, compiled the corrected builder, and restarted from the exact Response 66 baseline.",
            "validation_result":"The corrected baseline report table completed during the subsequent full rerun.",
            "data_quality_effect":"None; the failed report derivative existed only in a disposable working directory and was never emitted.",
            "next_checkpoint":"Complete all report, index, manifest, recovery-package, and clean-apply gates.",
        }
        qa_report_tuple_fix_event={
            "event_code":"V3-CP4-S2-REC-REPORT-QA-TUPLE-UNPACK-CORRECTED","occurred_at":NOW,
            "failed_step":"Generate the editable Checkpoint 1 database-workbook-application-publication QA table during the third disposable build.",
            "exact_error_or_reason":"ValueError: too many values to unpack (expected 2) because each QA row contained area, key result, and status while the loop accepted only area and result.",
            "intact_artifacts":"The Response 66 restore and project snapshot, copied synchronized database, application, workbook, publication, editable assembly, capability registry, evidence baseline, drift baseline, and prior QA remained intact.",
            "recovery_action":"Changed the QA report loop to unpack all three values and render the supplied status, then restarted the entire deterministic build from the exact Response 66 baseline.",
            "validation_result":"The corrected QA report table completed during the successful full rerun.",
            "data_quality_effect":"None; the unsuccessful report existed only in an automatically deleted temporary directory.",
            "next_checkpoint":"Complete report validation, indexes, manifests, recovery packaging, and clean-apply verification.",
        }
        recovery_events=[inspection_event,capabilities_event,evidence_event,app_event,package_event,generated_source_fix_event,report_tuple_fix_event,qa_report_tuple_fix_event]
'''
if old_events not in text:
    raise SystemExit("expected recovery_events assignment was not found")
text = text.replace(old_events, new_events, 1)
text = text.replace('RECOVERY_EVENTS_101_105.json', 'RECOVERY_EVENTS_101_108.json')
path.write_text(text, encoding="utf-8")
print("patched", path)
