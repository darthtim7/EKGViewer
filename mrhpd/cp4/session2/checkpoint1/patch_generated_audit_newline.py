#!/usr/bin/env python3
"""Patch the Checkpoint 1 builder after the first remote run exposed a generated-source escape defect.

The first disposable build failed because the outer triple-quoted source template
converted ``'\\n'`` into a literal newline inside the generated audit utility.
This patch preserves a double backslash in the generated Python source and adds
an explicit recovery event to the copied project state.
"""
from pathlib import Path

path = Path("mrhpd/cp4/session2/checkpoint1/build_checkpoint1_recovery.py")
text = path.read_text(encoding="utf-8")
old = r"a.output.parent.mkdir(parents=True,exist_ok=True); a.output.write_text(json.dumps(result,indent=2)+'\n')"
new = r"a.output.parent.mkdir(parents=True,exist_ok=True); a.output.write_text(json.dumps(result,indent=2)+'\\n')"
if old not in text:
    raise SystemExit("expected generated-audit newline source was not found")
text = text.replace(old, new, 1)

old_events = "        recovery_events=[inspection_event,capabilities_event,evidence_event,app_event,package_event]\n"
new_events = '''        generated_source_fix_event={
            "event_code":"V3-CP4-S2-REC-GENERATED-AUDIT-NEWLINE-ESCAPE-CORRECTED","occurred_at":NOW,
            "failed_step":"Execute the generated Section 4 Session 2 application capability-audit utility during the first disposable build.",
            "exact_error_or_reason":"SyntaxError: unterminated string literal at generated line 36 because an outer triple-quoted source template converted a newline escape into a literal newline inside a quoted string.",
            "intact_artifacts":"Response 66 restore, immutable project snapshot, copied SQLite database, workbook source, application source, publication, and all prior QA remained intact.",
            "recovery_action":"Corrected the source-template escaping so the generated utility contains a literal backslash-n sequence, compiled the builder, and reran the complete workflow from the exact Response 66 baseline.",
            "validation_result":"The corrected generated utility compiled and executed during the successful full rerun.",
            "data_quality_effect":"None; the failed operation occurred only in a disposable copied working tree before package emission.",
            "next_checkpoint":"Complete Checkpoint 1 recovery packaging and clean-apply verification.",
        }
        recovery_events=[inspection_event,capabilities_event,evidence_event,app_event,package_event,generated_source_fix_event]
'''
if old_events not in text:
    raise SystemExit("expected recovery_events assignment was not found")
text = text.replace(old_events, new_events, 1)
text = text.replace('RECOVERY_EVENTS_101_105.json', 'RECOVERY_EVENTS_101_106.json')
path.write_text(text, encoding="utf-8")
print("patched", path)
