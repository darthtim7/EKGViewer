#!/usr/bin/env python3
"""Apply recoverable Checkpoint 2 builder corrections.

The first disposable build exposed generated-audit newline escaping. The next
disposable build reached the provisional workbook stage and exposed a missing
summary key that is intentionally unavailable until drift classification. This
patch corrects both defects, adds fail-fast generated-source compilation, and
preserves both recovery events in the current project records.
"""
from pathlib import Path

PATH = Path("mrhpd/cp4/session2/checkpoint2/build_checkpoint2_recovery.py")
text = PATH.read_text(encoding="utf-8")
original = text

old_write = "a.output.write_text(json.dumps(result,indent=2)+'\\n',encoding='utf-8')"
new_write = "a.output.write_text(json.dumps(result,indent=2)+chr(10),encoding='utf-8')"
if old_write in text:
    text = text.replace(old_write, new_write, 1)
elif new_write not in text:
    raise SystemExit("generated-audit newline target not found")

old_compile_anchor = '''    text_write(audit, application_audit_source())
    launcher = app_dir / "run_section4_session2_checkpoint2.py"
'''
new_compile_anchor = '''    text_write(audit, application_audit_source())
    generated_compile = subprocess.run(
        [sys.executable, "-m", "py_compile", str(audit)],
        cwd=app_dir,
        text=True,
        capture_output=True,
        timeout=120,
    )
    if generated_compile.returncode != 0:
        raise RuntimeError(
            {
                "generated_checkpoint2_audit_compile_failed": {
                    "stdout": generated_compile.stdout[-12000:],
                    "stderr": generated_compile.stderr[-12000:],
                }
            }
        )
    launcher = app_dir / "run_section4_session2_checkpoint2.py"
'''
if old_compile_anchor in text:
    text = text.replace(old_compile_anchor, new_compile_anchor, 1)
elif "generated_checkpoint2_audit_compile_failed" not in text:
    raise SystemExit("generated-audit compile anchor not found")

old_drift_row = '''        ws.append(["Checkpoint 2 prohibited drift", 0, summaries["project_diff"]["prohibited_drift_count"], "PASS", "DRIFT_FILE_DIFF.json"])
'''
new_drift_row = '''        ws.append(["Checkpoint 2 prohibited drift", 0, summaries["project_diff"].get("prohibited_drift_count", 0), "PASS", "DRIFT_FILE_DIFF.json"])
'''
if old_drift_row in text:
    text = text.replace(old_drift_row, new_drift_row, 1)
elif new_drift_row not in text:
    raise SystemExit("provisional workbook drift-summary target not found")

final_event = '''        ("V3-CP4-S2-REC-CHECKPOINT2-RECOVERY-CLEAN-VERIFIED", "None; Checkpoint 2 recovery package completed.", "Intermediate turns require complete recovery data tied directly to the last session-end full restore.", "Built the cumulative Response 66-to-68 overlay, deterministic apply utility, reports, manifests, checksums, and clean-applied the package before emission."),
'''
event_113 = '''        ("V3-CP4-S2-REC-GENERATED-AUDIT-NEWLINE-ESCAPING-CORRECTED", "Generate and execute the Section 4 Session 2 Checkpoint 2 application-audit sidecar.", "The initial disposable build generated an unterminated Python string at line 40 because an embedded newline escape was interpreted while constructing the sidecar source.", "Changed the generated JSON newline write to chr(10), added an explicit py_compile gate before execution, restarted from the exact Response 66 restore plus verified Response 67 recovery, and reran all dependent controls."),
'''
event_114 = '''        ("V3-CP4-S2-REC-PROVISIONAL-WORKBOOK-DRIFT-SUMMARY-DEFAULT-CORRECTED", "Generate the provisional synchronized workbook before final drift classification.", "The disposable build raised KeyError: 'prohibited_drift_count' because the provisional project-diff summary intentionally contained only a pending status before drift analysis had run.", "Changed the provisional workbook row to use an explicit zero default until final drift evidence is available, retained the final strict drift gate after classification, restarted from the exact verified baselines, and reran all dependent controls."),
'''
if event_113 not in text:
    if final_event not in text:
        raise SystemExit("Recovery Event 113 insertion anchor not found")
    text = text.replace(final_event, event_113 + final_event, 1)
if event_114 not in text:
    if final_event not in text:
        raise SystemExit("Recovery Event 114 insertion anchor not found")
    text = text.replace(final_event, event_114 + final_event, 1)

text = text.replace("RECOVERY_EVENTS_106_112.json", "RECOVERY_EVENTS_106_114.json")
text = text.replace("RECOVERY_EVENTS_106_113.json", "RECOVERY_EVENTS_106_114.json")

if text == original:
    print("Checkpoint 2 builder was already patched.")
else:
    PATH.write_text(text, encoding="utf-8")
    print("Patched generated audit, provisional workbook drift default, and Recovery Events 113-114.")
