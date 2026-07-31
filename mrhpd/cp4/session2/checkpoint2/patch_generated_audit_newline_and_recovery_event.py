#!/usr/bin/env python3
"""Apply the recoverable Checkpoint 2 generated-audit correction.

The initial Checkpoint 2 run failed only in a disposable copied tree because an
embedded ``\n`` escape became a literal line break inside generated Python
source. This patch converts the generated write to ``chr(10)``, compiles the
generated audit before execution, and preserves the event in the checkpoint's
recovery history.
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

final_event = '''        ("V3-CP4-S2-REC-CHECKPOINT2-RECOVERY-CLEAN-VERIFIED", "None; Checkpoint 2 recovery package completed.", "Intermediate turns require complete recovery data tied directly to the last session-end full restore.", "Built the cumulative Response 66-to-68 overlay, deterministic apply utility, reports, manifests, checksums, and clean-applied the package before emission."),
'''
new_event = '''        ("V3-CP4-S2-REC-GENERATED-AUDIT-NEWLINE-ESCAPING-CORRECTED", "Generate and execute the Section 4 Session 2 Checkpoint 2 application-audit sidecar.", "The initial disposable build generated an unterminated Python string at line 40 because an embedded newline escape was interpreted while constructing the sidecar source.", "Changed the generated JSON newline write to chr(10), added an explicit py_compile gate before execution, restarted from the exact Response 66 restore plus verified Response 67 recovery, and reran all dependent controls."),
'''
if new_event not in text:
    if final_event not in text:
        raise SystemExit("recovery-event insertion anchor not found")
    text = text.replace(final_event, new_event + final_event, 1)

text = text.replace("RECOVERY_EVENTS_106_112.json", "RECOVERY_EVENTS_106_113.json")

if text == original:
    print("Checkpoint 2 builder was already patched.")
else:
    PATH.write_text(text, encoding="utf-8")
    print("Patched generated audit newline, compile gate, and Recovery Event 113.")
