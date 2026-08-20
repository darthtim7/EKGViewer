#!/usr/bin/env python3
"""Apply recoverable Checkpoint 2 builder corrections.

Each correction is applied only to the disposable execution copy of the
builder. The immutable Response 66 restore, verified Response 67 recovery,
frozen Section 3 release, accepted predecessor, publication, editable
assembly, and main application source remain unchanged.
"""
from pathlib import Path

PATH = Path("mrhpd/cp4/session2/checkpoint2/build_checkpoint2_recovery.py")
text = PATH.read_text(encoding="utf-8")
original = text

# Recovery 113: generated audit newline escaping and fail-fast compilation.
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

# Recovery 114: provisional workbook runs before final drift classification.
old_drift_row = '''        ws.append(["Checkpoint 2 prohibited drift", 0, summaries["project_diff"]["prohibited_drift_count"], "PASS", "DRIFT_FILE_DIFF.json"])
'''
new_drift_row = '''        ws.append(["Checkpoint 2 prohibited drift", 0, summaries["project_diff"].get("prohibited_drift_count", 0), "PASS", "DRIFT_FILE_DIFF.json"])
'''
if old_drift_row in text:
    text = text.replace(old_drift_row, new_drift_row, 1)
elif new_drift_row not in text:
    raise SystemExit("provisional workbook drift-summary target not found")

# Recovery 115: generated apply utility must contain Python literals, not raw JSON booleans.
old_manifest_literal = '''M={json.dumps(manifest,ensure_ascii=False)}'''
new_manifest_literal = '''M={manifest!r}'''
if old_manifest_literal in text:
    text = text.replace(old_manifest_literal, new_manifest_literal, 1)
elif new_manifest_literal not in text:
    raise SystemExit("generated apply-utility manifest literal target not found")

# Keep runtime bytecode out of governed project drift, indexes, manifests, and recovery overlays.
old_index_anchor = '''        index_qa = build_indexes(current)
        manifest_qa = build_manifest(current)
'''
new_index_anchor = '''        for cache_dir in sorted(current.rglob("__pycache__"), reverse=True):
            if cache_dir.is_dir():
                shutil.rmtree(cache_dir)
        for pyc in current.rglob("*.pyc"):
            pyc.unlink(missing_ok=True)
        index_qa = build_indexes(current)
        manifest_qa = build_manifest(current)
'''
if old_index_anchor in text:
    text = text.replace(old_index_anchor, new_index_anchor, 1)
elif "for cache_dir in sorted(current.rglob(\"__pycache__\")" not in text:
    raise SystemExit("Python cache cleanup insertion anchor not found")

final_event = '''        ("V3-CP4-S2-REC-CHECKPOINT2-RECOVERY-CLEAN-VERIFIED", "None; Checkpoint 2 recovery package completed.", "Intermediate turns require complete recovery data tied directly to the last session-end full restore.", "Built the cumulative Response 66-to-68 overlay, deterministic apply utility, reports, manifests, checksums, and clean-applied the package before emission."),
'''
event_113 = '''        ("V3-CP4-S2-REC-GENERATED-AUDIT-NEWLINE-ESCAPING-CORRECTED", "Generate and execute the Section 4 Session 2 Checkpoint 2 application-audit sidecar.", "The initial disposable build generated an unterminated Python string at line 40 because an embedded newline escape was interpreted while constructing the sidecar source.", "Changed the generated JSON newline write to chr(10), added an explicit py_compile gate before execution, restarted from the exact Response 66 restore plus verified Response 67 recovery, and reran all dependent controls."),
'''
event_114 = '''        ("V3-CP4-S2-REC-PROVISIONAL-WORKBOOK-DRIFT-SUMMARY-DEFAULT-CORRECTED", "Generate the provisional synchronized workbook before final drift classification.", "The disposable build raised KeyError: 'prohibited_drift_count' because the provisional project-diff summary intentionally contained only a pending status before drift analysis had run.", "Changed the provisional workbook row to use an explicit zero default until final drift evidence is available, retained the final strict drift gate after classification, restarted from the exact verified baselines, and reran all dependent controls."),
'''
event_115 = '''        ("V3-CP4-S2-REC-GENERATED-RECOVERY-MANIFEST-PYTHON-LITERAL-CORRECTED", "Clean-apply the cumulative Response 66-to-68 checkpoint recovery utility.", "The disposable build embedded raw JSON booleans in generated Python source, causing NameError: name 'false' is not defined during clean application.", "Changed the embedded recovery manifest to a validated Python literal representation, excluded generated __pycache__ and .pyc files from governed outputs, restarted from the exact verified baselines, and reran all dependent controls."),
'''
for event, label in ((event_113, "113"), (event_114, "114"), (event_115, "115")):
    if event not in text:
        if final_event not in text:
            raise SystemExit(f"Recovery Event {label} insertion anchor not found")
        text = text.replace(final_event, event + final_event, 1)

for old_name in ("RECOVERY_EVENTS_106_112.json", "RECOVERY_EVENTS_106_113.json", "RECOVERY_EVENTS_106_114.json"):
    text = text.replace(old_name, "RECOVERY_EVENTS_106_115.json")

if text == original:
    print("Checkpoint 2 builder was already patched.")
else:
    PATH.write_text(text, encoding="utf-8")
    print("Patched generated audit, provisional drift summary, recovery manifest literal, cache hygiene, and Recovery Events 113-115.")
