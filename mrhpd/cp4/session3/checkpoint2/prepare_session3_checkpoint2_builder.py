#!/usr/bin/env python3
"""Generate the Section 4 Session 3 Checkpoint 2 recovery builder.

The verified Session 2 Checkpoint 2 builder already implements cumulative
baseline restoration, predecessor-checkpoint application, copied mutation,
field/query/source/drift audits, workbook/application/publication controls,
tracking, reports, indexes, manifests, recovery overlay, and clean apply.
This adapter advances those mechanics to the exact Response 69 session-boundary
restore plus the clean-verified Response 70 Checkpoint 1 recovery, and emits
Response 71 / Session 3 / Checkpoint 2 semantics. A separate integration patch
adds final source-version, page-level publication, graphics, release-candidate
drift, and expanded workbook/application parity controls.
"""
from __future__ import annotations

import io
import tokenize
from pathlib import Path

SOURCE = Path("mrhpd/cp4/session2/checkpoint2/build_checkpoint2_recovery.py")
TARGET = Path("mrhpd/cp4/session3/checkpoint2/build_checkpoint2_recovery.py")


def replace_number_tokens(source: str, mapping: dict[str, str]) -> str:
    tokens = []
    for token in tokenize.generate_tokens(io.StringIO(source).readline):
        if token.type == tokenize.NUMBER and token.string in mapping:
            token = tokenize.TokenInfo(token.type, mapping[token.string], token.start, token.end, token.line)
        tokens.append(token)
    return tokenize.untokenize(tokens)


def replace_once(text: str, old: str, new: str, label: str) -> str:
    count = text.count(old)
    if count != 1:
        raise SystemExit(f"{label}: expected exactly one match, found {count}")
    return text.replace(old, new, 1)


def replace_required(text: str, old: str, new: str, label: str, minimum: int = 1) -> str:
    count = text.count(old)
    if count < minimum:
        raise SystemExit(f"{label}: expected at least {minimum} match(es), found {count}")
    return text.replace(old, new)


text = SOURCE.read_text(encoding="utf-8")

# Carry forward verified recoverable corrections from Session 2 Checkpoint 2.
text = text.replace(
    "a.output.write_text(json.dumps(result,indent=2)+'\\n',encoding='utf-8')",
    "a.output.write_text(json.dumps(result,indent=2)+chr(10),encoding='utf-8')",
    1,
)
compile_anchor = '''    text_write(audit, application_audit_source())
    launcher = app_dir / "run_section4_session2_checkpoint2.py"
'''
compile_block = '''    text_write(audit, application_audit_source())
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
if "generated_checkpoint2_audit_compile_failed" not in text:
    text = replace_once(text, compile_anchor, compile_block, "generated audit compile gate")
text = text.replace(
    'summaries["project_diff"]["prohibited_drift_count"]',
    'summaries["project_diff"].get("prohibited_drift_count", 0)',
)
text = text.replace("M={json.dumps(manifest,ensure_ascii=False)}", "M={manifest!r}")
index_anchor = '''        index_qa = build_indexes(current)
        manifest_qa = build_manifest(current)
'''
index_block = '''        for cache_dir in sorted(current.rglob("__pycache__"), reverse=True):
            if cache_dir.is_dir():
                shutil.rmtree(cache_dir)
        for pyc in current.rglob("*.pyc"):
            pyc.unlink(missing_ok=True)
        index_qa = build_indexes(current)
        manifest_qa = build_manifest(current)
'''
if "for cache_dir in sorted(current.rglob(\"__pycache__\")" not in text:
    text = replace_required(text, index_anchor, index_block, "Python cache hygiene")
final_event = '''        ("V3-CP4-S2-REC-CHECKPOINT2-RECOVERY-CLEAN-VERIFIED", "None; Checkpoint 2 recovery package completed.", "Intermediate turns require complete recovery data tied directly to the last session-end full restore.", "Built the cumulative Response 66-to-68 overlay, deterministic apply utility, reports, manifests, checksums, and clean-applied the package before emission."),
'''
carry_events = '''        ("V3-CP4-S2-REC-GENERATED-AUDIT-NEWLINE-ESCAPING-CORRECTED", "Generate and execute the Section 4 Session 2 Checkpoint 2 application-audit sidecar.", "The initial disposable build generated an unterminated Python string because an embedded newline escape was interpreted while constructing the sidecar source.", "Changed the generated JSON newline write to chr(10), added an explicit py_compile gate before execution, and reran all dependent controls."),
        ("V3-CP4-S2-REC-PROVISIONAL-WORKBOOK-DRIFT-SUMMARY-DEFAULT-CORRECTED", "Generate the provisional synchronized workbook before final drift classification.", "The disposable build raised KeyError because the provisional project-diff summary intentionally contained only a pending status.", "Changed the provisional workbook row to use an explicit zero default until final drift evidence is available and retained the final strict drift gate."),
        ("V3-CP4-S2-REC-GENERATED-RECOVERY-MANIFEST-PYTHON-LITERAL-CORRECTED", "Clean-apply the cumulative checkpoint recovery utility.", "The disposable build embedded raw JSON booleans in generated Python source.", "Changed the embedded recovery manifest to a validated Python literal representation and excluded runtime bytecode from governed outputs."),
'''
if "V3-CP4-S2-REC-GENERATED-AUDIT-NEWLINE-ESCAPING-CORRECTED" not in text:
    text = replace_once(text, final_event, carry_events + final_event, "Recovery Events 113-115")

# Advance current response 68 -> 71 before introducing the Response 70 predecessor.
for old, new in [
    ("Response 68", "Response 71"),
    ("response68", "response71"),
    ("RESPONSE68", "RESPONSE71"),
    ("Response_68", "Response_71"),
    ("RESPONSE_68", "RESPONSE_71"),
    ("R68", "R71"),
    ("r68", "r71"),
]:
    text = text.replace(old, new)
text = replace_number_tokens(text, {"68": "71"})

# Advance predecessor checkpoint 67 -> 70.
for old, new in [
    ("Response 67", "Response 70"),
    ("response67", "response70"),
    ("RESPONSE67", "RESPONSE70"),
    ("Response_67", "Response_70"),
    ("RESPONSE_67", "RESPONSE_70"),
    ("R67", "R70"),
    ("r67", "r70"),
]:
    text = text.replace(old, new)

# Advance immutable full-restore baseline 66 -> 69.
for old, new in [
    ("Response 66", "Response 69"),
    ("response66", "response69"),
    ("RESPONSE66", "RESPONSE69"),
    ("Response_66", "Response_69"),
    ("RESPONSE_66", "RESPONSE_69"),
    ("R66", "R69"),
    ("r66", "r69"),
]:
    text = text.replace(old, new)

# Advance Session 2 -> Session 3 and all governed identifiers.
for old, new in [
    ("Session 2", "Session 3"),
    ("session2", "session3"),
    ("SESSION2", "SESSION3"),
    ("SESSION_2", "SESSION_3"),
    ("section4_session2", "section4_session3"),
    ("Section 4 Session 2", "Section 4 Session 3"),
    ("S4S2", "S4S3"),
    ("CP4-S2", "CP4-S3"),
    ("cp4-s2", "cp4-s3"),
    ("cp4_s2", "cp4_s3"),
    ("session_2_of_3_complete", "session_3_of_3_complete"),
]:
    text = text.replace(old, new)

# Final-session recovery-event continuity and current checkpoint narrative.
for old_name in (
    "RECOVERY_EVENTS_106_112.json",
    "RECOVERY_EVENTS_106_113.json",
    "RECOVERY_EVENTS_106_114.json",
    "RECOVERY_EVENTS_106_115.json",
):
    text = text.replace(old_name, "RECOVERY_EVENTS_139_154.json")
text = text.replace("Capability Parity and Drift Baseline", "Release-Candidate Reconciliation and Page-Level QA")
text = text.replace("Capability Parity and Drift", "Release-Candidate Reconciliation")
text = text.replace("complete Session 3 restore", "independently verify and complete the Section 4 release")
text = text.replace("complete Session 3 restore.", "independently verify and complete the Section 4 release.")
text = text.replace("Checkpoint 3 ends Session 3 and emits a complete self-contained restore.", "Checkpoint 3 independently verifies and emits the complete Section 4 release candidate and self-contained restore.")

# Replace inherited response narrative with exact Response 71 semantics.
anchor = "\nSEARCH_TERMS = ["
override = r'''

RESPONSE71.update({
    "title": "Section 4 Session 3 final source, page-level publication, graphics, and cross-artifact reconciliation checkpoint",
    "goal": (
        "Continue from the clean-verified Response 70 checkpoint; complete the final authoritative source/version sweep; perform "
        "page-level QA across the 537-page publication; reconcile graphics provenance and placeholder status; resolve final "
        "database-workbook-application-publication drift; expand read-only application and workbook release-candidate parity; "
        "and emit deterministic Checkpoint 2 recovery through Response 71 without modifying immutable source artifacts."
    ),
    "summary": (
        "Restored the exact Response 69 baseline, clean-applied the Response 70 recovery, created a copied Response 71 working tree, "
        "completed detailed field/query/source audits, performed the final official-source metadata/version sweep, audited every "
        "publication page with rendered page metrics and selected raster proofs, reconciled graphics rights/readiness and final "
        "cross-artifact drift, expanded workbook and read-only application release-candidate controls, rebuilt tracking, reports, "
        "indexes, manifests and checksums, and clean-applied the emitted recovery package."
    ),
    "notes": "Checkpoint 2 of 3 is complete. Checkpoint 3 independently revalidates the release candidate, completes final custody/transport controls, and emits the Section 4 complete restore.",
})

NET_PROMPT = (
    "Continue the Human Pathogen Database from the exact verified Response 69 complete restore plus the clean-verified Response 70 "
    "Checkpoint 1 recovery. Work only in a copied tree. Add Response 71 and recovery history; complete the final authoritative "
    "source/version sweep; perform page-level rendered QA across the 537-page publication; reconcile graphics rights, provenance, "
    "alt-text and governed placeholder state; classify and resolve database-workbook-application-publication drift; expand workbook "
    "and read-only application release-candidate parity; preserve the accepted predecessor, frozen Section 3 release, main application "
    "source, publication and editable assembly; rebuild tracking, reports, Source Index, Bit Index, manifests, checksums and QA; and "
    "emit deterministic Checkpoint 2 recovery tied directly to the exact Response 69 complete restore."
)

NET_RESPONSE = (
    "Section 4 Session 3 Checkpoint 2 is complete through Response 71. The exact Response 69 restore and Response 70 recovery were "
    "verified before work began. The copied current project contains final source/version evidence, page-level publication render QA, "
    "graphics release-readiness and placeholder governance, resolved cross-artifact drift, expanded workbook controls, and a read-only "
    "release-candidate application audit. The accepted predecessor, frozen Section 3 release, main application source, 537-page "
    "publication and editable assembly remain unchanged. Continue proceeds to Checkpoint 3 of 3 for independent release verification "
    "and the complete Section 4 restore."
)
'''
if anchor not in text:
    raise SystemExit("SEARCH_TERMS insertion anchor not found")
text = text.replace(anchor, override + anchor, 1)

# Exact CLI/default labels and custody wording.
text = text.replace('default=Path("response69_artifacts")', 'default=Path("response69_artifacts")')
text = text.replace('default=Path("response70_artifact")', 'default=Path("response70_artifact")')
text = text.replace('default=Path("dist_cp4_s3_cp2")', 'default=Path("dist_cp4_s3_cp2")')
text = text.replace("Response 69-to-71", "Response 69-to-71")
text = text.replace("earlier Response 70 checkpoint package is not required", "earlier Response 70 checkpoint package is not required")

required = [
    "RESPONSE71",
    "reconstruct_response70",
    "cp1.reconstruct_response69",
    "section4_session3_field_coverage",
    "MRHPD-V3-CP4-S3-CP2",
    "--base-response69-restore",
    "--response69-dir",
    "--response70-dir",
    "session_3_of_3_complete",
    "RECOVERY_EVENTS_139_154.json",
]
missing = [marker for marker in required if marker not in text]
if missing:
    raise SystemExit({"missing_generated_markers": missing})

prohibited = [
    "RESPONSE68",
    "Response 68",
    "response68",
    "RESPONSE67",
    "Response 67",
    "response67",
    "section4_session2",
    "MRHPD-V3-CP4-S2-CP2",
    "S4S2",
    "--base-response66-restore",
    "--response66-dir",
    "--response67-dir",
]
remaining = [marker for marker in prohibited if marker in text]
if remaining:
    raise SystemExit({"prohibited_generated_markers": remaining})

TARGET.parent.mkdir(parents=True, exist_ok=True)
TARGET.write_text(text, encoding="utf-8")
print(f"Generated {TARGET} from {SOURCE} for Response 71 / Section 4 Session 3 Checkpoint 2.")
