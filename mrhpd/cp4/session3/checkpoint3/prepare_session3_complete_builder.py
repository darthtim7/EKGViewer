#!/usr/bin/env python3
"""Generate the final Session 3 / Section 4 complete-restore builder.

The adapter reads the clean-verified generated Session 2 complete-restore
builder, advances it to Response 72 and Session 3, connects independent final
release governance, and changes the terminal state from session-complete to
Remediation Section 4 complete. The generated builder is disposable.
"""
from __future__ import annotations

import io
import re
import tokenize
from pathlib import Path

SOURCE = Path("mrhpd/cp4/session2/checkpoint3/build_session2_complete_restore.py")
TARGET = Path("mrhpd/cp4/session3/checkpoint3/build_session3_complete_restore.py")


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


text = SOURCE.read_text(encoding="utf-8")

# Protect source references to the next session before advancing current
# Session 2 labels to Session 3.
next_placeholders = {
    "Remediation Section 4 of 5 Session 3 of 3": "__MRHPD_NEXT_SECTION_LONG__",
    "Section 4 Session 2 to Session 3 Handoff": "__MRHPD_HANDOFF_TITLE__",
    "SESSION_3_HANDOFF.md": "__MRHPD_HANDOFF_FILE__",
    "Session 3 begins from": "__MRHPD_NEXT_BEGINS_FROM__",
}
for old, placeholder in next_placeholders.items():
    text = text.replace(old, placeholder)

# Advance final response, checkpoint input, and base response in that order.
for old, new in [
    ("Response 69", "Response 72"), ("response69", "response72"), ("RESPONSE69", "RESPONSE72"), ("R69", "R72"), ("r69", "r72"),
    ("Response 68", "Response 71"), ("response68", "response71"), ("RESPONSE68", "RESPONSE71"), ("R68", "R71"), ("r68", "r71"),
    ("Response 66", "Response 69"), ("response66", "response69"), ("RESPONSE66", "RESPONSE69"), ("R66", "R69"), ("r66", "r69"),
]:
    text = text.replace(old, new)
text = replace_number_tokens(text, {"69": "72"})

# Advance current session identifiers.
for old, new in [
    ("Session 2", "Session 3"), ("session2", "session3"), ("SESSION_2", "SESSION_3"),
    ("session_2", "session_3"), ("SESSION2", "SESSION3"), ("CP4-S2", "CP4-S3"),
    ("S4S2", "S4S3"), ("S2-", "S3-"),
]:
    text = text.replace(old, new)

for placeholder, replacement in [
    ("__MRHPD_NEXT_SECTION_LONG__", "Remediation Section 5 of 5"),
    ("__MRHPD_HANDOFF_TITLE__", "Section 4 Complete to Section 5 Handoff"),
    ("__MRHPD_HANDOFF_FILE__", "SECTION_5_HANDOFF.md"),
    ("__MRHPD_NEXT_BEGINS_FROM__", "Remediation Section 5 begins from"),
]:
    text = text.replace(placeholder, replacement)

# Exact source identities.
text = text.replace("BASE_RESPONSE69_BYTES = 177_617_796", "BASE_RESPONSE69_BYTES = 179_612_090")
text = text.replace(
    'BASE_RESPONSE69_SHA256 = "38c8fa08763d5698217ce33a2bbe1e889e726087575b14fb31086f38cfe1300f"',
    'BASE_RESPONSE69_SHA256 = "31e4ba64c7a36870ebeb01e4c88109d512a498ff069f44edfb48ba141044ebcb"',
)
text = text.replace("CP2_RECOVERY_BYTES = 18_318_469", "CP2_RECOVERY_BYTES = 20_512_775")
text = text.replace(
    'CP2_RECOVERY_SHA256 = "b466c463c55dc95d2ac780ff78755f5ae09fa19d9a6be4fb97e914af7568adbe"',
    'CP2_RECOVERY_SHA256 = "08c83d06485479a9c495b153d0a7c0d27c986feb5af127c7b2998da6895ee12f"',
)

# Replace response/tracking narrative and the new recoverable-build events.
response_and_events = r'''RESPONSE72 = {
    "response_key": "R72",
    "response_number": 72,
    "response_label": "72",
    "branch_id": "mainline",
    "canonical_current": 1,
    "response_date": NOW,
    "major_topic": "Human Pathogen Database remediation",
    "title": "Section 4 final release, Session 3 complete restore, and Section 5 handoff",
    "goal": (
        "Independently reconstruct and verify the Response 71 release candidate, complete Session 3 and Remediation Section 4, "
        "and emit the final self-contained Section 4 restore through Response 72."
    ),
    "raw_prompt": RAW_PROMPT,
    "raw_response": "[PRE-EMISSION RESPONSE; final user-visible response is represented by the source-supported summary]",
    "summary": (
        "Reconstructed the exact Response 71 release candidate from the governed Response 69 baseline, independently revalidated "
        "database, workbook, application, publication, source, graphics, drift, tracking, recovery, index, manifest and archive controls, "
        "completed Session 3 and Remediation Section 4, and emitted a clean-extraction-tested self-contained restore through Response 72."
    ),
    "state": "section_complete_continue_required",
    "coverage": "exact raw prompt plus source-supported response summary",
    "fidelity_classification": "source_verified_prompt_and_summary",
    "source_id": "CURRENT-CONVERSATION-R72",
    "source_path": "Current conversation turn and Remediation Section 4 final complete restore",
    "notes": "Remediation Section 4 of 5 is complete. Continue begins Remediation Section 5 of 5.",
}

RECOVERY_EVENTS = [
    {
        "event_number": 159,
        "event_code": "V3-CP4-S3-REC-RESPONSE71-CANDIDATE-APPLIED-AND-VERIFIED",
        "occurred_at": NOW,
        "failed_step": "None; newest verified release-candidate recovery was applied.",
        "exact_error_or_reason": "Checkpoint 3 resumed from the exact Response 69 complete restore and verified Response 71 cumulative recovery package rather than reconstructing state from conversation text.",
        "intact_artifacts": "Response 69 complete restore and project identity, Response 71 recovery, accepted predecessor, frozen Section 3 release, 537-page publication, editable assembly, and main application source.",
        "recovery_action": "Verified input byte counts and SHA-256 identities, executed the embedded deterministic recovery utility, and required the resulting canonical project to satisfy all Response 71 release-candidate gates.",
        "validation_result": "The exact Response 71 candidate was reproduced before final synchronization.",
        "data_quality_effect": "None.",
        "next_checkpoint": "Run independent Checkpoint 3 acceptance.",
    },
    {
        "event_number": 160,
        "event_code": "V3-CP4-S3-REC-INDEPENDENT-DATABASE-ACCEPTANCE-PASSED",
        "occurred_at": NOW,
        "failed_step": "None; independent database and lineage acceptance passed.",
        "exact_error_or_reason": "Final release required independent verification of integrity, foreign keys, response lineage, coverage, governance, source, graphics, page and drift records.",
        "intact_artifacts": "All accepted and frozen source artifacts remained immutable.",
        "recovery_action": "Audited physical and logical tables, response/checkpoint state, field/query/source governance, release governance, risk controls, ten source-version controls, 537 page-QA rows, graphics governance, and fourteen drift domains.",
        "validation_result": "Every independent database acceptance control passed.",
        "data_quality_effect": "Final acceptance metadata added; no clinical content altered.",
        "next_checkpoint": "Run independent publication and workbook acceptance.",
    },
    {
        "event_number": 161,
        "event_code": "V3-CP4-S3-REC-INDEPENDENT-537-PAGE-RENDER-AUDIT-PASSED",
        "occurred_at": NOW,
        "failed_step": "None; all publication pages independently rendered and inspected.",
        "exact_error_or_reason": "Checkpoint 3 could not rely only on Checkpoint 2 self-attestation.",
        "intact_artifacts": "The governed 537-page publication remained byte-identical.",
        "recovery_action": "Used independent dual text extraction, orientation-neutral geometry, full low-resolution pixel census, render-object evidence, and selected raster proofs across all pages.",
        "validation_result": "537 of 537 pages passed and remained searchable.",
        "data_quality_effect": "Independent QA evidence added only.",
        "next_checkpoint": "Synchronize final workbook and release state.",
    },
    {
        "event_number": 162,
        "event_code": "V3-CP4-S3-REC-FINAL-WORKBOOK-AND-APPLICATION-AUDIT-PASSED",
        "occurred_at": NOW,
        "failed_step": "None; final derivative and read-only application controls passed.",
        "exact_error_or_reason": "The final release required direct review surfaces and a read-only audit of the completed database, workbook, publication, editable assembly, and application invariants.",
        "intact_artifacts": "Main application source, publication, and editable assembly retained governed hashes.",
        "recovery_action": "Added final Checkpoint 3 and Section 4 workbook sheets, checked all formulas, generated and executed the final read-only application audit, and persisted the results.",
        "validation_result": "Workbook and final application audit passed.",
        "data_quality_effect": "Governed final-release surfaces added.",
        "next_checkpoint": "Close Session 3 and Remediation Section 4.",
    },
    {
        "event_number": 163,
        "event_code": "V3-CP4-S3-REC-SESSION3-AND-SECTION4-FINAL-STATE-COMPLETE",
        "occurred_at": NOW,
        "failed_step": "None; terminal Section 4 state synchronized.",
        "exact_error_or_reason": "Checkpoint 3 is the terminal checkpoint of Session 3 and Remediation Section 4.",
        "intact_artifacts": "Accepted predecessor and frozen Section 3 release remained unchanged.",
        "recovery_action": "Recorded Response 72, Checkpoint 3 session-complete state, Session 3 release state, final Section 4 release state, risk dispositions, tracking, reports and Section 5 handoff.",
        "validation_result": "Checkpoint 3 of 3, Session 3 of 3, and Remediation Section 4 of 5 are complete.",
        "data_quality_effect": "Final release and handoff metadata added.",
        "next_checkpoint": "Build final indexes, manifest, project archive and complete restore.",
    },
    {
        "event_number": 164,
        "event_code": "V3-CP4-S3-REC-SESSION3-DERIVATIVE-COMPACTION-VERIFIED",
        "occurred_at": NOW,
        "failed_step": "None; superseded Session 3 derivatives were compacted under explicit superset controls.",
        "exact_error_or_reason": "The complete restore must remain below the governed 180 MiB ceiling without filler or loss of canonical data.",
        "intact_artifacts": "Canonical clinical data, source evidence, tracking, recovery history, immutable publication artifacts, current workbook and current database remained intact.",
        "recovery_action": "Removed only equivalence-verified superseded checkpoint database/workbook snapshots and reconstructible prior index, manifest, render-QA and raster-proof derivatives, then rebuilt current indexes and manifests.",
        "validation_result": "Compaction register, current indexes, manifests and clean-project verification passed.",
        "data_quality_effect": "No canonical row, clinical claim, source record or immutable publication artifact removed.",
        "next_checkpoint": "Build the self-contained complete restore.",
    },
    {
        "event_number": 165,
        "event_code": "V3-CP4-S3-REC-SECTION4-COMPLETE-PROJECT-CLEAN-VERIFIED",
        "occurred_at": NOW,
        "failed_step": "None; complete project archive passed clean extraction.",
        "exact_error_or_reason": "A Section 4 terminal checkpoint requires a portable current project, not only a differential overlay.",
        "intact_artifacts": "All current governed project surfaces were included.",
        "recovery_action": "Built the Response 72 complete project archive and reran database, workbook, application, publication, source, graphics, drift, tracking, index and manifest gates from a clean extraction.",
        "validation_result": "Complete project archive passed.",
        "data_quality_effect": "None.",
        "next_checkpoint": "Build and verify the complete restore package.",
    },
    {
        "event_number": 166,
        "event_code": "V3-CP4-S3-REC-SECTION4-COMPLETE-RESTORE-CLEAN-VERIFIED",
        "occurred_at": NOW,
        "failed_step": "None; final self-contained restore passed its embedded verifier.",
        "exact_error_or_reason": "The terminal Section 4 output must require no prior project file, cloud artifact, or conversation reconstruction.",
        "intact_artifacts": "The final complete project snapshot and all controlling QA/report/identity files were preserved.",
        "recovery_action": "Built the complete restore, embedded verification and restoration controls, clean-extracted it, and executed its restore verifier.",
        "validation_result": "Self-contained restore passed.",
        "data_quality_effect": "None.",
        "next_checkpoint": "Create transport volumes and custody copies.",
    },
    {
        "event_number": 167,
        "event_code": "V3-CP4-S3-REC-SECTION4-TRANSPORT-AND-HANDOFF-PREPARED",
        "occurred_at": NOW,
        "failed_step": "None; connector-compatible transport and Section 5 handoff prepared.",
        "exact_error_or_reason": "The final restore must remain downloadable and recoverable through persistent custody.",
        "intact_artifacts": "Complete restore, controls, manifest, checksums, reassembly utility and final reports.",
        "recovery_action": "Split the complete restore into the minimum two sub-100-MiB transport volumes, verified every wrapper, and prepared final controls and Section 5 handoff.",
        "validation_result": "Transport and handoff controls passed.",
        "data_quality_effect": "None.",
        "next_checkpoint": "Remediation Section 5 of 5.",
    },
]

NET_PROMPT = (
    "Continue the Human Pathogen Database from the newest verified checkpoint. Reconstruct the exact Response 71 release candidate "
    "from the governed Response 69 complete restore, independently revalidate all clinical-data, evidence, graphics, publication, workbook, "
    "application, tracking, recovery, Source Index, Bit Index, manifest and archive controls, complete Session 3 and Remediation Section 4, "
    "and emit a clean-extraction-tested self-contained restore through Response 72 without modifying accepted or frozen artifacts."
)

NET_RESPONSE = (
    "Remediation Section 4 of 5 is complete through Response 72. The exact Response 71 candidate was reconstructed from the governed "
    "Response 69 baseline and independently revalidated. Checkpoint 3 of 3 and Session 3 of 3 are complete; database, workbook, application, "
    "publication, source-version, graphics, drift, tracking, recovery, index, manifest, project-archive and complete-restore gates passed. "
    "The 537-page publication, editable assembly, accepted predecessor, frozen Section 3 release, and main application source remain unchanged. "
    "Continue begins Remediation Section 5 of 5."
)
'''
pattern = r"RESPONSE72 = \{.*?\n\}\n\nRECOVERY_EVENTS = \[.*?\n\]\n\nNET_PROMPT = \(.*?\n\)\n\nNET_RESPONSE = \(.*?\n\)\n"
text, substitutions = re.subn(pattern, response_and_events, text, count=1, flags=re.S)
if substitutions != 1:
    raise SystemExit(f"response/event/tracking block replacement count: {substitutions}")

# Load the independent final-governance module.
if "import importlib.util\n" not in text:
    text = replace_once(text, "import hashlib\n", "import hashlib\nimport importlib.util\n", "importlib integration")
module_anchor = 'PROJECT_VERSION = "3.0.0a"\n'
module_block = '''GOVERNANCE_PATH = Path(__file__).with_name("session3_checkpoint3_governance.py")
_governance_spec = importlib.util.spec_from_file_location("mrhpd_cp4_s3_checkpoint3_governance", GOVERNANCE_PATH)
if _governance_spec is None or _governance_spec.loader is None:
    raise RuntimeError(f"Unable to load final Section 4 governance module: {GOVERNANCE_PATH}")
cp3gov = importlib.util.module_from_spec(_governance_spec)
_governance_spec.loader.exec_module(cp3gov)

'''
if "GOVERNANCE_PATH = Path(__file__).with_name(\"session3_checkpoint3_governance.py\")" not in text:
    text = replace_once(text, module_anchor, module_block + module_anchor, "final governance module loader")

# Replace the main finalization order so independent acceptance is prepared,
# generic Session 3 state is closed, and final Section 4 state is then declared.
main_pattern = re.compile(
    r'        provisional_qa\["workbook"\]=workbook_qa\n'
    r'.*?'
    r'        database_qa\["sha256"\]=sha256_file\(db\); database_qa\["bytes"\]=db\.stat\(\)\.st_size\n',
    flags=re.S,
)
main_replacement = '''        provisional_qa["workbook"]=workbook_qa
        section4_candidate=cp3gov.prepare_final_verification(
            project,db,workbook,database_qa,workbook_qa,application_qa,publication_qa,
            generated_at=NOW,response_number=72
        )
        workbook_qa=section4_candidate["workbook_qa"]
        application_qa=section4_candidate["application_qa"]
        publication_qa=section4_candidate["publication_qa"]
        provisional_qa.update({"workbook":workbook_qa,"application":application_qa,"publication":publication_qa,"section4_candidate":section4_candidate["qa"]})
        final_db=finalize_database_status(db,workbook_qa,application_qa,publication_qa)
        database_qa.update(final_db)
        section4_final=cp3gov.complete_section4_release(
            project,db,workbook,database_qa,workbook_qa,application_qa,publication_qa,
            candidate=section4_candidate,generated_at=NOW
        )
        database_qa.update(section4_final["database_qa"])
        workbook_qa=section4_final["workbook_qa"]
        application_qa=section4_final["application_qa"]
        publication_qa=section4_final["publication_qa"]
        provisional_qa.update({"database":database_qa,"workbook":workbook_qa,"application":application_qa,"publication":publication_qa,"section4_final_release":section4_final["qa"]})
        tracking_files=build_tracking_documents(project,db,provisional_qa)
        report_files=build_report(project,provisional_qa)
        report_files.extend(section4_final["report_files"])
        database_qa["sha256"]=sha256_file(db); database_qa["bytes"]=db.stat().st_size
'''
text, substitutions = main_pattern.subn(main_replacement, text, count=1)
if substitutions != 1:
    raise SystemExit(f"main finalization replacement count: {substitutions}")

# Surface the final release throughout session, QA and summary objects.
text = text.replace(
    '"database":database_qa,"workbook":workbook_qa,"application":application_qa,"publication":publication_qa,\n            "accepted_predecessor_mutated":False',
    '"database":database_qa,"workbook":workbook_qa,"application":application_qa,"publication":publication_qa,"section4_final_release":section4_final["qa"],\n            "accepted_predecessor_mutated":False',
)
text = text.replace('"reports":len(report_files),', '"reports":len(report_files),"section4_final_release":section4_final["qa"],', 1)
if '"publication":publication_qa,"compaction":compaction_qa,' in text:
    text = text.replace('"publication":publication_qa,"compaction":compaction_qa,', '"publication":publication_qa,"section4_final_release":section4_final["qa"],"compaction":compaction_qa,', 1)
elif '"publication":publication_qa,"indexes":index_qa,' in text:
    text = text.replace('"publication":publication_qa,"indexes":index_qa,', '"publication":publication_qa,"section4_final_release":section4_final["qa"],"indexes":index_qa,', 1)
else:
    raise SystemExit("summary final-release insertion anchor not found")

text = text.replace('"remediation_section_4_complete":False', '"remediation_section_4_complete":True')
text = text.replace('"remediation_section_4_complete": false', '"remediation_section_4_complete": true')
text = text.replace("- Remediation Section 4 of 5: CONTINUE", "- Remediation Section 4 of 5: COMPLETE")
text = text.replace("RECOVERY_EVENTS_116_124.json", "RECOVERY_EVENTS_159_167.json")
text = text.replace("RECOVERY_EVENTS_116_123.json", "RECOVERY_EVENTS_159_167.json")
text = text.replace("RECOVERY_EVENTS_116_118.json", "RECOVERY_EVENTS_159_167.json")
text = text.replace("events 116–124", "events 159–167")
text = text.replace("events 116-124", "events 159-167")
text = text.replace(
    "Remediation Section 4 of 5 Session 3 of 3 COMPLETE RESTORE THROUGH RESPONSE 72",
    "Remediation Section 4 of 5 COMPLETE Session 3 of 3 COMPLETE RESTORE THROUGH RESPONSE 72",
)

# Require the embedded clean-restore verifier to validate the final Section 4
# state in addition to generic Checkpoint 3 state.
verifier_anchor = "  if con.execute(\"SELECT state FROM section4_checkpoint WHERE checkpoint_code='MRHPD-V3-CP4-S3-CP3'\").fetchone()!=('session_complete',): raise SystemExit('Checkpoint 3 state failure')\n"
verifier_insert = verifier_anchor + "  if con.execute(\"SELECT state FROM section4_final_release WHERE release_code='MRHPD-V3-CP4-COMPLETE'\").fetchone()!=('section_complete',): raise SystemExit('Section 4 final release state failure')\n"
if "Section 4 final release state failure" not in text:
    text = replace_once(text, verifier_anchor, verifier_insert, "complete-restore final-release verifier")

# Prior Checkpoint 2 raster proofs are reconstructible after the independent
# Checkpoint 3 page audit and may be compacted with other superseded derivatives.
derivative_anchor = '        project / "Reports" / "Section 4 Session 3" / "Checkpoint 2" / "PDF Render QA",\n'
if derivative_anchor in text and '"Checkpoint 2" / "Publication Render Proofs"' not in text:
    text = text.replace(
        derivative_anchor,
        derivative_anchor + '        project / "QA" / "Section 4 Session 3" / "Checkpoint 2" / "Publication Render Proofs",\n',
        1,
    )

# Canonical output names and state keys.
text = text.replace("MRHPD_RESPONSE72_SESSION2_COMPLETE_BUILD_SUMMARY.json", "MRHPD_RESPONSE72_SESSION3_COMPLETE_BUILD_SUMMARY.json")
text = text.replace("session_2_of_3_complete", "session_3_of_3_complete")
text = text.replace("mrhpd-response72-session2-complete-build-summary-1.0", "mrhpd-response72-session3-complete-build-summary-1.0")
text = text.replace('"state":"session_complete_continue_required"', '"state":"section_complete_continue_required"')

required = [
    "RESPONSE72",
    "RECOVERY_EVENTS_159_167.json",
    "section3_checkpoint3_governance.py",
    "prepare_final_verification",
    "complete_section4_release",
    "section4_final_release",
    "MRHPD_RESPONSE72_SESSION3_COMPLETE_BUILD_SUMMARY.json",
    '"remediation_section_4_complete":True',
    "Remediation Section 5 of 5",
]
missing = [marker for marker in required if marker not in text]
if missing:
    raise SystemExit({"missing_final_builder_markers": missing})

for forbidden in [
    "BASE_RESPONSE66", "BASE_RESPONSE64", "response_key='R68'", "MRHPD-V3-CP4-S2-CP3",
    "Session 2 of 3 COMPLETE", "MRHPD_RESPONSE72_SESSION2_COMPLETE_BUILD_SUMMARY.json",
    "Remediation Section 4 of 5 Session 3 of 3\"",
]:
    if forbidden in text:
        raise SystemExit(f"forbidden stale control remains: {forbidden}")

TARGET.parent.mkdir(parents=True, exist_ok=True)
TARGET.write_text(text, encoding="utf-8")
print(f"Generated {TARGET} from {SOURCE} for Response 72 / Session 3 / Remediation Section 4 completion.")
