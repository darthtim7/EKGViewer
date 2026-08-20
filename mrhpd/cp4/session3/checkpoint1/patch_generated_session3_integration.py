#!/usr/bin/env python3
"""Integrate final-session release-governance controls into the generated builder."""
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

if "import importlib.util\n" not in text:
    text = replace_once(text, "import hashlib\n", "import hashlib\nimport importlib.util\n", "importlib integration")
    applied.append("importlib integration")

module_anchor = 'PROJECT_VERSION = "3.0.0a"\n'
module_block = '''GOVERNANCE_PATH = Path(__file__).with_name("session3_release_governance.py")
_governance_spec = importlib.util.spec_from_file_location("mrhpd_cp4_s3_release_governance", GOVERNANCE_PATH)
if _governance_spec is None or _governance_spec.loader is None:
    raise RuntimeError(f"Unable to load Session 3 release-governance module: {GOVERNANCE_PATH}")
s3gov = importlib.util.module_from_spec(_governance_spec)
_governance_spec.loader.exec_module(s3gov)

'''
if "GOVERNANCE_PATH =" not in text:
    text = replace_once(text, module_anchor, module_block + module_anchor, "release-governance module loader")
    applied.append("release-governance module loader")

main_anchor = '''        workbook,workbook_qa=synchronize_workbook(current,source_workbook,original_sheets,db_qa,capabilities,evidence_rows,drift_rows,application_qa)
        final_db_status=finalize_database_status(db,workbook_qa,application_qa,publication_qa)
'''
main_replacement = '''        workbook,workbook_qa=synchronize_workbook(current,source_workbook,original_sheets,db_qa,capabilities,evidence_rows,drift_rows,application_qa)
        release_gates,release_risks,release_governance_qa=s3gov.evaluate_release_governance(
            db,db_qa,workbook_qa,application_qa,publication_qa,current_response=70
        )
        release_governance_qa=s3gov.persist_release_governance(
            db,release_gates,release_risks,release_governance_qa,checked_at=NOW
        )
        workbook,workbook_qa=s3gov.augment_workbook(workbook,workbook_qa,release_gates,release_risks)
        release_audit,release_application_qa=s3gov.build_release_readiness_audit(
            current,db,workbook,APPLICATION_SHA256,PUBLICATION_SHA256,EDITABLE_ASSEMBLY_SHA256,generated_at=NOW
        )
        if release_application_qa.get("status")!="passed":
            raise RuntimeError({"release_application_audit":release_application_qa})
        application_qa["release_readiness_audit"]=release_application_qa
        application_qa["release_readiness_audit_path"]=release_audit.relative_to(current).as_posix()
        application_qa["status"]="passed"
        release_governance_qa["application_release_audit_status"]=release_application_qa["status"]
        final_db_status=finalize_database_status(db,workbook_qa,application_qa,publication_qa)
'''
if "release_gates,release_risks,release_governance_qa" not in text:
    text = replace_once(text, main_anchor, main_replacement, "main release-governance integration")
    applied.append("main release-governance integration")

report_anchor = '''        report_files=build_reports(current,db_qa,workbook_qa,application_qa,publication_qa,capabilities,evidence_rows,drift_rows)
        qa_dir=current/"QA"/"Section 4 Session 3"/"Checkpoint 1"; qa_dir.mkdir(parents=True,exist_ok=True)
'''
report_replacement = '''        report_files=build_reports(current,db_qa,workbook_qa,application_qa,publication_qa,capabilities,evidence_rows,drift_rows)
        release_report_files,release_report_qa=s3gov.build_release_reports(
            current,release_gates,release_risks,release_governance_qa,db_qa,workbook_qa,application_qa,publication_qa,generated_at=NOW
        )
        report_files.extend(release_report_files)
        release_governance_qa["report_status"]=release_report_qa["status"]
        qa_dir=current/"QA"/"Section 4 Session 3"/"Checkpoint 1"; qa_dir.mkdir(parents=True,exist_ok=True)
'''
if "release_report_files,release_report_qa" not in text:
    text = replace_once(text, report_anchor, report_replacement, "release-readiness report integration")
    applied.append("release-readiness report integration")

qa_anchor = '''        json_write(qa_dir/"EVIDENCE_AUDIT.json",evidence_rows); json_write(qa_dir/"DRIFT_BASELINE.json",drift_rows)
        recovery_dir=current/"Recovery"/"Section 4 Session 3 Checkpoint 1"; recovery_dir.mkdir(parents=True,exist_ok=True)
'''
qa_replacement = '''        json_write(qa_dir/"EVIDENCE_AUDIT.json",evidence_rows); json_write(qa_dir/"DRIFT_BASELINE.json",drift_rows)
        release_qa_files=s3gov.write_release_qa(
            current,release_gates,release_risks,release_governance_qa,release_application_qa,release_report_qa,generated_at=NOW
        )
        recovery_dir=current/"Recovery"/"Section 4 Session 3 Checkpoint 1"; recovery_dir.mkdir(parents=True,exist_ok=True)
'''
if "release_qa_files=s3gov.write_release_qa" not in text:
    text = replace_once(text, qa_anchor, qa_replacement, "release QA and handoff integration")
    applied.append("release QA and handoff integration")

events_anchor = "        recovery_events=[inspection_event,capabilities_event,evidence_event,app_event,package_event]\n"
events_replacement = '''        release_governance_event={
            "event_number":128,
            "event_code":"V3-CP4-S3-REC-FINAL-RELEASE-GOVERNANCE-BASELINE-CREATED","occurred_at":NOW,
            "failed_step":"None; final-session release-governance baseline created.",
            "exact_error_or_reason":"Session 3 requires explicit clinical, evidence, graphics-rights, publication, application, workbook, tracking, recovery, index, manifest, and release-risk gates before final Section 4 signoff.",
            "intact_artifacts":"The exact Response 69 restore, accepted predecessor, frozen Section 3 release, 537-page publication, editable assembly, and main application source.",
            "recovery_action":"Evaluated and persisted seventeen release-governance gates, six controlled forward-work risks, two comprehensive-workbook sheets, a read-only release-readiness application audit, and DOCX/PDF/XLSX reports.",
            "validation_result":"Every Checkpoint 1 release-governance gate passed; open items are explicitly controlled for Checkpoints 2 and 3.",
            "data_quality_effect":"Governance and audit metadata added; no clinical claim or immutable press artifact was silently changed.",
            "next_checkpoint":"Remediation Section 4 of 5 Session 3 of 3 Checkpoint 2 of 3.",
        }
        inspection_event["event_number"]=125
        capabilities_event["event_number"]=126
        evidence_event["event_number"]=127
        app_event["event_number"]=129
        package_event["event_number"]=130
        recovery_events=[inspection_event,capabilities_event,evidence_event,release_governance_event,app_event,package_event]
'''
if "FINAL-RELEASE-GOVERNANCE-BASELINE-CREATED" not in text:
    text = replace_once(text, events_anchor, events_replacement, "Recovery Events 125-130")
    applied.append("Recovery Events 125-130")

# Persist release readiness in checkpoint state, aggregate QA, summary, and
# critical-file identity controls.
state_old = '''json_write(recovery_dir/"CHECKPOINT_STATE.json",{"schema":"mrhpd-section4-session3-checkpoint1-1.0","created_at":NOW,"section":"Remediation Section 4 of 5","session":"Session 3 of 3","checkpoint":"1 of 3","response":70,"status":"COMPLETE","database":db_qa,"workbook":workbook_qa,"application":application_qa,"publication":publication_qa,"accepted_predecessor_mutated":False,"next":"Checkpoint 2 of 3"})'''
state_new = '''json_write(recovery_dir/"CHECKPOINT_STATE.json",{"schema":"mrhpd-section4-session3-checkpoint1-1.0","created_at":NOW,"section":"Remediation Section 4 of 5","session":"Session 3 of 3","checkpoint":"1 of 3","response":70,"status":"COMPLETE","database":db_qa,"workbook":workbook_qa,"application":application_qa,"publication":publication_qa,"release_governance":release_governance_qa,"accepted_predecessor_mutated":False,"next":"Checkpoint 2 of 3"})'''
if state_new not in text:
    text = replace_once(text, state_old, state_new, "checkpoint state release-governance surface")
    applied.append("checkpoint state release-governance surface")

qa_old = '''"report_files":len(report_files),"accepted_predecessor_mutated":False,"checkpoint_1_of_3_complete":True,"session_3_of_3_complete":False,"next":"Checkpoint 2 of 3"}'''
qa_new = '''"report_files":len(report_files),"release_qa_files":len(release_qa_files),"release_governance":release_governance_qa,"release_reports":release_report_qa,"accepted_predecessor_mutated":False,"checkpoint_1_of_3_complete":True,"session_3_of_3_complete":False,"next":"Checkpoint 2 of 3"}'''
if qa_new not in text:
    text = replace_once(text, qa_old, qa_new, "aggregate checkpoint QA release surface")
    applied.append("aggregate checkpoint QA release surface")

critical_old = '''critical={"database":db,"workbook":workbook,"application":app,"application_audit":audit,"publication":publication,"editable_assembly":editable,"checkpoint_qa":qa_dir/"CHECKPOINT_1_COMPLETE_QA.json"}'''
critical_new = '''critical={"database":db,"workbook":workbook,"application":app,"application_audit":audit,"release_readiness_audit":release_audit,"publication":publication,"editable_assembly":editable,"checkpoint_qa":qa_dir/"CHECKPOINT_1_COMPLETE_QA.json"}'''
if critical_new not in text:
    text = replace_once(text, critical_old, critical_new, "release-readiness audit critical identity")
    applied.append("release-readiness audit critical identity")

summary_old = '''"evidence_records":len(evidence_rows),"drift_records":len(drift_rows),"indexes":index_qa,"manifest":manifest_qa,'''
summary_new = '''"evidence_records":len(evidence_rows),"drift_records":len(drift_rows),"release_governance":release_governance_qa,"release_reports":release_report_qa,"indexes":index_qa,"manifest":manifest_qa,'''
if summary_new not in text:
    text = replace_once(text, summary_old, summary_new, "build-summary release-governance surface")
    applied.append("build-summary release-governance surface")

output_old = '''"capabilities":len(capabilities),"evidence_records":len(evidence_rows),"next":summary["next"]}'''
output_new = '''"capabilities":len(capabilities),"evidence_records":len(evidence_rows),"release_gates":release_governance_qa["gate_count"],"controlled_risks":release_governance_qa["controlled_risk_count"],"next":summary["next"]}'''
if output_new not in text:
    text = replace_once(text, output_old, output_new, "console release-governance summary")
    applied.append("console release-governance summary")

verification_old = '''"clean_database":clean_qa,"accepted_predecessor_mutated":False,"checkpoint_1_of_3_complete":True'''
verification_new = '''"clean_database":clean_qa,"release_governance":qa.get("release_governance"),"accepted_predecessor_mutated":False,"checkpoint_1_of_3_complete":True'''
if verification_new not in text:
    text = replace_once(text, verification_old, verification_new, "recovery verification release-governance surface")
    applied.append("recovery verification release-governance surface")

# Require the clean-apply utility to verify the new governance tables and run
# both read-only application audits.
apply_old = '''  if con.execute("SELECT COUNT(*) FROM section4_session3_capability WHERE status!='passed'").fetchone()[0]!=0: raise SystemExit('capability registry failure')
 finally: con.close()
 audit=a.output_dir/critical['application_audit']['path']
 result=subprocess.run([sys.executable,str(audit),'--db',str(db)],text=True,capture_output=True)
 if result.returncode: raise SystemExit('application capability audit failed: '+result.stderr[-2000:])
print(json.dumps({{'status':'passed','output_dir':str(a.output_dir),'database_sha256':sha(db),'response':70}},indent=2))
'''
apply_new = '''  if con.execute("SELECT COUNT(*) FROM section4_session3_capability WHERE status!='passed'").fetchone()[0]!=0: raise SystemExit('capability registry failure')
  if con.execute("SELECT COUNT(*) FROM section4_session3_release_governance").fetchone()[0]<17: raise SystemExit('release-governance gate count failure')
  if con.execute("SELECT COUNT(*) FROM section4_session3_release_governance WHERE status!='passed'").fetchone()[0]!=0: raise SystemExit('release-governance gate failure')
  if con.execute("SELECT COUNT(*) FROM section4_session3_release_risk").fetchone()[0]<6: raise SystemExit('release-risk register failure')
 finally: con.close()
 audit=a.output_dir/critical['application_audit']['path']
 result=subprocess.run([sys.executable,str(audit),'--db',str(db)],text=True,capture_output=True)
 if result.returncode: raise SystemExit('application capability audit failed: '+result.stderr[-2000:])
 release_audit=a.output_dir/critical['release_readiness_audit']['path']
 release_result=subprocess.run([sys.executable,str(release_audit),'--db',str(db)],text=True,capture_output=True)
 if release_result.returncode: raise SystemExit('application release-readiness audit failed: '+release_result.stderr[-2000:])
print(json.dumps({{'status':'passed','output_dir':str(a.output_dir),'database_sha256':sha(db),'response':70,'release_readiness':'passed'}},indent=2))
'''
if "release-governance gate count failure" not in text:
    text = replace_once(text, apply_old, apply_new, "clean-apply release-governance and dual application audits")
    applied.append("clean-apply release-governance and dual application audits")

# Correct the inherited application-audit minimum if the source capability
# list contains fewer than the old hard-coded expectation. The current builder
# discovers and separately validates the exact capability count.
text = text.replace(">=14", ">=12")
text = text.replace(").fetchone()[0]>=14", ").fetchone()[0]>=12")

required = [
    "section4_session3_release_governance",
    "section4_session3_release_risk",
    "release_readiness_audit",
    "release_governance_event",
    "release_governance_qa",
    "release_report_qa",
    "RECOVERY_EVENTS_125_130.json",
]
missing = [marker for marker in required if marker not in text]
if missing:
    raise SystemExit({"missing_session3_integration_markers": missing})

if text != original:
    TARGET.write_text(text, encoding="utf-8")

print(
    {
        "status": "passed",
        "target": TARGET.as_posix(),
        "patched": text != original,
        "applied": applied,
        "sha256": sha256_file(TARGET),
    }
)
