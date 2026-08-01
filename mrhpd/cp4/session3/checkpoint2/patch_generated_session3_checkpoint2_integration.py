#!/usr/bin/env python3
"""Integrate Session 3 Checkpoint 2 release-candidate controls."""
from __future__ import annotations

import hashlib
from pathlib import Path

TARGET = Path("mrhpd/cp4/session3/checkpoint2/build_checkpoint2_recovery.py")


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

import_anchor = '''spec.loader.exec_module(cp1)

PROJECT_VERSION = "3.0.0a"
'''
import_block = '''spec.loader.exec_module(cp1)

S3CP2_PATH = Path(__file__).with_name("session3_checkpoint2_governance.py")
s3cp2_spec = importlib.util.spec_from_file_location("mrhpd_cp4_s3_cp2_governance", S3CP2_PATH)
if s3cp2_spec is None or s3cp2_spec.loader is None:
    raise RuntimeError(f"Unable to load Session 3 Checkpoint 2 governance module: {S3CP2_PATH}")
s3cp2 = importlib.util.module_from_spec(s3cp2_spec)
s3cp2_spec.loader.exec_module(s3cp2)

PROJECT_VERSION = "3.0.0a"
'''
if "mrhpd_cp4_s3_cp2_governance" not in text:
    text = replace_once(text, import_anchor, import_block, "Checkpoint 2 governance module loader")
    applied.append("governance module loader")

event_anchor = '''        events = recovery_events()
'''
event_block = '''        events = recovery_events()
        events.extend(s3cp2.recovery_events(NOW))
'''
if "events.extend(s3cp2.recovery_events(NOW))" not in text:
    text = replace_once(text, event_anchor, event_block, "Recovery Events 149-154 integration")
    applied.append("Recovery Events 149-154")

application_anchor = '''        application_qa["final_audit_stdout"] = audit_result.stdout[-15000:]
        application_qa["status"] = "passed"

        tracking_files = build_tracking(current, db)
'''
application_block = '''        application_qa["final_audit_stdout"] = audit_result.stdout[-15000:]
        application_qa["status"] = "passed"

        release_candidate = s3cp2.prepare_release_candidate(
            current,
            db,
            workbook,
            app_paths["application"],
            current_artifacts["publication"],
            current_artifacts["editable_assembly"],
            generated_at=NOW,
            checkpoint_code="MRHPD-V3-CP4-S3-CP2",
            response_number=71,
            base_workbook_qa=workbook_qa,
            base_application_qa=application_qa,
            base_publication_qa=publication_qa,
        )
        workbook_qa = release_candidate["workbook_qa"]
        application_qa = release_candidate["application_qa"]
        publication_qa = release_candidate["publication_qa"]
        current_artifacts["workbook"] = workbook
        drift_rows = critical_drift_rows(baseline, current_artifacts, immutable_response70, current)
        insert_drift_rows(db, drift_rows)
        finalize_database(db, workbook_qa, application_qa, publication_qa)
        db_qa = database_qa(db, current)
        state_data = load_json(state)
        state_data["database_sha256"] = db_qa["sha256"]
        state_data["workbook_sha256"] = workbook_qa["sha256"]
        state_data["checkpoint2_release_candidate"] = release_candidate["qa"]
        json_write(state, state_data)
        audit_result = subprocess.run(
            [sys.executable, str(app_paths["audit"]), "--db", str(db), "--app", str(app_paths["application"]), "--output", str(app_paths["audit_output"])],
            text=True,
            capture_output=True,
            timeout=300,
        )
        if audit_result.returncode != 0:
            raise RuntimeError({"post_release_candidate_application_audit_failed": {"stdout": audit_result.stdout[-20000:], "stderr": audit_result.stderr[-12000:]}})
        application_qa["post_release_candidate_audit_stdout"] = audit_result.stdout[-15000:]
        application_qa["status"] = "passed"

        tracking_files = build_tracking(current, db)
'''
if "release_candidate = s3cp2.prepare_release_candidate" not in text:
    text = replace_once(text, application_anchor, application_block, "Release-candidate orchestration")
    applied.append("release-candidate orchestration")

report_anchor = '''        report_files = build_reports(current, field_rows, query_rows, governance_rows, drift_rows, summaries)
'''
report_block = '''        report_files = build_reports(current, field_rows, query_rows, governance_rows, drift_rows, summaries)
        report_files.extend(release_candidate["report_files"])
        report_files.extend(release_candidate["sample_proofs"])
'''
if "report_files.extend(release_candidate" not in text:
    text = replace_once(text, report_anchor, report_block, "Release-candidate reports and raster proofs")
    applied.append("release-candidate reports and raster proofs")

# Add release-candidate evidence to checkpoint state, checkpoint QA, and build summary.
text = text.replace(
    '                "drift": project_diff_qa,\n                "accepted_predecessor_mutated": False,',
    '                "drift": project_diff_qa,\n                "release_candidate": release_candidate["qa"],\n                "accepted_predecessor_mutated": False,',
)
text = text.replace(
    '            "project_diff": project_diff_qa,\n            "indexes": index_qa,',
    '            "project_diff": project_diff_qa,\n            "release_candidate": release_candidate["qa"],\n            "indexes": index_qa,',
)
if '"release_candidate": release_candidate["qa"]' not in text:
    raise SystemExit("release-candidate QA surface insertion failed")
applied.append("checkpoint state, QA, and summary surfaces")

critical_anchor = '''        recovery, verification = build_recovery_package(base_restore, base_project_archive, immutable_response69, current, critical, report_files, checkpoint_qa, args.dist, work)
'''
critical_block = '''        critical.update(release_candidate["critical_paths"])
        recovery, verification = build_recovery_package(base_restore, base_project_archive, immutable_response69, current, critical, report_files, checkpoint_qa, args.dist, work)
'''
if "critical.update(release_candidate" not in text:
    text = replace_once(text, critical_anchor, critical_block, "Critical release-candidate identities")
    applied.append("critical release-candidate identities")

apply_anchor = '''        if con.execute("SELECT COUNT(*) FROM section4_session3_drift_resolution WHERE checkpoint_code='MRHPD-V3-CP4-S3-CP2' AND status!='passed'").fetchone()[0]!=0: raise SystemExit('drift resolution failure')
'''
apply_block = '''        if con.execute("SELECT COUNT(*) FROM section4_session3_drift_resolution WHERE checkpoint_code='MRHPD-V3-CP4-S3-CP2' AND status!='passed'").fetchone()[0]!=0: raise SystemExit('drift resolution failure')
        if con.execute("SELECT COUNT(*) FROM section4_session3_source_version_sweep WHERE checkpoint_code='MRHPD-V3-CP4-S3-CP2' AND status!='passed'").fetchone()[0]!=0: raise SystemExit('source version sweep failure')
        if con.execute("SELECT COUNT(*) FROM section4_session3_publication_page_qa WHERE checkpoint_code='MRHPD-V3-CP4-S3-CP2' AND status='passed'").fetchone()[0]!=537: raise SystemExit('publication page-level QA failure')
        if con.execute("SELECT COUNT(*) FROM section4_session3_graphics_release_audit WHERE checkpoint_code='MRHPD-V3-CP4-S3-CP2' AND status!='passed'").fetchone()[0]!=0: raise SystemExit('graphics release audit failure')
        if con.execute("SELECT COUNT(*) FROM section4_session3_cross_artifact_drift WHERE checkpoint_code='MRHPD-V3-CP4-S3-CP2' AND status='passed'").fetchone()[0]<14: raise SystemExit('final cross-artifact drift failure')
        if con.execute("SELECT state FROM section4_session3_checkpoint2_release_candidate WHERE checkpoint_code='MRHPD-V3-CP4-S3-CP2'").fetchone()!=('checkpoint_complete',): raise SystemExit('release-candidate checkpoint state failure')
'''
if "source version sweep failure" not in text:
    text = replace_once(text, apply_anchor, apply_block, "Clean-apply release-candidate gates")
    applied.append("clean-apply release-candidate gates")

# Extend console summary with the principal Checkpoint 2 gates.
console_anchor = '''                    "source_governance_tables": governance_summary["audited_tables"],
                    "next": summary["next"],
'''
console_block = '''                    "source_governance_tables": governance_summary["audited_tables"],
                    "final_source_controls": release_candidate["qa"]["source_version_sweep"]["sources"],
                    "publication_page_qa": release_candidate["qa"]["publication_page_qa"]["page_count"],
                    "graphics_assets": release_candidate["qa"]["graphics_release_audit"]["total_assets"],
                    "final_drift_domains": release_candidate["qa"]["cross_artifact_drift"]["domains"],
                    "next": summary["next"],
'''
if "final_source_controls" not in text:
    text = replace_once(text, console_anchor, console_block, "Console release-candidate summary")
    applied.append("console release-candidate summary")

required = [
    "mrhpd_cp4_s3_cp2_governance",
    "events.extend(s3cp2.recovery_events(NOW))",
    "release_candidate = s3cp2.prepare_release_candidate",
    "report_files.extend(release_candidate",
    "critical.update(release_candidate",
    "source version sweep failure",
    "publication page-level QA failure",
    "final_source_controls",
]
missing = [marker for marker in required if marker not in text]
if missing:
    raise SystemExit({"missing_checkpoint2_integration_markers": missing})

if text != original:
    TARGET.write_text(text, encoding="utf-8")
print({"status": "passed", "target": TARGET.as_posix(), "applied": applied, "sha256": sha256_file(TARGET)})
