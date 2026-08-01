#!/usr/bin/env python3
from __future__ import annotations

import hashlib
import json
import re
import sys
import zipfile
from pathlib import Path, PurePosixPath


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def verify_zip(path: Path) -> list[str]:
    with zipfile.ZipFile(path) as zf:
        bad = zf.testzip()
        names = zf.namelist()
        duplicates = len(names) - len(set(names))
        unsafe = []
        filler = []
        for name in names:
            pp = PurePosixPath(name.replace("\\", "/"))
            if pp.is_absolute() or ".." in pp.parts or re.match(r"^[A-Za-z]:", name):
                unsafe.append(name)
            if re.search(r"(^|/)(filler|padding|dummy_payload|artificial_inflation)(/|$)", name, re.I):
                filler.append(name)
    if bad or duplicates or unsafe or filler:
        raise RuntimeError({
            "zip_verification_failed": {
                "crc_error": bad,
                "duplicates": duplicates,
                "unsafe_paths": unsafe,
                "filler_members": filler,
            }
        })
    return names


def main() -> None:
    dist = Path(sys.argv[1] if len(sys.argv) > 1 else "dist_cp5_s3_cp2")
    deliveries = list(dist.glob("MRHPD v3.0.0a Response 83 Section 5 Session 3 Checkpoint 2 Recovery Package *.zip"))
    if len(deliveries) != 1:
        raise RuntimeError({"delivery_candidates": [str(path) for path in deliveries]})
    delivery = deliveries[0]
    names = verify_zip(delivery)
    required_tokens = [
        "RECOVERY DATA THROUGH RESPONSE 83",
        "Recovery Verification.json",
        "BUILD_SUMMARY.json",
        "Exact File Names.txt",
        ".sha256.txt",
        "Evidence Ingestion and Release Candidate Report.docx",
        "Evidence Ingestion and Release Candidate Report.pdf",
        "Release Candidate Register.xlsx",
        "Evidence Ingestion and Governed Correction Cycle",
    ]
    missing = [token for token in required_tokens if not any(token in name for name in names)]
    if missing:
        raise RuntimeError({"missing_delivery_controls": missing})

    summary_path = dist / "MRHPD_RESPONSE83_SECTION5_SESSION3_CHECKPOINT2_BUILD_SUMMARY.json"
    verification_path = dist / "MRHPD v3.0.0a Response 83 Checkpoint 2 Recovery Verification.json"
    summary = json.loads(summary_path.read_text(encoding="utf-8"))
    verification = json.loads(verification_path.read_text(encoding="utf-8"))

    database = summary.get("database", {})
    workbook = summary.get("workbook", {})
    application = summary.get("application", {})
    evidence = summary.get("evidence_ingestion", {})
    correction = summary.get("correction_cycle", {})
    gates = summary.get("release_candidate_gates", {})
    recovery = summary.get("recovery", {})

    checks = {
        "summary_status": summary.get("status") == "passed",
        "response": summary.get("response") == 83,
        "checkpoint": summary.get("checkpoint") == "Checkpoint 2 of 3",
        "database": database.get("integrity") == "ok" and database.get("foreign_key_violations") == 0 and database.get("response83_records") == 1 and database.get("evidence_records") == 10 and database.get("correction_records") == 10 and database.get("failed_gates") == 0,
        "workbook": workbook.get("current_sheet_count", 0) >= 145 and workbook.get("formula_error_count") == 0 and not workbook.get("lost_sheets") and workbook.get("status") == "passed",
        "application": application.get("status") == "passed" and application.get("response83") == 1 and application.get("evidence") == 10 and application.get("corrections") == 10 and application.get("failed_gates") == 0 and application.get("main_application_unchanged") is True,
        "evidence_boundary": evidence.get("new_item_level_evidence") == 0 and evidence.get("unsupported_claims") == 0,
        "correction_boundary": correction.get("content_corrections_triggered") == 0,
        "release_gates": gates.get("status") == "passed" and gates.get("passed", 0) >= 20 and gates.get("controlled_pending", 0) >= 1,
        "recovery": recovery.get("status") == "passed" and recovery.get("clean_apply", {}).get("status") == "passed" and recovery.get("checkpoint_2_of_3_complete") is True,
        "verification": verification.get("status") == "passed" and verification.get("clean_apply", {}).get("status") == "passed",
        "unsupported_provider_claim": summary.get("provider_approval_claimed", False) is False and recovery.get("provider_approval_claimed", False) is False,
        "unsupported_proof_claim": summary.get("physical_proof_completion_claimed", False) is False and recovery.get("physical_proof_completion_claimed", False) is False,
        "user_upload": summary.get("user_upload_required") is False and recovery.get("user_upload_required") is False,
    }
    if not all(checks.values()):
        raise RuntimeError({"independent_output_gate_failed": checks})

    result = {
        "schema": "mrhpd-response83-checkpoint2-independent-verification-1.0",
        "status": "passed",
        "delivery": delivery.name,
        "bytes": delivery.stat().st_size,
        "sha256": sha256_file(delivery),
        "members": len(names),
        "database_tables": database.get("table_count"),
        "workbook_sheets": workbook.get("current_sheet_count"),
        "evidence_records": database.get("evidence_records"),
        "correction_records": database.get("correction_records"),
        "new_external_evidence": evidence.get("new_item_level_evidence"),
        "content_corrections_triggered": correction.get("content_corrections_triggered"),
        "provider_approval_claimed": False,
        "physical_proof_completion_claimed": False,
        "user_upload_required": False,
        "checks": checks,
    }
    output = dist / "MRHPD v3.0.0a Response 83 Checkpoint 2 Independent Verification.json"
    output.write_text(json.dumps(result, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(result, indent=2))


if __name__ == "__main__":
    main()
