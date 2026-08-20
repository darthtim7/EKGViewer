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
        if zf.testzip() is not None:
            raise RuntimeError("ZIP CRC failure")
        names = zf.namelist()
        if len(names) != len(set(names)):
            raise RuntimeError("duplicate ZIP members")
        for name in names:
            pp = PurePosixPath(name.replace("\\", "/"))
            if pp.is_absolute() or ".." in pp.parts or re.match(r"^[A-Za-z]:", name):
                raise RuntimeError(f"unsafe ZIP member: {name}")
            if re.search(r"(^|/)(filler|padding|dummy_payload|artificial_inflation)(/|$)", name, re.I):
                raise RuntimeError(f"artificial filler member: {name}")
    return names


def main() -> None:
    dist = Path(sys.argv[1] if len(sys.argv) > 1 else "dist_cp5_s3_cp1")
    deliveries = list(dist.glob("MRHPD v3.0.0a Response 82 Section 5 Session 3 Checkpoint 1 Recovery Package *.zip"))
    if len(deliveries) != 1:
        raise RuntimeError({"delivery_candidates": [str(path) for path in deliveries]})
    delivery = deliveries[0]
    names = verify_zip(delivery)
    required = [
        "RECOVERY DATA THROUGH RESPONSE 82",
        "Recovery Verification.json",
        "BUILD_SUMMARY.json",
        "Exact File Names.txt",
        ".sha256.txt",
        "Final Release Intake Report.pdf",
        "Final Release Intake Report.docx",
        "Final Release Register.xlsx",
        "Final Release Readiness Map",
    ]
    missing = [token for token in required if not any(token in name for name in names)]
    if missing:
        raise RuntimeError({"missing_delivery_controls": missing})
    summary_path = dist / "MRHPD_RESPONSE82_SECTION5_SESSION3_CHECKPOINT1_BUILD_SUMMARY.json"
    summary = json.loads(summary_path.read_text(encoding="utf-8"))
    gates = {
        "summary": summary.get("status") == "passed",
        "clean_apply": summary.get("recovery", {}).get("clean_apply", {}).get("status") == "passed",
        "database": summary.get("database", {}).get("integrity") == "ok" and summary.get("database", {}).get("foreign_key_violations") == 0 and summary.get("database", {}).get("response82_records") == 1,
        "checkpoint": summary.get("database", {}).get("checkpoint_state") == "checkpoint_complete" and summary.get("database", {}).get("internal_acceptance_state") == "passed",
        "unsupported_claims": summary.get("database", {}).get("unsupported_provider_claims") == 0 and summary.get("database", {}).get("unsupported_proof_claims") == 0,
        "workbook": summary.get("workbook", {}).get("current_sheet_count", 0) >= 137 and summary.get("workbook", {}).get("formula_error_count") == 0 and summary.get("workbook", {}).get("extension_preservation", {}).get("status") == "passed",
        "application": summary.get("application", {}).get("status") == "passed" and summary.get("application", {}).get("main_application_unchanged") is True,
        "external_boundary": summary.get("external_evidence", {}).get("status") == "passed" and summary.get("external_evidence", {}).get("unsupported_provider_claims") == 0 and summary.get("external_evidence", {}).get("unsupported_proof_claims") == 0,
        "reports": summary.get("reports", {}).get("status") == "passed",
        "index": summary.get("index", {}).get("status") == "passed" and summary.get("index", {}).get("bit_index_integrity") == "ok",
        "manifest": summary.get("manifest_records", 0) > 900,
        "user_upload": summary.get("user_upload_required") is False,
        "checkpoint_state_summary": summary.get("checkpoint_1_of_3_complete") is True and summary.get("session_3_of_3_complete") is False and summary.get("remediation_section_5_complete") is False,
    }
    if not all(gates.values()):
        raise RuntimeError({"independent_output_gate_failed": gates})
    result = {
        "status": "passed",
        "delivery": delivery.name,
        "bytes": delivery.stat().st_size,
        "sha256": sha256_file(delivery),
        "members": len(names),
        "database_tables": summary["database"]["table_count"],
        "workbook_sheets": summary["workbook"]["current_sheet_count"],
        "passed_gates": summary["final_release_gates"]["passed"],
        "controlled_pending_gates": summary["final_release_gates"]["controlled_pending"],
        "provider_approval_claimed": False,
        "physical_proof_completion_claimed": False,
        "user_upload_required": False,
        "checkpoint_1_of_3_complete": True,
        "next": "Checkpoint 2 of 3 - evidence ingestion and governed correction cycle",
    }
    output = dist / "MRHPD v3.0.0a Response 82 Checkpoint 1 Independent Verification.json"
    output.write_text(json.dumps(result, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(result, indent=2))


if __name__ == "__main__":
    main()
