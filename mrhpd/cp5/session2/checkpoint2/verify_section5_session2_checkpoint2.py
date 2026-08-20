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
            raise RuntimeError(f"ZIP CRC failure: {path}")
        names = zf.namelist()
        if len(names) != len(set(names)):
            raise RuntimeError(f"duplicate ZIP members: {path}")
        for name in names:
            pp = PurePosixPath(name.replace("\\", "/"))
            if pp.is_absolute() or ".." in pp.parts or re.match(r"^[A-Za-z]:", name):
                raise RuntimeError(f"unsafe ZIP member: {name}")
            if any(token in name.lower() for token in ("filler", "padding", "dummy_payload", "artificial_inflation")):
                raise RuntimeError(f"artificial filler member: {name}")
    return names


def main() -> None:
    dist = Path(sys.argv[1] if len(sys.argv) > 1 else "dist_cp5_s2_cp2")
    deliveries = list(dist.glob("MRHPD v3.0.0a Response 80 Section 5 Session 2 Checkpoint 2 Recovery Package *.zip"))
    if len(deliveries) != 1:
        raise RuntimeError({"delivery_candidates": [str(path) for path in deliveries]})
    delivery = deliveries[0]
    names = verify_zip(delivery)
    required = [
        "Recovery Verification.json",
        "BUILD_SUMMARY.json",
        "Exact File Names.txt",
        ".sha256.txt",
        "Provider Evidence and Proof Readiness Report.pdf",
        "Provider Evidence and Proof Readiness Report.docx",
        "Provider Evidence and Proof Readiness Register.xlsx",
        "Provider Evidence Boundary and Proof Readiness",
    ]
    missing = [token for token in required if not any(token in name for name in names)]
    if missing:
        raise RuntimeError({"missing_delivery_controls": missing})
    summary_path = dist / "MRHPD_RESPONSE80_SECTION5_SESSION2_CHECKPOINT2_BUILD_SUMMARY.json"
    verification_path = dist / "MRHPD v3.0.0a Response 80 Checkpoint 2 Recovery Verification.json"
    summary = json.loads(summary_path.read_text(encoding="utf-8"))
    verification = json.loads(verification_path.read_text(encoding="utf-8"))
    gates = {
        "summary": summary.get("status") == "passed",
        "verification": verification.get("status") == "passed",
        "clean_apply": verification.get("clean_apply", {}).get("status") == "passed",
        "database": summary.get("database", {}).get("integrity") == "ok" and summary.get("database", {}).get("foreign_key_violations") == 0,
        "response80": summary.get("database", {}).get("response80_records") == 1,
        "workbook": summary.get("workbook", {}).get("current_sheet_count", 0) >= 122 and summary.get("workbook", {}).get("formula_error_count") == 0,
        "provider_evidence": summary.get("provider_evidence", {}).get("records", 0) >= 12,
        "approval_boundary": summary.get("provider_approval_claimed") is False and summary.get("database", {}).get("provider_approval_claims") == 0,
        "proof_boundary": summary.get("physical_proof_ordered") is False,
        "proof_readiness": summary.get("proof_readiness", {}).get("records", 0) >= 16 and summary.get("proof_readiness", {}).get("external_pending", 0) >= 4,
        "issues": summary.get("conversion_issue_taxonomy", {}).get("records", 0) >= 12 and summary.get("conversion_issue_taxonomy", {}).get("observed_issues") == 0,
        "publication": summary.get("publication", {}).get("pages") == 537 and summary.get("publication", {}).get("searchable_pages") == 537 and summary.get("publication", {}).get("unchanged") is True,
        "print_interior": summary.get("print_interior", {}).get("pages") == 538 and summary.get("print_interior", {}).get("unchanged") is True,
        "cover": summary.get("cover", {}).get("pixels") == [5554, 3375] and summary.get("cover", {}).get("unchanged") is True,
        "application": summary.get("application", {}).get("status") == "passed" and summary.get("application", {}).get("main_application_unchanged") is True,
        "index": summary.get("index", {}).get("status") == "passed" and summary.get("index", {}).get("bit_index_integrity") == "ok",
        "user_upload": summary.get("user_upload_required") is False,
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
        "provider_evidence_records": summary["provider_evidence"]["records"],
        "proof_readiness_records": summary["proof_readiness"]["records"],
        "provider_approval_claimed": False,
        "physical_proof_ordered": False,
        "user_upload_required": False,
        "checkpoint_2_of_3_complete": True,
    }
    print(json.dumps(result, indent=2))


if __name__ == "__main__":
    main()
