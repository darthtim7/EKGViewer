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


def verify_zip(path: Path) -> dict:
    with zipfile.ZipFile(path) as zf:
        if zf.testzip() is not None:
            raise RuntimeError(f"ZIP CRC failure: {path}")
        names = zf.namelist()
        if len(names) != len(set(names)):
            raise RuntimeError(f"Duplicate ZIP members: {path}")
        for name in names:
            pp = PurePosixPath(name.replace("\\", "/"))
            if pp.is_absolute() or ".." in pp.parts or re.match(r"^[A-Za-z]:", name):
                raise RuntimeError(f"Unsafe ZIP member: {name}")
    return {"name": path.name, "bytes": path.stat().st_size, "sha256": sha256_file(path), "members": len(names)}


def main() -> None:
    dist = Path(sys.argv[1] if len(sys.argv) > 1 else "dist_cp5_s1_cp3")
    summary_path = dist / "MRHPD_RESPONSE77_SECTION5_SESSION1_COMPLETE_BUILD_SUMMARY.json"
    verification_path = dist / "MRHPD v3.0.0a Response 77 Complete Restore Verification.json"
    exact_names = dist / "MRHPD v3.0.0a Response 77 Exact File Names.txt"
    if not summary_path.exists() or not verification_path.exists() or not exact_names.exists():
        raise RuntimeError("Required Response 77 controls are missing")
    summary = json.loads(summary_path.read_text(encoding="utf-8"))
    verification = json.loads(verification_path.read_text(encoding="utf-8"))
    project_archives = list(dist.glob("*COMPLETE PROJECT THROUGH RESPONSE 77*.zip"))
    restores = list(dist.glob("*COMPLETE RESTORE THROUGH RESPONSE 77*.zip"))
    controls = list(dist.glob("MRHPD v3.0.0a Response 77 Section 5 Session 1 Complete Verification and Controls.zip"))
    volumes = sorted(dist.glob("MRHPD v3.0.0a Response 77 Complete Restore Drive Volume * of *.zip"))
    if len(project_archives) != 1 or len(restores) != 1 or len(controls) != 1 or len(volumes) < 2:
        raise RuntimeError({"output_candidates": {"projects": [str(p) for p in project_archives], "restores": [str(p) for p in restores], "controls": [str(p) for p in controls], "volumes": [str(p) for p in volumes]}})
    project_qa = verify_zip(project_archives[0])
    restore_qa = verify_zip(restores[0])
    controls_qa = verify_zip(controls[0])
    volume_qa = [verify_zip(path) for path in volumes]
    expected_part_count = summary["transport"]["part_count"]
    gates = {
        "summary": summary.get("status") == "passed",
        "verification": verification.get("status") == "passed",
        "self_contained": verification.get("self_contained") is True and verification.get("requires_other_project_files") is False and verification.get("requires_conversation_reconstruction") is False,
        "session_complete": summary.get("checkpoint_3_of_3_complete") is True and summary.get("session_1_of_3_complete") is True and summary.get("remediation_section_5_complete") is False,
        "database": summary.get("database", {}).get("integrity") == "ok" and summary.get("database", {}).get("foreign_key_violations") == 0 and summary.get("database", {}).get("response77_records") == 1,
        "workbook": summary.get("workbook", {}).get("current_sheet_count", 0) >= 106 and summary.get("workbook", {}).get("formula_error_count") == 0,
        "application": summary.get("application", {}).get("status") == "passed" and summary.get("application", {}).get("audit", {}).get("status") == "passed",
        "publication": summary.get("publication", {}).get("digital_pages") == 537 and summary.get("publication", {}).get("print_pages") == 538 and summary.get("publication", {}).get("print_searchable_source_pages") == 537,
        "indexes": summary.get("indexes", {}).get("bit_index_integrity") == "ok",
        "project_clean": summary.get("project_archive", {}).get("clean_project_verification", {}).get("status") == "passed",
        "restore_clean": summary.get("complete_restore", {}).get("clean_restore_verifier") == "passed",
        "transport_count": len(volumes) == expected_part_count,
        "transport_sizes": all(row["bytes"] < 104_857_600 for row in volume_qa),
        "controls": controls_qa["members"] >= 6,
        "user_upload": summary.get("user_upload_required") is False,
    }
    if not all(gates.values()):
        raise RuntimeError({"independent_response77_gate_failed": gates})
    result = {
        "status": "passed",
        "project_archive": project_qa,
        "complete_restore": restore_qa,
        "volumes": volume_qa,
        "controls": controls_qa,
        "database_tables": summary["database"]["table_count"],
        "workbook_sheets": summary["workbook"]["current_sheet_count"],
        "print_pages": summary["publication"]["print_pages"],
        "searchable_pages": summary["publication"]["print_searchable_source_pages"],
        "checkpoint_3_of_3_complete": True,
        "session_1_of_3_complete": True,
        "remediation_section_5_complete": False,
        "user_upload_required": False,
        "next": summary["next"],
    }
    print(json.dumps(result, indent=2))


if __name__ == "__main__":
    main()
