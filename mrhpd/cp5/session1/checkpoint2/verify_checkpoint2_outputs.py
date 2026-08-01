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
    return names


def main() -> None:
    dist = Path(sys.argv[1] if len(sys.argv) > 1 else "dist_cp5_s1_cp2")
    deliveries = list(dist.glob("MRHPD v3.0.0a Response 76 Section 5 Session 1 Checkpoint 2 Recovery Package *.zip"))
    if len(deliveries) != 1:
        raise RuntimeError({"delivery_candidates": [str(path) for path in deliveries]})
    delivery = deliveries[0]
    names = verify_zip(delivery)
    required = [
        "Recovery Verification.json",
        "BUILD_SUMMARY.json",
        "Exact File Names.txt",
        ".sha256.txt",
        "Print Production Candidate Report.pdf",
        "Print Production Candidate Report.docx",
        "Print Production Register.xlsx",
    ]
    missing = [token for token in required if not any(token in name for name in names)]
    if missing:
        raise RuntimeError({"missing_delivery_controls": missing})
    summary = json.loads((dist / "MRHPD_RESPONSE76_SECTION5_CHECKPOINT2_BUILD_SUMMARY.json").read_text(encoding="utf-8"))
    gates = {
        "summary": summary.get("status") == "passed",
        "clean_apply": summary.get("recovery", {}).get("clean_apply", {}).get("status") == "passed",
        "database": summary.get("database", {}).get("integrity") == "ok" and summary.get("database", {}).get("foreign_key_violations") == 0,
        "workbook": summary.get("workbook", {}).get("current_sheet_count", 0) >= 100 and summary.get("workbook", {}).get("formula_error_count") == 0,
        "interior": summary.get("print_interior", {}).get("output_page_count") == 538 and summary.get("print_interior", {}).get("searchable_pages") == 537 and summary.get("print_interior", {}).get("text_mismatch_pages") == 0,
        "cover": summary.get("cover", {}).get("pixel_width") == 5554 and summary.get("cover", {}).get("pixel_height") == 3375 and not summary.get("cover", {}).get("alpha_present"),
        "preflight": summary.get("preflight", {}).get("hard_failures") == 0,
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
        "print_pages": summary["print_interior"]["output_page_count"],
        "cover_pixels": [summary["cover"]["pixel_width"], summary["cover"]["pixel_height"]],
        "user_upload_required": False,
    }
    print(json.dumps(result, indent=2))


if __name__ == "__main__":
    main()
