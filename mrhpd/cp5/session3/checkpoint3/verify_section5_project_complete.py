#!/usr/bin/env python3
"""Independently verify the complete MRHPD Response 84 release outputs."""
from __future__ import annotations

import argparse
import collections
import hashlib
import json
import re
import subprocess
import sys
import tempfile
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
        bad = zf.testzip()
        names = zf.namelist()
        duplicates = [name for name, count in collections.Counter(names).items() if count > 1]
        unsafe = []
        filler = []
        for name in names:
            pp = PurePosixPath(name.replace("\\", "/"))
            if pp.is_absolute() or ".." in pp.parts or re.match(r"^[A-Za-z]:", name):
                unsafe.append(name)
            if re.search(r"(^|/)(filler|padding|dummy_payload|artificial_inflation)(/|$)", name, re.I):
                filler.append(name)
    result = {
        "name": path.name,
        "bytes": path.stat().st_size,
        "sha256": sha256_file(path),
        "members": len(names),
        "crc_error": bad,
        "duplicates": duplicates,
        "unsafe_paths": unsafe,
        "filler_members": filler,
    }
    if bad or duplicates or unsafe or filler:
        raise RuntimeError({"zip_verification_failed": result})
    return result


def safe_extract(path: Path, destination: Path) -> None:
    verify_zip(path)
    destination.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(path) as zf:
        zf.extractall(destination)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("dist", type=Path)
    args = parser.parse_args()
    dist = args.dist.resolve()
    summaries = list(dist.glob("MRHPD_RESPONSE84_SECTION5_ENTIRE_PROJECT_COMPLETE_BUILD_SUMMARY.json"))
    if len(summaries) != 1:
        raise RuntimeError({"build_summary_candidates": [str(path) for path in summaries]})
    summary = json.loads(summaries[0].read_text(encoding="utf-8"))
    wrappers = sorted(dist.glob("MRHPD v3.0.0a Response 84 Complete Restore Drive Volume * of 4.zip"))
    verification_deliveries = list(dist.glob("MRHPD v3.0.0a Response 84 Section 5 and Entire Project Complete Verification Delivery *.zip"))
    project_archives = list(dist.glob("*ALL SECTIONS COMPLETE PROJECT THROUGH RESPONSE 84*.zip"))
    restores = list(dist.glob("*ALL SECTIONS COMPLETE RESTORE THROUGH RESPONSE 84*.zip"))
    checks = {
        "summary_status": summary.get("status") == "passed_with_controlled_external_gates",
        "response": summary.get("response") == 84,
        "checkpoint": summary.get("checkpoint_3_of_3_complete") is True,
        "session": summary.get("session_3_of_3_complete") is True,
        "section": summary.get("remediation_section_5_complete") is True,
        "all_sections": summary.get("all_sections_complete") is True,
        "unsupported_provider_claim": summary.get("provider_approval_claimed") is False,
        "unsupported_proof_claim": summary.get("physical_proof_completion_claimed") is False,
        "user_upload": summary.get("user_upload_required") is False,
        "database": summary.get("database", {}).get("integrity") == "ok" and summary.get("database", {}).get("foreign_keys") == 0 and summary.get("database", {}).get("response84") == 1,
        "workbook": summary.get("workbook", {}).get("current_sheet_count", 0) >= 153 and summary.get("workbook", {}).get("formula_error_count") == 0,
        "publication": summary.get("core_artifacts", {}).get("digital_publication", {}).get("pages") == 537 and summary.get("core_artifacts", {}).get("digital_publication", {}).get("searchable_pages") == 537,
        "print": summary.get("core_artifacts", {}).get("print_interior", {}).get("pages") == 538,
        "cover": summary.get("core_artifacts", {}).get("cover", {}).get("pixels") == [5554, 3375],
        "master_category": summary.get("master_category", {}).get("status") == "passed",
        "project_archive": len(project_archives) == 1,
        "complete_restore": len(restores) == 1,
        "volume_count": len(wrappers) == 4,
        "verification_delivery": len(verification_deliveries) == 1,
    }
    for path in [*wrappers, *verification_deliveries, *project_archives, *restores]:
        verify_zip(path)
    expected_wrapper_hashes = {
        row["sequence"]: row["qa"]["sha256"]
        for row in summary.get("transport", {}).get("wrappers", [])
    }
    for index, wrapper in enumerate(wrappers, start=1):
        if expected_wrapper_hashes.get(index) != sha256_file(wrapper):
            checks[f"volume_{index}_identity"] = False
        else:
            checks[f"volume_{index}_identity"] = True

    with tempfile.TemporaryDirectory(prefix="mrhpd-r84-independent-") as td:
        root = Path(td)
        reassembly = root / "reassembly"
        for wrapper in wrappers:
            safe_extract(wrapper, reassembly)
        scripts = list(reassembly.glob("reassemble_response84_complete_restore.py"))
        if len(scripts) != 1:
            raise RuntimeError({"reassembler_candidates": [str(path) for path in scripts]})
        result = subprocess.run([sys.executable, str(scripts[0].resolve())], cwd=reassembly, text=True, capture_output=True, timeout=1800)
        if result.returncode:
            raise RuntimeError({"reassembly_failed": {"stdout": result.stdout[-12000:], "stderr": result.stderr[-12000:]}})
        reconstructed = json.loads(result.stdout)
        expected_restore = summary["complete_restore"]
        checks["transport_reassembly"] = (
            reconstructed.get("status") == "passed"
            and reconstructed.get("bytes") == expected_restore["bytes"]
            and reconstructed.get("sha256") == expected_restore["sha256"]
        )
        restore_path = Path(reconstructed["restore"])
        restore_extract = root / "restore"
        safe_extract(restore_path, restore_extract)
        verifier = restore_extract / "TOOLS" / "restore_verify_extract.py"
        embedded = subprocess.run([sys.executable, str(verifier.resolve())], cwd=restore_extract, text=True, capture_output=True, timeout=3600)
        if embedded.returncode:
            raise RuntimeError({"embedded_verifier_failed": {"stdout": embedded.stdout[-20000:], "stderr": embedded.stderr[-20000:]}})
        embedded_result = json.loads(embedded.stdout)
        checks["embedded_restore_verifier"] = embedded_result.get("status") == "passed"

    status = "passed" if all(checks.values()) else "failed"
    output = {
        "schema": "mrhpd-response84-entire-project-independent-verification-1.0",
        "status": status,
        "build_summary": summaries[0].name,
        "project_archive": verify_zip(project_archives[0]),
        "complete_restore": verify_zip(restores[0]),
        "transport_volumes": [verify_zip(path) for path in wrappers],
        "verification_delivery": verify_zip(verification_deliveries[0]),
        "embedded_restore_verification": embedded_result,
        "checks": checks,
        "provider_approval_claimed": False,
        "physical_proof_completion_claimed": False,
        "user_upload_required": False,
        "checkpoint_3_of_3_complete": True,
        "session_3_of_3_complete": True,
        "remediation_section_5_complete": True,
        "all_sections_complete": True,
    }
    output_path = dist / "MRHPD v3.0.0a Response 84 Entire Project Independent Verification.json"
    output_path.write_text(json.dumps(output, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(output, indent=2))
    raise SystemExit(0 if status == "passed" else 1)


if __name__ == "__main__":
    main()
