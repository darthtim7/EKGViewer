#!/usr/bin/env python3
from __future__ import annotations

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


def verify_zip(path: Path) -> dict[str, object]:
    with zipfile.ZipFile(path) as zf:
        bad = zf.testzip()
        names = zf.namelist()
        unsafe = []
        for name in names:
            pp = PurePosixPath(name.replace("\\", "/"))
            if pp.is_absolute() or ".." in pp.parts or re.match(r"^[A-Za-z]:", name):
                unsafe.append(name)
    result = {
        "name": path.name,
        "bytes": path.stat().st_size,
        "sha256": sha256_file(path),
        "members": len(names),
        "crc_error": bad,
        "duplicates": len(names) - len(set(names)),
        "unsafe_paths": unsafe,
    }
    if bad or result["duplicates"] or unsafe:
        raise RuntimeError({"zip_verification_failed": result})
    return result


def safe_extract(path: Path, destination: Path) -> None:
    verify_zip(path)
    destination.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(path) as zf:
        zf.extractall(destination)


def main() -> None:
    dist = Path(sys.argv[1] if len(sys.argv) > 1 else "dist_cp5_s2_cp3")
    summaries = list(dist.glob("MRHPD_RESPONSE81_SECTION5_SESSION2_COMPLETE_BUILD_SUMMARY.json"))
    if len(summaries) != 1:
        raise RuntimeError({"summary_candidates": [str(path) for path in summaries]})
    summary = json.loads(summaries[0].read_text(encoding="utf-8"))
    if summary.get("status") != "passed_with_controlled_external_gates":
        raise RuntimeError({"summary_status": summary.get("status")})
    gates = {
        "response": summary.get("response") == 81,
        "database": summary.get("database", {}).get("integrity") == "ok" and summary.get("database", {}).get("foreign_keys") == 0 and summary.get("database", {}).get("response81") == 1,
        "workbook": summary.get("workbook", {}).get("current_sheet_count", 0) >= 129 and summary.get("workbook", {}).get("formula_error_count") == 0,
        "digital": summary.get("publication", {}).get("digital_publication", {}).get("pages") == 537 and summary.get("publication", {}).get("digital_publication", {}).get("searchable_pages") == 537,
        "print": summary.get("publication", {}).get("print_interior", {}).get("pages") == 538 and summary.get("publication", {}).get("print_interior", {}).get("searchable_source_pages") == 537 and summary.get("publication", {}).get("print_interior", {}).get("terminal_blank") is True,
        "cover": summary.get("publication", {}).get("cover", {}).get("pixels") == [5554, 3375] and summary.get("publication", {}).get("cover", {}).get("mode") == "RGB" and summary.get("publication", {}).get("cover", {}).get("alpha") is False,
        "application": summary.get("application", {}).get("status") == "passed" and summary.get("application", {}).get("main_application_unchanged") is True,
        "clean_project": summary.get("clean_project", {}).get("status") == "passed" and summary.get("clean_project", {}).get("manifest_mismatches") == 0,
        "restore": summary.get("embedded_restore_verification", {}).get("status") == "passed",
        "provider": summary.get("provider_approval_claimed") is False,
        "proof": summary.get("physical_proof_ordered") is False,
        "user_upload": summary.get("user_upload_required") is False,
        "session": summary.get("checkpoint_3_of_3_complete") is True and summary.get("session_2_of_3_complete") is True,
    }
    if not all(gates.values()):
        raise RuntimeError({"independent_gate_failure": gates})
    project_archives = list(dist.glob("*COMPLETE PROJECT THROUGH RESPONSE 81*.zip"))
    restores = list(dist.glob("*COMPLETE RESTORE THROUGH RESPONSE 81*.zip"))
    wrappers = sorted(dist.glob("MRHPD v3.0.0a Response 81 Complete Restore Drive Volume * of 3.zip"))
    verification_deliveries = list(dist.glob("MRHPD v3.0.0a Response 81 Section 5 Session 2 Complete Verification Delivery *.zip"))
    if len(project_archives) != 1 or len(restores) != 1 or len(wrappers) != 3 or len(verification_deliveries) != 1:
        raise RuntimeError({
            "project_archives": [str(path) for path in project_archives],
            "restores": [str(path) for path in restores],
            "wrappers": [str(path) for path in wrappers],
            "verification_deliveries": [str(path) for path in verification_deliveries],
        })
    project_qa = verify_zip(project_archives[0])
    restore_qa = verify_zip(restores[0])
    wrapper_qa = [verify_zip(path) for path in wrappers]
    verification_qa = verify_zip(verification_deliveries[0])
    with tempfile.TemporaryDirectory(prefix="mrhpd-r81-independent-transport-") as td:
        root = Path(td)
        for wrapper in wrappers:
            safe_extract(wrapper, root)
        reassembler = root / "reassemble_response81_complete_restore.py"
        result = subprocess.run([sys.executable, str(reassembler.resolve())], cwd=root, text=True, capture_output=True, timeout=1200)
        if result.returncode:
            raise RuntimeError({"transport_reassembly_failed": {"stdout": result.stdout[-8000:], "stderr": result.stderr[-8000:]}})
        reconstructed = root / restores[0].name
        if reconstructed.stat().st_size != restores[0].stat().st_size or sha256_file(reconstructed) != sha256_file(restores[0]):
            raise RuntimeError("independent transport identity mismatch")
        extract_root = root / "restore"
        safe_extract(reconstructed, extract_root)
        verifier = extract_root / "TOOLS" / "restore_verify_extract.py"
        verify_result = subprocess.run([sys.executable, str(verifier.resolve())], cwd=extract_root, text=True, capture_output=True, timeout=2400)
        if verify_result.returncode:
            raise RuntimeError({"embedded_verifier_failed": {"stdout": verify_result.stdout[-12000:], "stderr": verify_result.stderr[-12000:]}})
        embedded = json.loads(verify_result.stdout)
    result = {
        "status": "passed_with_controlled_external_gates",
        "gates": gates,
        "project_archive": project_qa,
        "complete_restore": restore_qa,
        "transport_volumes": wrapper_qa,
        "verification_delivery": verification_qa,
        "embedded_verification": embedded,
        "provider_preview": "controlled_pending",
        "physical_proof": "controlled_pending",
        "user_upload_required": False,
        "checkpoint_3_of_3_complete": True,
        "session_2_of_3_complete": True,
        "next": "Remediation Section 5 of 5 Session 3 of 3 Checkpoint 1 of 3",
    }
    print(json.dumps(result, indent=2))


if __name__ == "__main__":
    main()
