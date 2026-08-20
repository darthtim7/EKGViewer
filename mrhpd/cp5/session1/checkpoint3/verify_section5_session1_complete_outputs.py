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
        if bad:
            raise RuntimeError({"zip_crc_error": bad, "file": str(path)})
        if len(names) != len(set(names)):
            raise RuntimeError({"duplicate_members": str(path)})
        for name in names:
            pp = PurePosixPath(name.replace("\\", "/"))
            if pp.is_absolute() or ".." in pp.parts or re.match(r"^[A-Za-z]:", name):
                raise RuntimeError({"unsafe_member": name, "file": str(path)})
            if re.search(r"(^|/)(filler|padding|dummy_payload|artificial_inflation)(/|$)", name, re.I):
                raise RuntimeError({"filler_member": name, "file": str(path)})
    return {
        "name": path.name,
        "bytes": path.stat().st_size,
        "sha256": sha256_file(path),
        "members": len(names),
        "crc": "passed",
    }


def safe_extract(path: Path, destination: Path) -> None:
    destination.mkdir(parents=True, exist_ok=True)
    verify_zip(path)
    with zipfile.ZipFile(path) as zf:
        zf.extractall(destination)


def main() -> None:
    dist = Path(sys.argv[1] if len(sys.argv) > 1 else "dist_cp5_s1_cp3")
    volumes = sorted(dist.glob("MRHPD v3.0.0a Response 78 Complete Restore Drive Volume * of *.zip"))
    if len(volumes) < 2:
        raise RuntimeError({"volume_candidates": [str(path) for path in volumes]})
    verification_deliveries = list(dist.glob("MRHPD v3.0.0a Response 78 Section 5 Session 1 Complete Verification Delivery *.zip"))
    if len(verification_deliveries) != 1:
        raise RuntimeError({"verification_delivery_candidates": [str(path) for path in verification_deliveries]})
    summary_path = dist / "MRHPD_RESPONSE78_SECTION5_SESSION1_COMPLETE_BUILD_SUMMARY.json"
    summary = json.loads(summary_path.read_text(encoding="utf-8"))
    if summary.get("status") != "passed_with_controlled_external_gates":
        raise RuntimeError({"summary_status": summary.get("status")})
    gates = {row["gate_key"]: row["status"] for row in summary["acceptance_gates"]}
    required_passed = [key for key, value in gates.items() if key not in {"external_preview", "physical_proof"} and value != "passed"]
    if required_passed:
        raise RuntimeError({"nonpassing_internal_gates": required_passed})
    if gates.get("external_preview") != "controlled_pending" or gates.get("physical_proof") != "controlled_pending":
        raise RuntimeError({"external_gate_boundary": {"preview": gates.get("external_preview"), "proof": gates.get("physical_proof")}})
    if summary["database"]["integrity"] != "ok" or summary["database"]["foreign_keys"] != 0:
        raise RuntimeError("database summary gate failed")
    if summary["database"]["response77"] != 1 or summary["database"]["response78"] != 1:
        raise RuntimeError("response lineage gate failed")
    if summary["workbook"]["formula_error_count"] != 0 or summary["workbook"]["current_sheet_count"] < 107:
        raise RuntimeError("workbook gate failed")
    if summary["print_surfaces"]["digital"]["pages"] != 537:
        raise RuntimeError("digital publication page gate failed")
    if summary["print_surfaces"]["print_interior"]["pages"] != 538:
        raise RuntimeError("print interior page gate failed")
    if summary["print_surfaces"]["cover_png"]["pixels"] != [5554, 3375]:
        raise RuntimeError("cover dimension gate failed")
    volume_qa = [verify_zip(path) for path in volumes]
    verification_qa = verify_zip(verification_deliveries[0])
    with tempfile.TemporaryDirectory(prefix="mrhpd-r78-independent-") as td:
        root = Path(td)
        for path in volumes:
            safe_extract(path, root)
        scripts = list(root.glob("reassemble_response78_complete_restore.py"))
        if len(scripts) != 1:
            raise RuntimeError({"reassembler_candidates": [str(path) for path in scripts]})
        result = subprocess.run([sys.executable, str(scripts[0])], cwd=root, text=True, capture_output=True, timeout=1800)
        if result.returncode:
            raise RuntimeError({"reassembly_failed": {"stdout": result.stdout[-20000:], "stderr": result.stderr[-20000:]}})
        reassembly = json.loads(result.stdout)
        restore = root / reassembly["restore"]
        restore_qa = verify_zip(restore)
        extract = root / "restore"
        safe_extract(restore, extract)
        verifier = extract / "TOOLS" / "restore_verify_extract.py"
        result2 = subprocess.run([sys.executable, str(verifier)], cwd=extract, text=True, capture_output=True, timeout=2400)
        if result2.returncode:
            raise RuntimeError({"embedded_verifier_failed": {"stdout": result2.stdout[-20000:], "stderr": result2.stderr[-20000:]}})
        embedded = json.loads(result2.stdout)
        if embedded.get("status") != "passed":
            raise RuntimeError({"embedded_gate": embedded})
    result = {
        "status": "passed_with_controlled_external_gates",
        "volumes": volume_qa,
        "verification_delivery": verification_qa,
        "reassembly": reassembly,
        "complete_restore": restore_qa,
        "embedded_verifier": embedded,
        "database_tables": summary["database"]["tables"],
        "workbook_sheets": summary["workbook"]["current_sheet_count"],
        "digital_pages": 537,
        "print_pages": 538,
        "cover_pixels": [5554, 3375],
        "provider_preview": "controlled_pending_session2",
        "physical_proof": "controlled_pending_session2",
        "user_upload_required": False,
        "next": "Remediation Section 5 of 5 Session 2 of 3",
    }
    print(json.dumps(result, indent=2))


if __name__ == "__main__":
    main()
