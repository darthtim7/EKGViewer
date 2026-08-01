#!/usr/bin/env python3
"""Prepare exact Response 69 and Response 71 inputs for final Section 4 verification.

The utility reconstructs the governed Response 69 complete restore from its two
transport volumes and extracts the exact Response 71 Checkpoint 2 cumulative
recovery package. All work occurs in disposable directories; no accepted or
frozen artifact is edited in place.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import re
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path, PurePosixPath

BASE_BYTES = 179_612_090
BASE_SHA256 = "31e4ba64c7a36870ebeb01e4c88109d512a498ff069f44edfb48ba141044ebcb"
RECOVERY_BYTES = 20_512_775
RECOVERY_SHA256 = "08c83d06485479a9c495b153d0a7c0d27c986feb5af127c7b2998da6895ee12f"


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def safe_extract(path: Path, target: Path) -> None:
    target.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(path) as zf:
        names = zf.namelist()
        if len(names) != len(set(names)):
            raise RuntimeError({"duplicate_zip_members": path.name})
        for name in names:
            posix = PurePosixPath(name.replace("\\", "/"))
            if posix.is_absolute() or ".." in posix.parts or re.match(r"^[A-Za-z]:", name):
                raise RuntimeError({"unsafe_zip_path": name, "archive": path.name})
        bad = zf.testzip()
        if bad:
            raise RuntimeError({"zip_crc_failure": bad, "archive": path.name})
        zf.extractall(target)


def reconstruct_base(response69_dir: Path, output_dir: Path, work: Path) -> Path:
    wrappers = sorted(response69_dir.rglob("*Complete Restore Drive Volume * of 2.zip"))
    if len(wrappers) != 2:
        raise RuntimeError({"response69_wrappers": [str(path) for path in wrappers]})

    staging = work / "response69_wrappers"
    for index, wrapper in enumerate(wrappers, 1):
        safe_extract(wrapper, staging / f"volume{index}")

    flat = work / "response69_flat"
    flat.mkdir(parents=True)
    seen: dict[str, str] = {}
    for root in sorted(staging.iterdir()):
        for path in root.rglob("*"):
            if not path.is_file():
                continue
            digest = sha256_file(path)
            destination = flat / path.name
            if destination.exists():
                if seen[path.name] != digest:
                    raise RuntimeError({"duplicate_volume_control_mismatch": path.name})
            else:
                shutil.copy2(path, destination)
                seen[path.name] = digest

    utilities = list(flat.glob("reassemble_response69_complete_restore.py"))
    if len(utilities) != 1:
        raise RuntimeError({"response69_reassembly_utilities": [str(path) for path in utilities]})
    result = subprocess.run(
        [sys.executable, str(utilities[0].resolve())],
        cwd=flat,
        text=True,
        capture_output=True,
        timeout=1800,
    )
    if result.returncode != 0:
        raise RuntimeError(
            {
                "response69_reassembly_failed": {
                    "stdout": result.stdout[-20000:],
                    "stderr": result.stderr[-20000:],
                }
            }
        )

    restores = [
        path
        for path in flat.glob("*.zip")
        if path.stat().st_size == BASE_BYTES and sha256_file(path) == BASE_SHA256
    ]
    if len(restores) != 1:
        raise RuntimeError(
            {
                "response69_restore_candidates": [
                    {
                        "name": path.name,
                        "bytes": path.stat().st_size,
                        "sha256": sha256_file(path),
                    }
                    for path in flat.glob("*.zip")
                ]
            }
        )

    output_dir.mkdir(parents=True, exist_ok=True)
    destination = output_dir / restores[0].name
    shutil.copy2(restores[0], destination)
    return destination


def extract_recovery(response71_dir: Path, output_dir: Path) -> tuple[Path, Path]:
    wrappers = list(response71_dir.rglob("*Response 71 Checkpoint 2 Recovery Delivery.zip"))
    if len(wrappers) != 1:
        raise RuntimeError({"response71_delivery_wrappers": [str(path) for path in wrappers]})
    safe_extract(wrappers[0], output_dir)
    inner = [
        path
        for path in output_dir.rglob("*RECOVERY DATA THROUGH RESPONSE 71*.zip")
        if path.stat().st_size == RECOVERY_BYTES and sha256_file(path) == RECOVERY_SHA256
    ]
    if len(inner) != 1:
        raise RuntimeError(
            {
                "response71_recovery_candidates": [
                    {
                        "name": path.name,
                        "bytes": path.stat().st_size,
                        "sha256": sha256_file(path),
                    }
                    for path in output_dir.rglob("*.zip")
                ]
            }
        )
    return wrappers[0], inner[0]


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--response69-dir", type=Path, required=True)
    parser.add_argument("--response71-dir", type=Path, required=True)
    parser.add_argument("--base-output", type=Path, required=True)
    parser.add_argument("--recovery-output", type=Path, required=True)
    parser.add_argument("--work", type=Path, default=Path("prepared_session3_input_work"))
    args = parser.parse_args()

    for path in (args.work, args.base_output, args.recovery_output):
        if path.exists():
            shutil.rmtree(path)
    args.work.mkdir(parents=True)

    base = reconstruct_base(args.response69_dir, args.base_output, args.work)
    delivery, recovery = extract_recovery(args.response71_dir, args.recovery_output)
    print(
        json.dumps(
            {
                "status": "passed",
                "base_restore": {
                    "name": base.name,
                    "bytes": base.stat().st_size,
                    "sha256": sha256_file(base),
                },
                "checkpoint2_delivery": {
                    "name": delivery.name,
                    "bytes": delivery.stat().st_size,
                    "sha256": sha256_file(delivery),
                },
                "checkpoint2_recovery": {
                    "name": recovery.name,
                    "bytes": recovery.stat().st_size,
                    "sha256": sha256_file(recovery),
                },
                "accepted_predecessor_mutated": False,
            },
            indent=2,
        )
    )


if __name__ == "__main__":
    main()
