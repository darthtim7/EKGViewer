#!/usr/bin/env python3
"""Automated, resumable recovery for the MRHPD Response 72 complete restore.

Place this launcher in the same directory tree as the two downloaded Response 72
volume files and run it. The launcher identifies files by ZIP contents and
cryptographic identity rather than requiring the user to rename or manually
identify artifact-envelope filenames.

The default run performs all feasible recovery steps:
1. Discover the two downloaded GitHub artifact envelopes or inner governed
   volume wrappers.
2. Safely extract and verify the two inner governed volume wrappers.
3. Verify the transport manifest and both raw parts.
4. Reassemble and SHA-256 verify the complete Response 72 restore.
5. Safely extract the restore package.
6. Run the embedded restore verifier.
7. Use the embedded verifier to extract the complete current project.
8. Persist machine-readable status, checkpoints, hashes, logs, and exact names.

The process is idempotent. Verified work is reused, while incomplete derivatives
created by this launcher are regenerated without modifying the source downloads.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import shutil
import stat
import subprocess
import sys
import time
import traceback
import zipfile
from pathlib import Path, PurePosixPath
from typing import Any, Iterable

LAUNCHER_VERSION = "1.0.0"
SCHEMA = "mrhpd-response72-automated-recovery-1.0"

EXPECTED_RESTORE_NAME = (
    "Medical References - Human Pathogen Database v3.0.0a Part 8 of 8 "
    "Remediation Section 4 of 5 COMPLETE Session 3 of 3 COMPLETE RESTORE "
    "THROUGH RESPONSE 72 2026-08-01 0117 UTC.zip"
)
EXPECTED_RESTORE_BYTES = 159_186_352
EXPECTED_RESTORE_SHA256 = "cb6d2de9bb351a4ff580e8ac0ac071a774670974098da88be822d64b437b25ce"
EXPECTED_PROJECT_ARCHIVE_NAME = (
    "Medical References - Human Pathogen Database v3.0.0a Part 8 of 8 "
    "Remediation Section 4 of 5 Session 3 of 3 COMPLETE PROJECT THROUGH "
    "RESPONSE 72 2026-08-01 0117 UTC.zip"
)
EXPECTED_PROJECT_ARCHIVE_BYTES = 159_865_032
EXPECTED_PROJECT_ARCHIVE_SHA256 = "88b3a6fab6e1106b2942b92fbe5b10c9b06ffe6f15963a7f0c308203dcb6beb5"

OUTER_DOWNLOAD_NAMES = {
    1: "MRHPD v3.0.0a Response 72 Complete Restore Volume 1 of 2.zip",
    2: "MRHPD v3.0.0a Response 72 Complete Restore Volume 2 of 2.zip",
}
INNER_VOLUME_NAMES = {
    1: "MRHPD v3.0.0a Response 72 Complete Restore Drive Volume 1 of 2.zip",
    2: "MRHPD v3.0.0a Response 72 Complete Restore Drive Volume 2 of 2.zip",
}
VERIFICATION_DOWNLOAD_NAME = "MRHPD v3.0.0a Response 72 Section 4 Complete Verification Delivery.zip"
TRANSPORT_MANIFEST_NAME = "MRHPD_RESPONSE72_COMPLETE_RESTORE_TRANSPORT_MANIFEST.json"
REASSEMBLY_UTILITY_NAME = "reassemble_response72_complete_restore.py"
STATUS_NAME = "MRHPD Response 72 Automated Recovery Status.json"
LOG_NAME = "MRHPD Response 72 Automated Recovery Log.txt"
EXACT_NAMES_NAME = "MRHPD Response 72 Exact File Names.json"


def utc_now() -> str:
    return time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime())


def sha256_file(path: Path, chunk_size: int = 1024 * 1024) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(chunk_size), b""):
            digest.update(block)
    return digest.hexdigest()


def json_write(path: Path, value: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_name(path.name + ".partial")
    temporary.write_text(
        json.dumps(value, indent=2, ensure_ascii=False, default=str) + "\n",
        encoding="utf-8",
    )
    os.replace(temporary, path)


def text_append(path: Path, value: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("a", encoding="utf-8") as handle:
        handle.write(value.rstrip() + "\n")


def file_identity(path: Path) -> dict[str, Any]:
    return {
        "name": path.name,
        "path": str(path.resolve()),
        "bytes": path.stat().st_size,
        "sha256": sha256_file(path),
    }


def exact_file_names() -> dict[str, Any]:
    return {
        "schema": "mrhpd-response72-exact-file-names-1.0",
        "launcher_version": LAUNCHER_VERSION,
        "downloaded_artifact_envelopes": [
            OUTER_DOWNLOAD_NAMES[1],
            OUTER_DOWNLOAD_NAMES[2],
            VERIFICATION_DOWNLOAD_NAME,
        ],
        "inner_governed_volume_wrappers": [
            INNER_VOLUME_NAMES[1],
            INNER_VOLUME_NAMES[2],
        ],
        "transport_controls": [
            TRANSPORT_MANIFEST_NAME,
            REASSEMBLY_UTILITY_NAME,
        ],
        "reconstructed_complete_restore": EXPECTED_RESTORE_NAME,
        "embedded_complete_project_archive": EXPECTED_PROJECT_ARCHIVE_NAME,
        "automated_recovery_outputs": [
            STATUS_NAME,
            LOG_NAME,
            EXACT_NAMES_NAME,
            "RESTORE_PACKAGE",
            "CURRENT_PROJECT",
        ],
    }


class RecoveryError(RuntimeError):
    pass


class RecoverySession:
    def __init__(self, search_root: Path, output_root: Path, *, keep_work: bool = False) -> None:
        self.search_root = search_root.resolve()
        self.output_root = output_root.resolve()
        self.work_root = self.output_root / "_AUTOMATION_WORK"
        self.keep_work = keep_work
        self.status_path = self.output_root / STATUS_NAME
        self.log_path = self.output_root / LOG_NAME
        self.state: dict[str, Any] = {
            "schema": SCHEMA,
            "launcher_version": LAUNCHER_VERSION,
            "started_at": utc_now(),
            "updated_at": utc_now(),
            "status": "running",
            "search_root": str(self.search_root),
            "output_root": str(self.output_root),
            "source_downloads_modified": False,
            "expected_restore": {
                "name": EXPECTED_RESTORE_NAME,
                "bytes": EXPECTED_RESTORE_BYTES,
                "sha256": EXPECTED_RESTORE_SHA256,
            },
            "expected_project_archive": {
                "name": EXPECTED_PROJECT_ARCHIVE_NAME,
                "bytes": EXPECTED_PROJECT_ARCHIVE_BYTES,
                "sha256": EXPECTED_PROJECT_ARCHIVE_SHA256,
            },
            "exact_file_names": exact_file_names(),
            "checkpoints": [],
            "artifacts": {},
        }

    def initialize(self) -> None:
        self.output_root.mkdir(parents=True, exist_ok=True)
        self.work_root.mkdir(parents=True, exist_ok=True)
        json_write(self.output_root / EXACT_NAMES_NAME, exact_file_names())
        text_append(
            self.log_path,
            f"[{utc_now()}] Started MRHPD Response 72 automated recovery launcher v{LAUNCHER_VERSION}.",
        )
        self.persist()

    def persist(self) -> None:
        self.state["updated_at"] = utc_now()
        json_write(self.status_path, self.state)

    def log(self, message: str) -> None:
        line = f"[{utc_now()}] {message}"
        print(line, flush=True)
        text_append(self.log_path, line)

    def checkpoint(self, name: str, status: str, details: Any) -> None:
        record = {
            "sequence": len(self.state["checkpoints"]) + 1,
            "name": name,
            "status": status,
            "recorded_at": utc_now(),
            "details": details,
        }
        self.state["checkpoints"].append(record)
        self.persist()
        self.log(f"Checkpoint {record['sequence']}: {name} — {status}.")

    def fail(self, exc: BaseException) -> None:
        self.state["status"] = "failed"
        self.state["failed_at"] = utc_now()
        self.state["error"] = {
            "type": type(exc).__name__,
            "message": str(exc),
            "traceback": traceback.format_exc(),
            "intact_artifacts": self.state.get("artifacts", {}),
            "recovery_boundary": (
                "Downloaded source files were opened read-only and were not modified. "
                "Rerunning the launcher resumes from verified derivatives or regenerates only incomplete outputs."
            ),
        }
        self.persist()
        self.log(f"FAILED: {type(exc).__name__}: {exc}")

    def complete(self) -> None:
        self.state["status"] = "passed"
        self.state["completed_at"] = utc_now()
        self.persist()
        self.log("Automated recovery completed successfully.")


def safe_member_path(name: str) -> PurePosixPath:
    normalized = name.replace("\\", "/")
    path = PurePosixPath(normalized)
    if path.is_absolute() or ".." in path.parts or re.match(r"^[A-Za-z]:", normalized):
        raise RecoveryError(f"Unsafe ZIP member path: {name!r}")
    return path


def is_zip_symlink(info: zipfile.ZipInfo) -> bool:
    mode = (info.external_attr >> 16) & 0o170000
    return mode == stat.S_IFLNK


def verify_zip(path: Path) -> dict[str, Any]:
    if not path.is_file():
        raise RecoveryError(f"ZIP file not found: {path}")
    with zipfile.ZipFile(path) as archive:
        corrupt = archive.testzip()
        if corrupt is not None:
            raise RecoveryError(f"ZIP CRC failure in {path.name}: {corrupt}")
        names = archive.namelist()
        if len(names) != len(set(names)):
            raise RecoveryError(f"Duplicate ZIP member names in {path.name}")
        file_count = 0
        total_uncompressed = 0
        for info in archive.infolist():
            safe_member_path(info.filename)
            if is_zip_symlink(info):
                raise RecoveryError(f"Symbolic-link ZIP member prohibited: {info.filename}")
            if not info.is_dir():
                file_count += 1
                total_uncompressed += info.file_size
    return {
        **file_identity(path),
        "zip_crc": "passed",
        "member_count": len(names),
        "file_count": file_count,
        "uncompressed_bytes": total_uncompressed,
        "path_safety": "passed",
    }


def stream_extract_member(
    archive: zipfile.ZipFile,
    info: zipfile.ZipInfo,
    destination_root: Path,
) -> Path:
    member = safe_member_path(info.filename)
    target = destination_root.joinpath(*member.parts)
    if info.is_dir():
        target.mkdir(parents=True, exist_ok=True)
        return target
    if is_zip_symlink(info):
        raise RecoveryError(f"Symbolic-link ZIP member prohibited: {info.filename}")
    target.parent.mkdir(parents=True, exist_ok=True)
    temporary = target.with_name(target.name + ".partial")
    if temporary.exists():
        temporary.unlink()
    with archive.open(info, "r") as source, temporary.open("wb") as output:
        shutil.copyfileobj(source, output, length=1024 * 1024)
    if temporary.stat().st_size != info.file_size:
        temporary.unlink(missing_ok=True)
        raise RecoveryError(f"Extracted size mismatch for {info.filename}")
    if target.exists():
        if target.stat().st_size == temporary.stat().st_size and sha256_file(target) == sha256_file(temporary):
            temporary.unlink()
            return target
        temporary.unlink(missing_ok=True)
        raise RecoveryError(f"Conflicting duplicate extracted member: {info.filename}")
    os.replace(temporary, target)
    return target


def safe_extract_zip(path: Path, destination: Path) -> list[Path]:
    verify_zip(path)
    destination.mkdir(parents=True, exist_ok=True)
    extracted: list[Path] = []
    with zipfile.ZipFile(path) as archive:
        for info in archive.infolist():
            extracted.append(stream_extract_member(archive, info, destination))
    return extracted


def iter_candidate_zips(search_root: Path, output_root: Path) -> list[Path]:
    candidates: list[Path] = []
    for path in search_root.rglob("*.zip"):
        try:
            resolved = path.resolve()
        except OSError:
            continue
        if output_root == resolved or output_root in resolved.parents:
            continue
        if path.is_file():
            candidates.append(path)
    return sorted(set(candidates), key=lambda item: (len(item.parts), item.name.lower(), str(item)))


def archive_has_transport_signature(names: Iterable[str]) -> bool:
    basenames = {PurePosixPath(name.replace("\\", "/")).name for name in names}
    return (
        TRANSPORT_MANIFEST_NAME in basenames
        and any(name.endswith(".part001") or name.endswith(".part002") for name in basenames)
    )


def classify_direct_wrapper(path: Path) -> int | None:
    with zipfile.ZipFile(path) as archive:
        basenames = [PurePosixPath(name.replace("\\", "/")).name for name in archive.namelist()]
    has_1 = any(name.endswith(".part001") for name in basenames)
    has_2 = any(name.endswith(".part002") for name in basenames)
    if has_1 and not has_2:
        return 1
    if has_2 and not has_1:
        return 2
    for sequence, exact_name in INNER_VOLUME_NAMES.items():
        if path.name == exact_name:
            return sequence
    return None


def register_wrapper(
    wrappers: dict[int, Path],
    sequence: int,
    path: Path,
) -> None:
    current = wrappers.get(sequence)
    if current is None:
        wrappers[sequence] = path
        return
    if current.resolve() == path.resolve():
        return
    if current.stat().st_size == path.stat().st_size and sha256_file(current) == sha256_file(path):
        return
    raise RecoveryError(
        f"More than one nonidentical candidate was found for governed volume {sequence}: "
        f"{current} and {path}"
    )


def discover_volume_wrappers(session: RecoverySession) -> dict[int, Path]:
    candidates = iter_candidate_zips(session.search_root, session.output_root)
    if not candidates:
        raise RecoveryError(
            f"No ZIP files were found under {session.search_root}. Place both downloaded Response 72 volumes "
            "in this directory tree and rerun."
        )
    wrappers: dict[int, Path] = {}
    extracted_root = session.work_root / "DISCOVERED_INNER_WRAPPERS"
    extracted_root.mkdir(parents=True, exist_ok=True)
    inspected: list[dict[str, Any]] = []

    for candidate in candidates:
        try:
            qa = verify_zip(candidate)
        except Exception as exc:
            inspected.append({"path": str(candidate), "status": "ignored_invalid_zip", "error": str(exc)})
            continue
        with zipfile.ZipFile(candidate) as archive:
            names = archive.namelist()
            basenames: dict[str, list[zipfile.ZipInfo]] = {}
            for info in archive.infolist():
                base = PurePosixPath(info.filename.replace("\\", "/")).name
                basenames.setdefault(base, []).append(info)

            if archive_has_transport_signature(names):
                sequence = classify_direct_wrapper(candidate)
                if sequence is not None:
                    register_wrapper(wrappers, sequence, candidate)
                    inspected.append({"path": str(candidate), "status": "direct_governed_wrapper", "sequence": sequence, "qa": qa})

            for sequence, exact_name in INNER_VOLUME_NAMES.items():
                for info in basenames.get(exact_name, []):
                    target = extracted_root / exact_name
                    extracted = stream_extract_member(archive, info, extracted_root)
                    if extracted != target:
                        target.parent.mkdir(parents=True, exist_ok=True)
                        if not target.exists():
                            shutil.copy2(extracted, target)
                    verify_zip(target)
                    register_wrapper(wrappers, sequence, target)
                    inspected.append(
                        {
                            "path": str(candidate),
                            "status": "artifact_envelope_with_governed_wrapper",
                            "sequence": sequence,
                            "inner": file_identity(target),
                            "qa": qa,
                        }
                    )

    missing = [sequence for sequence in (1, 2) if sequence not in wrappers]
    if missing:
        raise RecoveryError(
            "Unable to discover both governed Response 72 volume wrappers by content. "
            f"Missing volume sequence(s): {missing}. Candidates inspected: {len(inspected)}"
        )
    session.state["artifacts"]["discovery"] = {
        "candidate_zip_count": len(candidates),
        "inspected": inspected,
        "governed_volume_wrappers": {
            str(sequence): verify_zip(path) for sequence, path in sorted(wrappers.items())
        },
    }
    session.checkpoint("discover_exact_governed_volume_wrappers", "passed", session.state["artifacts"]["discovery"])
    return wrappers


def extract_transport(session: RecoverySession, wrappers: dict[int, Path]) -> Path:
    transport_root = session.work_root / "TRANSPORT"
    transport_root.mkdir(parents=True, exist_ok=True)
    wrapper_qa: dict[str, Any] = {}
    for sequence, wrapper in sorted(wrappers.items()):
        wrapper_qa[str(sequence)] = verify_zip(wrapper)
        safe_extract_zip(wrapper, transport_root)

    exact_manifest = list(transport_root.rglob(TRANSPORT_MANIFEST_NAME))
    if len(exact_manifest) == 1:
        manifest_path = exact_manifest[0]
    else:
        compatible: list[Path] = []
        for candidate in transport_root.rglob("*.json"):
            try:
                value = json.loads(candidate.read_text(encoding="utf-8-sig"))
            except Exception:
                continue
            if isinstance(value, dict) and isinstance(value.get("restore"), dict) and isinstance(value.get("parts"), list):
                compatible.append(candidate)
        if len(compatible) != 1:
            raise RecoveryError(
                f"Expected one transport manifest; exact={len(exact_manifest)}, compatible={len(compatible)}"
            )
        manifest_path = compatible[0]

    session.state["artifacts"]["transport_extraction"] = {
        "transport_root": str(transport_root),
        "wrapper_qa": wrapper_qa,
        "manifest": file_identity(manifest_path),
    }
    session.checkpoint("extract_transport_controls_and_parts", "passed", session.state["artifacts"]["transport_extraction"])
    return manifest_path


def load_and_verify_transport_manifest(
    session: RecoverySession,
    manifest_path: Path,
) -> tuple[dict[str, Any], list[Path]]:
    manifest = json.loads(manifest_path.read_text(encoding="utf-8-sig"))
    if not isinstance(manifest, dict):
        raise RecoveryError("Transport manifest root must be an object")
    restore = manifest.get("restore")
    parts = manifest.get("parts")
    if not isinstance(restore, dict) or not isinstance(parts, list) or len(parts) != 2:
        raise RecoveryError("Transport manifest must define one restore and exactly two parts")
    if restore.get("name") != EXPECTED_RESTORE_NAME:
        raise RecoveryError(
            f"Unexpected restore filename in transport manifest: {restore.get('name')!r}"
        )
    if int(restore.get("bytes", -1)) != EXPECTED_RESTORE_BYTES:
        raise RecoveryError("Transport manifest restore byte count does not match the governed Response 72 identity")
    if str(restore.get("sha256", "")).lower() != EXPECTED_RESTORE_SHA256:
        raise RecoveryError("Transport manifest restore SHA-256 does not match the governed Response 72 identity")

    verified_parts: list[Path] = []
    part_results: list[dict[str, Any]] = []
    seen_sequences: set[int] = set()
    for row in sorted(parts, key=lambda item: int(item.get("sequence", 0))):
        sequence = int(row.get("sequence", 0))
        if sequence not in (1, 2) or sequence in seen_sequences:
            raise RecoveryError(f"Invalid or duplicate transport part sequence: {sequence}")
        seen_sequences.add(sequence)
        name = str(row.get("name", ""))
        matches = list(manifest_path.parent.rglob(name))
        if len(matches) != 1:
            raise RecoveryError(f"Expected exactly one extracted raw part named {name!r}; found {len(matches)}")
        path = matches[0]
        identity = file_identity(path)
        if identity["bytes"] != int(row.get("bytes", -1)):
            raise RecoveryError(f"Byte-count mismatch for transport part {sequence}")
        if identity["sha256"] != str(row.get("sha256", "")).lower():
            raise RecoveryError(f"SHA-256 mismatch for transport part {sequence}")
        verified_parts.append(path)
        part_results.append({"sequence": sequence, **identity})

    if seen_sequences != {1, 2}:
        raise RecoveryError(f"Transport part sequence set is incomplete: {sorted(seen_sequences)}")
    session.state["artifacts"]["transport_manifest_verification"] = {
        "manifest": file_identity(manifest_path),
        "restore": restore,
        "parts": part_results,
        "status": "passed",
    }
    session.checkpoint(
        "verify_transport_manifest_and_raw_parts",
        "passed",
        session.state["artifacts"]["transport_manifest_verification"],
    )
    return manifest, verified_parts


def verify_expected_restore(path: Path) -> dict[str, Any]:
    identity = file_identity(path)
    if path.name != EXPECTED_RESTORE_NAME:
        raise RecoveryError(f"Unexpected complete-restore filename: {path.name}")
    if identity["bytes"] != EXPECTED_RESTORE_BYTES:
        raise RecoveryError(
            f"Complete restore byte count mismatch: {identity['bytes']} != {EXPECTED_RESTORE_BYTES}"
        )
    if identity["sha256"] != EXPECTED_RESTORE_SHA256:
        raise RecoveryError(
            f"Complete restore SHA-256 mismatch: {identity['sha256']} != {EXPECTED_RESTORE_SHA256}"
        )
    identity["zip_qa"] = verify_zip(path)
    return identity


def reassemble_restore(
    session: RecoverySession,
    manifest: dict[str, Any],
    parts: list[Path],
) -> Path:
    destination = session.output_root / EXPECTED_RESTORE_NAME
    if destination.exists():
        existing = verify_expected_restore(destination)
        session.state["artifacts"]["complete_restore"] = existing
        session.checkpoint("reuse_verified_complete_restore", "passed", existing)
        return destination

    temporary = destination.with_name(destination.name + ".partial")
    if temporary.exists():
        temporary.unlink()
    with temporary.open("wb") as output:
        for part in parts:
            with part.open("rb") as source:
                shutil.copyfileobj(source, output, length=1024 * 1024)
    os.replace(temporary, destination)
    identity = verify_expected_restore(destination)
    if identity["bytes"] != int(manifest["restore"]["bytes"]):
        raise RecoveryError("Reassembled restore does not match transport-manifest byte count")
    if identity["sha256"] != str(manifest["restore"]["sha256"]).lower():
        raise RecoveryError("Reassembled restore does not match transport-manifest SHA-256")
    session.state["artifacts"]["complete_restore"] = identity
    session.checkpoint("reassemble_and_verify_complete_restore", "passed", identity)
    return destination


def prepare_clean_directory(path: Path, marker_name: str) -> Path:
    marker = path / marker_name
    if path.exists():
        if marker.exists():
            shutil.rmtree(path)
        else:
            alternate = path.with_name(path.name + " RECOVERED " + time.strftime("%Y%m%d-%H%M%S"))
            return prepare_clean_directory(alternate, marker_name)
    path.mkdir(parents=True, exist_ok=True)
    marker.write_text(
        "Created by the MRHPD Response 72 automated recovery launcher.\n",
        encoding="utf-8",
    )
    return path


def find_restore_package_root(extracted_root: Path) -> Path:
    direct_verifier = extracted_root / "TOOLS" / "restore_verify_extract.py"
    if direct_verifier.is_file():
        return extracted_root
    matches = list(extracted_root.rglob("TOOLS/restore_verify_extract.py"))
    if len(matches) != 1:
        raise RecoveryError(f"Expected exactly one embedded restore verifier; found {len(matches)}")
    return matches[0].parent.parent


def extract_restore_package(session: RecoverySession, restore: Path) -> Path:
    package_root = session.output_root / "RESTORE_PACKAGE"
    state_record = session.state.get("artifacts", {}).get("restore_package")
    if package_root.exists() and state_record and state_record.get("status") == "passed":
        root = find_restore_package_root(package_root)
        session.checkpoint("reuse_verified_restore_package", "passed", state_record)
        return root
    package_root = prepare_clean_directory(package_root, ".mrhpd_response72_restore_package")
    safe_extract_zip(restore, package_root)
    root = find_restore_package_root(package_root)
    record = {
        "status": "passed",
        "extraction_root": str(package_root),
        "package_root": str(root),
        "embedded_verifier": str(root / "TOOLS" / "restore_verify_extract.py"),
    }
    session.state["artifacts"]["restore_package"] = record
    session.checkpoint("safely_extract_complete_restore_package", "passed", record)
    return root


def run_command(
    session: RecoverySession,
    label: str,
    command: list[str],
    cwd: Path,
    timeout: int = 1800,
) -> dict[str, Any]:
    session.log(f"Running {label}: {' '.join(command)}")
    result = subprocess.run(
        command,
        cwd=cwd,
        text=True,
        capture_output=True,
        timeout=timeout,
        check=False,
    )
    stdout_path = session.output_root / f"{label} stdout.txt"
    stderr_path = session.output_root / f"{label} stderr.txt"
    stdout_path.write_text(result.stdout or "", encoding="utf-8")
    stderr_path.write_text(result.stderr or "", encoding="utf-8")
    record = {
        "label": label,
        "command": command,
        "cwd": str(cwd),
        "returncode": result.returncode,
        "stdout": str(stdout_path),
        "stderr": str(stderr_path),
    }
    if result.returncode != 0:
        raise RecoveryError(
            f"{label} failed with exit code {result.returncode}. "
            f"See {stdout_path.name} and {stderr_path.name}."
        )
    return record


def run_embedded_verifier(
    session: RecoverySession,
    package_root: Path,
    *,
    extract_project: bool,
) -> Path | None:
    verifier = package_root / "TOOLS" / "restore_verify_extract.py"
    if not verifier.is_file():
        raise RecoveryError(f"Embedded restore verifier not found: {verifier}")
    verify_record = run_command(
        session,
        "Embedded Restore Verification",
        [sys.executable, str(verifier)],
        package_root,
    )
    session.state["artifacts"]["embedded_restore_verification"] = {
        "status": "passed",
        **verify_record,
    }
    session.checkpoint(
        "run_embedded_self_contained_restore_verifier",
        "passed",
        session.state["artifacts"]["embedded_restore_verification"],
    )
    if not extract_project:
        return None

    current_project = session.output_root / "CURRENT_PROJECT"
    existing = session.state.get("artifacts", {}).get("current_project")
    if current_project.exists() and existing and existing.get("status") == "passed":
        session.checkpoint("reuse_verified_extracted_current_project", "passed", existing)
        return current_project

    partial = session.output_root / "CURRENT_PROJECT_PARTIAL"
    partial = prepare_clean_directory(partial, ".mrhpd_response72_current_project_partial")
    extract_record = run_command(
        session,
        "Embedded Current Project Extraction",
        [sys.executable, str(verifier), "--extract-project-to", str(partial)],
        package_root,
        timeout=3600,
    )
    if current_project.exists():
        if (current_project / ".mrhpd_response72_current_project").exists():
            shutil.rmtree(current_project)
        else:
            current_project = current_project.with_name(
                current_project.name + " RECOVERED " + time.strftime("%Y%m%d-%H%M%S")
            )
    os.replace(partial, current_project)
    (current_project / ".mrhpd_response72_current_project").write_text(
        "Verified and extracted by the MRHPD Response 72 automated recovery launcher.\n",
        encoding="utf-8",
    )
    project_archives = list(current_project.rglob(EXPECTED_PROJECT_ARCHIVE_NAME))
    record: dict[str, Any] = {
        "status": "passed",
        "path": str(current_project),
        "extraction_command": extract_record,
        "expected_project_archive_name": EXPECTED_PROJECT_ARCHIVE_NAME,
        "expected_project_archive_present_in_extracted_project": bool(project_archives),
    }
    if project_archives:
        archive_identity = file_identity(project_archives[0])
        record["embedded_project_archive"] = archive_identity
    session.state["artifacts"]["current_project"] = record
    session.checkpoint("verify_and_extract_complete_current_project", "passed", record)
    return current_project


def recover(
    search_root: Path,
    output_root: Path,
    *,
    keep_work: bool,
    extract_project: bool,
) -> dict[str, Any]:
    session = RecoverySession(search_root, output_root, keep_work=keep_work)
    session.initialize()
    try:
        wrappers = discover_volume_wrappers(session)
        manifest_path = extract_transport(session, wrappers)
        manifest, parts = load_and_verify_transport_manifest(session, manifest_path)
        restore = reassemble_restore(session, manifest, parts)
        package_root = extract_restore_package(session, restore)
        current_project = run_embedded_verifier(
            session,
            package_root,
            extract_project=extract_project,
        )
        session.state["artifacts"]["final"] = {
            "complete_restore": verify_expected_restore(restore),
            "restore_package_root": str(package_root),
            "current_project": str(current_project) if current_project else None,
            "source_downloads_modified": False,
        }
        if not keep_work and session.work_root.exists():
            shutil.rmtree(session.work_root)
        session.complete()
        return session.state
    except BaseException as exc:
        session.fail(exc)
        raise


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "Automatically discover, verify, reassemble, and extract the MRHPD Response 72 complete restore."
        )
    )
    parser.add_argument(
        "--search-root",
        type=Path,
        default=Path(__file__).resolve().parent,
        help="Directory tree containing both downloaded volume files. Defaults to the launcher directory.",
    )
    parser.add_argument(
        "--output-root",
        type=Path,
        default=None,
        help=(
            "Recovery output directory. Defaults to 'MRHPD Response 72 Automated Recovery Output' "
            "under --search-root."
        ),
    )
    parser.add_argument(
        "--verify-only",
        action="store_true",
        help="Reassemble and run the embedded verifier, but do not extract the complete current project.",
    )
    parser.add_argument(
        "--keep-work",
        action="store_true",
        help="Retain disposable transport extraction work after a successful run.",
    )
    parser.add_argument(
        "--print-file-names",
        action="store_true",
        help="Print the exact governed filenames and exit.",
    )
    return parser


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()
    if args.print_file_names:
        print(json.dumps(exact_file_names(), indent=2, ensure_ascii=False))
        return 0
    search_root = args.search_root.resolve()
    output_root = (
        args.output_root.resolve()
        if args.output_root is not None
        else search_root / "MRHPD Response 72 Automated Recovery Output"
    )
    try:
        state = recover(
            search_root,
            output_root,
            keep_work=args.keep_work,
            extract_project=not args.verify_only,
        )
    except BaseException as exc:
        print(f"\nMRHPD Response 72 recovery failed: {type(exc).__name__}: {exc}", file=sys.stderr)
        print(f"Status and logs are preserved under: {output_root}", file=sys.stderr)
        return 1
    print("\nMRHPD Response 72 recovery: PASSED")
    print(f"Complete restore: {output_root / EXPECTED_RESTORE_NAME}")
    if not args.verify_only:
        print(f"Current project: {output_root / 'CURRENT_PROJECT'}")
    print(f"Status: {output_root / STATUS_NAME}")
    print(json.dumps({
        "status": state.get("status"),
        "launcher_version": LAUNCHER_VERSION,
        "complete_restore_name": EXPECTED_RESTORE_NAME,
        "complete_restore_bytes": EXPECTED_RESTORE_BYTES,
        "complete_restore_sha256": EXPECTED_RESTORE_SHA256,
        "output_root": str(output_root),
    }, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
