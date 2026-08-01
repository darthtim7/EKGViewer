#!/usr/bin/env python3
"""Create a four-volume Google Drive transport for the complete Response 84 project archive."""
from __future__ import annotations

import argparse
import hashlib
import json
import math
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path, PurePosixPath

PROJECT_BYTES = 334_963_130
PROJECT_SHA256 = "9641f41060433dfb02a9dcf5c899b8c0218f6ce085d490846c57872b588cfc2c"
PROJECT_NAME = (
    "Medical References - Human Pathogen Database v3.0.0a Part 8 of 8 "
    "Remediation Section 5 of 5 Session 3 of 3 ALL SECTIONS COMPLETE PROJECT "
    "THROUGH RESPONSE 84 2026-08-01 1816 UTC.zip"
)
VOLUME_COUNT = 4
MAX_PART_BYTES = 92 * 1024 * 1024


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
        if len(names) != len(set(names)):
            raise RuntimeError("duplicate ZIP members")
        for name in names:
            pp = PurePosixPath(name.replace("\\", "/"))
            if pp.is_absolute() or ".." in pp.parts:
                raise RuntimeError("unsafe ZIP path: " + name)
    if bad:
        raise RuntimeError("ZIP CRC failure: " + str(path))
    return {
        "name": path.name,
        "bytes": path.stat().st_size,
        "sha256": sha256_file(path),
        "members": len(names),
        "crc_error": bad,
    }


def write_text(path: Path, text: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text.rstrip() + "\n", encoding="utf-8")


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--project-archive", type=Path, required=True)
    parser.add_argument("--dist", type=Path, default=Path("dist_response84_project_transport"))
    args = parser.parse_args()
    if args.project_archive.name != PROJECT_NAME:
        raise RuntimeError({"project_archive_name": args.project_archive.name, "expected": PROJECT_NAME})
    if args.project_archive.stat().st_size != PROJECT_BYTES or sha256_file(args.project_archive) != PROJECT_SHA256:
        raise RuntimeError({
            "project_archive_identity": {
                "bytes": args.project_archive.stat().st_size,
                "sha256": sha256_file(args.project_archive),
                "expected_bytes": PROJECT_BYTES,
                "expected_sha256": PROJECT_SHA256,
            }
        })
    project_qa = verify_zip(args.project_archive)
    if args.dist.exists():
        shutil.rmtree(args.dist)
    args.dist.mkdir(parents=True)
    part_size = math.ceil(PROJECT_BYTES / VOLUME_COUNT)
    if part_size > MAX_PART_BYTES:
        raise RuntimeError({"part_size": part_size, "limit": MAX_PART_BYTES})
    parts: list[Path] = []
    with args.project_archive.open("rb") as source:
        written = 0
        for sequence in range(1, VOLUME_COUNT + 1):
            remaining = PROJECT_BYTES - written
            target_size = remaining if sequence == VOLUME_COUNT else min(part_size, remaining)
            part = args.dist / f"{PROJECT_NAME}.part{sequence:03d}"
            with part.open("wb") as target:
                left = target_size
                while left:
                    block = source.read(min(1024 * 1024, left))
                    if not block:
                        raise RuntimeError("unexpected end of project archive")
                    target.write(block)
                    left -= len(block)
            written += part.stat().st_size
            parts.append(part)
    if sum(path.stat().st_size for path in parts) != PROJECT_BYTES:
        raise RuntimeError("split byte total mismatch")
    part_rows = [
        {
            "sequence": sequence,
            "name": path.name,
            "bytes": path.stat().st_size,
            "sha256": sha256_file(path),
        }
        for sequence, path in enumerate(parts, start=1)
    ]
    manifest = {
        "schema": "mrhpd-response84-complete-project-transport-1.0",
        "project_archive": project_qa,
        "volume_count": VOLUME_COUNT,
        "parts": part_rows,
        "all_volumes_required": True,
        "user_upload_required": False,
        "all_sections_complete": True,
    }
    manifest_path = args.dist / "MRHPD_RESPONSE84_COMPLETE_PROJECT_TRANSPORT_MANIFEST.json"
    manifest_path.write_text(json.dumps(manifest, indent=2) + "\n", encoding="utf-8")
    reassembler = args.dist / "reassemble_response84_complete_project.py"
    write_text(reassembler, f'''#!/usr/bin/env python3
import hashlib,json
from pathlib import Path
MANIFEST={manifest!r}
def sha(path):
 h=hashlib.sha256()
 with path.open('rb') as handle:
  for block in iter(lambda:handle.read(1024*1024),b''): h.update(block)
 return h.hexdigest()
def main():
 root=Path.cwd(); output=root/MANIFEST['project_archive']['name']
 with output.open('wb') as target:
  for row in MANIFEST['parts']:
   part=root/row['name']
   if part.stat().st_size!=row['bytes'] or sha(part)!=row['sha256']: raise RuntimeError({{'part_identity_failed':row['name']}})
   with part.open('rb') as source:
    for block in iter(lambda:source.read(1024*1024),b''): target.write(block)
 expected=MANIFEST['project_archive']
 if output.stat().st_size!=expected['bytes'] or sha(output)!=expected['sha256']: raise RuntimeError('reassembled project identity failed')
 print(json.dumps({{'status':'passed','project_archive':str(output),'bytes':output.stat().st_size,'sha256':sha(output)}},indent=2))
if __name__=='__main__': main()
''')
    wrappers = []
    for sequence, part in enumerate(parts, start=1):
        readme = args.dist / f"README_RESPONSE84_COMPLETE_PROJECT_VOLUME_{sequence}.txt"
        write_text(
            readme,
            f"Human Pathogen Database Response 84 complete project volume {sequence} of {VOLUME_COUNT}. "
            "Extract all four governed volume ZIPs into the same otherwise-empty directory, then run "
            "reassemble_response84_complete_project.py. All four volumes are required.",
        )
        wrapper = args.dist / f"MRHPD v3.0.0a Response 84 Complete Project Drive Volume {sequence} of {VOLUME_COUNT}.zip"
        with zipfile.ZipFile(wrapper, "w", compression=zipfile.ZIP_STORED, allowZip64=True) as zf:
            zf.write(part, part.name)
            zf.write(manifest_path, manifest_path.name)
            zf.write(reassembler, reassembler.name)
            zf.write(readme, readme.name)
        wrappers.append({"sequence": sequence, "qa": verify_zip(wrapper), "path": wrapper})
    with tempfile.TemporaryDirectory(prefix="mrhpd-r84-project-transport-") as td:
        root = Path(td)
        for row in wrappers:
            with zipfile.ZipFile(row["path"]) as zf:
                zf.extractall(root)
        result = subprocess.run([sys.executable, str((root / reassembler.name).resolve())], cwd=root, text=True, capture_output=True, timeout=1800)
        if result.returncode:
            raise RuntimeError({"reassembly_failed": {"stdout": result.stdout, "stderr": result.stderr}})
        reconstructed = json.loads(result.stdout)
        archive = Path(reconstructed["project_archive"])
        verify_zip(archive)
        if archive.stat().st_size != PROJECT_BYTES or sha256_file(archive) != PROJECT_SHA256:
            raise RuntimeError("independent reassembly identity mismatch")
    summary = {
        "schema": "mrhpd-response84-complete-project-transport-verification-1.0",
        "status": "passed",
        "project_archive": project_qa,
        "wrappers": [{"sequence": row["sequence"], "qa": row["qa"]} for row in wrappers],
        "reassembly": reconstructed,
        "all_volumes_required": True,
        "user_upload_required": False,
        "all_sections_complete": True,
    }
    summary_path = args.dist / "MRHPD v3.0.0a Response 84 Complete Project Transport Verification.json"
    summary_path.write_text(json.dumps(summary, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(summary, indent=2))


if __name__ == "__main__":
    main()
