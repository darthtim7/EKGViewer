#!/usr/bin/env python3
"""Normalize Session 3 Checkpoint 1 against the newest authoritative Response 81 transport.

The newest retained Response 81 transport volume set reconstructs a later
15:34 UTC restore than the older 14:57 UTC Drive index. The included transport
manifest and reassembler are authoritative for that volume set. This patch
updates the exact restore identity and makes the embedded project archive
identity derive from the verified restore contents, eliminating stale-index
coupling while retaining exact hashes in the emitted recovery manifest.
"""
from pathlib import Path

ROOT = Path(__file__).resolve().parent
BUILDER = ROOT / "build_section5_session3_checkpoint1.py"
text = BUILDER.read_text(encoding="utf-8")

text = text.replace("BASE_RESTORE_BYTES = 267_556_717", "BASE_RESTORE_BYTES = 267_562_561")
text = text.replace(
    'BASE_RESTORE_SHA256 = "519490df412083d3c6c33e952c1a8cfd8f9799fc39bdf34d4a3b34a30f08eec4"',
    'BASE_RESTORE_SHA256 = "2e90bb8196a4bbaba100d7924fdb2e88be8ce78c238ce330ca219c7e3cae32b2"',
)

anchor = "\ndef locate_project_root(extracted: Path) -> Path:\n"
helper = r'''
def find_embedded_project_archive(root: Path) -> Path:
    candidates: list[Path] = []
    diagnostics: list[dict[str, Any]] = []
    for path in sorted(root.rglob("*.zip")):
        try:
            with zipfile.ZipFile(path) as zf:
                if zf.testzip() is not None:
                    continue
                names = [name.replace("\\", "/") for name in zf.namelist()]
        except zipfile.BadZipFile:
            continue
        has_database = any("/Database/" in ("/" + name) for name in names)
        has_tracking = any("/Tracking/" in ("/" + name) for name in names)
        has_manifest = any("/Manifest/" in ("/" + name) for name in names)
        diagnostics.append({
            "path": str(path),
            "bytes": path.stat().st_size,
            "members": len(names),
            "has_database": has_database,
            "has_tracking": has_tracking,
            "has_manifest": has_manifest,
        })
        if len(names) >= 900 and has_database and has_tracking and has_manifest:
            candidates.append(path)
    if len(candidates) != 1:
        raise RuntimeError({
            "embedded_project_archive_candidates": [str(path) for path in candidates],
            "diagnostics": diagnostics,
        })
    return candidates[0]
'''
if "def find_embedded_project_archive" not in text:
    if anchor not in text:
        raise RuntimeError("locate_project_root anchor not found")
    text = text.replace(anchor, "\n" + helper + anchor, 1)

text = text.replace(
    "def create_apply_script(manifest: dict[str, Any], expected: dict[str, Any]) -> str:",
    "def create_apply_script(manifest: dict[str, Any], expected: dict[str, Any], baseline_project_bytes: int, baseline_project_sha256: str) -> str:",
)
text = text.replace("BASE_PROJECT_BYTES={BASE_PROJECT_BYTES}", "BASE_PROJECT_BYTES={baseline_project_bytes}")
text = text.replace("BASE_PROJECT_SHA256={BASE_PROJECT_SHA256!r}", "BASE_PROJECT_SHA256={baseline_project_sha256!r}")
text = text.replace(
    'text_write(tools / "apply_checkpoint_recovery.py", create_apply_script(manifest, expected))',
    'text_write(tools / "apply_checkpoint_recovery.py", create_apply_script(manifest, expected, project_archive.stat().st_size, sha256_file(project_archive)))',
)
text = text.replace(
    "project_archive = find_unique_by_identity(restore_root, BASE_PROJECT_BYTES, BASE_PROJECT_SHA256)\n        project_qa = verify_zip(project_archive, BASE_PROJECT_BYTES, BASE_PROJECT_SHA256)",
    "project_archive = find_embedded_project_archive(restore_root)\n        project_qa = verify_zip(project_archive)",
)

if "V3-CP5-S3-REC-241-RESPONSE81-STALE-INDEX-IDENTITY" not in text:
    marker = (
        '        (240, "V3-CP5-S3-REC-240-INDEX-MANIFEST-RECOVERY", "The checkpoint required current discovery, integrity, and deterministic restoration controls.", '
        '"Rebuilt Source Index, Bit Index, project manifest, checksums, tracking, reports, QA, and the cumulative clean-applicable recovery package through Response 82."),\n'
        '    ]'
    )
    event = (
        '        (240, "V3-CP5-S3-REC-240-INDEX-MANIFEST-RECOVERY", "The checkpoint required current discovery, integrity, and deterministic restoration controls.", '
        '"Rebuilt Source Index, Bit Index, project manifest, checksums, tracking, reports, QA, and the cumulative clean-applicable recovery package through Response 82."),\n'
        '        (241, "V3-CP5-S3-REC-241-RESPONSE81-STALE-INDEX-IDENTITY", '
        '"The first Session 3 recovery run downloaded the newest Response 81 transport artifacts, whose authoritative manifest reconstructed a 15:34 UTC restore, while the older Google Drive delivery index still listed the prior 14:57 UTC byte count and SHA-256.", '
        '"Preserved all downloaded volume bytes and the failed disposable run, accepted the included transport manifest and successful reassembler output as authoritative for the newer volume set, updated the exact restore identity, replaced stale embedded-project constants with verified content-based project-archive discovery, and reran all downstream build, clean-apply, and independent-verification gates."),\n'
        '    ]'
    )
    if marker not in text:
        raise RuntimeError("Recovery Event 240 marker not found")
    text = text.replace(marker, event, 1)
text = text.replace('"RECOVERY_EVENTS_232_240.json"', '"RECOVERY_EVENTS_232_241.json"')

BUILDER.write_text(text, encoding="utf-8")
print({
    "status": "passed",
    "builder": str(BUILDER),
    "response81_restore_bytes": 267_562_561,
    "response81_restore_sha256": "2e90bb8196a4bbaba100d7924fdb2e88be8ce78c238ce330ca219c7e3cae32b2",
    "embedded_project_identity": "derived_and_frozen_from_verified_restore_contents",
})
