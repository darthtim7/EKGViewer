#!/usr/bin/env python3
"""Patch the Session 1 completion builder with a package-manifest-verified overlay fallback.

Checkpoint 2 was independently package-verified, but its embedded apply utility can
reject a late-synchronized pointer file because the earlier overlay row retained the
pre-synchronization hash. This patch never trusts the stale row: it verifies every
actual overlay byte against PACKAGE_MANIFEST.json, reconstructs the exact base tree,
applies the package-verified overlay, and then requires the governed Checkpoint 2
database/workbook/application/publication identities before continuing.
"""
from pathlib import Path

path = Path(__file__).with_name("build_session1_complete_restore.py")
text = path.read_text(encoding="utf-8")
old = '''    if result.returncode != 0:
        raise RuntimeError({"checkpoint2_apply_failed": apply_record})

    db = find_by_hash(output, "*.sqlite", CP2_DATABASE_SHA256)
'''
new = '''    overlay_fallback = None
    if result.returncode != 0:
        # The Checkpoint 2 package itself passed CRC, package-manifest and checksum
        # verification. Recover from a stale overlay-row hash by validating the
        # actual overlay bytes against the package-level manifest, then require the
        # exact governed Checkpoint 2 output identities below.
        baseline = json.loads((recovery_package / "BASELINE_IDENTITY.json").read_text(encoding="utf-8"))
        overlay_manifest = json.loads((recovery_package / "CHECKPOINT_RECOVERY_MANIFEST.json").read_text(encoding="utf-8"))
        package_manifest = json.loads((recovery_package / "PACKAGE_MANIFEST.json").read_text(encoding="utf-8"))
        package_rows = {row["path"]: row for row in package_manifest["files"]}
        package_mismatches = []
        for rel, row in package_rows.items():
            p = recovery_package / rel
            if not p.exists() or p.stat().st_size != row["bytes"] or sha256_file(p) != row["sha256"]:
                package_mismatches.append(rel)
        if package_mismatches:
            raise RuntimeError({"checkpoint2_package_manifest_mismatches": package_mismatches[:30], "original_apply": apply_record})

        fallback_restore_root = work / "fallback_response64"
        safe_extract(base_restore, fallback_restore_root)
        expected_final = baseline["frozen_final_section3_release"]
        finals = [
            p for p in fallback_restore_root.rglob("*.zip")
            if p.stat().st_size == int(expected_final["bytes"]) and sha256_file(p) == expected_final["sha256"]
        ]
        if len(finals) != 1:
            raise RuntimeError({"checkpoint2_fallback_final_release_matches": [str(p) for p in finals], "original_apply": apply_record})
        project_extract = work / "fallback_section3_project"
        safe_extract(finals[0], project_extract)
        dirs = [p for p in project_extract.iterdir() if p.is_dir()]
        files = [p for p in project_extract.iterdir() if p.is_file()]
        source_project = dirs[0] if len(dirs) == 1 and not files else project_extract
        if output.exists():
            shutil.rmtree(output)
        shutil.copytree(source_project, output)

        overlay_root = recovery_package / "OVERLAY" / overlay_manifest["project_root_name"]
        stale_overlay_rows = []
        copied = 0
        for row in overlay_manifest["files"]:
            src = overlay_root / row["path"]
            package_rel = src.relative_to(recovery_package).as_posix()
            package_row = package_rows.get(package_rel)
            if not package_row:
                raise RuntimeError({"checkpoint2_overlay_not_in_package_manifest": package_rel})
            actual = {"bytes": src.stat().st_size, "sha256": sha256_file(src)}
            if actual["bytes"] != package_row["bytes"] or actual["sha256"] != package_row["sha256"]:
                raise RuntimeError({"checkpoint2_overlay_package_identity_failure": package_rel, "actual": actual, "package": package_row})
            if actual["bytes"] != row["bytes"] or actual["sha256"] != row["sha256"]:
                stale_overlay_rows.append({"path": row["path"], "overlay_row": {"bytes": row["bytes"], "sha256": row["sha256"]}, "actual_package_verified": actual})
            dst = output / row["path"]
            dst.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(src, dst)
            copied += 1
        for rel in overlay_manifest.get("deleted_paths", []):
            target = output / rel
            if target.is_file() or target.is_symlink():
                target.unlink()
            elif target.is_dir():
                shutil.rmtree(target)
        overlay_fallback = {
            "status": "passed",
            "reason": result.stderr.strip() or "embedded apply utility returned nonzero",
            "package_manifest_records_verified": len(package_rows),
            "overlay_files_copied": copied,
            "stale_overlay_rows": stale_overlay_rows,
            "deleted_paths_applied": len(overlay_manifest.get("deleted_paths", [])),
        }
        RECOVERY_EVENTS.append({
            "event_number": 98,
            "event_code": "V3-CP4-S1-REC-CHECKPOINT2-STALE-OVERLAY-ROW-RECOVERED",
            "occurred_at": NOW,
            "failed_step": "Execute the embedded Checkpoint 2 recovery utility without an overlay-row discrepancy.",
            "exact_error_or_reason": result.stderr.strip() or "The embedded utility returned a nonzero status.",
            "intact_artifacts": "The complete Response 64 restore, independently verified Checkpoint 2 package, package manifest, all overlay bytes, accepted predecessor and frozen Section 3 release remained intact.",
            "recovery_action": "Verified the complete Checkpoint 2 package against PACKAGE_MANIFEST.json, reconstructed the immutable base, applied every package-verified overlay file, recorded any stale overlay-row metadata, and required all governed Checkpoint 2 output hashes before continuing.",
            "validation_result": "Recovery passes only when the exact Checkpoint 2 database, workbook, application, publication and editable-assembly identities are reproduced.",
            "data_quality_effect": "None; the stale metadata row was not trusted and no accepted artifact was changed.",
            "next_checkpoint": "Continue Session 1 completion after exact Checkpoint 2 state verification.",
        })

    db = find_by_hash(output, "*.sqlite", CP2_DATABASE_SHA256)
'''
if old not in text:
    raise SystemExit("apply_checkpoint2 patch anchor not found")
text = text.replace(old, new, 1)
old2 = '''        "apply_utility": apply_record,
        "database": db.relative_to(output).as_posix(),
'''
new2 = '''        "apply_utility": apply_record,
        "package_manifest_overlay_fallback": overlay_fallback,
        "database": db.relative_to(output).as_posix(),
'''
if old2 not in text:
    raise SystemExit("apply record patch anchor not found")
text = text.replace(old2, new2, 1)
path.write_text(text, encoding="utf-8")
print({"checkpoint2_overlay_fallback_patch": "applied", "builder": str(path)})
