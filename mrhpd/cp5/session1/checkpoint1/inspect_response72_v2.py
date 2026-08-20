#!/usr/bin/env python3
from __future__ import annotations

import json
import zipfile
from pathlib import Path

import inspect_response72 as base


def expand_until_part(source_dir: Path, sequence: int, work: Path) -> tuple[Path, list[Path]]:
    suffix = f".part{sequence:03d}"
    direct = list(source_dir.rglob(f"*{suffix}"))
    if len(direct) == 1:
        manifests = list(source_dir.rglob("*TRANSPORT_MANIFEST.json"))
        return direct[0], manifests
    if len(direct) > 1:
        raise RuntimeError({"duplicate_direct_parts": [str(p) for p in direct]})

    queue = list(source_dir.rglob("*.zip"))
    seen: set[Path] = set()
    manifests: list[Path] = []
    attempt = 0
    while queue:
        archive = queue.pop(0)
        resolved = archive.resolve()
        if resolved in seen:
            continue
        seen.add(resolved)
        attempt += 1
        destination = work / f"sequence_{sequence}_layer_{attempt}"
        base.safe_extract(archive, destination)
        manifests.extend(destination.rglob("*TRANSPORT_MANIFEST.json"))
        parts = list(destination.rglob(f"*{suffix}"))
        if len(parts) == 1:
            return parts[0], manifests
        if len(parts) > 1:
            raise RuntimeError({"duplicate_nested_parts": [str(p) for p in parts]})
        queue.extend(destination.rglob("*.zip"))
        if attempt > 12:
            raise RuntimeError("nested ZIP depth/volume search exceeded")
    raise RuntimeError({"missing_part": suffix, "source_dir": str(source_dir)})


def reconstruct_restore(volume1_dir: Path, volume2_dir: Path, work: Path) -> Path:
    part1, manifests1 = expand_until_part(volume1_dir, 1, work / "search1")
    part2, manifests2 = expand_until_part(volume2_dir, 2, work / "search2")
    manifest_paths = manifests1 + manifests2 + list(volume1_dir.rglob("*TRANSPORT_MANIFEST.json")) + list(volume2_dir.rglob("*TRANSPORT_MANIFEST.json"))
    if not manifest_paths:
        raise RuntimeError("transport manifest absent")
    manifest = json.loads(manifest_paths[0].read_text(encoding="utf-8-sig"))
    output = work / manifest["restore"]["name"]
    with output.open("wb") as dst:
        for part in (part1, part2):
            with part.open("rb") as src:
                for block in iter(lambda: src.read(1024 * 1024), b""):
                    dst.write(block)
    identity = {"bytes": output.stat().st_size, "sha256": base.sha256_file(output)}
    if identity["bytes"] != base.RESTORE_BYTES or identity["sha256"] != base.RESTORE_SHA256:
        raise RuntimeError({"restore_identity": identity})
    with zipfile.ZipFile(output) as zf:
        base.safe_members(zf)
    return output


base.reconstruct_restore = reconstruct_restore

if __name__ == "__main__":
    base.main()
