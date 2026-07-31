#!/usr/bin/env python3
"""Prepare exact Response 66 and Response 68 inputs for Session 2 completion."""
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

BASE_BYTES = 177_617_796
BASE_SHA256 = "38c8fa08763d5698217ce33a2bbe1e889e726087575b14fb31086f38cfe1300f"
RECOVERY_BYTES = 18_318_469
RECOVERY_SHA256 = "b466c463c55dc95d2ac780ff78755f5ae09fa19d9a6be4fb97e914af7568adbe"
ADAPTER_PATH = Path("mrhpd/cp4/session2/checkpoint3/prepare_session2_complete_builder.py")


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


def reconstruct_base(response66_dir: Path, output_dir: Path, work: Path) -> Path:
    wrappers = sorted(response66_dir.rglob("*Complete Restore Drive Volume * of 2.zip"))
    if len(wrappers) != 2:
        raise RuntimeError({"response66_wrappers": [str(path) for path in wrappers]})
    staging = work / "response66_wrappers"
    for index, wrapper in enumerate(wrappers, 1):
        safe_extract(wrapper, staging / f"volume{index}")
    flat = work / "response66_flat"
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
    utilities = list(flat.glob("reassemble_response66_complete_restore.py"))
    if len(utilities) != 1:
        raise RuntimeError({"response66_reassembly_utilities": [str(path) for path in utilities]})
    result = subprocess.run(
        [sys.executable, str(utilities[0].resolve())],
        cwd=flat,
        text=True,
        capture_output=True,
        timeout=1800,
    )
    if result.returncode != 0:
        raise RuntimeError({"response66_reassembly_failed": {"stdout": result.stdout[-15000:], "stderr": result.stderr[-15000:]}})
    restores = [
        path
        for path in flat.glob("*.zip")
        if path.stat().st_size == BASE_BYTES and sha256_file(path) == BASE_SHA256
    ]
    if len(restores) != 1:
        raise RuntimeError({
            "response66_restore_candidates": [
                {"name": path.name, "bytes": path.stat().st_size, "sha256": sha256_file(path)}
                for path in flat.glob("*.zip")
            ]
        })
    output_dir.mkdir(parents=True, exist_ok=True)
    destination = output_dir / restores[0].name
    shutil.copy2(restores[0], destination)
    return destination


def extract_recovery(response68_dir: Path, output_dir: Path) -> Path:
    wrappers = list(response68_dir.rglob("*Response 68 Checkpoint 2 Recovery Delivery.zip"))
    if len(wrappers) != 1:
        raise RuntimeError({"response68_delivery_wrappers": [str(path) for path in wrappers]})
    safe_extract(wrappers[0], output_dir)
    inner = [
        path
        for path in output_dir.rglob("*RECOVERY DATA THROUGH RESPONSE 68*.zip")
        if path.stat().st_size == RECOVERY_BYTES and sha256_file(path) == RECOVERY_SHA256
    ]
    if len(inner) != 1:
        raise RuntimeError({
            "response68_recovery_candidates": [
                {"name": path.name, "bytes": path.stat().st_size, "sha256": sha256_file(path)}
                for path in output_dir.rglob("*.zip")
            ]
        })
    return inner[0]


def patch_builder_adapter(path: Path = ADAPTER_PATH) -> dict[str, object]:
    """Correct the nested generated-verifier string without altering governed logic."""
    text = path.read_text(encoding="utf-8")
    original = text
    old_open = "    verifier = f'''#!/usr/bin/env python3\n"
    new_open = '    verifier = f"""#!/usr/bin/env python3\n'
    old_close = "\n'''\n    text_write(tools / \"restore_verify_extract.py\", verifier)\n"
    new_close = '\n"""\n    text_write(tools / "restore_verify_extract.py", verifier)\n'
    if old_open in text:
        text = text.replace(old_open, new_open, 1)
    elif new_open not in text:
        raise RuntimeError("generated verifier opening delimiter target not found")
    if old_close in text:
        text = text.replace(old_close, new_close, 1)
    elif new_close not in text:
        raise RuntimeError("generated verifier closing delimiter target not found")
    if text != original:
        path.write_text(text, encoding="utf-8")
    return {
        "status": "passed",
        "path": path.as_posix(),
        "patched": text != original,
        "sha256": sha256_file(path),
    }


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--response66-dir", type=Path, required=True)
    parser.add_argument("--response68-dir", type=Path, required=True)
    parser.add_argument("--base-output", type=Path, required=True)
    parser.add_argument("--recovery-output", type=Path, required=True)
    parser.add_argument("--work", type=Path, default=Path("prepared_input_work"))
    args = parser.parse_args()
    if args.work.exists():
        shutil.rmtree(args.work)
    if args.base_output.exists():
        shutil.rmtree(args.base_output)
    if args.recovery_output.exists():
        shutil.rmtree(args.recovery_output)
    args.work.mkdir(parents=True)
    base = reconstruct_base(args.response66_dir, args.base_output, args.work)
    recovery = extract_recovery(args.response68_dir, args.recovery_output)
    adapter = patch_builder_adapter()
    print(json.dumps({
        "status": "passed",
        "base_restore": {"name": base.name, "bytes": base.stat().st_size, "sha256": sha256_file(base)},
        "checkpoint2_recovery": {"name": recovery.name, "bytes": recovery.stat().st_size, "sha256": sha256_file(recovery)},
        "builder_adapter": adapter,
    }, indent=2))


if __name__ == "__main__":
    main()
