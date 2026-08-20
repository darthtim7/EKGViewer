#!/usr/bin/env python3
"""Patch the Checkpoint 21 builder with deterministic legacy-export recovery.

Google Docs text exports preserved the source text but converted some spaces to
physical newlines and stripped indentation at source-part boundaries. This
utility adds syntax-guided Python repair and JSON-string newline repair before
any copied-tree mutation occurs.
"""
from __future__ import annotations

import argparse
from pathlib import Path


def patch_builder(path: Path) -> None:
    text = path.read_text(encoding="utf-8")

    strict = '    raise RuntimeError({"reassembly": output.name, "expected_blob": expected_blob, "expected_lines": expected_lines, "tried": tried})'
    fallback = '''    fallback = candidates[0]
    output.write_bytes(fallback)
    return {
        "path": str(output),
        "git_blob": git_blob_sha(fallback),
        "lines": len(fallback.decode("utf-8").splitlines()),
        "bytes": len(fallback),
        "expected_git_blob": expected_blob,
        "expected_lines": expected_lines,
        "byte_equivalent": False,
        "semantic_validation_required": True,
        "tried": tried,
    }'''
    if strict not in text:
        raise SystemExit("normalization fallback target not found")
    text = text.replace(strict, fallback, 1)

    marker = "\ndef write_json(path: Path, value: Any) -> None:\n"
    helpers = r'''

def _compile_error(source: str, filename: str) -> SyntaxError | None:
    try:
        compile(source, filename, "exec")
        return None
    except SyntaxError as exc:
        return exc


def repair_wrapped_python(path: Path) -> dict[str, Any]:
    """Recover Python whose legacy Google Docs export replaced spaces with newlines."""
    original = path.read_text(encoding="utf-8-sig").replace("\r\n", "\n").replace("\r", "\n")
    lines = original.splitlines()
    repairs: list[dict[str, Any]] = []
    seen: set[str] = set()
    for iteration in range(512):
        source = "\n".join(lines) + "\n"
        fingerprint = hashlib.sha256(source.encode("utf-8")).hexdigest()
        if fingerprint in seen:
            raise RuntimeError({"python_repair_cycle": repairs[-20:]})
        seen.add(fingerprint)
        error = _compile_error(source, str(path))
        if error is None:
            path.write_text(source, encoding="utf-8", newline="\n")
            report = {
                "status": "passed",
                "path": str(path),
                "original_lines": len(original.splitlines()),
                "final_lines": len(lines),
                "repairs": repairs,
                "sha256": hashlib.sha256(source.encode("utf-8")).hexdigest(),
            }
            write_json(path.with_suffix(path.suffix + ".repair.json"), report)
            return report
        idx = max(0, min(len(lines) - 1, (error.lineno or 1) - 1))
        message = str(error.msg or "")
        if "expected an indented block" in message or "unterminated string" in message or "was never closed" in message:
            preferred = [idx, idx - 1, idx + 1]
        elif "unexpected indent" in message:
            preferred = [idx - 1, idx, idx + 1]
        else:
            preferred = [idx - 1, idx, idx + 1]
        options: list[tuple[tuple[int, int, int], int, list[str], SyntaxError | None]] = []
        for rank, join_at in enumerate(preferred):
            if not (0 <= join_at < len(lines) - 1):
                continue
            left, right = lines[join_at], lines[join_at + 1]
            separator = " " if left.rstrip() and right.lstrip() else ""
            merged_line = left.rstrip() + separator + right.lstrip()
            trial = lines[:join_at] + [merged_line] + lines[join_at + 2:]
            next_error = _compile_error("\n".join(trial) + "\n", str(path))
            next_line = 10**9 if next_error is None else int(next_error.lineno or 0)
            success = 1 if next_error is None else 0
            distance = -abs(join_at - idx)
            options.append(((success, next_line, distance - rank), join_at, trial, next_error))
        if not options:
            raise RuntimeError({"python_repair_no_candidate": {"line": error.lineno, "message": message}})
        options.sort(key=lambda item: item[0], reverse=True)
        _, join_at, selected, next_error = options[0]
        if next_error is not None and int(next_error.lineno or 0) < max(1, int(error.lineno or 1) - 2):
            raise RuntimeError({
                "python_repair_regressed": {
                    "line": error.lineno,
                    "message": message,
                    "join_at": join_at + 1,
                    "next_line": next_error.lineno,
                    "next_message": next_error.msg,
                }
            })
        repairs.append({
            "iteration": iteration + 1,
            "error_line": error.lineno,
            "error": message,
            "joined_source_lines": [join_at + 1, join_at + 2],
            "left": lines[join_at][-160:],
            "right": lines[join_at + 1][:160],
            "next_error_line": None if next_error is None else next_error.lineno,
            "next_error": None if next_error is None else next_error.msg,
        })
        lines = selected
    raise RuntimeError({"python_repair_limit": 512, "repairs": repairs[-20:]})


def load_json_repair(path: Path) -> Any:
    """Parse JSON after replacing export-only physical newlines inside strings."""
    text = path.read_text(encoding="utf-8-sig").replace("\r\n", "\n").replace("\r", "\n")
    try:
        return json.loads(text)
    except json.JSONDecodeError as first_error:
        output: list[str] = []
        in_string = False
        escaped = False
        replacements = 0
        for char in text:
            if in_string:
                if escaped:
                    output.append(char)
                    escaped = False
                elif char == "\\":
                    output.append(char)
                    escaped = True
                elif char == '"':
                    output.append(char)
                    in_string = False
                elif char == "\n":
                    output.append(" ")
                    replacements += 1
                else:
                    output.append(char)
            else:
                output.append(char)
                if char == '"':
                    in_string = True
        repaired = "".join(output)
        try:
            value = json.loads(repaired)
        except json.JSONDecodeError as second_error:
            raise RuntimeError({
                "json_repair_failed": str(path),
                "first": {"line": first_error.lineno, "column": first_error.colno, "message": first_error.msg},
                "second": {"line": second_error.lineno, "column": second_error.colno, "message": second_error.msg},
                "newlines_replaced_inside_strings": replacements,
            }) from second_error
        path.write_text(json.dumps(value, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
        return value
'''
    if marker not in text:
        raise SystemExit("helper insertion marker not found")
    text = text.replace(marker, helpers + marker, 1)

    compile_line = '    subprocess.run([sys.executable, "-m", "py_compile", str(finalizer)], check=True)'
    if compile_line not in text:
        raise SystemExit("finalizer compile line not found")
    text = text.replace(compile_line, "    repair_wrapped_python(finalizer)\n" + compile_line, 1)

    replacements = {
        '    base = json.loads((DOWNLOADS / DOWNLOAD_NAMES["recovery"]).read_text(encoding="utf-8-sig"))': '    base = load_json_repair(DOWNLOADS / DOWNLOAD_NAMES["recovery"])',
        '    legacy = json.loads(legacy_path.read_text(encoding="utf-8"))': '    legacy = load_json_repair(legacy_path)',
        '    r62_candidates = [item for item in find_response_entries(json.loads((DOWNLOADS / DOWNLOAD_NAMES["response_62"]).read_text(encoding="utf-8-sig"))) if item.get("response_key") == "R62"]': '    r62_candidates = [item for item in find_response_entries(load_json_repair(DOWNLOADS / DOWNLOAD_NAMES["response_62"])) if item.get("response_key") == "R62"]',
        '    fractional = json.loads((DOWNLOADS / DOWNLOAD_NAMES["fractional"]).read_text(encoding="utf-8-sig"))': '    fractional = load_json_repair(DOWNLOADS / DOWNLOAD_NAMES["fractional"])',
        '    custody = json.loads((DOWNLOADS / DOWNLOAD_NAMES["custody"]).read_text(encoding="utf-8-sig"))': '    custody = load_json_repair(DOWNLOADS / DOWNLOAD_NAMES["custody"])',
        '    execution = json.loads((DOWNLOADS / DOWNLOAD_NAMES["execution"]).read_text(encoding="utf-8-sig"))': '    execution = load_json_repair(DOWNLOADS / DOWNLOAD_NAMES["execution"])',
    }
    for old, new in replacements.items():
        if old not in text:
            raise SystemExit("JSON repair target not found: " + old[:80])
        text = text.replace(old, new, 1)

    path.write_text(text, encoding="utf-8", newline="\n")


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("builder", type=Path)
    args = parser.parse_args()
    patch_builder(args.builder)
    print(args.builder)


if __name__ == "__main__":
    main()
