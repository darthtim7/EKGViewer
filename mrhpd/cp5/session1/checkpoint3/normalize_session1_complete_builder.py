#!/usr/bin/env python3
from __future__ import annotations

import hashlib
from pathlib import Path

HERE = Path(__file__).resolve().parent
TARGET = HERE / "build_section5_session1_complete.py"


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    digest.update(path.read_bytes())
    return digest.hexdigest()


def replace_once(text: str, old: str, new: str, label: str) -> tuple[str, bool]:
    if new in text:
        return text, False
    count = text.count(old)
    if count != 1:
        raise RuntimeError({"normalization_anchor_failure": label, "count": count})
    return text.replace(old, new, 1), True


def main() -> None:
    text = TARGET.read_text(encoding="utf-8")
    applied: list[str] = []

    text, changed = replace_once(
        text,
        '''    wb.calculation.fullCalcOnLoad = True
    wb.calculation.forceFullCalc = True
    wb.save(destination)
''',
        '''    # Some inherited workbooks legitimately omit calcPr. Initialize it
    # rather than assuming openpyxl returned a calculation-properties object.
    if wb.calculation is None:
        from openpyxl.workbook.properties import CalcProperties
        wb.calculation = CalcProperties(calcMode="auto")
    wb.calculation.fullCalcOnLoad = True
    wb.calculation.forceFullCalc = True
    wb.save(destination)
''',
        "initialize missing workbook calculation properties",
    )
    if changed:
        applied.append("initialize missing workbook calculation properties")

    text, changed = replace_once(
        text,
        '''    pk_columns = {row[1] for row in info if row[5]}
    source = con.execute(f"SELECT * FROM {table} WHERE {where_column}=?", (source_value,)).fetchone()
''',
        '''    # Exclude only integer rowid-style primary keys. Text business keys
    # such as checkpoint_code must remain in the cloned insert.
    pk_columns = {row[1] for row in info if row[5] and str(row[2] or "").upper() == "INTEGER"}
    source = con.execute(f"SELECT * FROM {table} WHERE {where_column}=?", (source_value,)).fetchone()
''',
        "preserve text business-key primary columns during row cloning",
    )
    if changed:
        applied.append("preserve text business-key primary columns during row cloning")

    text, changed = replace_once(
        text,
        '''    for row in rows:
        ws.append([json.dumps(row.get(header), ensure_ascii=False) if isinstance(row.get(header), (list, dict)) else row.get(header) for header in headers])
''',
        '''    for row in rows:
        values = []
        for header in headers:
            value = row.get(header)
            if isinstance(value, (list, dict, tuple, set)):
                value = json.dumps(value, ensure_ascii=False, default=str)
            elif isinstance(value, Path):
                value = value.as_posix()
            elif not isinstance(value, (str, int, float, bool, type(None))):
                value = str(value)
            values.append(value)
        ws.append(values)
''',
        "normalize path and structured values before Excel insertion",
    )
    if changed:
        applied.append("normalize path and structured values before Excel insertion")

    text, changed = replace_once(
        text,
        '''                (RELEASE_CODE, key, path.relative_to(path.parents[len(path.parts) - len(path.parts)] if False else destination.parents[1]).as_posix() if False else str(path), path.stat().st_size, sha256_file(path), int(immutable), "passed", now_iso),
''',
        '''                (RELEASE_CODE, key, path.relative_to(destination.parents[1]).as_posix(), path.stat().st_size, sha256_file(path), int(immutable), "passed", now_iso),
''',
        "store governed project-relative freeze paths",
    )
    if changed:
        applied.append("store governed project-relative freeze paths")

    old_manifest_block = '''        index_result = build_indexes(current_project, now_iso)
        manifest_path, checksums_path, manifest_rows = build_manifest(current_project, now_iso)
        final_qa["indexes"] = index_result["qa"]
        final_qa["manifest_records"] = len(manifest_rows)
        json_write(qa_root / "SECTION5_SESSION1_COMPLETE_QA.json", final_qa)
        # Rebuild indexes and manifest after the final QA record is frozen.
        index_result = build_indexes(current_project, now_iso)
        manifest_path, checksums_path, manifest_rows = build_manifest(current_project, now_iso)
        final_qa["indexes"] = index_result["qa"]
        final_qa["manifest_records"] = len(manifest_rows)
        json_write(qa_root / "SECTION5_SESSION1_COMPLETE_QA_EXTERNAL.json", final_qa)

'''
    new_manifest_block = '''        # Freeze the last in-project QA bytes before generating the controlling
        # indexes and manifest. The fully populated final QA remains an external
        # restore/delivery control and is not written back into the manifested tree.
        final_qa["indexes"] = {"status": "pending_final_rebuild"}
        final_qa["manifest_records"] = None
        json_write(qa_root / "SECTION5_SESSION1_COMPLETE_QA.json", final_qa)
        index_result = build_indexes(current_project, now_iso)
        manifest_path, checksums_path, manifest_rows = build_manifest(current_project, now_iso)
        final_qa["indexes"] = index_result["qa"]
        final_qa["manifest_records"] = len(manifest_rows)

'''
    text, changed = replace_once(
        text,
        old_manifest_block,
        new_manifest_block,
        "freeze in-project bytes before final index and manifest",
    )
    if changed:
        applied.append("freeze in-project bytes before final index and manifest")

    TARGET.write_text(text, encoding="utf-8")
    print({
        "status": "passed",
        "target": str(TARGET),
        "applied": applied,
        "sha256": sha256_file(TARGET),
    })


if __name__ == "__main__":
    main()
