#!/usr/bin/env python3
from __future__ import annotations

from collections import Counter
from pathlib import Path
import argparse
import hashlib
import json
import sqlite3

from openpyxl import load_workbook
from pypdf import PdfReader


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def inspect_project(base: Path, output: Path) -> dict:
    dirs = [path for path in base.iterdir() if path.is_dir()]
    project = dirs[0] if len(dirs) == 1 else base
    files = sorted(path for path in project.rglob("*") if path.is_file())
    by_ext = Counter((path.suffix.lower() or "<none>") for path in files)

    sqlite_candidates = sorted(
        project.rglob("*.sqlite"),
        key=lambda path: (path.stat().st_mtime_ns, path.stat().st_size),
        reverse=True,
    )
    db_summaries: list[dict] = []
    for db in sqlite_candidates:
        try:
            con = sqlite3.connect(f"file:{db}?mode=ro", uri=True)
            integrity = con.execute("PRAGMA integrity_check").fetchone()[0]
            fk = list(con.execute("PRAGMA foreign_key_check"))
            tables = [
                row[0]
                for row in con.execute(
                    "SELECT name FROM sqlite_master "
                    "WHERE type='table' AND name NOT LIKE 'sqlite_%' ORDER BY name"
                )
            ]
            views = [
                row[0]
                for row in con.execute(
                    "SELECT name FROM sqlite_master WHERE type='view' ORDER BY name"
                )
            ]
            selected: dict[str, dict] = {}
            for name in tables:
                lname = name.lower()
                if any(
                    term in lname
                    for term in (
                        "section4",
                        "response",
                        "recovery",
                        "checkpoint",
                        "release",
                        "tracking",
                        "source",
                        "graphic",
                        "publication",
                        "cover",
                        "print",
                        "workbook",
                        "application",
                    )
                ):
                    cols = [
                        row[1]
                        for row in con.execute(f'PRAGMA table_info("{name}")')
                    ]
                    try:
                        count = con.execute(
                            f'SELECT COUNT(*) FROM "{name}"'
                        ).fetchone()[0]
                    except Exception:
                        count = None
                    selected[name] = {"columns": cols, "row_count": count}
            db_summaries.append(
                {
                    "path": db.relative_to(project).as_posix(),
                    "bytes": db.stat().st_size,
                    "sha256": sha256(db),
                    "integrity": integrity,
                    "foreign_key_violations": len(fk),
                    "table_count": len(tables),
                    "view_count": len(views),
                    "tables": tables,
                    "selected_tables": selected,
                }
            )
            con.close()
        except Exception as exc:
            db_summaries.append(
                {"path": db.relative_to(project).as_posix(), "error": repr(exc)}
            )

    xlsx_candidates = sorted(
        project.rglob("*.xlsx"), key=lambda path: path.stat().st_size, reverse=True
    )
    workbook_summaries: list[dict] = []
    for wb_path in xlsx_candidates[:16]:
        try:
            wb = load_workbook(wb_path, read_only=True, data_only=False)
            sheets = list(wb.sheetnames)
            formulas = 0
            for ws in wb.worksheets:
                for row in ws.iter_rows():
                    for cell in row:
                        if isinstance(cell.value, str) and cell.value.startswith("="):
                            formulas += 1
            workbook_summaries.append(
                {
                    "path": wb_path.relative_to(project).as_posix(),
                    "bytes": wb_path.stat().st_size,
                    "sha256": sha256(wb_path),
                    "sheet_count": len(sheets),
                    "sheets": sheets,
                    "formula_cells": formulas,
                }
            )
            wb.close()
        except Exception as exc:
            workbook_summaries.append(
                {"path": wb_path.relative_to(project).as_posix(), "error": repr(exc)}
            )

    pdf_summaries: list[dict] = []
    for pdf in sorted(
        project.rglob("*.pdf"), key=lambda path: path.stat().st_size, reverse=True
    )[:32]:
        try:
            reader = PdfReader(str(pdf))
            pdf_summaries.append(
                {
                    "path": pdf.relative_to(project).as_posix(),
                    "bytes": pdf.stat().st_size,
                    "sha256": sha256(pdf),
                    "pages": len(reader.pages),
                }
            )
        except Exception as exc:
            pdf_summaries.append(
                {"path": pdf.relative_to(project).as_posix(), "error": repr(exc)}
            )

    cover_terms = (
        "cover",
        "spine",
        "wrap",
        "front",
        "back",
        "barcode",
        "printer",
        "press",
        "trim",
        "bleed",
        "cmyk",
        "icc",
    )
    cover_files: list[dict] = []
    for path in files:
        rel = path.relative_to(project).as_posix()
        if any(term in rel.lower() for term in cover_terms):
            cover_files.append(
                {"path": rel, "bytes": path.stat().st_size, "sha256": sha256(path)}
            )

    instruction_files: list[dict] = []
    for path in files:
        if path.name.lower() in ("instructions.txt", "instructions.md"):
            instruction_files.append(
                {
                    "path": path.relative_to(project).as_posix(),
                    "bytes": path.stat().st_size,
                    "sha256": sha256(path),
                }
            )

    top_level = []
    for path in sorted(
        project.iterdir(), key=lambda item: (not item.is_dir(), item.name.lower())
    ):
        top_level.append(
            {
                "name": path.name,
                "type": "directory" if path.is_dir() else "file",
                "bytes": None if path.is_dir() else path.stat().st_size,
            }
        )

    result = {
        "schema": "mrhpd-section5-intake-inspection-1.0",
        "status": "passed",
        "project_root": project.as_posix(),
        "file_count": len(files),
        "total_bytes": sum(path.stat().st_size for path in files),
        "extension_counts": dict(sorted(by_ext.items())),
        "top_level": top_level,
        "instruction_files": instruction_files,
        "sqlite": db_summaries,
        "workbooks": workbook_summaries,
        "pdfs": pdf_summaries,
        "cover_and_print_related_files": cover_files,
        "accepted_predecessor_mutated": False,
    }
    output.mkdir(parents=True, exist_ok=True)
    (output / "SECTION5_INTAKE_INSPECTION.json").write_text(
        json.dumps(result, indent=2, ensure_ascii=False), encoding="utf-8"
    )
    (output / "SECTION5_PROJECT_PATHS.txt").write_text(
        "\n".join(path.relative_to(project).as_posix() for path in files),
        encoding="utf-8",
    )
    return result


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--project", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    result = inspect_project(args.project, args.output)
    print(
        json.dumps(
            {
                "status": result["status"],
                "project_files": result["file_count"],
                "sqlite_candidates": len(result["sqlite"]),
                "workbook_candidates": len(result["workbooks"]),
                "pdf_candidates": len(result["pdfs"]),
                "cover_print_related": len(result["cover_and_print_related_files"]),
            },
            indent=2,
        )
    )


if __name__ == "__main__":
    main()
