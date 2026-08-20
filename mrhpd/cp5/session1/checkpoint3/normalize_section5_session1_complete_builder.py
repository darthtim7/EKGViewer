#!/usr/bin/env python3
from pathlib import Path

path = Path(__file__).resolve().parent / "build_section5_session1_complete_restore.py"
text = path.read_text(encoding="utf-8")

old = '''        project_zip_qa = zip_tree(current_project, project_zip)
        if project_zip_qa["bytes"] >= MAX_ARCHIVE_BYTES:
            raise RuntimeError({"complete_project_exceeds_180_mib_after_compaction": project_zip_qa})
        clean_project = verify_project_clean_extract(project_zip, manifest_path.relative_to(current_project).as_posix(), {"workbook_sheets": workbook_qa["current_sheet_count"]})
        update_gate(gates, "project_archive", "passed", f"{project_zip_qa['bytes']} bytes; clean extraction passed", now_iso)
        persist_final_gates(current_db, gates)
        # Freeze final gate state and rebuild manifest/project archive one final time.
        manifest_path, checksums_path, manifest_rows = build_project_manifest(current_project, now_iso)
        project_zip_qa = zip_tree(current_project, project_zip)
        clean_project = verify_project_clean_extract(project_zip, manifest_path.relative_to(current_project).as_posix(), {"workbook_sheets": workbook_qa["current_sheet_count"]})
        direct_controls = [docx_report, pdf_report, xlsx_register, freeze_figure, compaction_json, compaction_csv, final_qa_path, manifest_path, checksums_path]
        restore_result = build_complete_restore(project_zip, manifest_path.relative_to(current_project).as_posix(), workbook_qa["current_sheet_count"], args.dist, stamp, direct_controls)
        update_gate(gates, "restore", "passed", f"{restore_result['qa']['bytes']} bytes; embedded verifier passed", now_iso)
        transport = build_transport(restore_result["restore"], args.dist)
        update_gate(gates, "transport", "passed", f"{len(transport['wrappers'])} volumes; reassembly passed", now_iso)
        persist_final_gates(current_db, gates)
'''
new = '''        # Preliminary complete-artifact cycle. It proves that the terminal project,
        # self-contained restore, and transport design all work before the final
        # pass/freeze states are written into the copied project database.
        project_zip_qa = zip_tree(current_project, project_zip)
        if project_zip_qa["bytes"] >= MAX_ARCHIVE_BYTES:
            raise RuntimeError({"complete_project_exceeds_180_mib_after_compaction": project_zip_qa})
        clean_project = verify_project_clean_extract(
            project_zip,
            manifest_path.relative_to(current_project).as_posix(),
            {"workbook_sheets": workbook_qa["current_sheet_count"]},
        )
        direct_controls = [
            docx_report, pdf_report, xlsx_register, freeze_figure,
            compaction_json, compaction_csv, final_qa_path,
            manifest_path, checksums_path,
        ]
        restore_result = build_complete_restore(
            project_zip,
            manifest_path.relative_to(current_project).as_posix(),
            workbook_qa["current_sheet_count"],
            args.dist,
            stamp,
            direct_controls,
        )
        transport = build_transport(restore_result["restore"], args.dist)

        # Only after the preliminary external-object cycle passes are the three
        # terminal packaging gates marked passed. Evidence points to the external
        # verification sidecar so the final project does not contain circular
        # self-hashes. Then the project, restore, and transport are rebuilt and
        # verified one final time without any subsequent project mutation.
        update_gate(gates, "project_archive", "passed", "External clean-extraction verifier passed; final identity is recorded in the Response 78 verification sidecar.", now_iso)
        update_gate(gates, "restore", "passed", "Embedded self-contained restore verifier passed; final identity is recorded in the Response 78 verification sidecar.", now_iso)
        update_gate(gates, "transport", "passed", f"{len(transport['wrappers'])} transport volumes reassembled exactly; final identities are recorded in the Response 78 verification sidecar.", now_iso)
        persist_final_gates(current_db, gates)
        qa_payload["acceptance_gates"] = gates
        qa_payload["status"] = "passed_with_controlled_external_gates"
        json_write(final_qa_path, qa_payload)

        manifest_path, checksums_path, manifest_rows = build_project_manifest(current_project, now_iso)
        project_zip_qa = zip_tree(current_project, project_zip)
        if project_zip_qa["bytes"] >= MAX_ARCHIVE_BYTES:
            raise RuntimeError({"final_complete_project_exceeds_180_mib": project_zip_qa})
        clean_project = verify_project_clean_extract(
            project_zip,
            manifest_path.relative_to(current_project).as_posix(),
            {"workbook_sheets": workbook_qa["current_sheet_count"]},
        )
        direct_controls = [
            docx_report, pdf_report, xlsx_register, freeze_figure,
            compaction_json, compaction_csv, final_qa_path,
            manifest_path, checksums_path,
        ]
        restore_result = build_complete_restore(
            project_zip,
            manifest_path.relative_to(current_project).as_posix(),
            workbook_qa["current_sheet_count"],
            args.dist,
            stamp,
            direct_controls,
        )
        transport = build_transport(restore_result["restore"], args.dist)
'''
if old not in text:
    raise SystemExit("terminal freeze block not found")
text = text.replace(old, new, 1)

# Convert long PDF-table text to Paragraph flowables so report cells wrap within
# their governed column widths rather than expanding beyond the page.
anchor = '''    small = ParagraphStyle("MRSmall", parent=body, fontSize=7.6, leading=9.4)
    document = SimpleDocTemplate'''
replacement = '''    small = ParagraphStyle("MRSmall", parent=body, fontSize=7.6, leading=9.4)
    def pdf_cell(value: Any, style: ParagraphStyle = small) -> Paragraph:
        escaped = str(value).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        return Paragraph(escaped, style)
    document = SimpleDocTemplate'''
if anchor not in text:
    raise SystemExit("PDF cell anchor not found")
text = text.replace(anchor, replacement, 1)
text = text.replace(
    '''    table = Table(terminal_data, colWidths=[1.3 * inch, 5.75 * inch], repeatRows=1)''',
    '''    terminal_data = [[pdf_cell(cell) for cell in row] for row in terminal_data]
    table = Table(terminal_data, colWidths=[1.3 * inch, 5.75 * inch], repeatRows=1)''',
    1,
)
text = text.replace(
    '''    gate_table = Table(gate_data, colWidths=[1.55 * inch, 1.1 * inch, 4.35 * inch], repeatRows=1)''',
    '''    gate_data = [[pdf_cell(cell) for cell in row] for row in gate_data]
    gate_table = Table(gate_data, colWidths=[1.55 * inch, 1.1 * inch, 4.35 * inch], repeatRows=1)''',
    1,
)
text = text.replace(
    '''    event_table = Table(event_data, colWidths=[0.55 * inch, 3.1 * inch, 3.35 * inch], repeatRows=1)''',
    '''    event_data = [[pdf_cell(cell) for cell in row] for row in event_data]
    event_table = Table(event_data, colWidths=[0.55 * inch, 3.1 * inch, 3.35 * inch], repeatRows=1)''',
    1,
)

path.write_text(text, encoding="utf-8")
print({"status": "passed", "builder": str(path)})
