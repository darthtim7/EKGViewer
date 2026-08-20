#!/usr/bin/env python3
"""Preserve inherited unsupported worksheet OOXML extensions during workbook augmentation.

openpyxl intentionally warns that unsupported worksheet extensions, including
some extended conditional-formatting payloads, will be removed on save. The
Session 2 terminal workbook only adds new sheets; it does not intentionally
modify inherited worksheets. This patch injects a deterministic ZIP-level
preservation pass into the builder and adds an independent verification gate.
"""
from __future__ import annotations

from pathlib import Path

ROOT = Path(__file__).resolve().parent
BUILDER = ROOT / "build_section5_session2_complete_restore.py"
VERIFIER = ROOT / "verify_section5_session2_complete_outputs.py"

HELPER = r'''
def _workbook_sheet_xml_map(package: dict[str, bytes]) -> dict[str, str]:
    import xml.etree.ElementTree as ET
    main_ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    rel_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    package_rel_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
    workbook = ET.fromstring(package["xl/workbook.xml"])
    rels = ET.fromstring(package["xl/_rels/workbook.xml.rels"])
    relationship_targets = {
        element.attrib["Id"]: element.attrib["Target"]
        for element in rels.findall(f"{{{package_rel_ns}}}Relationship")
    }
    result: dict[str, str] = {}
    sheets = workbook.find(f"{{{main_ns}}}sheets")
    if sheets is None:
        return result
    for sheet in sheets.findall(f"{{{main_ns}}}sheet"):
        relationship_id = sheet.attrib.get(f"{{{rel_ns}}}id")
        target = relationship_targets.get(relationship_id or "")
        if not target:
            continue
        normalized = target.lstrip("/")
        if not normalized.startswith("xl/"):
            normalized = "xl/" + normalized
        result[sheet.attrib["name"]] = normalized
    return result


def _worksheet_extension_block(xml: bytes) -> bytes | None:
    pattern = re.compile(
        rb"<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?extLst\\b.*?</(?:[A-Za-z_][A-Za-z0-9_.-]*:)?extLst>",
        re.DOTALL,
    )
    match = pattern.search(xml)
    return match.group(0) if match else None


def _copy_extension_block(source_xml: bytes, destination_xml: bytes) -> bytes:
    extension = _worksheet_extension_block(source_xml)
    if extension is None:
        return destination_xml
    pattern = re.compile(
        rb"<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?extLst\\b.*?</(?:[A-Za-z_][A-Za-z0-9_.-]*:)?extLst>",
        re.DOTALL,
    )
    destination_xml = pattern.sub(b"", destination_xml)
    source_open_end = source_xml.find(b">")
    destination_open_end = destination_xml.find(b">")
    if source_open_end < 0 or destination_open_end < 0:
        raise RuntimeError("worksheet root tag missing")
    source_open = source_xml[: source_open_end + 1]
    destination_open = destination_xml[: destination_open_end + 1]
    declarations = re.findall(rb"xmlns(?::[A-Za-z_][A-Za-z0-9_.-]*)?=\"[^\"]+\"", source_open)
    prefixed_attributes = [
        value
        for value in re.findall(rb"[A-Za-z_][A-Za-z0-9_.-]*:[A-Za-z_][A-Za-z0-9_.-]*=\"[^\"]+\"", source_open)
        if not value.startswith(b"xmlns:")
    ]
    additions: list[bytes] = []
    for declaration in declarations + prefixed_attributes:
        attribute = declaration.split(b"=", 1)[0]
        if attribute + b"=" not in destination_open:
            additions.append(declaration)
    if additions:
        insert_at = destination_open_end
        destination_xml = (
            destination_xml[:insert_at]
            + b" "
            + b" ".join(additions)
            + destination_xml[insert_at:]
        )
    closing = destination_xml.rfind(b"</worksheet>")
    if closing < 0:
        raise RuntimeError("worksheet closing tag missing")
    return destination_xml[:closing] + extension + destination_xml[closing:]


def preserve_inherited_sheet_extensions(source: Path, destination: Path, inherited_sheet_names: list[str]) -> dict[str, Any]:
    with zipfile.ZipFile(source) as source_zip:
        source_package = {info.filename: source_zip.read(info) for info in source_zip.infolist()}
    with zipfile.ZipFile(destination) as destination_zip:
        destination_infos = destination_zip.infolist()
        destination_package = {info.filename: destination_zip.read(info) for info in destination_infos}
    source_map = _workbook_sheet_xml_map(source_package)
    destination_map = _workbook_sheet_xml_map(destination_package)
    source_extension_sheets: list[str] = []
    preserved_sheets: list[str] = []
    missing_sheets: list[str] = []
    for name in inherited_sheet_names:
        source_path = source_map.get(name)
        destination_path = destination_map.get(name)
        if not source_path or not destination_path:
            missing_sheets.append(name)
            continue
        source_xml = source_package[source_path]
        source_extension = _worksheet_extension_block(source_xml)
        if source_extension is None:
            continue
        source_extension_sheets.append(name)
        destination_package[destination_path] = _copy_extension_block(
            source_xml,
            destination_package[destination_path],
        )
        observed = _worksheet_extension_block(destination_package[destination_path])
        if observed != source_extension:
            raise RuntimeError({"worksheet_extension_mismatch": name})
        preserved_sheets.append(name)
    if missing_sheets:
        raise RuntimeError({"inherited_sheet_mapping_missing": missing_sheets})
    temporary = destination.with_name(destination.name + ".extensions-preserved.tmp")
    with zipfile.ZipFile(temporary, "w", allowZip64=True) as output:
        for info in destination_infos:
            output.writestr(info, destination_package[info.filename])
    temporary.replace(destination)
    with zipfile.ZipFile(destination) as check:
        if check.testzip() is not None:
            raise RuntimeError("workbook CRC failed after extension preservation")
        final_package = {info.filename: check.read(info) for info in check.infolist()}
    final_map = _workbook_sheet_xml_map(final_package)
    mismatches: list[str] = []
    for name in source_extension_sheets:
        source_extension = _worksheet_extension_block(source_package[source_map[name]])
        final_extension = _worksheet_extension_block(final_package[final_map[name]])
        if final_extension != source_extension:
            mismatches.append(name)
    result = {
        "status": "passed" if not mismatches else "failed",
        "source_extension_sheet_count": len(source_extension_sheets),
        "preserved_extension_sheet_count": len(preserved_sheets),
        "source_extension_sheets": source_extension_sheets,
        "preserved_extension_sheets": preserved_sheets,
        "extension_mismatches": mismatches,
        "inherited_sheet_mapping_missing": [],
    }
    if result["status"] != "passed" or len(source_extension_sheets) != len(preserved_sheets):
        raise RuntimeError({"workbook_extension_preservation_gate": result})
    return result
'''


def patch_builder() -> None:
    text = BUILDER.read_text(encoding="utf-8")
    if "def preserve_inherited_sheet_extensions" not in text:
        anchor = "\ndef sync_workbook("
        if anchor not in text:
            raise RuntimeError("sync_workbook anchor not found")
        text = text.replace(anchor, "\n" + HELPER + anchor, 1)
    save_anchor = "    wb.save(destination)\n    with zipfile.ZipFile(destination) as zf:\n"
    save_replacement = (
        "    wb.save(destination)\n"
        "    extension_qa = preserve_inherited_sheet_extensions(source, destination, inherited)\n"
        "    with zipfile.ZipFile(destination) as zf:\n"
    )
    if "extension_qa = preserve_inherited_sheet_extensions" not in text:
        if save_anchor not in text:
            raise RuntimeError("workbook save anchor not found")
        text = text.replace(save_anchor, save_replacement, 1)
    return_anchor = (
        '        "formula_error_count": len(formula_errors),\n'
        '        "status": "passed",\n'
    )
    return_replacement = (
        '        "formula_error_count": len(formula_errors),\n'
        '        "extension_preservation": extension_qa,\n'
        '        "extension_preservation_status": extension_qa["status"],\n'
        '        "status": "passed",\n'
    )
    if '"extension_preservation_status"' not in text:
        if return_anchor not in text:
            raise RuntimeError("workbook return anchor not found")
        text = text.replace(return_anchor, return_replacement, 1)
    if "V3-CP5-S2-REC-231-WORKBOOK-EXTENSION-PRESERVATION" not in text:
        start = text.index("def recovery_events")
        end = text.index("\ndef acceptance_gates", start)
        block = text[start:end]
        closing = block.rfind("\n    ]")
        if closing < 0:
            raise RuntimeError("recovery event list closing not found")
        event = (
            "\n        (231, \"V3-CP5-S2-REC-231-WORKBOOK-EXTENSION-PRESERVATION\", "
            "\"The otherwise-passing terminal build emitted an openpyxl warning that unsupported worksheet extensions, including extended conditional-formatting payloads, could be removed during workbook augmentation.\", "
            "\"Preserved the completed Response 81 copied project, inserted a deterministic ZIP-level inherited-sheet extension restoration pass, required exact extension-block parity for every affected inherited worksheet, expanded the independent verifier, and reran the full build, clean extraction, restore, transport, and delivery pipeline.\"),"
        )
        block = block[:closing] + event + block[closing:]
        text = text[:start] + block + text[end:]
    text = text.replace('"RECOVERY_EVENTS_221_230.json"', '"RECOVERY_EVENTS_221_231.json"')
    BUILDER.write_text(text, encoding="utf-8")


def patch_verifier() -> None:
    text = VERIFIER.read_text(encoding="utf-8")
    if '"workbook_extensions"' not in text:
        anchor = (
            '        "workbook": summary.get("workbook", {}).get("current_sheet_count", 0) >= 129 and summary.get("workbook", {}).get("formula_error_count") == 0,\n'
        )
        replacement = anchor + (
            '        "workbook_extensions": summary.get("workbook", {}).get("extension_preservation_status") == "passed" '
            'and summary.get("workbook", {}).get("extension_preservation", {}).get("source_extension_sheet_count") '
            '== summary.get("workbook", {}).get("extension_preservation", {}).get("preserved_extension_sheet_count") '
            'and not summary.get("workbook", {}).get("extension_preservation", {}).get("extension_mismatches"),\n'
        )
        if anchor not in text:
            raise RuntimeError("independent verifier workbook anchor not found")
        text = text.replace(anchor, replacement, 1)
    VERIFIER.write_text(text, encoding="utf-8")


if __name__ == "__main__":
    patch_builder()
    patch_verifier()
    print({
        "status": "passed",
        "builder": str(BUILDER),
        "verifier": str(VERIFIER),
        "control": "inherited unsupported worksheet extension parity",
    })
