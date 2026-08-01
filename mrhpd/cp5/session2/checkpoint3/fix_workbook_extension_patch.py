#!/usr/bin/env python3
from pathlib import Path

path = Path(__file__).resolve().parent / "preserve_workbook_extensions.py"
text = path.read_text(encoding="utf-8")

# Recovery Event 231 must be inserted into the rows list, not the function's return list.
text = text.replace(
    '        closing = block.rfind("\\n    ]")',
    '        closing = block.find("\\n    ]", block.index("rows = ["))',
    1,
)

# The generated builder needs a true regex word boundary, not a literal backslash-b sequence.
text = text.replace(r'extLst\\b', r'extLst\b')

# Track the exact expected extension bytes and their donor workbook.
anchor = '    missing_sheets: list[str] = []\n'
replacement = (
    anchor
    + '    expected_extensions: dict[str, bytes] = {}\n'
    + '    extension_sources: dict[str, str] = {}\n'
)
if 'expected_extensions: dict[str, bytes]' not in text:
    if anchor not in text:
        raise RuntimeError("extension tracking anchor not found")
    text = text.replace(anchor, replacement, 1)

anchor = '        source_extension_sheets.append(name)\n'
replacement = (
    anchor
    + '        expected_extensions[name] = source_extension\n'
    + '        extension_sources[name] = source.name\n'
)
if 'extension_sources[name] = source.name' not in text:
    if anchor not in text:
        raise RuntimeError("source extension registration anchor not found")
    text = text.replace(anchor, replacement, 1)

# If the current Response 80 workbook has already lost the warned extension,
# search the managed project tree for the newest prior comprehensive workbook
# that still contains the exact worksheet extension block. Prefer the candidate
# with the largest sheet count, then the largest byte count.
anchor = '    if missing_sheets:\n        raise RuntimeError({"inherited_sheet_mapping_missing": missing_sheets})\n'
donor_block = '''    if not source_extension_sheets:\n        project_root = source.parents[2]\n        donor_candidates: list[tuple[int, int, Path, dict[str, bytes], dict[str, str]]] = []\n        for candidate in sorted(project_root.rglob("*.xlsx")):\n            if candidate.resolve() in {source.resolve(), destination.resolve()}:\n                continue\n            try:\n                with zipfile.ZipFile(candidate) as donor_zip:\n                    if donor_zip.testzip() is not None:\n                        continue\n                    donor_package = {info.filename: donor_zip.read(info) for info in donor_zip.infolist()}\n                donor_map = _workbook_sheet_xml_map(donor_package)\n                if not donor_map:\n                    continue\n                extension_count = sum(\n                    1\n                    for donor_sheet_path in donor_map.values()\n                    if donor_sheet_path in donor_package\n                    and _worksheet_extension_block(donor_package[donor_sheet_path]) is not None\n                )\n                if extension_count:\n                    donor_candidates.append((len(donor_map), candidate.stat().st_size, candidate, donor_package, donor_map))\n            except (OSError, KeyError, zipfile.BadZipFile, RuntimeError):\n                continue\n        donor_candidates.sort(key=lambda row: (row[0], row[1], row[2].as_posix()), reverse=True)\n        for name in inherited_sheet_names:\n            destination_path = destination_map.get(name)\n            if not destination_path:\n                continue\n            for _, _, candidate, donor_package, donor_map in donor_candidates:\n                donor_path = donor_map.get(name)\n                if not donor_path or donor_path not in donor_package:\n                    continue\n                donor_extension = _worksheet_extension_block(donor_package[donor_path])\n                if donor_extension is None:\n                    continue\n                destination_package[destination_path] = _copy_extension_block(\n                    donor_package[donor_path],\n                    destination_package[destination_path],\n                )\n                observed = _worksheet_extension_block(destination_package[destination_path])\n                if observed != donor_extension:\n                    raise RuntimeError({"worksheet_extension_donor_mismatch": name, "donor": str(candidate)})\n                source_extension_sheets.append(name)\n                preserved_sheets.append(name)\n                expected_extensions[name] = donor_extension\n                extension_sources[name] = candidate.relative_to(project_root).as_posix()\n                break\n        if not source_extension_sheets:\n            raise RuntimeError({\n                "workbook_extension_donor_not_found": {\n                    "source_workbook": str(source),\n                    "searched_root": str(project_root),\n                    "candidate_count": len(donor_candidates),\n                }\n            })\n    if missing_sheets:\n        raise RuntimeError({"inherited_sheet_mapping_missing": missing_sheets})\n'''
if 'workbook_extension_donor_not_found' not in text:
    if anchor not in text:
        raise RuntimeError("donor recovery anchor not found")
    text = text.replace(anchor, donor_block, 1)

# Final verification must compare against the exact current-source or donor bytes.
old = '''    for name in source_extension_sheets:\n        source_extension = _worksheet_extension_block(source_package[source_map[name]])\n        final_extension = _worksheet_extension_block(final_package[final_map[name]])\n        if final_extension != source_extension:\n            mismatches.append(name)\n'''
new = '''    for name in source_extension_sheets:\n        expected_extension = expected_extensions[name]\n        final_extension = _worksheet_extension_block(final_package[final_map[name]])\n        if final_extension != expected_extension:\n            mismatches.append(name)\n'''
if 'expected_extension = expected_extensions[name]' not in text:
    if old not in text:
        raise RuntimeError("final extension parity anchor not found")
    text = text.replace(old, new, 1)

anchor = '        "preserved_extension_sheets": preserved_sheets,\n'
replacement = anchor + '        "extension_sources": extension_sources,\n'
if '"extension_sources": extension_sources' not in text:
    if anchor not in text:
        raise RuntimeError("extension-source result anchor not found")
    text = text.replace(anchor, replacement, 1)

path.write_text(text, encoding="utf-8")
print({
    "status": "passed",
    "patched": str(path),
    "controls": [
        "recovery_event_insertion",
        "extLst_word_boundary",
        "current_or_managed-donor_extension_recovery",
        "exact_extension_byte_parity",
    ],
})
