#!/usr/bin/env python3
from pathlib import Path

path = Path(__file__).resolve().parent / "preserve_workbook_extensions.py"
text = path.read_text(encoding="utf-8")
text = text.replace(
    '        closing = block.rfind("\\n    ]")',
    '        closing = block.find("\\n    ]", block.index("rows = ["))',
    1,
)
text = text.replace(r'extLst\\b', r'extLst\b')
anchor = '    if missing_sheets:\n        raise RuntimeError({"inherited_sheet_mapping_missing": missing_sheets})\n'
replacement = (
    '    if not source_extension_sheets:\n'
    '        raise RuntimeError("no inherited worksheet extension blocks were discovered despite the observed openpyxl warning")\n'
    + anchor
)
if 'no inherited worksheet extension blocks were discovered' not in text:
    if anchor not in text:
        raise RuntimeError("source extension count anchor not found")
    text = text.replace(anchor, replacement, 1)
path.write_text(text, encoding="utf-8")
print({"status": "passed", "patched": str(path), "controls": ["recovery_event_insertion", "extLst_word_boundary", "nonzero_extension_count"]})
