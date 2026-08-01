#!/usr/bin/env python3
"""Normalize the terminal Response 84 builder against the proven Checkpoint 1 workbook-extension helper."""
from pathlib import Path

HERE = Path(__file__).resolve().parent
BUILDER = HERE / "build_section5_project_complete.py"
text = BUILDER.read_text(encoding="utf-8")

# Checkpoint 2 imports Checkpoint 1 as ``cp1``. The extension-preservation
# implementation is defined there and is therefore available as ``cp2.cp1``.
# Keep the terminal builder coupled to that verified implementation rather than
# duplicating XML package surgery.
verified_call = "extension_qa = cp2.cp1.preserve_inherited_sheet_extensions(source, destination, inherited)"
for stale_call in (
    "extension_qa = cp2.preserve_inherited_sheet_extensions(source, destination, inherited)",
    "extension_qa = s2_complete.preserve_inherited_sheet_extensions(source, destination, inherited)",
):
    if stale_call in text:
        text = text.replace(stale_call, verified_call, 1)

if verified_call not in text:
    raise RuntimeError("verified workbook-extension helper call missing")

BUILDER.write_text(text, encoding="utf-8")
print({
    "status": "passed",
    "builder": str(BUILDER),
    "workbook_extension_helper": "build_section5_session3_checkpoint2.cp1.preserve_inherited_sheet_extensions",
})
