#!/usr/bin/env python3
"""Normalize the terminal Response 84 builder against the proven workbook-extension helper."""
from pathlib import Path

HERE = Path(__file__).resolve().parent
BUILDER = HERE / "build_section5_project_complete.py"
text = BUILDER.read_text(encoding="utf-8")

path_anchor = 'CP2_DIR = HERE.parent / "checkpoint2"\n'
path_insertion = (
    'CP2_DIR = HERE.parent / "checkpoint2"\n'
    'S2_COMPLETE_DIR = HERE.parents[1] / "session2" / "checkpoint3"\n'
)
if 'S2_COMPLETE_DIR = HERE.parents[1] / "session2" / "checkpoint3"' not in text:
    if path_anchor not in text:
        raise RuntimeError("CP2_DIR anchor missing")
    text = text.replace(path_anchor, path_insertion, 1)

sys_path_anchor = (
    'if str(CP2_DIR) not in sys.path:\n'
    '    sys.path.insert(0, str(CP2_DIR))\n'
    'import build_section5_session3_checkpoint2 as cp2  # noqa: E402\n'
)
sys_path_replacement = (
    'if str(CP2_DIR) not in sys.path:\n'
    '    sys.path.insert(0, str(CP2_DIR))\n'
    'if str(S2_COMPLETE_DIR) not in sys.path:\n'
    '    sys.path.insert(0, str(S2_COMPLETE_DIR))\n'
    'import build_section5_session3_checkpoint2 as cp2  # noqa: E402\n'
    'import build_section5_session2_complete_restore as s2_complete  # noqa: E402\n'
)
if 'import build_section5_session2_complete_restore as s2_complete' not in text:
    if sys_path_anchor not in text:
        raise RuntimeError("terminal import anchor missing")
    text = text.replace(sys_path_anchor, sys_path_replacement, 1)

old_call = 'extension_qa = cp2.preserve_inherited_sheet_extensions(source, destination, inherited)'
new_call = 'extension_qa = s2_complete.preserve_inherited_sheet_extensions(source, destination, inherited)'
if old_call in text:
    text = text.replace(old_call, new_call, 1)
elif new_call not in text:
    raise RuntimeError("workbook extension helper call missing")

BUILDER.write_text(text, encoding="utf-8")
print({
    "status": "passed",
    "builder": str(BUILDER),
    "workbook_extension_helper": "build_section5_session2_complete_restore.preserve_inherited_sheet_extensions",
})
