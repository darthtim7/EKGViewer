#!/usr/bin/env python3
from pathlib import Path

path = Path(__file__).resolve().with_name("build_section5_checkpoint2.py")
text = path.read_text(encoding="utf-8")
start_old = '        row = {\n            **transform_rows[index] if index < 537 else {'
start_new = '        row = {\n            **(transform_rows[index] if index < 537 else {'
end_old = '            },\n            "searchable": bool(text.strip()),'
end_new = '            }),\n            "searchable": bool(text.strip()),'
if start_old in text:
    text = text.replace(start_old, start_new, 1)
if end_old in text:
    text = text.replace(end_old, end_new, 1)
path.write_text(text, encoding="utf-8")
print({"status": "passed", "path": str(path)})
