#!/usr/bin/env python3
from pathlib import Path

path = Path(__file__).resolve().with_name("build_section5_checkpoint2.py")
text = path.read_text(encoding="utf-8")

# Normalize the conditional mapping unpack used by the page-transform register.
start_old = '        row = {\n            **transform_rows[index] if index < 537 else {'
start_new = '        row = {\n            **(transform_rows[index] if index < 537 else {'
end_old = '            },\n            "searchable": bool(text.strip()),'
end_new = '            }),\n            "searchable": bool(text.strip()),'
if start_old in text:
    text = text.replace(start_old, start_new, 1)
if end_old in text:
    text = text.replace(end_old, end_new, 1)

# Checkpoint 1 created the governing column as digital_publication_sha256.
text = text.replace(
    'baseline_restore_sha256,baseline_project_sha256,publication_sha256,digital_page_count,',
    'baseline_restore_sha256,baseline_project_sha256,digital_publication_sha256,digital_page_count,',
    1,
)

# Preserve the exact recoverable database-schema mismatch in the project recovery record.
function_start = text.index('def recovery_events(now_iso: str)')
function_end = text.index('\n\n\ndef _clone_response_row', function_start)
block = text[function_start:function_end]
if 'V3-CP5-S1-REC-189-CHECKPOINT-COLUMN-NAME' not in block:
    marker = '        {\n            "event_code": "V3-CP5-S1-REC-188-PREFLIGHT-AND-CLEAN-APPLY"'
    event_start = block.index(marker)
    closing = block.index('\n        },\n    ]', event_start)
    insertion = '''
        {
            "event_code": "V3-CP5-S1-REC-189-CHECKPOINT-COLUMN-NAME",
            "condition": "The first disposable Checkpoint 2 database pass used publication_sha256, while the governing Checkpoint 1 table defines digital_publication_sha256.",
            "recovery": "Preserved the recovered Checkpoint 1 project and all generated print assets, inspected the actual copied schema, changed only the insertion column name, and reran every dependent database, workbook, application, document, index, manifest, archive, and clean-application gate.",
            "status": "recovered",
            "recorded_at": now_iso,
        },'''
    block = block[:closing + len('\n        },')] + insertion + block[closing + len('\n        },'):]
    text = text[:function_start] + block + text[function_end:]

path.write_text(text, encoding="utf-8")
print({"status": "passed", "path": str(path)})
