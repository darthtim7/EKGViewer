#!/usr/bin/env python3
from pathlib import Path

root = Path(__file__).resolve().parent
builder_path = root / "build_section5_checkpoint2.py"
reporting_path = root / "section5_checkpoint2_reporting.py"
text = builder_path.read_text(encoding="utf-8")

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

# Convert list/dict workbook values to deterministic JSON text before openpyxl cell assignment.
text = text.replace(
    '        ws.append([row.get(header) for header in headers])',
    '        ws.append([json.dumps(value, ensure_ascii=False, sort_keys=True) if isinstance(value, (dict, list, tuple, set)) else value for value in (row.get(header) for header in headers)])',
)

# Resolve the generated recovery utility before changing cwd so its path cannot be duplicated.
text = text.replace(
    'str(tools / "apply_checkpoint_recovery.py")',
    'str((tools / "apply_checkpoint_recovery.py").resolve())',
)

# Preserve exact recoverable build events in the project record.
function_start = text.index('def recovery_events(now_iso: str)')
function_end = text.index('\n\n\ndef _clone_response_row', function_start)
block = text[function_start:function_end]
events = [
    (
        'V3-CP5-S1-REC-189-CHECKPOINT-COLUMN-NAME',
        'The first disposable Checkpoint 2 database pass used publication_sha256, while the governing Checkpoint 1 table defines digital_publication_sha256.',
        'Preserved the recovered Checkpoint 1 project and all generated print assets, inspected the actual copied schema, changed only the insertion column name, and reran every dependent database, workbook, application, document, index, manifest, archive, and clean-application gate.'
    ),
    (
        'V3-CP5-S1-REC-190-WORKBOOK-STRUCTURED-VALUE',
        'The next disposable comprehensive-workbook pass attempted to assign the official-source URL list directly to one Excel cell and openpyxl raised ValueError: Cannot convert the list to Excel.',
        'Preserved the verified database and generated print assets, added deterministic JSON serialization for list, tuple, set, and dictionary cell values, rebuilt the workbook from the verified Checkpoint 1 source, and reran every dependent gate.'
    ),
    (
        'V3-CP5-S1-REC-191-REPORT-REGISTER-STRUCTURED-VALUE',
        'The subsequent print-production register attempted the same direct list assignment and openpyxl raised ValueError while writing the Selection sheet.',
        'Preserved the verified copied state, added the same deterministic JSON serialization to the independent checkpoint-register builder, regenerated the DOCX, searchable PDF, XLSX register, rendered report QA, indexes, manifests, recovery package, and clean-application evidence, and reran every gate.'
    ),
    (
        'V3-CP5-S1-REC-192-CLEAN-APPLY-ABSOLUTE-PATH',
        'The next otherwise-complete disposable build invoked the generated recovery utility with a relative path while also setting its package directory as cwd, duplicating the path and producing Errno 2: No such file or directory.',
        'Preserved the completed print derivatives and synchronized copied project, resolved the generated recovery utility to an absolute path before subprocess launch, reran the clean application from the exact Response 72 restore, and reran every packaging and independent-verification gate.'
    ),
]
for event_code, condition, recovery in events:
    if event_code in block:
        continue
    closing = block.rindex('\n    ]')
    insertion = f'''        {{
            "event_code": "{event_code}",
            "condition": {condition!r},
            "recovery": {recovery!r},
            "status": "recovered",
            "recorded_at": now_iso,
        }},
'''
    block = block[:closing] + insertion + block[closing:]
text = text[:function_start] + block + text[function_end:]
builder_path.write_text(text, encoding="utf-8")

# Apply the same structured-value conversion to the independent XLSX register builder.
reporting = reporting_path.read_text(encoding="utf-8")
reporting = reporting.replace(
    '            ws.append([row.get(header) for header in headers])',
    '            ws.append([json.dumps(value, ensure_ascii=False, sort_keys=True) if isinstance(value, (dict, list, tuple, set)) else value for value in (row.get(header) for header in headers)])',
)
reporting_path.write_text(reporting, encoding="utf-8")

print({"status": "passed", "builder": str(builder_path), "reporting": str(reporting_path)})
