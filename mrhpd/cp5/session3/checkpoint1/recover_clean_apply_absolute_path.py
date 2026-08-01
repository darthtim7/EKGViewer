#!/usr/bin/env python3
"""Resolve the Response 81 baseline before the clean-apply cwd changes."""
from pathlib import Path

ROOT = Path(__file__).resolve().parent
BUILDER = ROOT / "build_section5_session3_checkpoint1.py"
text = BUILDER.read_text(encoding="utf-8")
old = 'str(baseline_restore), "--output-dir", str(output)'
new = 'str(baseline_restore.resolve()), "--output-dir", str(output)'
if old not in text and new not in text:
    raise RuntimeError("clean-apply baseline-path anchor not found")
text = text.replace(old, new)

if "V3-CP5-S3-REC-242-CLEAN-APPLY-ABSOLUTE-BASELINE" not in text:
    start = text.index("def recovery_events(now_iso: str)")
    rows_start = text.index("    rows = [", start)
    rows_end = text.index("\n    ]", rows_start)
    event = (
        '\n        (242, "V3-CP5-S3-REC-242-CLEAN-APPLY-ABSOLUTE-BASELINE", '
        '"The next otherwise-complete disposable build invoked the generated clean-apply utility from its package directory while passing the Response 81 baseline as a repository-relative path, producing FileNotFoundError before verification began.", '
        '"Preserved the synchronized copied project, reports, indexes, manifests, and recovery package; resolved the Response 81 baseline to an absolute path before the subprocess changed its working directory; and reran clean application, independent output verification, packaging, and upload gates."),'
    )
    text = text[:rows_end] + event + text[rows_end:]
text = text.replace('"RECOVERY_EVENTS_232_241.json"', '"RECOVERY_EVENTS_232_242.json"')

BUILDER.write_text(text, encoding="utf-8")
print({"status": "passed", "builder": str(BUILDER), "clean_apply_baseline": "absolute", "recovery_event": 242})
