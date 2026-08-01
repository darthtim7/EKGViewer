#!/usr/bin/env python3
from __future__ import annotations

import json
import sqlite3
import subprocess
import sys
from pathlib import Path

import build_section5_checkpoint1 as base


def recovered_events(now_iso: str):
    rows = base.recovery_events(now_iso)
    rows.append({
        "event_code": "V3-CP5-S1-REC-178-APPLICATION-AUDIT-PROJECT-ROOT",
        "condition": "The first copied-tree application audit resolved the project root one parent too high and failed with sqlite3.OperationalError: unable to open database file.",
        "recovery": "Preserved the exact baseline, corrected the generated audit root from parents[3] to parents[2], restarted from the verified Response 72 restore, and reran all dependent database, workbook, report, index, manifest, package, and clean-apply gates.",
        "status": "recovered",
        "recorded_at": now_iso,
    })
    return rows


def recovered_application_surfaces(project: Path, db_path: Path, app_path: Path, now_iso: str):
    root = project / "App" / "Section 5 Session 1 Checkpoint 1"
    root.mkdir(parents=True, exist_ok=True)
    db_rel = db_path.relative_to(project).as_posix()
    app_rel = app_path.relative_to(project).as_posix()
    pointer = root / "CURRENT_DATABASE.txt"
    base.text_write(pointer, db_rel + "\n")
    state = root / "CURRENT_PROJECT_STATE.json"
    base.json_write(state, {
        "schema": "mrhpd-section5-current-project-state-1.0",
        "response": 75,
        "section": base.SECTION_LABEL,
        "session": base.SESSION_LABEL,
        "checkpoint": base.CHECKPOINT_LABEL,
        "state": "checkpoint_complete",
        "database": db_rel,
        "main_application": app_rel,
        "main_application_sha256": base.sha256_file(app_path),
        "main_application_unchanged": True,
        "recorded_at": now_iso,
    })
    audit_script = root / "audit_section5_checkpoint1.py"
    base.text_write(audit_script, f'''#!/usr/bin/env python3
import json, sqlite3
from pathlib import Path
project = Path(__file__).resolve().parents[2]
db = project / {db_rel!r}
con = sqlite3.connect(db)
try:
    integrity = con.execute("PRAGMA integrity_check").fetchone()[0]
    fk = len(list(con.execute("PRAGMA foreign_key_check")))
    response = con.execute("SELECT COUNT(*) FROM thread_response_reconciliation_cp3 WHERE response_key='R75'").fetchone()[0]
    checkpoint = con.execute("SELECT state FROM section5_session1_checkpoint WHERE checkpoint_code='{base.CHECKPOINT_CODE}'").fetchone()
    scenarios = con.execute("SELECT COUNT(*) FROM section5_spine_scenario WHERE checkpoint_code='{base.CHECKPOINT_CODE}'").fetchone()[0]
    specs = con.execute("SELECT COUNT(*) FROM section5_print_provider_spec WHERE checkpoint_code='{base.CHECKPOINT_CODE}'").fetchone()[0]
finally:
    con.close()
result = {{"status":"passed" if integrity=="ok" and fk==0 and response==1 and checkpoint==("checkpoint_complete",) and scenarios>=5 and specs>=10 else "failed", "integrity":integrity,"foreign_key_violations":fk,"response75":response,"checkpoint":checkpoint,"scenarios":scenarios,"provider_specs":specs}}
print(json.dumps(result,indent=2))
raise SystemExit(0 if result["status"]=="passed" else 1)
''')
    result = subprocess.run([sys.executable, str(audit_script)], cwd=project, text=True, capture_output=True, timeout=120)
    output = root / "SECTION5_CHECKPOINT1_APPLICATION_AUDIT.json"
    if result.returncode:
        raise RuntimeError({"application_audit_failed": {"stdout": result.stdout, "stderr": result.stderr}})
    audit = json.loads(result.stdout)
    audit.update({"main_application_sha256": base.sha256_file(app_path), "main_application_unchanged": True})
    base.json_write(output, audit)
    return [pointer, state, audit_script, output]


base.recovery_events = recovered_events
base.write_application_surfaces = recovered_application_surfaces

if __name__ == "__main__":
    base.main()
