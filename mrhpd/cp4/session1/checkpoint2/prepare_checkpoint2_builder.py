#!/usr/bin/env python3
"""Normalize and apply verified recoveries to the transported Checkpoint 2 builder."""
from __future__ import annotations

from pathlib import Path

BUILDER = Path(__file__).with_name("build_checkpoint2_recovery.py")


def replace_once(text: str, old: str, new: str, label: str) -> str:
    if old not in text:
        raise SystemExit(f"Expected {label} source was not found; refusing an unverified patch")
    return text.replace(old, new, 1)


def main() -> None:
    lines = BUILDER.read_text(encoding="utf-8").splitlines(keepends=True)
    normalized_plus_lines = sum(1 for line in lines if line.startswith("+"))
    text = "".join(line[1:] if line.startswith("+") else line for line in lines)

    text = replace_once(
        text,
        '''    columns = table_columns(con, table)\n    values = {name: value for name, value in row.items() if name in columns and not name.endswith("_id")}\n''',
        '''    info = list(con.execute(f'PRAGMA table_info("{table}")'))\n    columns = [row[1] for row in info]\n    pk_columns = {row[1] for row in info if row[5]}\n    values = {name: value for name, value in row.items() if name in columns and name not in pk_columns}\n''',
        "schema-blind upsert helper",
    )

    text = replace_once(
        text,
        '''    from openpyxl.styles import Alignment\n    for column in ws.columns:\n        letter = column[0].column_letter\n''',
        '''    from openpyxl.styles import Alignment\n    from openpyxl.utils import get_column_letter\n    for column_index, column in enumerate(ws.columns, 1):\n        letter = get_column_letter(column_index)\n''',
        "merged-cell-sensitive workbook autosize helper",
    )

    text = replace_once(
        text,
        '''            metadata_updates = {\n                "version": PROJECT_VERSION,\n                "current_remediation_section": "Remediation Section 4 of 5",\n''',
        '''            metadata_updates = {\n                "version": PROJECT_VERSION,\n                "checkpoint_section": "Remediation Section 4 of 5",\n                "checkpoint_session": "Session 1 of 3",\n                "status": "remediation_section_4_session_1_checkpoint_2_complete",\n                "current_checkpoint": "MRHPD-V3-CP4-S1-CP2",\n                "current_resume_point": "Checkpoint 3 of 3 — complete Session 1 restore emission",\n                "current_remediation_section": "Remediation Section 4 of 5",\n''',
        "current metadata update block",
    )

    start = text.index("def update_application(mutable_root: Path, db: Path)")
    end = text.index("\ndef verify_publication(mutable_root: Path)", start)
    replacement_application = r'''def update_application(mutable_root: Path, db: Path) -> tuple[list[Path], dict[str, Any]]:
    import http.client
    import importlib.util
    import socket
    import time

    app_files = sorted(mutable_root.rglob("human_pathogen_app.py"))
    if len(app_files) != 1:
        raise RuntimeError({"human_pathogen_app_matches": [str(p) for p in app_files]})
    app = app_files[0]
    app_original_sha256 = sha256_file(app)

    # The resolver is intentionally database-agnostic. Its required --db
    # argument is the native configuration surface. Preserve the proven
    # clinical resolver and add a current-state launcher instead of embedding
    # a brittle database filename in the application source.
    state = {
        "schema": "mrhpd-current-application-state-1.0",
        "generated_at": NOW,
        "remediation_section": "4 of 5",
        "session": "1 of 3",
        "checkpoint": "2 of 3",
        "response": 65,
        "canonical_database": db.name,
        "database_relative_path": db.relative_to(mutable_root).as_posix(),
        "application": app.relative_to(mutable_root).as_posix(),
        "configuration_surface": "human_pathogen_app.py --db <SQLite database path>",
        "accepted_predecessor_mutated": False,
        "restore_policy": EMISSION_POLICY,
    }
    state_path = app.parent / "CURRENT_PROJECT_STATE.json"
    json_write(state_path, state)

    launcher_path = app.parent / "run_section4_session1_checkpoint2.py"
    launcher_text = (
        "#!/usr/bin/env python3\n"
        "from pathlib import Path\n"
        "import subprocess,sys\n"
        "HERE=Path(__file__).resolve().parent\n"
        "APP=HERE/'human_pathogen_app.py'\n"
        f"DB=HERE.parent/'Database'/{db.name!r}\n"
        "if not DB.exists(): raise SystemExit(f'Canonical database not found: {DB}')\n"
        "raise SystemExit(subprocess.call([sys.executable,str(APP),'--db',str(DB),*sys.argv[1:]]))\n"
    )
    text_write(launcher_path, launcher_text)
    database_pointer = app.parent / "CURRENT_DATABASE.txt"
    text_write(database_pointer, db.relative_to(mutable_root).as_posix())
    readme_path = app.parent / "README_SECTION4_SESSION1_CHECKPOINT2.md"
    text_write(
        readme_path,
        "# Human Pathogen Database local application — Section 4 Session 1 Checkpoint 2\n\n"
        f"Canonical database: `{db.name}`\n\n"
        "The resolver remains database-agnostic and read-only. Start the synchronized application with "
        "`python run_section4_session1_checkpoint2.py`; the launcher supplies the copied Section 4 SQLite database through "
        "the application's required `--db` argument. The frozen Section 3 release and accepted predecessor remain immutable."
    )

    # Preserve and rerun the complete legacy regression set against the frozen
    # source database that those scripts explicitly govern.
    legacy_results: list[dict[str, Any]] = []
    legacy_assertion_count = 0
    test_files = sorted(app.parent.glob("test*.py"))
    if not test_files:
        raise RuntimeError("No application regression tests found")
    for test in test_files:
        result = subprocess.run(
            [sys.executable, str(test)], cwd=app.parent,
            text=True, capture_output=True, timeout=300,
        )
        parsed = None
        try:
            parsed = json.loads(result.stdout)
        except Exception:
            parsed = None
        if isinstance(parsed, dict):
            legacy_assertion_count += int(parsed.get("test_count") or len(parsed.get("tests") or []))
        record = {
            "test": test.name,
            "returncode": result.returncode,
            "parsed_passed": parsed.get("passed") if isinstance(parsed, dict) else None,
            "parsed_test_count": parsed.get("test_count") if isinstance(parsed, dict) else None,
            "stdout_tail": result.stdout[-12000:],
            "stderr_tail": result.stderr[-12000:],
        }
        legacy_results.append(record)
        if result.returncode != 0 or (isinstance(parsed, dict) and parsed.get("passed") is False):
            raise RuntimeError({"legacy_application_test_failed": record})

    # Independently exercise the same resolver source against the new current
    # Section 4 database.
    spec = importlib.util.spec_from_file_location("mrhpd_section4_app", app)
    if spec is None or spec.loader is None:
        raise RuntimeError("Could not import the application for current-database regression testing")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    con = module.connect(db)
    try:
        direct_checks = {
            "database_integrity": con.execute("PRAGMA integrity_check").fetchone()[0] == "ok",
            "foreign_keys": len(con.execute("PRAGMA foreign_key_check").fetchall()) == 0,
            "response65_record": con.execute("SELECT COUNT(*) FROM thread_response_reconciliation_cp3 WHERE response_key='R65'").fetchone()[0] == 1,
            "fractional_prompts_64_1_to_64_3": con.execute("SELECT COUNT(*) FROM fractional_prompt_cp3 WHERE prompt_number IN ('64.1','64.2','64.3')").fetchone()[0] == 3,
            "restore_emission_policy": con.execute("SELECT COUNT(*) FROM restore_emission_policy WHERE effective_response=65").fetchone()[0] == 1,
            "section4_checkpoint": con.execute("SELECT COUNT(*) FROM section4_checkpoint WHERE checkpoint_code='MRHPD-V3-CP4-S1-CP2'").fetchone()[0] == 1,
            "search_strep": bool(module.search_database(con, "Strep")["results"]),
            "sbsec_disambiguation": module.search_database(con, "strep bovis").get("resolver", {}).get("behavior") == "show_disambiguation",
            "crypto_disambiguation": len(module.search_database(con, "Crypto")["results"]) >= 2,
            "pathways": len(module.list_pathways(con)) > 0,
            "dimensions": len(module.list_dimensions(con)) > 0,
            "contexts": len(module.list_contexts(con)) > 0,
            "antibiograms": len(module.list_antibiograms(con)) > 0,
            "page_map": len(module.cp3_final_page_map(con)) == 10 and sum(r["page_count"] for r in module.cp3_final_page_map(con)) == 537,
            "cross_references": len(module.cp3_current_cross_references(con)) == 12,
        }
    finally:
        con.close()
    if not all(direct_checks.values()):
        raise RuntimeError({"current_database_direct_regression_failed": direct_checks})

    # Rerun loopback HTTP and security controls against the new database.
    port_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    port_socket.bind(("127.0.0.1", 0))
    port = port_socket.getsockname()[1]
    port_socket.close()
    proc = subprocess.Popen(
        [sys.executable, str(app), "--db", str(db), "--host", "127.0.0.1", "--port", str(port), "--app-dir", str(app.parent)],
        cwd=app.parent, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True,
    )
    http_checks: dict[str, bool] = {}
    http_details: dict[str, Any] = {}
    try:
        deadline = time.time() + 25
        while time.time() < deadline:
            if proc.poll() is not None:
                output = proc.stdout.read() if proc.stdout else ""
                raise RuntimeError({"application_server_exited": proc.returncode, "output": output[-8000:]})
            try:
                with socket.create_connection(("127.0.0.1", port), timeout=0.5):
                    break
            except OSError:
                time.sleep(0.1)
        else:
            raise RuntimeError("Current Section 4 application server did not start")

        def request(path: str):
            client = http.client.HTTPConnection("127.0.0.1", port, timeout=15)
            client.request("GET", path)
            response = client.getresponse()
            body = response.read()
            headers = dict(response.getheaders())
            client.close()
            try:
                payload = json.loads(body) if body else None
            except Exception:
                payload = body.decode("utf-8", "replace")
            return response.status, headers, payload

        status, headers, payload = request("/")
        http_checks["index"] = status == 200
        http_checks["security_headers"] = headers.get("X-Content-Type-Options") == "nosniff" and headers.get("X-Frame-Options") == "DENY"
        status, headers, payload = request("/api/health")
        http_checks["health"] = status == 200 and isinstance(payload, dict) and payload.get("integrity") == "ok"
        status, headers, payload = request("/api/search?q=Strep")
        http_checks["search_strep"] = status == 200 and isinstance(payload, dict) and bool(payload.get("results"))
        status, headers, payload = request("/api/search?q=strep%20bovis")
        http_checks["search_sbsec"] = status == 200 and isinstance(payload, dict) and payload.get("resolver", {}).get("behavior") == "show_disambiguation"
        status, headers, payload = request("/api/search?q=Medical%20Absolutes")
        http_checks["restricted_source_not_searchable"] = status == 200 and isinstance(payload, dict) and payload.get("results") == []
        status, headers, payload = request("/api/not-a-route")
        http_checks["unknown_route"] = status == 404
        http_details = {"port": port, "checks": http_checks}
    finally:
        proc.terminate()
        try:
            proc.wait(timeout=5)
        except subprocess.TimeoutExpired:
            proc.kill()
            proc.wait(timeout=5)

    bad_bind = subprocess.run(
        [sys.executable, str(app), "--db", str(db), "--host", "0.0.0.0", "--port", str(port + 1), "--app-dir", str(app.parent)],
        cwd=app.parent, text=True, capture_output=True, timeout=15,
    )
    http_checks["non_loopback_rejected"] = bad_bind.returncode != 0 and "loopback" in (bad_bind.stdout + bad_bind.stderr).lower()
    if not all(http_checks.values()):
        raise RuntimeError({"current_database_http_regression_failed": http_checks, "details": http_details})

    qa = {
        "status": "passed",
        "application": app.relative_to(mutable_root).as_posix(),
        "application_sha256": sha256_file(app),
        "application_source_preserved": sha256_file(app) == app_original_sha256,
        "native_database_configuration": "required --db argument",
        "canonical_database": db.relative_to(mutable_root).as_posix(),
        "canonical_database_referenced": db.name in launcher_path.read_text(encoding="utf-8") and state["canonical_database"] == db.name,
        "launcher": launcher_path.relative_to(mutable_root).as_posix(),
        "state_file": state_path.relative_to(mutable_root).as_posix(),
        "database_pointer": database_pointer.relative_to(mutable_root).as_posix(),
        "legacy_tests": legacy_results,
        "legacy_test_files": len(legacy_results),
        "legacy_assertion_count": legacy_assertion_count,
        "current_direct_checks": direct_checks,
        "current_direct_check_count": len(direct_checks),
        "current_http_checks": http_checks,
        "current_http_check_count": len(http_checks),
        "all_returncodes_zero": all(item["returncode"] == 0 for item in legacy_results),
    }
    if not qa["canonical_database_referenced"] or not qa["application_source_preserved"]:
        raise RuntimeError({"application_sync_verification_failed": qa})
    return [app, state_path, database_pointer, launcher_path, readme_path, *test_files], qa
'''
    text = text[:start] + replacement_application + text[end:]

    BUILDER.write_text(text, encoding="utf-8")
    print({
        "normalized_leading_plus_lines": normalized_plus_lines,
        "schema_aware_upsert_patch": "applied",
        "merged_cell_autosize_patch": "applied",
        "current_metadata_patch": "applied",
        "native_app_configuration_patch": "applied",
        "builder": str(BUILDER),
    })


if __name__ == "__main__":
    main()
