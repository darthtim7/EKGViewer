#!/usr/bin/env python3
"""Add the explicit application aggregate expected by checkpoint reports."""
from pathlib import Path

path = Path(__file__).with_name("build_checkpoint2_recovery.py")
text = path.read_text(encoding="utf-8")
old = '''        "current_http_checks": http_checks,\n        "current_http_check_count": len(http_checks),\n        "all_returncodes_zero": all(item["returncode"] == 0 for item in legacy_results),\n'''
new = '''        "current_http_checks": http_checks,\n        "current_http_check_count": len(http_checks),\n        "test_count": legacy_assertion_count + len(direct_checks) + len(http_checks),\n        "all_returncodes_zero": all(item["returncode"] == 0 for item in legacy_results),\n'''
if old not in text:
    raise SystemExit("Expected application QA block was not found; refusing an unverified patch")
path.write_text(text.replace(old, new, 1), encoding="utf-8")
print({"application_aggregate_test_count_patch": "applied", "builder": str(path)})
