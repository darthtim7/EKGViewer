#!/usr/bin/env python3
"""Replace one brittle whole-block substitution in the Session 3 patcher.

The generated clean-apply utility may contain harmless formatting differences
from its Session 2 source. This preparatory patch changes the integration
patcher to insert the release-governance checks at two small, semantic anchors
instead of replacing the entire utility tail verbatim.
"""
from __future__ import annotations

import hashlib
from pathlib import Path

PATCHER = Path("mrhpd/cp4/session3/checkpoint1/patch_generated_session3_integration.py")
START = "# Require the clean-apply utility to verify the new governance tables and run\n"
END = "# Correct the inherited application-audit minimum if the source capability\n"


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


text = PATCHER.read_text(encoding="utf-8")
start = text.find(START)
end = text.find(END)
if start < 0 or end < 0 or end <= start:
    raise SystemExit({"integration_patcher_section_not_found": {"start": start, "end": end}})

replacement = r'''# Require the clean-apply utility to verify the new governance tables and run
# both read-only application audits. Use small semantic anchors rather than a
# brittle replacement of the whole generated utility tail.
capability_anchor = ''' + '"""' + r'''  if con.execute("SELECT COUNT(*) FROM section4_session3_capability WHERE status!='passed'").fetchone()[0]!=0: raise SystemExit('capability registry failure')
''' + '"""' + r'''
capability_expansion = capability_anchor + ''' + '"""' + r'''  if con.execute("SELECT COUNT(*) FROM section4_session3_release_governance").fetchone()[0]<17: raise SystemExit('release-governance gate count failure')
  if con.execute("SELECT COUNT(*) FROM section4_session3_release_governance WHERE status!='passed'").fetchone()[0]!=0: raise SystemExit('release-governance gate failure')
  if con.execute("SELECT COUNT(*) FROM section4_session3_release_risk").fetchone()[0]<6: raise SystemExit('release-risk register failure')
''' + '"""' + r'''
if "release-governance gate count failure" not in text:
    text = replace_once(text, capability_anchor, capability_expansion, "clean-apply release-governance table gates")
    applied.append("clean-apply release-governance table gates")

application_anchor = ''' + '"""' + r''' if result.returncode: raise SystemExit('application capability audit failed: '+result.stderr[-2000:])
''' + '"""' + r'''
application_expansion = application_anchor + ''' + '"""' + r''' release_audit=a.output_dir/critical['release_readiness_audit']['path']
 release_result=subprocess.run([sys.executable,str(release_audit),'--db',str(db)],text=True,capture_output=True)
 if release_result.returncode: raise SystemExit('application release-readiness audit failed: '+release_result.stderr[-2000:])
''' + '"""' + r'''
if "application release-readiness audit failed" not in text:
    text = replace_once(text, application_anchor, application_expansion, "clean-apply release-readiness application audit")
    applied.append("clean-apply release-readiness application audit")

'''

updated = text[:start] + replacement + text[end:]
PATCHER.write_text(updated, encoding="utf-8")
print(
    {
        "status": "passed",
        "patcher": PATCHER.as_posix(),
        "sha256": sha256_file(PATCHER),
        "strategy": "semantic-anchor insertion",
    }
)
