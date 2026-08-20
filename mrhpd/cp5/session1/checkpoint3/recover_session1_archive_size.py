#!/usr/bin/env python3
from pathlib import Path

path = Path(__file__).resolve().parent / "build_section5_session1_complete_restore.py"
text = path.read_text(encoding="utf-8")

# Extend compaction to checkpoint-specific report, QA, data, artwork, application,
# and source-control derivatives that are superseded by the terminal database,
# workbook, report, QA, Source Index, Bit Index, manifest, and recovery package.
anchor = '''        "Tracking/Prompt Response/Through Response 76/**",
    ]'''
replacement = '''        "Tracking/Prompt Response/Through Response 76/**",
        "Reports/Section 5 Session 1/Checkpoint 1/**",
        "Reports/Section 5 Session 1/Checkpoint 2/**",
        "QA/Section 5 Session 1/Checkpoint 1/**",
        "QA/Section 5 Session 1/Checkpoint 2/**",
        "Artwork/Section 5 Print Production/Checkpoint 1/**",
        "Artwork/Section 5 Print Production/Checkpoint 2/**",
        "Data/Section 5 Session 1 Checkpoint 1/**",
        "Data/Section 5 Session 1 Checkpoint 2/**",
        "App/Section 5 Session 1 Checkpoint 1/**",
        "App/Section 5 Session 1 Checkpoint 2/**",
        "Sources/Print Production/Response 75/**",
        "Sources/Print Production/Response 76/**",
    ]'''
if anchor not in text:
    raise SystemExit("compaction pattern anchor not found")
text = text.replace(anchor, replacement, 1)

# Record the exact recoverable archive-size condition in the project history.
event_anchor = '''        {
            "event_number": 198,
            "event_code": "V3-CP5-S1-REC-198-SELF-CONTAINED-RESTORE-AND-TRANSPORT",
            "condition": "The session boundary requires a complete restore requiring no earlier checkpoint, cloud artifact, or conversation reconstruction.",
            "recovery": "Embedded the complete current project archive with deterministic verification and extraction tools, clean-tested the restore, and divided it into the minimum two connector-compatible transport volumes with hashes and automated reassembly.",
            "status": "recovered",
            "recorded_at": now_iso,
        },
    ]'''
event_replacement = '''        {
            "event_number": 198,
            "event_code": "V3-CP5-S1-REC-198-SELF-CONTAINED-RESTORE-AND-TRANSPORT",
            "condition": "The session boundary requires a complete restore requiring no earlier checkpoint, cloud artifact, or conversation reconstruction.",
            "recovery": "Embedded the complete current project archive with deterministic verification and extraction tools, clean-tested the restore, and divided it into the minimum two connector-compatible transport volumes with hashes and automated reassembly.",
            "status": "recovered",
            "recorded_at": now_iso,
        },
        {
            "event_number": 199,
            "event_code": "V3-CP5-S1-REC-199-INITIAL-TERMINAL-ARCHIVE-OVER-180-MIB",
            "condition": "The first otherwise-valid terminal project ZIP contained 861 safe, unique, non-filler members but measured 222,738,173 bytes, exceeding the governed 180 MiB ceiling.",
            "recovery": "Preserved the copied terminal state and all canonical clinical/publication/print assets; broadened compaction only to superseded checkpoint-specific report, QA, data, artwork, application-pointer, source-control, index, manifest, tracking, database, and workbook derivatives with terminal-equivalence evidence; then rebuilt every downstream index, manifest, archive, restore, and transport gate.",
            "status": "recovered",
            "recorded_at": now_iso,
        },
    ]'''
if event_anchor not in text:
    raise SystemExit("recovery event anchor not found")
text = text.replace(event_anchor, event_replacement, 1)

path.write_text(text, encoding="utf-8")
print({"status": "passed", "builder": str(path)})
