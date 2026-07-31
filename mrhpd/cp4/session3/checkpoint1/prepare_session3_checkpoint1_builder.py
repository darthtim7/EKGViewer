#!/usr/bin/env python3
"""Generate the Section 4 Session 3 Checkpoint 1 recovery builder.

The verified Session 2 Checkpoint 1 builder already implements the immutable
baseline, copied working tree, database/workbook/application/publication,
tracking, report, index, manifest, overlay, clean-apply, and recovery lanes.
This adapter advances those governed mechanics to the exact Response 69
session-boundary restore and Response 70 / Session 3 labels. A separate
integration patch adds the final-session release-governance controls.
"""
from __future__ import annotations

import io
import tokenize
from pathlib import Path

SOURCE = Path("mrhpd/cp4/session2/checkpoint1/build_checkpoint1_recovery.py")
TARGET = Path("mrhpd/cp4/session3/checkpoint1/build_checkpoint1_recovery.py")


def replace_number_tokens(source: str, mapping: dict[str, str]) -> str:
    tokens = []
    for token in tokenize.generate_tokens(io.StringIO(source).readline):
        if token.type == tokenize.NUMBER and token.string in mapping:
            token = tokenize.TokenInfo(token.type, mapping[token.string], token.start, token.end, token.line)
        tokens.append(token)
    return tokenize.untokenize(tokens)


def replace_required(text: str, old: str, new: str, label: str, minimum: int = 1) -> str:
    count = text.count(old)
    if count < minimum:
        raise SystemExit(f"{label}: expected at least {minimum} match(es), found {count}")
    return text.replace(old, new)


text = SOURCE.read_text(encoding="utf-8")

# Advance the current response first so newly introduced Response 69 baseline
# labels are not accidentally advanced again.
for old, new in [
    ("Response 67", "Response 70"),
    ("response67", "response70"),
    ("RESPONSE67", "RESPONSE70"),
    ("Response_67", "Response_70"),
    ("RESPONSE_67", "RESPONSE_70"),
    ("R67", "R70"),
    ("r67", "r70"),
]:
    text = text.replace(old, new)
text = replace_number_tokens(text, {"67": "70"})

# Advance the current session and its governed identifiers.
for old, new in [
    ("Session 2", "Session 3"),
    ("session2", "session3"),
    ("SESSION2", "SESSION3"),
    ("SESSION_2", "SESSION_3"),
    ("section4_session2", "section4_session3"),
    ("Section 4 Session 2", "Section 4 Session 3"),
    ("S4S2", "S4S3"),
    ("CP4-S2", "CP4-S3"),
    ("cp4-s2", "cp4-s3"),
    ("cp4_s2", "cp4_s3"),
    ("session_2_of_3_complete", "session_3_of_3_complete"),
]:
    text = text.replace(old, new)

# Advance the immutable baseline from Response 66 to the exact verified
# Response 69 session-boundary restore.
for old, new in [
    ("Response 66", "Response 69"),
    ("response66", "response69"),
    ("RESPONSE66", "RESPONSE69"),
    ("Response_66", "Response_69"),
    ("RESPONSE_66", "RESPONSE_69"),
    ("R66", "R69"),
    ("r66", "r69"),
]:
    text = text.replace(old, new)

text = replace_required(text, "BASE_RESTORE_BYTES = 177_617_796", "BASE_RESTORE_BYTES = 179_612_090", "Response 69 restore byte identity")
text = replace_required(
    text,
    'BASE_RESTORE_SHA256 = "38c8fa08763d5698217ce33a2bbe1e889e726087575b14fb31086f38cfe1300f"',
    'BASE_RESTORE_SHA256 = "31e4ba64c7a36870ebeb01e4c88109d512a498ff069f44edfb48ba141044ebcb"',
    "Response 69 restore SHA-256",
)
text = replace_required(text, "BASE_PROJECT_BYTES = 169_294_854", "BASE_PROJECT_BYTES = 182_223_875", "Response 69 project byte identity")
text = replace_required(
    text,
    'BASE_PROJECT_SHA256 = "b59e5265c0515a5dbaadf55b631a37c581b828b1a37857ee3322cda532125cc4"',
    'BASE_PROJECT_SHA256 = "aa9d563357c375fd45b1e90e0e7f3465ebb855f1d4f997f817f46389ceb2b438"',
    "Response 69 project SHA-256",
)

# Final-session naming and learned recovery-event continuity.
text = text.replace("Capability Parity and Drift Baseline", "Final-Release Governance and Readiness Baseline")
text = text.replace("capability-parity and drift-baseline", "final-release governance and readiness baseline")
text = text.replace("RECOVERY_EVENTS_101_105.json", "RECOVERY_EVENTS_125_130.json")
text = text.replace("inspected 186 tables, 43-sheet current workbook", "inspected the 194-table canonical database and 58-sheet current workbook")
text = text.replace("five S4S3 workbook sheets", "five baseline S4S3 workbook sheets")

# Replace the inherited narrative with the exact current-state Net prompt,
# response, goal, and checkpoint semantics before capability definitions are
# evaluated.
anchor = "\nCAPABILITY_DEFINITIONS = ["
override = r'''

RESPONSE70.update({
    "title": "Section 4 Session 3 final-release governance and readiness baseline checkpoint",
    "goal": (
        "Begin the final Section 4 session from the exact verified Response 69 complete restore; establish explicit release-level "
        "clinical, evidence, graphics, publication, application, workbook, tracking, recovery, index, and manifest governance; "
        "preserve all immutable source artifacts; and emit deterministic Checkpoint 1 recovery through Response 70."
    ),
    "summary": (
        "Reconstructed and verified the exact Response 69 complete restore, created a copied Session 3 working tree, synchronized "
        "Response 70, established final-release governance gates and a controlled forward-work risk register, preserved the main "
        "application source and 537-page publication/editable assembly, expanded the comprehensive workbook and read-only audit "
        "surfaces, rebuilt tracking, reports, Source Index, Bit Index, manifests and checksums, and clean-applied the emitted recovery package."
    ),
    "notes": "Checkpoint 1 of 3 is complete. Checkpoint 2 performs final source/version, drift, publication, graphics, application, and workbook reconciliation; Checkpoint 3 emits the independently verified Section 4 release.",
})

NET_PROMPT = (
    "Continue the Human Pathogen Database from the exact verified Response 69 complete restore stored in persistent Google Drive custody. "
    "Begin Remediation Section 4 of 5 Session 3 of 3 in a copied working tree. Add Response 70 and recovery history; establish final-release "
    "clinical, evidence, graphics-rights, publication-navigation, application, workbook, tracking, recovery, Source Index, Bit Index, manifest, "
    "checksum, and QA governance; preserve the accepted predecessor, frozen Section 3 release, byte-identical main application, 537-page "
    "publication, and editable assembly; and emit deterministic Checkpoint 1 recovery tied directly to the exact Response 69 restore."
)

NET_RESPONSE = (
    "Section 4 Session 3 Checkpoint 1 is complete through Response 70. The exact Response 69 session-boundary restore and project snapshot "
    "were verified before work began. The copied current project now contains explicit final-release governance gates, controlled forward-work "
    "risks, expanded workbook surfaces, a read-only release-readiness application audit, current tracking, reports, indexes, manifests, and "
    "recovery controls. The main application source, 537-page publication, editable assembly, accepted predecessor, and frozen Section 3 "
    "release remain unchanged. Continue proceeds to Checkpoint 2 of 3."
)
'''
if anchor not in text:
    raise SystemExit("capability-definition insertion anchor not found")
text = text.replace(anchor, override + anchor, 1)

# Correct the current input/output defaults and display strings that are not
# covered by the semantic replacements above.
text = text.replace('default=Path("response69_artifacts")', 'default=Path("response69_artifacts")')
text = text.replace('default=Path("dist_cp4_s3_cp1")', 'default=Path("dist_cp4_s3_cp1")')
text = text.replace("Section 3 release remains byte-identical", "Section 3 release remains byte-identical")

required_markers = [
    "RESPONSE70",
    'BASE_RESTORE_BYTES = 179_612_090',
    'BASE_PROJECT_BYTES = 182_223_875',
    "reconstruct_response69",
    "section4_session3_checkpoint",
    "MRHPD-V3-CP4-S3-CP1",
    "S4S3 Capability",
    "RECOVERY DATA THROUGH RESPONSE 70",
    "--base-response69-restore",
    "session_3_of_3_complete",
]
missing = [marker for marker in required_markers if marker not in text]
if missing:
    raise SystemExit({"missing_generated_markers": missing})

prohibited_markers = [
    "RESPONSE67",
    "Response 67",
    "response67",
    "section4_session2",
    "MRHPD-V3-CP4-S2-CP1",
    "S4S2",
    "--base-response66-restore",
    "reconstruct_response66",
]
remaining = [marker for marker in prohibited_markers if marker in text]
if remaining:
    raise SystemExit({"prohibited_generated_markers": remaining})

TARGET.parent.mkdir(parents=True, exist_ok=True)
TARGET.write_text(text, encoding="utf-8")
print(f"Generated {TARGET} from {SOURCE} for Response 70 / Section 4 Session 3 Checkpoint 1.")
