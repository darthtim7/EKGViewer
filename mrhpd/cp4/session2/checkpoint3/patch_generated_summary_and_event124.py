#!/usr/bin/env python3
"""Patch the disposable Response 69 builder after generation.

The prior run completed the governed project and complete-restore build, but the
workflow verifier expected a Session 2 summary filename while the mechanically
adapted builder retained ``SESSION1`` in that one uppercase filename token.
This patch normalizes the filename, corrects the remaining machine-readable
session-completion key, and records Recovery Event 124 before the builder is
executed again from the exact verified inputs.
"""
from __future__ import annotations

import re
from pathlib import Path

PATH = Path("mrhpd/cp4/session2/checkpoint3/build_session2_complete_restore.py")

EVENT_124 = '''    {
        "event_number": 124,
        "event_code": "V3-CP4-S2-REC-VERIFICATION-SUMMARY-FILENAME-EXPECTATION-CORRECTED",
        "occurred_at": NOW,
        "failed_step": "Read the Response 69 Session 2 build summary during the post-build workflow verification step.",
        "exact_error_or_reason": "The complete project, complete restore, both transport volumes, and controls were built successfully, but the mechanically adapted builder retained SESSION1 in the uppercase build-summary filename while the verifier expected SESSION2.",
        "intact_artifacts": "The 182,222,320-byte complete project archive, 179,611,198-byte complete restore, both 89,809,675-byte transport wrappers, database, workbook, application, 537-page publication, editable assembly, reports, indexes, manifests, and clean-restore evidence remained intact in the disposable run.",
        "recovery_action": "Normalized the emitted build-summary filename to MRHPD_RESPONSE69_SESSION2_COMPLETE_BUILD_SUMMARY.json, corrected the remaining session_1_of_3_complete machine key, regenerated from the exact verified Response 66 and Response 68 inputs, and reran all build, verification, transport, and upload gates.",
        "validation_result": "Pending the rerun; successful completion requires the canonical summary filename, complete workflow verification, and artifact upload.",
        "data_quality_effect": "None; filename and machine-readable session-label correction only.",
        "next_checkpoint": "Store the verified Response 69 restore volumes and controls in Google Drive, then begin Remediation Section 4 Session 3 of 3.",
    },
'''


def main() -> None:
    text = PATH.read_text(encoding="utf-8")
    original = text

    # Normalize the one uppercase no-underscore SESSION1 filename token that
    # was not covered by the mechanical session relabeling pass.
    text = text.replace(
        "MRHPD_RESPONSE69_SESSION1_COMPLETE_BUILD_SUMMARY.json",
        "MRHPD_RESPONSE69_SESSION2_COMPLETE_BUILD_SUMMARY.json",
    )

    # Correct any remaining machine-readable completion key in the generated
    # Response 69 output surface.
    text = text.replace("session_1_of_3_complete", "session_2_of_3_complete")

    if "V3-CP4-S2-REC-VERIFICATION-SUMMARY-FILENAME-EXPECTATION-CORRECTED" not in text:
        anchor = "]\n\nNET_PROMPT ="
        if anchor not in text:
            raise SystemExit("Recovery Event 124 insertion anchor not found")
        text = text.replace(anchor, EVENT_124 + anchor, 1)

    text = text.replace("RECOVERY_EVENTS_116_123.json", "RECOVERY_EVENTS_116_124.json")

    # Fail fast if the canonical summary output is still absent or stale.
    canonical = "MRHPD_RESPONSE69_SESSION2_COMPLETE_BUILD_SUMMARY.json"
    stale = "MRHPD_RESPONSE69_SESSION1_COMPLETE_BUILD_SUMMARY.json"
    if canonical not in text or stale in text:
        raise SystemExit({"canonical_present": canonical in text, "stale_present": stale in text})

    if text == original:
        print("Generated Response 69 builder was already normalized.")
    else:
        PATH.write_text(text, encoding="utf-8")
        print("Normalized Response 69 summary filename, session key, and Recovery Event 124.")


if __name__ == "__main__":
    main()
