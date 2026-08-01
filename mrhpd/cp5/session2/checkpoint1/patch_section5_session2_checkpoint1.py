#!/usr/bin/env python3
from pathlib import Path

path = Path(__file__).resolve().parent / "build_section5_session2_checkpoint1.py"
text = path.read_text(encoding="utf-8")

# Carry the independent-verifier recoveries into every current project surface.
function_start = text.index("def recovery_rows(now_iso: str)")
function_end = text.index("\n\ndef clone_response_row", function_start)
block = text[function_start:function_end]
events = [
    (
        "V3-CP5-S2-REC-209-INDEPENDENT-VERIFIER-NESTED-REPORT-PATH",
        "The first independent output verifier searched the distribution root for report files that were intentionally nested inside the current-turn delivery wrapper and raised StopIteration after the substantive build and clean-application gates had passed.",
        "Preserved the completed copied project, recovery ZIP, reports, database, workbook, indexes, manifest, and clean-application evidence; changed only the independent verifier to safely extract the delivery wrapper before opening the report PDF and workbook; then reran the complete build, clean application, independent verification, and packaging workflow.",
    ),
    (
        "V3-CP5-S2-REC-210-RECOVERY-PATCH-LIST-ANCHOR",
        "The first recovery-patch implementation selected the final list terminator in recovery_rows rather than the events-list terminator, inserted the new tuple into the return comprehension, and py_compile raised SyntaxError before the build began.",
        "Preserved the original builder and all verified prior artifacts, changed the patch to anchor to the first events-list terminator after 'events = [', added a compile gate, and restarted the complete workflow from the immutable Response 77 baseline.",
    ),
]
insertions = []
for event_code, condition, recovery in events:
    if event_code not in block:
        insertions.append(f"        ({event_code!r}, {condition!r}, {recovery!r}),\n")
if insertions:
    events_start = block.index("    events = [")
    closing = block.index("\n    ]", events_start)
    block = block[:closing] + "".join(insertions) + block[closing:]
    text = text[:function_start] + block + text[function_end:]

# Make all long ReportLab table cells wrap as Paragraph flowables rather than
# drawing unbounded single-line strings across adjacent columns.
mr_body = '    styles.add(ParagraphStyle(name="MRBody", parent=styles["BodyText"], fontName="Helvetica", fontSize=8.5, leading=11, textColor=colors.HexColor("#24323D"), spaceAfter=5))\n'
mr_cell = mr_body + '    styles.add(ParagraphStyle(name="MRCell", parent=styles["BodyText"], fontName="Helvetica", fontSize=6.3, leading=7.4, textColor=colors.HexColor("#24323D"), spaceAfter=0))\n'
if 'name="MRCell"' not in text:
    if mr_body not in text:
        raise SystemExit("MRBody style anchor not found")
    text = text.replace(mr_body, mr_cell, 1)

for data_name in ("disposition", "preview_data", "proof_data", "check_data"):
    marker = f"    table=Table({data_name},"
    insertion = f"    {data_name} = [{data_name}[0]] + [[Paragraph(str(cell), styles[\"MRCell\"]) for cell in row] for row in {data_name}[1:]]\n"
    if insertion.strip() not in text:
        if marker not in text:
            raise SystemExit(f"table anchor not found: {data_name}")
        text = text.replace(marker, insertion + marker, 1)

path.write_text(text, encoding="utf-8")
print({
    "status": "passed",
    "builder": str(path),
    "recovery_events": [event[0] for event in events],
    "wrapped_tables": ["disposition", "preview_data", "proof_data", "check_data"],
})
