#!/usr/bin/env python3
from pathlib import Path

builder = Path(__file__).resolve().parent / "build_section5_session2_complete_restore.py"
text = builder.read_text(encoding="utf-8")
code = "V3-CP5-S2-REC-232-WORKBOOK-EXTENSION-NAMESPACE-ROOT"
if code not in text:
    start = text.index("def recovery_events")
    end = text.index("\ndef acceptance_gates", start)
    block = text[start:end]
    closing = block.find("\n    ]", block.index("rows = ["))
    if closing < 0:
        raise RuntimeError("recovery event rows closing not found")
    event = (
        "\n        (232, \"V3-CP5-S2-REC-232-WORKBOOK-EXTENSION-NAMESPACE-ROOT\", "
        "\"The first donor-extension recovery placed namespace declarations relative to the XML declaration rather than the worksheet root, and openpyxl subsequently raised xml.etree.ElementTree.ParseError: unbound prefix while validating the regenerated workbook.\", "
        "\"Preserved the verified Response 80 source and donor workbook, changed namespace discovery and insertion to the actual worksheet root element, retained exact donor extension-block parity, and reran workbook validation plus every dependent project, archive, restore, transport, and delivery gate.\"),"
    )
    block = block[:closing] + event + block[closing:]
    text = text[:start] + block + text[end:]
text = text.replace('"RECOVERY_EVENTS_221_231.json"', '"RECOVERY_EVENTS_221_232.json"')
builder.write_text(text, encoding="utf-8")
print({"status": "passed", "builder": str(builder), "event": code})
