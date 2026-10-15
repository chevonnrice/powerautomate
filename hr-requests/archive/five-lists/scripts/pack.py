"""Packs ../sourcecode into ../solution/hr-change-orchestration.zip (unmanaged solution).

Produces the same layout as `pac solution pack` for flows + environment variables, so it
can be used where the Power Platform CLI is not installed. If you have pac, prefer:

    pac solution pack --zipfile solution/hr-change-orchestration.zip --folder sourcecode
"""
import glob
import os
import re
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, "..", "sourcecode")
OUT = os.path.join(HERE, "..", "solution", "hr-change-orchestration.zip")
BOM = "﻿"


def read(path):
    with open(path, encoding="utf-8-sig") as f:
        return f.read()


workflow_xml = []
for data in sorted(glob.glob(os.path.join(SRC, "Workflows", "*.json.data.xml"))):
    body = re.sub(r"^<\?xml[^>]*\?>\s*", "", read(data))
    body = body.replace(' xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"', "")
    workflow_xml.append("\n".join("    " + line for line in body.splitlines()))

customizations = read(os.path.join(SRC, "Other", "Customizations.xml")).replace(
    "<Workflows />", "<Workflows>\n" + "\n".join(workflow_xml) + "\n  </Workflows>"
)

env_files = sorted(glob.glob(os.path.join(SRC, "environmentvariabledefinitions", "*", "*")))
overrides = "".join(
    f'<Override PartName="/{os.path.relpath(p, SRC).replace(os.sep, "/")}" ContentType="application/octet-stream" />'
    for p in env_files
    if p.endswith(".xml")
)
content_types = (
    BOM + '<?xml version="1.0" encoding="utf-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="xml" ContentType="text/xml" /><Default Extension="json" ContentType="application/octet-stream" />'
    + overrides
    + "</Types>"
)

os.makedirs(os.path.dirname(OUT), exist_ok=True)
with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("customizations.xml", BOM + customizations.lstrip(BOM))
    z.writestr("solution.xml", BOM + read(os.path.join(SRC, "Other", "Solution.xml")))
    for wf in sorted(glob.glob(os.path.join(SRC, "Workflows", "*.json"))):
        z.write(wf, "Workflows/" + os.path.basename(wf))
    for p in env_files:
        z.write(p, os.path.relpath(p, SRC).replace(os.sep, "/"))
    z.writestr("[Content_Types].xml", content_types)

print(f"wrote {os.path.normpath(OUT)}")
