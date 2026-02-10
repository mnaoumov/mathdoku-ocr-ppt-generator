"""
One-time helper: creates vba_template.pptm with the VBA module embedded.

Run this once (requires PowerPoint installed):
    python create_vba_template.py

The generated vba_template.pptm is then used by make_mathdoku_pptx.py
to produce .pptm files with the macro already inside.

NOTE: PowerPoint's Trust Center must allow VBA project access:
  File > Options > Trust Center > Trust Center Settings >
  Macro Settings > "Trust access to the VBA project object model"
"""

from __future__ import annotations

import os
import sys
import time
import zipfile
from io import BytesIO

_CUSTOM_UI_XML = b"""\
<?xml version="1.0" encoding="UTF-8"?>
<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui">
  <ribbon>
    <tabs>
      <tab id="mathdokuTab" label="Mathdoku">
        <group id="grpSolve" label="Solve">
          <button id="btnEnterValue" label="Enter Value"
                  onAction="RibbonEnterFinalValue"
                  size="large"/>
          <button id="btnEditCandidates" label="Edit Candidates"
                  onAction="RibbonEditCellCandidates"
                  size="large"/>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>
"""


def _inject_ribbon_xml(pptm_path: str) -> None:
    """Embed Ribbon XML (customUI.xml) into a .pptm for a custom ribbon tab."""
    raw = open(pptm_path, "rb").read()
    buf = BytesIO(raw)
    out = BytesIO()
    with zipfile.ZipFile(buf, "r") as zin, zipfile.ZipFile(out, "w") as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "[Content_Types].xml":
                data = data.replace(
                    b"</Types>",
                    b'<Override PartName="/customUI/customUI.xml"'
                    b' ContentType="application/xml"/>\n</Types>',
                )
            elif item.filename == "_rels/.rels":
                data = data.replace(
                    b"</Relationships>",
                    b'<Relationship Id="rCustomUI"'
                    b' Type="http://schemas.microsoft.com/office/2006/relationships/ui/extensibility"'
                    b' Target="customUI/customUI.xml"/>\n</Relationships>',
                )
            zout.writestr(item, data)
        zout.writestr("customUI/customUI.xml", _CUSTOM_UI_XML)
    with open(pptm_path, "wb") as f:
        f.write(out.getvalue())


def main() -> None:
    bas_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "MathdokuCandidatesMacro.bas")
    if not os.path.isfile(bas_path):
        print(f"Error: {bas_path} not found")
        raise SystemExit(1)

    template_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "vba_template.pptm")

    import win32com.client  # type: ignore[import-untyped]

    pptApp = win32com.client.Dispatch("PowerPoint.Application")
    pptApp.Visible = True

    try:
        prs = pptApp.Presentations.Add()

        # Import the VBA module
        prs.VBProject.VBComponents.Import(bas_path)

        # Remove default blank slide if present
        while prs.Slides.Count > 0:
            prs.Slides(1).Delete()

        # Save as macro-enabled .pptm (25 = ppSaveAsOpenXMLPresentationMacroEnabled)
        prs.SaveAs(template_path, 25)
        prs.Close()
        print(f"Created {template_path}")
    finally:
        pptApp.Quit()

    # Inject Ribbon XML for QAT buttons after PowerPoint saves the file
    _inject_ribbon_xml(template_path)
    print("Injected Ribbon XML (QAT buttons) into template.")


if __name__ == "__main__":
    main()
