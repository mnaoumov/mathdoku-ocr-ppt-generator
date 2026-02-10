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

        # Inject Presentation_Open into the document module so toolbar auto-appears.
        # Find the document module by type (100 = vbext_ct_Document) since the name
        # varies by PowerPoint version/locale.
        for i in range(1, prs.VBProject.VBComponents.Count + 1):
            comp = prs.VBProject.VBComponents.Item(i)
            if comp.Type == 100:  # vbext_ct_Document
                comp.CodeModule.InsertLines(
                    comp.CodeModule.CountOfLines + 1,
                    "Private Sub Presentation_Open()\r\n"
                    "    SetupToolbar\r\n"
                    "End Sub\r\n",
                )
                break

        # Remove default blank slide if present
        while prs.Slides.Count > 0:
            prs.Slides(1).Delete()

        # Save as macro-enabled .pptm (25 = ppSaveAsOpenXMLPresentationMacroEnabled)
        prs.SaveAs(template_path, 25)
        prs.Close()
        print(f"Created {template_path}")
    finally:
        pptApp.Quit()


if __name__ == "__main__":
    main()
