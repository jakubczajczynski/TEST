# @GHInput: FilePath (object) 
# @GHInput: SheetName (object) 
# @GHInput: CellRange (object) 
# @GHInput: FontName (object) 
# @GHInput: FontSize (object) 
# @GHInput: Bold (object) 
# @GHInput: Italic (object) 
# @GHInput: Color (object) 
# @GHInput: Trigger (object) 
# @GHOutput: styled (object) 

"""
name: Excel Text Styler
description: Apply font styling to a given Excel range via xlwings.
inputs:
  - name: FilePath
    type: str
    description: Full path to the Excel file
  - name: SheetName
    type: str
    description: Name of the sheet (e.g. "Sheet1")
  - name: CellRange
    type: str
    description: Excel range address, e.g. "A1:C3"
  - name: FontName
    type: str
    default: null
    description: Font family name; leave unconnected to keep existing
  - name: FontSize
    type: float
    default: null
    description: Size in points; leave unconnected to keep existing
  - name: Bold
    type: bool
    default: null
    description: Bold text; leave unconnected to keep existing
  - name: Italic
    type: bool
    default: null
    description: Italic text; leave unconnected to keep existing
  - name: Color
    type: Color
    default: null
    description: Grasshopper Colour Swatch (System.Drawing.Color) or (R,G,B) tuple; leave unconnected to keep existing
  - name: Trigger
    type: bool
    default: True
    description: When True, apply the styling
outputs:
  - name: Styled
    type: bool
    description: True if styling was applied successfully
"""

import xlwings as xw
import os
import System

def stylize_text(
    filepath,
    sheet_name,
    cell_range,
    font_name=None,
    font_size=None,
    bold=None,
    italic=None,
    color=None
):
    # Open or attach to workbook
    if not os.path.exists(filepath):
        raise FileNotFoundError(f"Excel file not found: {filepath}")
    try:
        wb = xw.Book(filepath)
    except Exception:
        wb = xw.apps.active.books.open(filepath)

    # Access sheet and range
    sht = wb.sheets[sheet_name]
    rng = sht.range(cell_range)
    font = rng.api.Font

    # Conditionally apply only those styles with inputs
    if font_name is not None:
        font.Name = font_name
    if font_size is not None:
        font.Size = font_size
    if bold is not None:
        font.Bold = bool(bold)
    if italic is not None:
        font.Italic = bool(italic)
    if color is not None:
        # Handle GH Colour Swatch (System.Drawing.Color) or tuple
        try:
            # .R/.G/.B exist on System.Drawing.Color
            r, g, b = color.R, color.G, color.B
        except Exception:
            # Assume it's a tuple/list
            r, g, b = color
        font.Color = int(r) + int(g) * 256 + int(b) * 256**2

    return True

# Main execution
if Trigger:
    try:
        Styled = stylize_text(
            FilePath,
            SheetName,
            CellRange,
            FontName,
            FontSize,
            Bold,
            Italic,
            Color
        )
    except Exception as e:
        Styled = False
        print(f"Error styling text: {e}")
else:
    Styled = False