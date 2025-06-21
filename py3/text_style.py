# @GHInput: FilePath (object) 
# @GHInput: SheetName (object) 
# @GHInput: CellRange (object) 
# @GHInput: FontName (object) 
# @GHInput: FontSize (object) 
# @GHInput: Bold (object) 
# @GHInput: Italic (object) 
# @GHInput: UnderlineType (object) 
# @GHInput: Strikethrough (object) 
# @GHInput: Color (object) 
# @GHInput: Trigger (object) 
# @GHOutput: Styled (object) 

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
  - name: UnderlineType
    type: int
    default: null
    description: 0 = remove underline; 1 = single underline; 2 = double underline; leave unconnected to keep existing
  - name: Strikethrough
    type: bool
    default: null
    description: Strikethrough text; leave unconnected to keep existing
  - name: Color
    type: Color
    default: null
    description: GH Colour Swatch or (R,G,B) tuple or normalized float tuple; leave unconnected to keep existing
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
    underline_type=None,
    strikethrough=None,
    color=None
):
    # Open or attach to workbook
    if not os.path.exists(filepath):
        raise FileNotFoundError(f"Excel file not found: {filepath}")
    try:
        wb = xw.Book(filepath)
    except Exception:
        wb = xw.apps.active.books.open(filepath)

    sht = wb.sheets[sheet_name]
    rng = sht.range(cell_range)
    com_font = rng.api.Font

    # Name, size, bold, italic
    if font_name is not None:
        com_font.Name = font_name
    if font_size is not None:
        com_font.Size = font_size
    if bold is not None:
        com_font.Bold = bool(bold)
    if italic is not None:
        com_font.Italic = bool(italic)

    # Underline Type: 0 = none; 1 = single; 2 = double
    if underline_type is not None:
        if underline_type == 2:
            com_font.Underline = -4119  # xlUnderlineStyleDouble
        elif underline_type == 1:
            com_font.Underline = 2      # xlUnderlineStyleSingle
        else:
            com_font.Underline = -4142  # xlUnderlineStyleNone

    # Strikethrough
    if strikethrough is not None:
        com_font.Strikethrough = bool(strikethrough)

    # Color: System.Drawing.Color, int tuple, or normalized float tuple
    if color is not None:
        try:
            r, g, b = color.R, color.G, color.B
        except Exception:
            r, g, b = color
        # Scale floats 0–1 to 0–255
        if all(isinstance(c, float) and c <= 1.0 for c in (r, g, b)):
            r, g, b = [int(c * 255) for c in (r, g, b)]
        # Use xlwings wrapper for correct COM interop
        rng.font.color = (r, g, b)

    return True

if Trigger:
    try:
        Styled = stylize_text(
            FilePath, SheetName, CellRange,
            FontName, FontSize, Bold, Italic,
            UnderlineType, Strikethrough, Color
        )
    except Exception as e:
        Styled = False
        print(f"Error styling text: {e}")
else:
    Styled = False