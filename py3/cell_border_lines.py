# @GHInput: FilePath (object) 
# @GHInput: SheetName (object) 
# @GHInput: CellRange (object) 
# @GHInput: TopEdgeFormat (object) 
# @GHInput: TopEdgeColor (object) 
# @GHInput: TopEdgeWeight (object) 
# @GHInput: BottomEdgeFormat (object) 
# @GHInput: BottomEdgeColor (object) 
# @GHInput: BottomEdgeWeight (object) 
# @GHInput: LeftEdgeFormat (object) 
# @GHInput: LeftEdgeColor (object) 
# @GHInput: LeftEdgeWeight (object) 
# @GHInput: RightEdgeFormat (object) 
# @GHInput: RightEdgeColor (object) 
# @GHInput: RightEdgeWeight (object) 
# @GHInput: InnerVColor (object) 
# @GHInput: InnerVFormat (object) 
# @GHInput: InnerHFormat (object) 
# @GHInput: InnerHColor (object) 
# @GHInput: InnerHorizontalWeight (object) 
# @GHInput: DiagDownFormat (object) 
# @GHInput: DiagDownColor (object) 
# @GHInput: DiagonalDownWeight (object) 
# @GHInput: DiagUpFormat (object) 
# @GHInput: DiagUpColor (object) 
# @GHInput: DiagonalUpWeight (object) 
# @GHInput: Trigger (object) 
# @GHOutput: Success (object) 

# Script for Grasshopper Python Component to customize Excel cell borders using xlwings
#
# Inputs:
#   FilePath              : str or None   - Full path to the Excel workbook
#   SheetName             : str or None   - Name of the worksheet
#   CellRange             : str or None   - Excel range address, e.g. "A1:D10"
#   Trigger               : bool or None  - When True, apply formatting; if False or None, skip
#   TopEdgeFormat         : int or None   - Border style code for top edge (0-15)
#   TopEdgeColor          : Color or tuple or None - Grasshopper Color Swatch or RGB tuple
#   BottomEdgeFormat      : int or None   - Border style code for bottom edge (0-15)
#   BottomEdgeColor       : Color or tuple or None
#   LeftEdgeFormat        : int or None   - Border style code for left edge (0-15)
#   LeftEdgeColor         : Color or tuple or None
#   RightEdgeFormat       : int or None   - Border style code for right edge (0-15)
#   RightEdgeColor        : Color or tuple or None
#   InnerVFormat          : int or None   - Border style code for inner vertical edges (0-15)
#   InnerVColor           : Color or tuple or None
#   InnerHFormat          : int or None   - Border style code for inner horizontal edges (0-15)
#   InnerHColor           : Color or tuple or None
#   DiagDownFormat        : int or None   - Border style code for diagonal-down edges (0-15)
#   DiagDownColor         : Color or tuple or None
#   DiagUpFormat          : int or None   - Border style code for diagonal-up edges (0-15)
#   DiagUpColor           : Color or tuple or None
#
# Border style codes:
#   0  = No border
#   1  = Continuous, Hairline
#   2  = Continuous, Thin
#   3  = Continuous, Medium
#   4  = Continuous, Thick
#   5  = Dashed, Thin
#   6  = Dashed, Medium
#   7  = Dotted, Thin
#   8  = Dash-Dot, Thin
#   9  = Dash-Dot-Dot, Thin
#   10 = Slant-Dash-Dot, Thin
#   11 = Double, Thin
#   12 = Gray25 pattern (shaded), Medium
#   13 = Gray50 pattern (shaded)
#   14 = Gray75 pattern (shaded)
#   15 = Automatic (Excel automatic)
#
# Output:
#   Success               : bool          - True if formatting applied without errors

import os
import xlwings as xw
from win32com.client import constants as c

# Map style code to (LineStyle, Weight)
FORMAT_MAP = {
    0:  (c.xlLineStyleNone,      None),
    1:  (c.xlContinuous,         c.xlHairline),
    2:  (c.xlContinuous,         c.xlThin),
    3:  (c.xlContinuous,         c.xlMedium),
    4:  (c.xlContinuous,         c.xlThick),
    5:  (c.xlDash,               c.xlThin),
    6:  (c.xlDash,               c.xlMedium),
    7:  (c.xlDot,                c.xlThin),
    8:  (c.xlDashDot,            c.xlThin),
    9:  (c.xlDashDotDot,         c.xlThin),
    10: (c.xlSlantDashDot,       c.xlThin),
    11: (c.xlDouble,             c.xlThin),
    12: (c.xlGray25,             c.xlMedium),
    13: (c.xlGray50,             c.xlThin),
    14: (c.xlGray75,             c.xlThin),
    15: (c.xlAutomatic,          c.xlThin)
}

Success = False


def normalize_color(col):
    """
    Normalize GH Color Swatch or tuple to an (R,G,B) tuple of ints 0-255.
    Supports .R,.G,.B attributes or tuple/list of 3 ints/floats.
    """
    if col is None:
        return None
    # GH swatch
    if hasattr(col, 'R') and hasattr(col, 'G') and hasattr(col, 'B'):
        r, g, b = col.R, col.G, col.B
    else:
        try:
            r, g, b = col
        except Exception:
            return None
    # scale floats in 0-1 to 0-255
    if isinstance(r, float) and r <= 1.0:
        r, g, b = int(r*255), int(g*255), int(b*255)
    # clamp
    r = max(0, min(255, int(r)))
    g = max(0, min(255, int(g)))
    b = max(0, min(255, int(b)))
    return (r, g, b)


def set_border(rng, edge_const, fmt, color_val):
    """
    Apply style and color to one border edge, overwriting previous settings.
    fmt: int code (0-15); color_val: GH swatch or tuple or None
    """
    if fmt is None:
        return
    try:
        code = int(fmt)
    except Exception:
        return
    code = max(0, min(15, code))
    line_style, weight = FORMAT_MAP[code]
    border = rng.Borders(edge_const)
    # clear previous style
    border.LineStyle = c.xlLineStyleNone
    if code == 0:
        return
    # apply new style and weight
    border.LineStyle = line_style
    if weight is not None:
        border.Weight = weight
    # apply color
    rgb = normalize_color(color_val)
    if rgb:
        r, g, b = rgb
        # Excel expects R + G*256 + B*65536
        border.Color = r + (g << 8) + (b << 16)
    else:
        border.ColorIndex = c.xlColorIndexAutomatic

if Trigger:
    try:
        if not os.path.exists(FilePath):
            raise FileNotFoundError(f"Excel not found: {FilePath}")
        try:
            wb = xw.Book(FilePath)
        except Exception:
            wb = xw.apps.active.books.open(FilePath)
        sht = wb.sheets[SheetName]
        com_rng = sht.range(CellRange).api

        edges = [
            (c.xlEdgeTop,        TopEdgeFormat,   TopEdgeColor),
            (c.xlEdgeBottom,     BottomEdgeFormat,BottomEdgeColor),
            (c.xlEdgeLeft,       LeftEdgeFormat,  LeftEdgeColor),
            (c.xlEdgeRight,      RightEdgeFormat, RightEdgeColor),
            (c.xlInsideVertical, InnerVFormat,    InnerVColor),
            (c.xlInsideHorizontal,InnerHFormat,   InnerHColor),
            (c.xlDiagonalDown,   DiagDownFormat,  DiagDownColor),
            (c.xlDiagonalUp,     DiagUpFormat,    DiagUpColor)
        ]
        for edge_const, fmt, clr in edges:
            set_border(com_rng, edge_const, fmt, clr)

        wb.save()
        Success = True
    except Exception as e:
        print(f"Error applying borders: {e}")
        Success = False
else:
    Success = False