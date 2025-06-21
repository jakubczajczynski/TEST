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
# @GHInput: InnerVerticalFormat (object) 
# @GHInput: InnerVerticalColor (object) 
# @GHInput: InnerVFormat (object) 
# @GHInput: InnerHFormat (object) 
# @GHInput: InnerHorizontalColor (object) 
# @GHInput: InnerHorizontalWeight (object) 
# @GHInput: DiagDownFormat (object) 
# @GHInput: DiagonalDownColor (object) 
# @GHInput: DiagonalDownWeight (object) 
# @GHInput: DiagUpFormat (object) 
# @GHInput: DiagonalUpColor (object) 
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
#   TopEdgeFormat         : int or None   - Border format code for top edge (0-14)
#   BottomEdgeFormat      : int or None   - Border format code for bottom edge (0-14)
#   LeftEdgeFormat        : int or None   - Border format code for left edge (0-14)
#   RightEdgeFormat       : int or None   - Border format code for right edge (0-14)
#   InnerVFormat          : int or None   - Border format code for inner vertical edges (0-14)
#   InnerHFormat          : int or None   - Border format code for inner horizontal edges (0-14)
#   DiagDownFormat        : int or None   - Border format code for diagonal-down edges (0-14)
#   DiagUpFormat          : int or None   - Border format code for diagonal-up edges (0-14)
#
# Border format codes:
#   0  = No border
#   1  = Continuous, Hairline
#   2  = Continuous, Thin
#   3  = Continuous, Medium
#   4  = Continuous, Thick
#   5  = Dashed, Thin
#   6  = Dotted, Thin
#   7  = Dash-Dot, Thin
#   8  = Dash-Dot-Dot, Thin
#   9  = Slant-Dash-Dot, Thin
#   10 = Double, Thin
#   11 = Gray25 pattern (shaded)
#   12 = Gray50 pattern (shaded)
#   13 = Gray75 pattern (shaded)
#   14 = Automatic (as in GUI automatic color/style)
#
# Output:
#   Success               : bool          - True if formatting applied without errors

import os
import xlwings as xw
from win32com.client import constants as c

# Combined map: format code to (LineStyle constant, Weight constant)
FORMAT_MAP = {
    0:  (c.xlLineStyleNone,      None),
    1:  (c.xlContinuous,         c.xlHairline),
    2:  (c.xlContinuous,         c.xlThin),
    3:  (c.xlContinuous,         c.xlMedium),
    4:  (c.xlContinuous,         c.xlThick),
    5:  (c.xlDash,               c.xlThin),
    6:  (c.xlDot,                c.xlThin),
    7:  (c.xlDashDot,            c.xlThin),
    8:  (c.xlDashDotDot,         c.xlThin),
    9:  (c.xlSlantDashDot,       c.xlThin),
    10: (c.xlDouble,             c.xlThin),
    11: (c.xlGray25,             None),
    12: (c.xlGray50,             None),
    13: (c.xlGray75,             None),
    14: (c.xlAutomatic,          None)
}

Success = False

def set_border(rng, edge_const, fmt):
    """
    Apply combined border format code.
    fmt: None=skip; integer code maps via FORMAT_MAP.
    """
    if fmt is None:
        return
    try:
        code = int(fmt)
    except Exception:
        return
    # Clamp code to available keys
    if code not in FORMAT_MAP:
        code = 0
    line_style, weight = FORMAT_MAP[code]
    border = rng.Borders(edge_const)
    border.LineStyle = line_style
    if weight is not None:
        border.Weight = weight

if Trigger:
    try:
        if not os.path.exists(FilePath):
            raise FileNotFoundError(f"Excel file not found: {FilePath}")
        try:
            wb = xw.Book(FilePath)
        except Exception:
            wb = xw.apps.active.books.open(FilePath)
        sht = wb.sheets[SheetName]
        com_rng = sht.range(CellRange).api

        edges = [
            (c.xlEdgeTop,        TopEdgeFormat),
            (c.xlEdgeBottom,     BottomEdgeFormat),
            (c.xlEdgeLeft,       LeftEdgeFormat),
            (c.xlEdgeRight,      RightEdgeFormat),
            (c.xlInsideVertical, InnerVFormat),
            (c.xlInsideHorizontal,InnerHFormat),
            (c.xlDiagonalDown,   DiagDownFormat),
            (c.xlDiagonalUp,     DiagUpFormat)
        ]
        for edge_const, fmt in edges:
            set_border(com_rng, edge_const, fmt)

        wb.save()
        Success = True
    except Exception as e:
        print(f"Error applying borders: {e}")
        Success = False
else:
    Success = False