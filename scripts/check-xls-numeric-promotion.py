#!/usr/bin/env python3
"""Optional independent xlrd oracle for PI numeric promotion/creation (no installs).

Usage: python3 scripts/check-xls-numeric-promotion.py ORIGINAL.xls EDITED.xls ROW COL
       [--sheet 0] [--created --xf INDEX]
Requires an existing xlrd installation; opens files as data only.
"""
import argparse
import math
import xlrd

parser = argparse.ArgumentParser(description=__doc__)
parser.add_argument("original")
parser.add_argument("edited")
parser.add_argument("row", type=int)
parser.add_argument("col", type=int)
parser.add_argument("--sheet", type=int, default=0)
parser.add_argument("--created", action="store_true", help="Target was absent; expect a new NUMBER")
parser.add_argument("--xf", type=int, help="Required target XF index for --created")
args = parser.parse_args()

original = xlrd.open_workbook(args.original, formatting_info=True)
edited = xlrd.open_workbook(args.edited, formatting_info=True)
assert original.sheet_names() == edited.sheet_names(), "Sheet names/order changed"
checked = 0
target_seen = False
for sheet_index in range(original.nsheets):
    before = original.sheet_by_index(sheet_index)
    after = edited.sheet_by_index(sheet_index)
    assert (before.nrows, before.ncols) == (after.nrows, after.ncols), "Sheet extent changed"
    for row in range(before.nrows):
        for col in range(before.ncols):
            previous, current = before.cell(row, col), after.cell(row, col)
            position = (sheet_index, row, col)
            is_target = position == (args.sheet, args.row, args.col)
            if is_target and args.created:
                assert args.xf is not None, "--created requires --xf"
                assert previous.ctype == xlrd.XL_CELL_EMPTY, "Creation target already exists"
                assert current.ctype == xlrd.XL_CELL_NUMBER, "Created target is not NUMBER"
                assert current.xf_index == args.xf, "Created target XF mismatch"
            else:
                assert previous.ctype == current.ctype, ("Cell type changed", position)
                assert previous.xf_index == current.xf_index, ("XF changed", position)
            expected = math.pi if is_target else previous.value
            assert current.value == expected, ("Unexpected value", position)
            target_seen |= is_target
            checked += 1
assert target_seen, "Target cell not present"
print(f"xlrd {xlrd.__version__}: {checked} cell slots / {original.nsheets} sheets passed; "
      "exact PI target, other values, cell types and XF indices preserved.")
