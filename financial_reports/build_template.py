"""Create a blank Trading Comps EV-bridge workbook.

Layout (column A = labels, one company per column from B onwards):

    1  Company                 14 Minorities
    2  BS Currency             16 Shares (outstanding at BS date)
    4  Cash                    18 Total Shares            = row 16
    5  Short Term Debt         19 Price                   (entered manually)
    6  ST Financial Lease      20 Equity Value            = shares x price
    7  ST Operating Lease      22 Net Debt                = rows 4,5,6,8,9
    8  Long Term Debt          23 Net Debt Post IFRS16    = net debt + op leases
    9  LT Financial Lease      24 Other Adjustments       = rows 11-14
    10 LT Operating Lease      25 Enterprise Value        = 20 + 23 + 24
    11 Affiliates
    12 Pensions
    13 Preferred Equity

Usage:
    python financial_reports/build_template.py [--out templates/TradingComps.xlsx] [--columns 10]
"""

import argparse
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

NUM_FMT = '#,##0.0;(#,##0.0);"-"'
FONT = "Aptos"

INPUT_ROWS = {
    4: "Cash",
    5: "Short Term Debt",
    6: "ST Financial Lease",
    7: "ST Operating Lease",
    8: "Long Term Debt",
    9: "LT Financial Lease",
    10: "LT Operating Lease",
    11: "Affiliates",
    12: "Pensions",
    13: "Preferred Equity",
    14: "Minorities",
    16: "Shares",
    19: "Price",
}

FORMULA_ROWS = {
    18: ("Total Shares", "={c}16"),
    20: ("Equity Value", "={c}18*{c}19"),
    22: ("Net Debt", "=SUM({c}4:{c}6)+{c}8+{c}9"),
    23: ("Net Debt Post IFRS16", "={c}22+{c}7+{c}10"),
    24: ("Other Adjustments", "=SUM({c}11:{c}14)"),
    25: ("Enterprise Value", "={c}20+{c}23+{c}24"),
}


def build(path, n_columns):
    wb = Workbook()
    ws = wb.active
    ws.title = "Comps"

    bold = Font(name=FONT, size=10, bold=True)
    plain = Font(name=FONT, size=10)
    black = Font(name=FONT, size=10, color="FF000000")
    header_fill = PatternFill("solid", fgColor="FFD9E1F2")
    total_border = Border(top=Side(style="thin"))

    ws["A1"], ws["A2"], ws["A3"] = "Company", "BS Currency", "EV Bridge (m)"
    for r in (1, 2, 3):
        ws.cell(row=r, column=1).font = bold
    for r, label in INPUT_ROWS.items():
        ws.cell(row=r, column=1, value=label).font = plain
    for r, (label, _) in FORMULA_ROWS.items():
        ws.cell(row=r, column=1, value=label).font = bold

    ws.column_dimensions["A"].width = 24
    for i in range(2, 2 + n_columns):
        c = get_column_letter(i)
        ws.column_dimensions[c].width = 14
        for r in (1, 2):
            cell = ws.cell(row=r, column=i)
            cell.font = bold
            cell.fill = header_fill
        for r in INPUT_ROWS:
            ws.cell(row=r, column=i).number_format = NUM_FMT
        for r, (_, formula) in FORMULA_ROWS.items():
            cell = ws.cell(row=r, column=i, value=formula.format(c=c))
            cell.font = black
            cell.number_format = NUM_FMT
            if r in (22, 25):
                cell.border = total_border
    for r in (1, 2, 3):
        ws.cell(row=r, column=1).fill = header_fill

    ws.freeze_panes = "B4"
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--out", default="templates/TradingComps.xlsx")
    ap.add_argument("--columns", type=int, default=10, help="Number of company columns to pre-format")
    args = ap.parse_args()
    build(args.out, args.columns)
    print(f"Wrote {args.out}")


if __name__ == "__main__":
    main()
