#!/usr/bin/env python3
"""
Build the FREE "Simple Mileage Log" lead-magnet workbook.

Deliberately smaller SUBSET of the paid Real Estate Agent Commission &
Expense Tracker's Mileage & Net Profit Tracker workbook
(dist/p7/Mileage-Net-Profit-Tracker.xlsx): one flat log instead of a log +
Dashboard-with-chart, a simple Purpose dropdown instead of Purpose/Client
free text, no Start/End Location split (just From/To), and -- by design --
no odometer start/end reading columns at all. It exists to be downloaded for
free (no signup) and to earn links/Pinterest saves, and to funnel interested
users to the paid Real Estate Agent Commission & Expense Tracker (for
agents) or the Freelancer Finance Toolkit (for everyone else who needs
broader income/expense tracking, not just mileage).

Run with:    .venv/bin/python build/free_mileage.py
Verify with: .venv/bin/python build/free_mileage.py --verify

Follows SPEC.md: design system, Google Sheets compatibility rules, and the
shared Start Here sheet convention (via build/common.py).
"""
import re
import sys
from datetime import date
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

sys.path.insert(0, str(Path(__file__).resolve().parent))
import common  # noqa: E402

ROOT = Path(__file__).resolve().parent.parent
OUT_PATH = ROOT / "dist" / "free" / "GrowthChief-Simple-Mileage-Log.xlsx"

PRODUCT_NAME = "Simple Mileage Log (Free)"

ETSY_P7_URL = "https://growthchief.etsy.com/listing/4562585345"
ETSY_TOOLKIT_URL = "https://growthchief.etsy.com/listing/4560433852"
ETSY_CUSTOM_URL = "https://growthchief.etsy.com/listing/4564738495"

link_font = Font(name=common.FONT_NAME, size=11, bold=True, color=common.ACCENT_COLOR, underline="single")
section_font = Font(name=common.FONT_NAME, size=13, bold=True, color=common.HEADER_FILL_COLOR)

RATE_NUMBER_FMT = "#,##0.000"  # 3dp: a per-mile/km rate can be fractions of a cent
DISTANCE_FMT = "#,##0.0"

PURPOSES = ["Client visit", "Showing", "Meeting", "Errand", "Other"]

ML_HEADER_ROW = 4
ML_FIRST_ROW = 5
ML_TOTAL_ROWS = 200
ML_LAST_ROW = ML_FIRST_ROW + ML_TOTAL_ROWS - 1  # 204
PURPOSE_LIST_COL = "H"  # hidden-in-plain-sight editable Purpose list

# 6 sample rows: Date, Purpose, From, To, Distance
SAMPLE_ROWS = [
    (date(2026, 9, 1), "Client visit", "Home Office (example, replace me)", "142 Oak Street (example, replace me)", 8.4),
    (date(2026, 9, 2), "Showing", "Home Office (example, replace me)", "27 Birch Avenue (example, replace me)", 11.2),
    (date(2026, 9, 2), "Showing", "27 Birch Avenue (example, replace me)", "9 Maple Court (example, replace me)", 4.7),
    (date(2026, 9, 3), "Meeting", "Home Office (example, replace me)", "Downtown Title Co. (example, replace me)", 15.9),
    (date(2026, 9, 4), "Errand", "Home Office (example, replace me)", "Office Supply Store (example, replace me)", 3.1),
    (date(2026, 9, 5), "Other", "Home Office (example, replace me)", "Client Coffee Meeting (example, replace me)", 6.5),
]

RATE_SAMPLE = 0.67
DISTANCE_UNIT_SAMPLE = "mi"


# ---------------------------------------------------------------------------
# Start Here
# ---------------------------------------------------------------------------

def build_start_here_sheet(wb):
    steps = [
        "Set your currency symbol, your rate per mile/km, and your distance unit (mi or km) in Settings below -- every sheet updates automatically.",
        "Go to the Mileage Log tab and log every business drive as it happens: pick a purpose from the dropdown, then fill in the date, From, To, and Distance.",
        "Reimbursement calculates itself for every row -- Distance x your rate per mile/km. There's nothing to add up by hand.",
        "Check the Summary tab any time for your total distance and reimbursement, this month's numbers, and a breakdown by purpose.",
        "By design, there are no odometer start/end columns -- just From, To, and the distance you drove. Keeps logging fast; add your own odometer columns if your situation needs them.",
        "Works in Excel and Google Sheets as-is -- no macros, no add-ons, nothing to enable.",
    ]
    settings = [
        {
            "label": "Currency symbol",
            "value": "$",
            "name": "CurrencySymbol",
            "note": "Used in headers throughout this workbook. Doesn't convert currencies -- just changes the symbol shown.",
        },
        {
            "label": "Rate per mile/km",
            "value": RATE_SAMPLE,
            "name": "RatePerUnit",
            "number_format": RATE_NUMBER_FMT,
            "note": "Multiplies every row's Distance into a Reimbursement amount. This workbook doesn't know or set the correct rate for your country, employer, or tax year -- look up the rate that applies to you and enter it here.",
        },
        {
            "label": "Distance unit",
            "value": DISTANCE_UNIT_SAMPLE,
            "name": "DistanceUnit",
            "note": "Text only (e.g. \"mi\" or \"km\") -- shown in the Mileage Log and Summary labels. Doesn't convert between units; just log every row in the same one.",
        },
    ]
    result = common.build_start_here(
        wb,
        PRODUCT_NAME,
        steps,
        settings,
        guide_link_placeholder=(
            "This is the free version -- it's simple by design, so there's no separate guide. "
            "If anything's unclear, the download page has a full walkthrough."
        ),
    )
    ws = result["ws"]
    row = result["settings_range"][1] + 2

    header = ws.cell(row=row, column=2, value="Want more?")
    header.font = section_font
    row += 1

    intro = ws.cell(
        row=row, column=2,
        value=(
            "This free log covers one flat mileage log with a fixed 5-option purpose list and "
            "no dashboard. The full Real Estate Agent Commission & Expense Tracker builds on top "
            "of it with:"
        ),
    )
    intro.font = common.body_font
    intro.alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=4)
    ws.row_dimensions[row].height = 32
    row += 1

    bullets = [
        "A Mileage & Net Profit Tracker workbook with its own Dashboard: live tiles for Total Miles/KM Logged and Total Mileage Deduction, plus a Net Profit Summary chart that combines your mileage deduction with totals you bring in from the other two workbooks.",
        "Two sibling workbooks in the same bundle: a Commission Log with Gross/Net Commission and brokerage-split tracking, and a Business Expense Tracker -- so mileage, commissions, and expenses live together instead of in one sheet.",
        "Start/End Location columns and a 150-row log built specifically around showings, closings, and client meetings, rather than this free version's generic 5-option purpose list.",
    ]
    for b in bullets:
        c = ws.cell(row=row, column=2, value=f"-  {b}")
        c.font = common.body_font
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=4)
        ws.row_dimensions[row].height = 44
        row += 1

    row += 1
    link1 = ws.cell(row=row, column=2, value="See the full Real Estate Agent Commission & Expense Tracker on Etsy")
    link1.font = link_font
    link1.hyperlink = ETSY_P7_URL
    ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=4)
    row += 1

    not_agent = ws.cell(
        row=row, column=2,
        value=(
            "Not a real estate agent? If you need broader income and expense tracking rather "
            "than just mileage, the Freelancer Finance Toolkit covers invoicing, a 12-month "
            "income & expense tracker, a rate calculator, and a tax set-aside calculator."
        ),
    )
    not_agent.font = common.body_font
    not_agent.alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=4)
    ws.row_dimensions[row].height = 32
    row += 1

    link2 = ws.cell(row=row, column=2, value="See the full Freelancer Finance Toolkit on Etsy")
    link2.font = link_font
    link2.hyperlink = ETSY_TOOLKIT_URL
    ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=4)
    row += 1

    link3 = ws.cell(row=row, column=2, value="Need something custom built for your business? Custom spreadsheets, made to order")
    link3.font = link_font
    link3.hyperlink = ETSY_CUSTOM_URL
    ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=4)
    row += 2

    return ws


# ---------------------------------------------------------------------------
# Mileage Log
# ---------------------------------------------------------------------------

def build_mileage_log_sheet(wb):
    ws = common.new_data_sheet(wb, "Mileage Log", tab_color=common.TAB_TEAL, gridlines=False)
    common.set_col_widths(ws, {"A": 13, "B": 16, "C": 30, "D": 30, "E": 14, "F": 16, "G": 2, "H": 16})

    common.style_title(ws, "A1", "Mileage Log")
    sub = ws.cell(row=2, column=1, value=(
        "Log every business drive here. Pick a purpose from the dropdown -- Reimbursement "
        "calculates itself from your Distance and the rate you set on Start Here."
    ))
    sub.font = common.subtitle_font
    sub.alignment = Alignment(wrap_text=True)
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=6)
    ws.row_dimensions[2].height = 18

    common.style_header_row(ws, ML_HEADER_ROW, 1, 6)
    headers = [
        "Date",
        "Purpose",
        "From",
        "To",
        '="Distance (" & DistanceUnit & ")"',
        '="Reimbursement (" & CurrencySymbol & ")"',
    ]
    for col, val in enumerate(headers, start=1):
        ws.cell(row=ML_HEADER_ROW, column=col, value=val)

    for r in range(ML_FIRST_ROW, ML_LAST_ROW + 1):
        ws.cell(row=r, column=1).number_format = common.DATE_FMT
        ws.cell(row=r, column=5).number_format = DISTANCE_FMT
        ws.cell(row=r, column=6,
                value=f'=IF(E{r}="","",E{r}*RatePerUnit)').number_format = common.NUMBER_FMT

    for i, (d, purpose, frm, to, dist) in enumerate(SAMPLE_ROWS):
        r = ML_FIRST_ROW + i
        ws.cell(row=r, column=1, value=d)
        ws.cell(row=r, column=2, value=purpose)
        ws.cell(row=r, column=3, value=frm)
        ws.cell(row=r, column=4, value=to)
        ws.cell(row=r, column=5, value=dist)

    common.band_rows(ws, ML_FIRST_ROW, ML_LAST_ROW, 1, 6)
    for r in range(ML_FIRST_ROW, ML_LAST_ROW + 1):
        ws.cell(row=r, column=1).number_format = common.DATE_FMT
        ws.cell(row=r, column=5).number_format = DISTANCE_FMT
        ws.cell(row=r, column=6).number_format = common.NUMBER_FMT

    # Hidden-in-plain-sight, EDITABLE Purpose reference list -- source for
    # the Purpose dropdown, kept as a real contiguous range (not an inline
    # comma list) so a buyer can rename/add/remove purposes by editing these
    # cells directly and the dropdown just follows.
    ws.cell(row=1, column=8, value="Purposes (editable)").font = common.note_font
    for i, name in enumerate(PURPOSES):
        ws.cell(row=2 + i, column=8, value=name).font = common.note_font
    purpose_first = 2
    purpose_last = 2 + len(PURPOSES) - 1
    wb.defined_names["MileagePurposes"] = DefinedName(
        "MileagePurposes", attr_text=f"'Mileage Log'!${PURPOSE_LIST_COL}${purpose_first}:${PURPOSE_LIST_COL}${purpose_last}"
    )

    purpose_range = f"B{ML_FIRST_ROW}:B{ML_LAST_ROW}"
    dv = DataValidation(type="list", formula1="MileagePurposes", allow_blank=True,
                         showErrorMessage=True, showDropDown=False)
    dv.error = "Choose a purpose from the list, or edit the list in column H to add your own."
    dv.errorTitle = "Invalid purpose"
    ws.add_data_validation(dv)
    dv.add(purpose_range)

    common.freeze_header(ws, f"A{ML_FIRST_ROW}")
    ws.sheet_view.zoomScale = 100
    return ws


# ---------------------------------------------------------------------------
# Summary
# ---------------------------------------------------------------------------

def build_summary_sheet(wb):
    ws = common.new_data_sheet(wb, "Summary", tab_color=common.TAB_SLATE, gridlines=False)
    common.set_col_widths(ws, {"A": 30, "B": 20, "C": 3, "D": 30, "E": 20, "F": 20})

    common.style_title(ws, "A1", "Summary")
    sub = ws.cell(row=2, column=1, value="Auto-updates as you log drives on the Mileage Log tab.")
    sub.font = common.subtitle_font
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=6)
    ws.row_dimensions[2].height = 18

    date_range = f"'Mileage Log'!$A${ML_FIRST_ROW}:$A${ML_LAST_ROW}"
    dist_range = f"'Mileage Log'!$E${ML_FIRST_ROW}:$E${ML_LAST_ROW}"
    reimb_range = f"'Mileage Log'!$F${ML_FIRST_ROW}:$F${ML_LAST_ROW}"
    purpose_range = f"'Mileage Log'!$B${ML_FIRST_ROW}:$B${ML_LAST_ROW}"

    # This-month bounds: first-of-month through last-of-month, both derived
    # from EOMONTH(TODAY(),...) so the tiles always reflect the month the
    # file is opened in, with no hardcoded date.
    month_start = "EOMONTH(TODAY(),-1)+1"
    month_end = "EOMONTH(TODAY(),0)"

    def tile(label_row, value_row, first_col, last_col, label, formula, number_format):
        ws.merge_cells(start_row=label_row, start_column=first_col, end_row=label_row, end_column=last_col)
        lbl = ws.cell(row=label_row, column=first_col, value=label)
        lbl.font = Font(name=common.FONT_NAME, size=10, bold=True, color="374151")
        lbl.alignment = Alignment(horizontal="left", vertical="center", indent=1, wrap_text=True)
        for c in range(first_col, last_col + 1):
            ws.cell(row=label_row, column=c).fill = common.accent_light_fill

        ws.merge_cells(start_row=value_row, start_column=first_col, end_row=value_row, end_column=last_col)
        val = ws.cell(row=value_row, column=first_col, value=formula)
        val.font = Font(name=common.FONT_NAME, size=18, bold=True, color=common.ACCENT_COLOR)
        val.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        val.number_format = number_format
        for c in range(first_col, last_col + 1):
            cell = ws.cell(row=value_row, column=c)
            cell.fill = common.accent_light_fill
            cell.border = common.thin_border
        ws.row_dimensions[label_row].height = 26
        ws.row_dimensions[value_row].height = 30
        return val

    tile(4, 5, 1, 2, '="Total Distance (" & DistanceUnit & ")"',
         f"=SUM({dist_range})", DISTANCE_FMT)
    tile(4, 5, 4, 5, '="Total Reimbursement (" & CurrencySymbol & ")"',
         f"=SUM({reimb_range})", common.NUMBER_FMT)
    tile(7, 8, 1, 2, '="This Month Distance (" & DistanceUnit & ")"',
         f'=SUMIFS({dist_range},{date_range},">="&{month_start},{date_range},"<="&{month_end})',
         DISTANCE_FMT)
    tile(7, 8, 4, 5, '="This Month Reimbursement (" & CurrencySymbol & ")"',
         f'=SUMIFS({reimb_range},{date_range},">="&{month_start},{date_range},"<="&{month_end})',
         common.NUMBER_FMT)

    # By purpose table
    BP_HEADER_ROW = 11
    BP_FIRST_ROW = 12
    header = ws.cell(row=10, column=1, value="By purpose")
    header.font = section_font
    ws.merge_cells(start_row=10, start_column=1, end_row=10, end_column=3)

    common.style_header_row(ws, BP_HEADER_ROW, 1, 3)
    ws.cell(row=BP_HEADER_ROW, column=1, value="Purpose")
    ws.cell(row=BP_HEADER_ROW, column=2, value='="Distance (" & DistanceUnit & ")"')
    ws.cell(row=BP_HEADER_ROW, column=3, value='="Reimbursement (" & CurrencySymbol & ")"')

    for i in range(len(PURPOSES)):
        r = BP_FIRST_ROW + i
        purpose_cell_ref = f"'Mileage Log'!${PURPOSE_LIST_COL}${2 + i}"
        ws.cell(row=r, column=1, value=f"={purpose_cell_ref}")
        ws.cell(row=r, column=2, value=f"=SUMIF({purpose_range},A{r},{dist_range})")
        ws.cell(row=r, column=3, value=f"=SUMIF({purpose_range},A{r},{reimb_range})")
        ws.cell(row=r, column=2).number_format = DISTANCE_FMT
        ws.cell(row=r, column=3).number_format = common.NUMBER_FMT

    bp_last_row = BP_FIRST_ROW + len(PURPOSES) - 1
    common.band_rows(ws, BP_FIRST_ROW, bp_last_row, 1, 3)

    ws.sheet_view.zoomScale = 100
    return ws


# ---------------------------------------------------------------------------
# Build
# ---------------------------------------------------------------------------

def build():
    wb = Workbook()
    build_start_here_sheet(wb)
    build_mileage_log_sheet(wb)
    build_summary_sheet(wb)

    assert wb.sheetnames == ["Start Here", "Mileage Log", "Summary"], wb.sheetnames

    OUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    wb.save(OUT_PATH)
    print(f"Wrote {OUT_PATH} ({OUT_PATH.stat().st_size:,} bytes)")
    return OUT_PATH


# ---------------------------------------------------------------------------
# Verify
# ---------------------------------------------------------------------------

TASK_ALLOWED_FUNCTIONS = {
    "SUM", "SUMIF", "SUMIFS", "COUNTIF", "IF", "IFERROR", "TODAY", "DATE",
    "TEXT", "EOMONTH", "ROUND",
}


def verify():
    assert OUT_PATH.exists(), f"{OUT_PATH} does not exist -- run without --verify first"
    wb = load_workbook(OUT_PATH)

    # 1. Sheet names, exactly 3, in order.
    assert wb.sheetnames == ["Start Here", "Mileage Log", "Summary"], (
        f"unexpected sheet list: {wb.sheetnames}"
    )
    print("PASS: sheet names ==", wb.sheetnames)

    log = wb["Mileage Log"]
    summary = wb["Summary"]

    # 2. Defined names present.
    for name in ("CurrencySymbol", "RatePerUnit", "DistanceUnit", "MileagePurposes"):
        assert name in wb.defined_names, f"{name} defined name missing"
    print("PASS: CurrencySymbol, RatePerUnit, DistanceUnit, MileagePurposes defined names present")

    # 3. Mileage Log: 200 rows, reimbursement formula on every row.
    assert ML_LAST_ROW - ML_FIRST_ROW + 1 == 200, "Mileage Log must have 200 data rows"
    for r in range(ML_FIRST_ROW, ML_LAST_ROW + 1):
        v = log.cell(row=r, column=6).value
        assert isinstance(v, str) and v.startswith("=") and "RatePerUnit" in v, (
            f"Mileage Log!F{r} missing Reimbursement formula, got {v!r}"
        )
    print(f"PASS: all {ML_LAST_ROW - ML_FIRST_ROW + 1} Reimbursement formulas present, wired to RatePerUnit")

    # 4. Purpose list: 5 options, matches PURPOSES, dropdown sourced from it.
    purpose_col_values = [log.cell(row=2 + i, column=8).value for i in range(len(PURPOSES))]
    assert purpose_col_values == PURPOSES, f"Mileage Log!H2:H6 purpose list doesn't match PURPOSES: {purpose_col_values}"
    dvs = log.data_validations.dataValidation
    assert len(dvs) == 1, f"expected 1 data validation on Mileage Log, found {len(dvs)}"
    dv = dvs[0]
    assert dv.type == "list", f"unexpected validation type {dv.type!r}"
    assert dv.formula1 == "MileagePurposes", f"expected formula1='MileagePurposes', got {dv.formula1!r}"
    sqref = str(dv.sqref)
    assert f"B{ML_FIRST_ROW}:B{ML_LAST_ROW}" in sqref, f"Purpose dropdown range wrong: {sqref}"
    print("PASS: 5-option editable Purpose list present, dropdown sourced from MileagePurposes:", sqref)

    # 5. Sample data: exactly 6 filled rows.
    filled_rows = sum(1 for r in range(ML_FIRST_ROW, ML_LAST_ROW + 1) if log.cell(row=r, column=1).value is not None)
    assert filled_rows == len(SAMPLE_ROWS) == 6, f"expected 6 filled sample rows, found {filled_rows}"
    print(f"PASS: {filled_rows} sample rows filled")

    # 6. No odometer start/end columns -- only 6 real data columns (Date,
    #    Purpose, From, To, Distance, Reimbursement) before the spacer/list
    #    columns G/H.
    header_vals = [log.cell(row=ML_HEADER_ROW, column=c).value for c in range(1, 7)]
    for h in header_vals:
        assert h is not None, "Mileage Log header row missing a column"
    assert not any("odometer" in str(h).lower() for h in header_vals), "Mileage Log must not have odometer columns"
    print("PASS: Mileage Log has exactly 6 data columns, no odometer start/end columns")

    # 7. All cross-sheet refs sheet-prefixed: every Summary formula that
    #    reaches into Mileage Log must carry a literal 'Mileage Log'! prefix.
    range_re = re.compile(r"(?<![!\w$'])(\$?[A-Z]{1,3}\$?\d+:\$?[A-Z]{1,3}\$?\d+)")
    for row in summary.iter_rows():
        for c in row:
            v = c.value
            if not (isinstance(v, str) and v.startswith("=")):
                continue
            body = re.sub(r'"[^"]*"', "", v)
            for rng in range_re.findall(body):
                # every raw range on Summary must be immediately preceded by 'Mileage Log'!
                idx = body.find(rng)
                preceding = body[:idx]
                assert preceding.rstrip().endswith("'Mileage Log'!") or preceding.rstrip().endswith("Log'!"), (
                    f"Unprefixed cross-sheet range on Summary!{c.coordinate}: {v}"
                )
    print("PASS: every Summary range formula is sheet-prefixed to 'Mileage Log' (no unprefixed cross-sheet refs)")

    # 8. Whitelisted functions only.
    func_counter, disallowed_common, placeholder_hits = common.scan_workbook_formulas(wb)
    assert not disallowed_common, f"functions outside common.py's ALLOWED_FUNCTIONS: {disallowed_common}"
    disallowed_task = set(func_counter) - TASK_ALLOWED_FUNCTIONS
    assert not disallowed_task, f"functions outside the task brief's whitelist: {disallowed_task}"
    assert not placeholder_hits, f"leftover [placeholder] text found: {placeholder_hits}"
    print("PASS: functions used are all whitelisted:", sorted(func_counter))
    print("PASS: no leftover [placeholder] text")

    # 9. Summary tiles and by-purpose table formulas present.
    for cell_ref in ("A5", "D5", "A8", "D8"):
        v = summary[cell_ref].value
        assert isinstance(v, str) and v.startswith("="), f"Summary!{cell_ref} missing tile formula, got {v!r}"
    for i in range(len(PURPOSES)):
        r = 12 + i
        for col in (2, 3):
            v = summary.cell(row=r, column=col).value
            assert isinstance(v, str) and v.startswith("=SUMIF("), f"Summary!{summary.cell(row=r, column=col).coordinate} missing SUMIF formula"
    print("PASS: all 4 Summary tiles and 5-row by-purpose SUMIF table present")

    # 10. Settings sample values.
    rate_attr = wb.defined_names["RatePerUnit"].attr_text
    rate_sheet, rate_cell = rate_attr.split("!")
    rate_val = wb[rate_sheet.strip("'")][rate_cell.replace("$", "")].value
    assert rate_val == RATE_SAMPLE, f"RatePerUnit sample value wrong: {rate_val}"
    unit_attr = wb.defined_names["DistanceUnit"].attr_text
    unit_sheet, unit_cell = unit_attr.split("!")
    unit_val = wb[unit_sheet.strip("'")][unit_cell.replace("$", "")].value
    assert unit_val == DISTANCE_UNIT_SAMPLE, f"DistanceUnit sample value wrong: {unit_val}"
    print(f"PASS: RatePerUnit sample filled ({rate_val}), DistanceUnit sample filled ({unit_val!r})")

    print("\nAll --verify checks passed.")


if __name__ == "__main__":
    if "--verify" in sys.argv:
        verify()
    else:
        build()
