#!/usr/bin/env python3
"""
Build the FREE "Simple Income & Expense Tracker" lead-magnet workbook.

This is a deliberately simple, deliberately smaller SUBSET of the paid
Freelancer Income & Expense Tracker (dist/Freelancer-Income-Expense-Tracker.xlsx):
3 sheets instead of 14, one flat Transactions log instead of 12 monthly tabs,
8 fixed generic categories instead of user-editable 6+12 category lists, and
one Summary sheet instead of a full category-breakdown Dashboard with charts.
It exists to be downloaded for free (no signup) and to earn links/Pinterest
saves, and to funnel interested users to the paid Freelancer Finance Toolkit
and the custom-spreadsheet service.

Run with:    .venv/bin/python build/free_tracker.py
Verify with: .venv/bin/python build/free_tracker.py --verify

Follows SPEC.md: design system, Google Sheets compatibility rules, and the
shared Start Here sheet convention (via build/common.py).
"""
import re
import sys
from datetime import date
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment
from openpyxl.worksheet.datavalidation import DataValidation

sys.path.insert(0, str(Path(__file__).resolve().parent))
import common  # noqa: E402

ROOT = Path(__file__).resolve().parent.parent
OUT_PATH = ROOT / "dist" / "free" / "GrowthChief-Simple-Income-Expense-Tracker.xlsx"

PRODUCT_NAME = "Simple Income & Expense Tracker (Free)"

# ---------------------------------------------------------------------------
# Fixed vocab for the free version (deliberately not user-editable, unlike
# the paid product's editable 6+12 category lists on Start Here)
# ---------------------------------------------------------------------------
TYPES = ["Income", "Expense"]
CATEGORIES = [
    "Client Work", "Product Sales", "Software & Tools", "Marketing",
    "Equipment", "Travel", "Fees & Charges", "Other",
]

TRANS_FIRST_DATA_ROW = 5
TRANS_SAMPLE_ROWS = 5
TRANS_TOTAL_ROWS = 200
TRANS_LAST_DATA_ROW = TRANS_FIRST_DATA_ROW + TRANS_TOTAL_ROWS - 1  # 204
TRANS_HEADER_ROW = 4

# 5 sample rows, dated in the current real month so the Summary sheet's
# "this month" SUMIFS tiles are non-zero when opened/recalculated today,
# not just the all-time tiles.
SAMPLE_ROWS = [
    (date(2026, 9, 1), "Income", "Client Work", "Website project - deposit (example, replace me)", 500.0),
    (date(2026, 9, 3), "Expense", "Software & Tools", "Design software subscription (example, replace me)", 25.0),
    (date(2026, 9, 4), "Income", "Product Sales", "Digital template sale (example, replace me)", 45.5),
    (date(2026, 9, 5), "Expense", "Marketing", "Etsy ads (example, replace me)", 15.0),
    (date(2026, 9, 6), "Expense", "Fees & Charges", "Payment processor fees (example, replace me)", 8.25),
]

ETSY_TOOLKIT_URL = "https://growthchief.etsy.com/listing/4560433852"
ETSY_CUSTOM_URL = "https://growthchief.etsy.com/listing/4564738495"

link_font = Font(name=common.FONT_NAME, size=11, bold=True, color=common.ACCENT_COLOR, underline="single")
section_font = Font(name=common.FONT_NAME, size=13, bold=True, color=common.HEADER_FILL_COLOR)


# ---------------------------------------------------------------------------
# Start Here
# ---------------------------------------------------------------------------

def build_start_here_sheet(wb):
    steps = [
        "Set your currency symbol in Settings below if you don't use $ -- every sheet updates automatically.",
        "Go to the Transactions tab and log every payment and expense as it happens: pick Income or Expense, choose a category, and enter the amount.",
        "Delete the 5 example rows at the top of Transactions once you're comfortable, then keep logging your own.",
        "Check the Summary tab any time for your totals and this month's numbers -- it updates itself, no manual math.",
        "Works in Excel and Google Sheets as-is -- no macros, no add-ons, nothing to enable.",
    ]
    settings = [
        {
            "label": "Currency symbol",
            "value": "$",
            "name": "CurrencySymbol",
            "note": "Used in headers throughout this workbook. Doesn't convert currencies -- just changes the symbol shown.",
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
            "This free tracker covers the basics: log transactions, see your totals. "
            "The full Freelancer Finance Toolkit builds on top of it with:"
        ),
    )
    intro.font = common.body_font
    intro.alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=4)
    ws.row_dimensions[row].height = 32
    row += 1

    bullets = [
        "A category dashboard that breaks income and expenses down automatically, month by month and by category.",
        "A tax set-aside calculator that works out how much to move into savings from every payment you receive.",
        "A rate calculator that works out the hourly and day rate you need to charge to hit a take-home income goal.",
        "A client & invoice tracker with auto-numbered invoices, due dates, a status dropdown, and a dashboard of what's outstanding and overdue.",
    ]
    for b in bullets:
        c = ws.cell(row=row, column=2, value=f"-  {b}")
        c.font = common.body_font
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=4)
        ws.row_dimensions[row].height = 28
        row += 1

    row += 1
    link1 = ws.cell(row=row, column=2, value="See the full Freelancer Finance Toolkit on Etsy")
    link1.font = link_font
    link1.hyperlink = ETSY_TOOLKIT_URL
    row += 1
    link2 = ws.cell(row=row, column=2, value="Need something custom built for your business? Custom spreadsheets, made to order")
    link2.font = link_font
    link2.hyperlink = ETSY_CUSTOM_URL
    ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=4)
    row += 2

    return ws


# ---------------------------------------------------------------------------
# Transactions
# ---------------------------------------------------------------------------

def build_transactions_sheet(wb):
    ws = common.new_data_sheet(wb, "Transactions", tab_color=common.TAB_SLATE)
    common.set_col_widths(ws, {"A": 13, "B": 12, "C": 20, "D": 38, "E": 14})

    common.style_title(ws, "A1", "Transactions")
    sub = ws.cell(row=2, column=1, value=(
        "Log every payment and expense here. Pick Income or Expense, choose a category, "
        "and enter the amount -- the Summary tab totals it automatically."
    ))
    sub.font = common.subtitle_font
    sub.alignment = Alignment(wrap_text=True)
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=5)
    ws.row_dimensions[2].height = 18

    common.style_header_row(ws, TRANS_HEADER_ROW, 1, 5)
    headers = ["Date", "Type", "Category", "Description", '="Amount (" & CurrencySymbol & ")"']
    for col, val in enumerate(headers, start=1):
        ws.cell(row=TRANS_HEADER_ROW, column=col, value=val)

    # Sample rows
    for i, (d, t, cat, desc, amt) in enumerate(SAMPLE_ROWS):
        r = TRANS_FIRST_DATA_ROW + i
        ws.cell(row=r, column=1, value=d)
        ws.cell(row=r, column=2, value=t)
        ws.cell(row=r, column=3, value=cat)
        ws.cell(row=r, column=4, value=desc)
        ws.cell(row=r, column=5, value=amt)

    # Row banding + borders across the whole 200-row pre-formatted block
    common.band_rows(ws, TRANS_FIRST_DATA_ROW, TRANS_LAST_DATA_ROW, 1, 5)

    # Number formats across the whole pre-formatted block
    for r in range(TRANS_FIRST_DATA_ROW, TRANS_LAST_DATA_ROW + 1):
        ws.cell(row=r, column=1).number_format = common.DATE_FMT
        ws.cell(row=r, column=5).number_format = common.NUMBER_FMT_NEG_RED

    data_rows_range = f"A{TRANS_FIRST_DATA_ROW}:E{TRANS_LAST_DATA_ROW}"
    type_range = f"B{TRANS_FIRST_DATA_ROW}:B{TRANS_LAST_DATA_ROW}"
    category_range = f"C{TRANS_FIRST_DATA_ROW}:C{TRANS_LAST_DATA_ROW}"
    amount_range = f"E{TRANS_FIRST_DATA_ROW}:E{TRANS_LAST_DATA_ROW}"

    common.add_list_validation(ws, type_range, TYPES)
    common.add_list_validation(ws, category_range, CATEGORIES)
    common.apply_negative_red(ws, amount_range)

    common.freeze_header(ws, f"A{TRANS_FIRST_DATA_ROW}")
    ws.sheet_view.zoomScale = 100
    return ws


# ---------------------------------------------------------------------------
# Summary (tile styling per common.py)
# ---------------------------------------------------------------------------

TILE_VALUE_FONT = Font(name=common.FONT_NAME, size=20, bold=True, color=common.HEADER_FILL_COLOR)
TILE_LABEL_FONT = Font(name=common.FONT_NAME, size=11, bold=True, color="6B7280")


def _tile(ws, label_row, value_row, col, label, formula):
    lc = ws.cell(row=label_row, column=col, value=label)
    lc.font = TILE_LABEL_FONT
    lc.alignment = Alignment(horizontal="left")

    vc = ws.cell(row=value_row, column=col, value=formula)
    vc.font = TILE_VALUE_FONT
    vc.fill = common.accent_light_fill
    vc.border = common.thin_border
    vc.number_format = common.NUMBER_FMT_NEG_RED
    vc.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[value_row].height = 30
    return vc


def build_summary_sheet(wb):
    ws = common.new_data_sheet(wb, "Summary", tab_color=common.TAB_TEAL, gridlines=False)
    common.set_col_widths(ws, {"A": 26, "B": 26, "C": 26, "D": 4})

    common.style_title(ws, "A1", "Summary")
    sub = ws.cell(row=2, column=1, value=(
        "Totals below update automatically as you add rows on the Transactions tab. "
        "No manual math required."
    ))
    sub.font = common.subtitle_font
    sub.alignment = Alignment(wrap_text=True)
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=3)

    trans_amount = f"Transactions!$E${TRANS_FIRST_DATA_ROW}:$E${TRANS_LAST_DATA_ROW}"
    trans_type = f"Transactions!$B${TRANS_FIRST_DATA_ROW}:$B${TRANS_LAST_DATA_ROW}"
    trans_date = f"Transactions!$A${TRANS_FIRST_DATA_ROW}:$A${TRANS_LAST_DATA_ROW}"

    # --- All-time totals ---
    ws.cell(row=4, column=1, value="All-Time Totals").font = section_font

    income_formula = f'=SUMIF({trans_type},"Income",{trans_amount})'
    expense_formula = f'=SUMIF({trans_type},"Expense",{trans_amount})'
    net_formula = "=A6-B6"

    _tile(ws, 5, 6, 1, '="Total Income (" & CurrencySymbol & ")"', income_formula)
    _tile(ws, 5, 6, 2, '="Total Expenses (" & CurrencySymbol & ")"', expense_formula)
    net_cell = _tile(ws, 5, 6, 3, '="Net (" & CurrencySymbol & ")"', net_formula)

    # --- This month ---
    ws.cell(row=8, column=1,
            value='="This Month (" & TEXT(TODAY(),"mmmm yyyy") & ")"').font = section_font

    month_start = "(EOMONTH(TODAY(),-1)+1)"
    month_end_exclusive = "(EOMONTH(TODAY(),0)+1)"

    income_month_formula = (
        f'=SUMIFS({trans_amount},{trans_type},"Income",'
        f'{trans_date},">="&{month_start},{trans_date},"<"&{month_end_exclusive})'
    )
    expense_month_formula = (
        f'=SUMIFS({trans_amount},{trans_type},"Expense",'
        f'{trans_date},">="&{month_start},{trans_date},"<"&{month_end_exclusive})'
    )
    net_month_formula = "=A10-B10"

    _tile(ws, 9, 10, 1, '="Income This Month (" & CurrencySymbol & ")"', income_month_formula)
    _tile(ws, 9, 10, 2, '="Expenses This Month (" & CurrencySymbol & ")"', expense_month_formula)
    net_month_cell = _tile(ws, 9, 10, 3, '="Net This Month (" & CurrencySymbol & ")"', net_month_formula)

    common.apply_negative_red(ws, "A6:C6")
    common.apply_negative_red(ws, "A10:C10")

    note = ws.cell(row=12, column=1, value=(
        "Want category breakdowns, a tax set-aside calculator, a rate calculator, or client & "
        "invoice tracking? See the Start Here tab for what the full Freelancer Finance Toolkit adds."
    ))
    note.font = common.note_font
    note.alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=12, start_column=1, end_row=13, end_column=3)

    ws.sheet_view.zoomScale = 100
    return ws


# ---------------------------------------------------------------------------
# Build
# ---------------------------------------------------------------------------

def build():
    wb = Workbook()
    build_start_here_sheet(wb)
    build_transactions_sheet(wb)
    build_summary_sheet(wb)

    assert wb.sheetnames == ["Start Here", "Transactions", "Summary"], wb.sheetnames

    OUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    wb.save(OUT_PATH)
    print(f"Wrote {OUT_PATH} ({OUT_PATH.stat().st_size:,} bytes)")
    return OUT_PATH


# ---------------------------------------------------------------------------
# Verify
# ---------------------------------------------------------------------------

# The task brief's own whitelist for this workbook (stricter/narrower than
# common.py's workbook-wide ALLOWED_FUNCTIONS, which also permits things
# like RANK/VLOOKUP used by other products). Checked in addition to
# common.scan_workbook_formulas so an accidental use of, say, AVERAGEIF
# (fine elsewhere in the toolkit, not asked for here) would still be caught.
TASK_ALLOWED_FUNCTIONS = {
    "SUM", "SUMIF", "SUMIFS", "COUNTIF", "IF", "TODAY", "DATE", "YEAR",
    "MONTH", "EOMONTH", "TEXT", "IFERROR",
}

_RANGE_RE = re.compile(r"(?<![!\w$'])(\$?[A-Z]{1,3}\$?\d+:\$?[A-Z]{1,3}\$?\d+)")


def _strip_strings(formula):
    return re.sub(r'"[^"]*"', "", formula)


def verify():
    assert OUT_PATH.exists(), f"{OUT_PATH} does not exist -- run without --verify first"
    wb = load_workbook(OUT_PATH)

    # 1. Sheet names, exactly 3, in order.
    assert wb.sheetnames == ["Start Here", "Transactions", "Summary"], (
        f"unexpected sheet list: {wb.sheetnames}"
    )
    print("PASS: sheet names ==", wb.sheetnames)

    # 2. Formulas present on Transactions and Summary.
    trans = wb["Transactions"]
    summary = wb["Summary"]

    trans_formula_cells = [
        c for row in trans.iter_rows() for c in row
        if isinstance(c.value, str) and c.value.startswith("=")
    ]
    assert trans_formula_cells, "Transactions has no formulas (expected the Amount header formula)"
    print(f"PASS: Transactions has {len(trans_formula_cells)} formula cell(s)")

    summary_formula_cells = [
        c for row in summary.iter_rows() for c in row
        if isinstance(c.value, str) and c.value.startswith("=")
    ]
    assert len(summary_formula_cells) >= 8, (
        f"expected >=8 formulas on Summary (6 tiles + 2 headers), found {len(summary_formula_cells)}"
    )
    print(f"PASS: Summary has {len(summary_formula_cells)} formula cell(s)")

    required_tile_cells = ["A6", "B6", "C6", "A10", "B10", "C10"]
    for ref in required_tile_cells:
        v = summary[ref].value
        assert isinstance(v, str) and v.startswith("="), f"Summary!{ref} missing formula, got {v!r}"
    print("PASS: all 6 Summary tile formulas present:", required_tile_cells)

    # 3. Data validations present on Transactions (Type + Category dropdowns).
    dvs = trans.data_validations.dataValidation
    assert len(dvs) == 2, f"expected 2 data validations on Transactions, found {len(dvs)}"
    sqrefs = {str(dv.sqref) for dv in dvs}
    assert any("B5:B204" in s for s in sqrefs), f"Type dropdown range missing, sqrefs={sqrefs}"
    assert any("C5:C204" in s for s in sqrefs), f"Category dropdown range missing, sqrefs={sqrefs}"
    for dv in dvs:
        assert dv.type == "list", f"unexpected validation type {dv.type!r}"
    print("PASS: 2 list data validations present on Transactions:", sqrefs)

    # 4. No unprefixed cross-sheet ranges: every multi-cell range referenced
    #    in a Summary-sheet formula must be preceded by a sheet name + "!".
    unprefixed = []
    for row in summary.iter_rows():
        for c in row:
            v = c.value
            if not (isinstance(v, str) and v.startswith("=")):
                continue
            body = _strip_strings(v)
            for m in re.finditer(r"([A-Za-z0-9_']*!)?(\$?[A-Z]{1,3}\$?\d+:\$?[A-Z]{1,3}\$?\d+)", body):
                prefix, rng = m.groups()
                if not prefix:
                    unprefixed.append((c.coordinate, rng, v))
    assert not unprefixed, f"unprefixed cross-sheet range(s) found on Summary: {unprefixed}"
    print("PASS: no unprefixed multi-cell ranges on Summary")

    # 5. Whitelisted functions only -- both common.py's broad list and the
    #    task brief's narrower explicit list.
    func_counter, disallowed_common, placeholder_hits = common.scan_workbook_formulas(wb)
    assert not disallowed_common, f"functions outside common.py's ALLOWED_FUNCTIONS: {disallowed_common}"
    disallowed_task = set(func_counter) - TASK_ALLOWED_FUNCTIONS
    assert not disallowed_task, f"functions outside the task brief's whitelist: {disallowed_task}"
    assert not placeholder_hits, f"leftover [placeholder] text found: {placeholder_hits}"
    print("PASS: functions used are all whitelisted:", sorted(func_counter))
    print("PASS: no leftover [placeholder] text")

    # 6. Sanity: 200 pre-formatted data rows, 5 of them sample rows.
    filled = sum(
        1 for r in range(TRANS_FIRST_DATA_ROW, TRANS_LAST_DATA_ROW + 1)
        if trans.cell(row=r, column=1).value is not None
    )
    assert filled == TRANS_SAMPLE_ROWS, f"expected {TRANS_SAMPLE_ROWS} filled sample rows, found {filled}"
    total_data_rows = TRANS_LAST_DATA_ROW - TRANS_FIRST_DATA_ROW + 1
    assert total_data_rows == TRANS_TOTAL_ROWS, total_data_rows
    print(f"PASS: {TRANS_TOTAL_ROWS} pre-formatted Transactions rows, {filled} sample rows filled")

    print("\nAll --verify checks passed.")


if __name__ == "__main__":
    if "--verify" in sys.argv:
        verify()
    else:
        build()
