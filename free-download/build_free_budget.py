#!/usr/bin/env python3
"""
Build the FREE "Simple Monthly Budget Template" lead-magnet workbook.

Deliberately smaller SUBSET of the paid Personal Budget & Debt Payoff Planner
(dist/p4/Monthly-Budget-Tracker.xlsx + dist/p4/Debt-Payoff-Planner.xlsx):
one month instead of 12 monthly tabs, 20 fixed categories instead of
user-editable 6 income / 16 expense category lists, no Dashboard/charts, and
no debt-payoff planner at all. It exists to be downloaded for free (no
signup) and to earn links/Pinterest saves, and to funnel interested users to
the paid Personal Budget & Debt Payoff Planner.

Run with:    .venv/bin/python build/free_budget.py
Verify with: .venv/bin/python build/free_budget.py --verify

Follows SPEC.md: design system, Google Sheets compatibility rules, and the
shared Start Here sheet convention (via build/common.py).
"""
import re
import sys
from datetime import date
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule

sys.path.insert(0, str(Path(__file__).resolve().parent))
import common  # noqa: E402

ROOT = Path(__file__).resolve().parent.parent
OUT_PATH = ROOT / "dist" / "free" / "GrowthChief-Simple-Monthly-Budget.xlsx"

PRODUCT_NAME = "Simple Monthly Budget Template (Free)"

ETSY_P4_URL = "https://growthchief.etsy.com/listing/4560856884"
ETSY_CUSTOM_URL = "https://growthchief.etsy.com/listing/4564738495"

link_font = Font(name=common.FONT_NAME, size=11, bold=True, color=common.ACCENT_COLOR, underline="single")
section_font = Font(name=common.FONT_NAME, size=13, bold=True, color=common.HEADER_FILL_COLOR)

# ---------------------------------------------------------------------------
# Categories, grouped Housing / Utilities / Food / Transport / Debt /
# Savings / Personal. Each tuple is (category name, sample Planned amount).
# ---------------------------------------------------------------------------
GROUPS = [
    ("Housing", [
        ("Rent/Mortgage", 1200.0),
        ("Home Insurance", 40.0),
        ("Home Maintenance", 50.0),
    ]),
    ("Utilities", [
        ("Electricity & Gas", 90.0),
        ("Water", 30.0),
        ("Internet & Phone", 70.0),
    ]),
    ("Food", [
        ("Groceries", 350.0),
        ("Dining Out", 100.0),
    ]),
    ("Transport", [
        ("Car Payment", 280.0),
        ("Fuel & Transit", 120.0),
        ("Car Insurance", 90.0),
    ]),
    ("Debt", [
        ("Credit Card Payment", 150.0),
        ("Student Loan / Other Debt", 200.0),
    ]),
    ("Savings", [
        ("Emergency Fund", 150.0),
        ("Retirement", 200.0),
        ("Other Savings", 50.0),
    ]),
    ("Personal", [
        ("Entertainment", 60.0),
        ("Subscriptions", 35.0),
        ("Personal Care", 40.0),
        ("Miscellaneous", 60.0),
    ]),
]
CATEGORIES = [name for _, items in GROUPS for name, _ in items]
assert len(CATEGORIES) == 20, len(CATEGORIES)

BUDGET_HEADER_ROW = 6
BUDGET_FIRST_LAYOUT_ROW = 7
CAT_LIST_COL = "G"  # hidden-in-plain-sight reference list, source of the Spending Log dropdown

SL_HEADER_ROW = 4
SL_FIRST_DATA_ROW = 5
SL_TOTAL_ROWS = 150
SL_LAST_DATA_ROW = SL_FIRST_DATA_ROW + SL_TOTAL_ROWS - 1  # 154

# 8 sample rows, dated in the current real month. "Subscriptions" (planned
# 35) deliberately gets 42 of actual spend so the Budget sheet's red
# over-budget conditional formatting has something to show out of the box.
SAMPLE_LOG_ROWS = [
    (date(2026, 9, 1), "Rent/Mortgage", "September rent (example, replace me)", 1200.0),
    (date(2026, 9, 2), "Groceries", "Weekly grocery shop (example, replace me)", 95.0),
    (date(2026, 9, 3), "Groceries", "Grocery top-up (example, replace me)", 60.0),
    (date(2026, 9, 4), "Dining Out", "Dinner out (example, replace me)", 45.0),
    (date(2026, 9, 5), "Fuel & Transit", "Gas fill-up (example, replace me)", 55.0),
    (date(2026, 9, 5), "Entertainment", "Movie night (example, replace me)", 30.0),
    (date(2026, 9, 6), "Subscriptions", "Streaming services (example, replace me)", 42.0),
    (date(2026, 9, 6), "Dining Out", "Coffee shop (example, replace me)", 20.0),
]

MONTHLY_INCOME_SAMPLE = 4000.0


# ---------------------------------------------------------------------------
# Start Here
# ---------------------------------------------------------------------------

def build_start_here_sheet(wb):
    steps = [
        "Set your currency symbol and your monthly take-home income in Settings below -- every sheet updates automatically.",
        "Go to the Budget tab and adjust the Planned column for each category to match your own numbers.",
        "Log every purchase on the Spending Log tab as it happens: pick a category from the dropdown, then fill in the date, description, and amount.",
        "The Budget tab's Actual column totals the Spending Log for you automatically, category by category -- no manual math.",
        "Check the \"Left to assign\" tile on the Budget tab any time to see how much of your income isn't planned for yet.",
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
            "label": "Monthly income",
            "value": MONTHLY_INCOME_SAMPLE,
            "name": "MonthlyIncome",
            "number_format": common.NUMBER_FMT,
            "note": "Your total take-home pay for the month. Drives the % of Income column and the \"Left to assign\" tile on the Budget tab.",
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
            "This free template covers one month with a fixed set of categories. "
            "The full Personal Budget & Debt Payoff Planner builds on top of it with:"
        ),
    )
    intro.font = common.body_font
    intro.alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=4)
    ws.row_dimensions[row].height = 32
    row += 1

    bullets = [
        "A Monthly Budget Tracker with 12 monthly tabs (Jan-Dec) instead of one, plus editable income and expense category lists (6 income, 16 expense) instead of a fixed set.",
        "A Dashboard tab with a year-at-a-glance summary, charts, and best/worst-month highlights.",
        "A separate Debt Payoff Planner workbook: track up to 15 debts with live Snowball and Avalanche payoff-order rankings, a 100-row payoff log, and its own Dashboard.",
    ]
    for b in bullets:
        c = ws.cell(row=row, column=2, value=f"-  {b}")
        c.font = common.body_font
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=4)
        ws.row_dimensions[row].height = 30
        row += 1

    row += 1
    link1 = ws.cell(row=row, column=2, value="See the full Personal Budget & Debt Payoff Planner on Etsy")
    link1.font = link_font
    link1.hyperlink = ETSY_P4_URL
    row += 1
    link2 = ws.cell(row=row, column=2, value="Need something custom built for your business? Custom spreadsheets, made to order")
    link2.font = link_font
    link2.hyperlink = ETSY_CUSTOM_URL
    ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=4)
    row += 2

    return ws


# ---------------------------------------------------------------------------
# Budget
# ---------------------------------------------------------------------------

def _compute_layout():
    """Return (layout, item_rows, total_row).

    layout: list of ("group", row, group_name) | ("item", row, cat_name, planned)
    item_rows: list of row numbers, in the same order as CATEGORIES.
    total_row: row number of the Totals row.
    """
    layout = []
    item_rows = []
    row = BUDGET_FIRST_LAYOUT_ROW
    for gname, items in GROUPS:
        layout.append(("group", row, gname))
        row += 1
        for name, planned in items:
            layout.append(("item", row, name, planned))
            item_rows.append(row)
            row += 1
    total_row = row + 1  # one blank row, then Totals
    return layout, item_rows, total_row


def build_budget_sheet(wb):
    ws = common.new_data_sheet(wb, "Budget", tab_color=common.TAB_TEAL, gridlines=False)
    common.set_col_widths(ws, {"A": 26, "B": 14, "C": 14, "D": 14, "E": 13, "F": 2, "G": 26})

    layout, item_rows, total_row = _compute_layout()
    assert len(item_rows) == len(CATEGORIES)
    layout_first_row = BUDGET_FIRST_LAYOUT_ROW
    layout_last_row = item_rows[-1]

    common.style_title(ws, "A1", '="Budget -- " & TEXT(TODAY(),"mmmm yyyy")')
    sub = ws.cell(row=2, column=1, value=(
        "One month, one plan. Set your Planned amount per category; Actual pulls "
        "automatically from the Spending Log tab as you log purchases."
    ))
    sub.font = common.subtitle_font
    sub.alignment = Alignment(wrap_text=True)
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=5)
    ws.row_dimensions[2].height = 18

    # "Left to assign" tile
    tile_label = ws.cell(row=4, column=1, value='="Left to assign (" & CurrencySymbol & ")"')
    tile_label.font = Font(name=common.FONT_NAME, size=11, bold=True, color="6B7280")
    tile_value = ws.cell(row=5, column=1, value=f"=MonthlyIncome-B{total_row}")
    tile_value.font = Font(name=common.FONT_NAME, size=20, bold=True, color=common.HEADER_FILL_COLOR)
    tile_value.fill = common.accent_light_fill
    tile_value.border = common.thin_border
    tile_value.number_format = common.NUMBER_FMT_NEG_RED
    tile_value.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.merge_cells(start_row=4, start_column=1, end_row=4, end_column=2)
    ws.merge_cells(start_row=5, start_column=1, end_row=5, end_column=2)
    ws.row_dimensions[5].height = 30
    common.apply_negative_red(ws, f"A5:B5")

    # Header row
    common.style_header_row(ws, BUDGET_HEADER_ROW, 1, 5)
    headers = [
        "Category",
        '="Planned (" & CurrencySymbol & ")"',
        '="Actual (" & CurrencySymbol & ")"',
        '="Difference (" & CurrencySymbol & ")"',
        "% of Income",
    ]
    for col, val in enumerate(headers, start=1):
        ws.cell(row=BUDGET_HEADER_ROW, column=col, value=val)

    sl_cat_range = f"'Spending Log'!$B${SL_FIRST_DATA_ROW}:$B${SL_LAST_DATA_ROW}"
    sl_amt_range = f"'Spending Log'!$D${SL_FIRST_DATA_ROW}:$D${SL_LAST_DATA_ROW}"

    group_fill = common.accent_light_fill
    for entry in layout:
        if entry[0] == "group":
            _, row, gname = entry
            c = ws.cell(row=row, column=1, value=gname)
            c.font = Font(name=common.FONT_NAME, size=11, bold=True, color=common.HEADER_FILL_COLOR)
            c.fill = group_fill
            ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=5)
            for col in range(1, 6):
                ws.cell(row=row, column=col).fill = group_fill
                ws.cell(row=row, column=col).border = common.thin_border
        else:
            _, row, name, planned = entry
            ws.cell(row=row, column=1, value=name)
            ws.cell(row=row, column=2, value=planned)
            ws.cell(row=row, column=3, value=f"=SUMIF({sl_cat_range},$A{row},{sl_amt_range})")
            ws.cell(row=row, column=4, value=f"=B{row}-C{row}")
            ws.cell(row=row, column=5, value=f"=IFERROR(B{row}/MonthlyIncome,0)")

    common.band_rows(ws, layout_first_row, layout_last_row, 1, 5)
    # Re-apply group-header fill/merge styling that band_rows may have overwritten.
    for entry in layout:
        if entry[0] == "group":
            row = entry[1]
            for col in range(1, 6):
                ws.cell(row=row, column=col).fill = group_fill

    for row in item_rows:
        ws.cell(row=row, column=2).number_format = common.NUMBER_FMT_NEG_RED
        ws.cell(row=row, column=3).number_format = common.NUMBER_FMT_NEG_RED
        ws.cell(row=row, column=4).number_format = common.NUMBER_FMT_NEG_RED
        ws.cell(row=row, column=5).number_format = common.PERCENT_FMT

    # Conditional format: red when Actual > Planned.
    over_fill = PatternFill("solid", fgColor="FEE2E2")
    ws.conditional_formatting.add(
        f"C{layout_first_row}:C{layout_last_row}",
        FormulaRule(formula=[f"$C{layout_first_row}>$B{layout_first_row}"], fill=over_fill),
    )

    # Totals row
    total_label = ws.cell(row=total_row, column=1, value="Total")
    total_planned = ws.cell(row=total_row, column=2, value=f"=SUM(B{layout_first_row}:B{layout_last_row})")
    total_actual = ws.cell(row=total_row, column=3, value=f"=SUM(C{layout_first_row}:C{layout_last_row})")
    total_diff = ws.cell(row=total_row, column=4, value=f"=B{total_row}-C{total_row}")
    total_pct = ws.cell(row=total_row, column=5, value=f"=IFERROR(B{total_row}/MonthlyIncome,0)")
    for col in range(1, 6):
        cell = ws.cell(row=total_row, column=col)
        cell.fill = common.accent_fill
        cell.font = Font(name=common.FONT_NAME, size=11, bold=True, color=common.WHITE)
        cell.border = common.thin_border
    total_planned.number_format = common.NUMBER_FMT_NEG_RED
    total_actual.number_format = common.NUMBER_FMT_NEG_RED
    total_diff.number_format = common.NUMBER_FMT_NEG_RED
    total_pct.number_format = common.PERCENT_FMT
    ws.row_dimensions[total_row].height = 22

    # Hidden-in-plain-sight reference list: source for the Spending Log
    # Category dropdown (kept as a real contiguous range rather than an
    # inline comma list, since 20 category names exceed Excel's 255-char
    # inline data-validation limit).
    ws.cell(row=1, column=7, value="Categories").font = common.note_font
    for i, name in enumerate(CATEGORIES):
        ws.cell(row=2 + i, column=7, value=name).font = common.note_font
    cat_first = 2
    cat_last = 2 + len(CATEGORIES) - 1
    wb.defined_names["BudgetCategories"] = DefinedName(
        "BudgetCategories", attr_text=f"Budget!${CAT_LIST_COL}${cat_first}:${CAT_LIST_COL}${cat_last}"
    )

    ws.sheet_view.zoomScale = 100
    return ws, total_row


# ---------------------------------------------------------------------------
# Spending Log
# ---------------------------------------------------------------------------

def build_spending_log_sheet(wb):
    ws = common.new_data_sheet(wb, "Spending Log", tab_color=common.TAB_SLATE)
    common.set_col_widths(ws, {"A": 13, "B": 22, "C": 38, "D": 14})

    common.style_title(ws, "A1", "Spending Log")
    sub = ws.cell(row=2, column=1, value=(
        "Log every purchase here. Pick a category from the dropdown -- the Budget tab's "
        "Actual column totals it automatically."
    ))
    sub.font = common.subtitle_font
    sub.alignment = Alignment(wrap_text=True)
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=4)
    ws.row_dimensions[2].height = 18

    common.style_header_row(ws, SL_HEADER_ROW, 1, 4)
    headers = ["Date", "Category", "Description", '="Amount (" & CurrencySymbol & ")"']
    for col, val in enumerate(headers, start=1):
        ws.cell(row=SL_HEADER_ROW, column=col, value=val)

    for i, (d, cat, desc, amt) in enumerate(SAMPLE_LOG_ROWS):
        r = SL_FIRST_DATA_ROW + i
        ws.cell(row=r, column=1, value=d)
        ws.cell(row=r, column=2, value=cat)
        ws.cell(row=r, column=3, value=desc)
        ws.cell(row=r, column=4, value=amt)

    common.band_rows(ws, SL_FIRST_DATA_ROW, SL_LAST_DATA_ROW, 1, 4)

    for r in range(SL_FIRST_DATA_ROW, SL_LAST_DATA_ROW + 1):
        ws.cell(row=r, column=1).number_format = common.DATE_FMT
        ws.cell(row=r, column=4).number_format = common.NUMBER_FMT

    category_range = f"B{SL_FIRST_DATA_ROW}:B{SL_LAST_DATA_ROW}"
    dv = DataValidation(type="list", formula1="BudgetCategories", allow_blank=True,
                         showErrorMessage=True, showDropDown=False)
    dv.error = "Choose a category from the Budget tab's category list."
    dv.errorTitle = "Invalid category"
    ws.add_data_validation(dv)
    dv.add(category_range)

    common.freeze_header(ws, f"A{SL_FIRST_DATA_ROW}")
    ws.sheet_view.zoomScale = 100
    return ws


# ---------------------------------------------------------------------------
# Build
# ---------------------------------------------------------------------------

def build():
    wb = Workbook()
    build_start_here_sheet(wb)
    build_budget_sheet(wb)
    build_spending_log_sheet(wb)

    assert wb.sheetnames == ["Start Here", "Budget", "Spending Log"], wb.sheetnames

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

_RANGE_RE = re.compile(r"(?<![!\w$'])(\$?[A-Z]{1,3}\$?\d+:\$?[A-Z]{1,3}\$?\d+)")


def _strip_strings(formula):
    return re.sub(r'"[^"]*"', "", formula)


def verify():
    assert OUT_PATH.exists(), f"{OUT_PATH} does not exist -- run without --verify first"
    wb = load_workbook(OUT_PATH)

    # 1. Sheet names, exactly 3, in order.
    assert wb.sheetnames == ["Start Here", "Budget", "Spending Log"], (
        f"unexpected sheet list: {wb.sheetnames}"
    )
    print("PASS: sheet names ==", wb.sheetnames)

    budget = wb["Budget"]
    log = wb["Spending Log"]

    layout, item_rows, total_row = _compute_layout()
    layout_first_row = BUDGET_FIRST_LAYOUT_ROW
    layout_last_row = item_rows[-1]

    # 2. Formulas present on Budget and Spending Log.
    budget_formula_cells = [
        c for row in budget.iter_rows() for c in row
        if isinstance(c.value, str) and c.value.startswith("=")
    ]
    # 20 items x 3 formula cols (Actual, Difference, %ofIncome) + 4 totals
    # formulas + title + tile label + tile value + 4 header formulas = 69
    assert len(budget_formula_cells) >= 20 * 3 + 4, (
        f"expected at least {20*3+4} formulas on Budget, found {len(budget_formula_cells)}"
    )
    print(f"PASS: Budget has {len(budget_formula_cells)} formula cell(s)")

    log_formula_cells = [
        c for row in log.iter_rows() for c in row
        if isinstance(c.value, str) and c.value.startswith("=")
    ]
    assert log_formula_cells, "Spending Log has no formulas (expected the Amount header formula)"
    print(f"PASS: Spending Log has {len(log_formula_cells)} formula cell(s)")

    # 3. Required cells: Actual/Difference/%ofIncome per item row, Left to
    #    assign tile, Totals row.
    for row in item_rows:
        for col in (3, 4, 5):
            v = budget.cell(row=row, column=col).value
            assert isinstance(v, str) and v.startswith("="), (
                f"Budget!{budget.cell(row=row, column=col).coordinate} missing formula, got {v!r}"
            )
    assert isinstance(budget.cell(row=5, column=1).value, str), "Left to assign tile formula missing"
    for col in (2, 3, 4, 5):
        v = budget.cell(row=total_row, column=col).value
        assert isinstance(v, str) and v.startswith("="), f"Totals row col {col} missing formula, got {v!r}"
    print("PASS: all Budget item-row and totals formulas present")

    # 4. Data validation: exactly 1 on Spending Log (Category dropdown).
    dvs = log.data_validations.dataValidation
    assert len(dvs) == 1, f"expected 1 data validation on Spending Log, found {len(dvs)}"
    dv = dvs[0]
    assert dv.type == "list", f"unexpected validation type {dv.type!r}"
    assert dv.formula1 == "BudgetCategories", f"expected formula1='BudgetCategories', got {dv.formula1!r}"
    sqref = str(dv.sqref)
    assert f"B{SL_FIRST_DATA_ROW}:B{SL_LAST_DATA_ROW}" in sqref, f"Category dropdown range wrong: {sqref}"
    print("PASS: 1 list data validation present on Spending Log, sourced from BudgetCategories:", sqref)

    # 5. Defined names present.
    assert "CurrencySymbol" in wb.defined_names, "CurrencySymbol defined name missing"
    assert "MonthlyIncome" in wb.defined_names, "MonthlyIncome defined name missing"
    assert "BudgetCategories" in wb.defined_names, "BudgetCategories defined name missing"
    print("PASS: CurrencySymbol, MonthlyIncome, BudgetCategories defined names present")

    # 6. No unprefixed cross-sheet ranges: every SUMIF Actual formula on
    #    Budget (the only formulas that reach across sheets in this
    #    workbook) must carry a literal 'Spending Log'! prefix on both the
    #    criteria range and the sum range -- this is the exact bug class
    #    (an unprefixed range silently resolving to the formula's own
    #    sheet) that shipped silently in two prior products per
    #    build/audit_products.py's docstring.
    bad_actual = []
    for row in item_rows:
        v = budget.cell(row=row, column=3).value
        if v.count("'Spending Log'!") != 2:
            bad_actual.append((row, v))
    assert not bad_actual, f"Actual formulas missing sheet-prefixed Spending Log ranges: {bad_actual}"
    print("PASS: every Actual SUMIF is sheet-prefixed to 'Spending Log' (no unprefixed cross-sheet refs)")

    # 7. Whitelisted functions only.
    func_counter, disallowed_common, placeholder_hits = common.scan_workbook_formulas(wb)
    assert not disallowed_common, f"functions outside common.py's ALLOWED_FUNCTIONS: {disallowed_common}"
    disallowed_task = set(func_counter) - TASK_ALLOWED_FUNCTIONS
    assert not disallowed_task, f"functions outside the task brief's whitelist: {disallowed_task}"
    assert not placeholder_hits, f"leftover [placeholder] text found: {placeholder_hits}"
    print("PASS: functions used are all whitelisted:", sorted(func_counter))
    print("PASS: no leftover [placeholder] text")

    # 8. Sanity: category count, sample data.
    assert len(CATEGORIES) == 20, len(CATEGORIES)
    cat_col_values = [budget.cell(row=2 + i, column=7).value for i in range(20)]
    assert cat_col_values == CATEGORIES, "Budget!G2:G21 category list doesn't match CATEGORIES"
    print("PASS: 20 categories in Budget, matching the hidden reference list")

    planned_filled = sum(1 for row in item_rows if budget.cell(row=row, column=2).value not in (None, 0))
    assert planned_filled == 20, f"expected all 20 Planned cells filled, found {planned_filled}"
    print("PASS: all 20 Planned amounts filled with sample data")

    mi_attr = wb.defined_names["MonthlyIncome"].attr_text  # e.g. "'Start Here'!$C$15"
    mi_sheet, mi_cell = mi_attr.split("!")
    mi_sheet = mi_sheet.strip("'")
    income_val = wb[mi_sheet][mi_cell.replace("$", "")].value
    assert income_val == MONTHLY_INCOME_SAMPLE, f"MonthlyIncome sample value wrong: {income_val}"
    print(f"PASS: MonthlyIncome sample filled ({income_val})")

    filled_log_rows = sum(
        1 for r in range(SL_FIRST_DATA_ROW, SL_LAST_DATA_ROW + 1)
        if log.cell(row=r, column=1).value is not None
    )
    assert filled_log_rows == len(SAMPLE_LOG_ROWS), (
        f"expected {len(SAMPLE_LOG_ROWS)} filled sample log rows, found {filled_log_rows}"
    )
    total_log_rows = SL_LAST_DATA_ROW - SL_FIRST_DATA_ROW + 1
    assert total_log_rows == SL_TOTAL_ROWS, total_log_rows
    print(f"PASS: {SL_TOTAL_ROWS} pre-formatted Spending Log rows, {filled_log_rows} sample rows filled")

    print("\nAll --verify checks passed.")


if __name__ == "__main__":
    if "--verify" in sys.argv:
        verify()
    else:
        build()
