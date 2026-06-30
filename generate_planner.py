"""
Indian Wedding Planner - Excel Workbook Generator
==================================================
Generates a comprehensive Indian Wedding Planner (.xlsx) for a
Baniya groom x Maharashtrian bride wedding.

Run:
    pip install openpyxl
    python generate_planner.py

Output:
    Indian_Wedding_Planner.xlsx
"""

import openpyxl
from openpyxl.styles import (
    PatternFill, Font, Alignment, Border, Side, Protection
)
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule, PatternFill as RulePatternFill
from openpyxl.chart import PieChart, Reference
from openpyxl.chart.series import DataPoint

# ─────────────────────────────────────────────
# THEME COLORS
# ─────────────────────────────────────────────
GOLD        = "FFD700"
IVORY       = "FFFFF0"
MAROON      = "800000"
LIGHT_GOLD  = "FFF8DC"
WHITE       = "FFFFFF"
GREEN       = "90EE90"
RED         = "FF6B6B"
ORANGE      = "FFD580"
DARK_GOLD   = "C8A400"
HEADER_TEXT = "FFFFFF"

# ─────────────────────────────────────────────
# FILL HELPERS
# ─────────────────────────────────────────────
def fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

def thin_border():
    s = Side(style="thin")
    return Border(left=s, right=s, top=s, bottom=s)

def thick_border():
    s = Side(style="medium")
    return Border(left=s, right=s, top=s, bottom=s)

def header_font(size=11, color=HEADER_TEXT):
    return Font(name="Calibri", bold=True, size=size, color=color)

def data_font(size=10):
    return Font(name="Calibri", size=size)

def title_font(size=18):
    return Font(name="Calibri", bold=True, size=size, color=MAROON)

def center():
    return Alignment(horizontal="center", vertical="center", wrap_text=True)

def left_align():
    return Alignment(horizontal="left", vertical="center", wrap_text=True)

def currency_fmt():
    return '₹#,##0'

def pct_fmt():
    return '0%'


# ─────────────────────────────────────────────
# STYLE A HEADER ROW
# ─────────────────────────────────────────────
def style_header_row(ws, row, col_count, start_col=1):
    for c in range(start_col, start_col + col_count):
        cell = ws.cell(row=row, column=c)
        cell.fill = fill(MAROON)
        cell.font = header_font()
        cell.alignment = center()
        cell.border = thin_border()

def style_data_row(ws, row, col_count, start_col=1, alternate=False):
    bg = LIGHT_GOLD if alternate else IVORY
    for c in range(start_col, start_col + col_count):
        cell = ws.cell(row=row, column=c)
        cell.fill = fill(bg)
        cell.font = data_font()
        cell.alignment = left_align()
        cell.border = thin_border()

def set_col_widths(ws, widths):
    """widths: list of (col_letter_or_idx, width)"""
    for col, w in widths:
        if isinstance(col, int):
            col = get_column_letter(col)
        ws.column_dimensions[col].width = w

def freeze(ws, cell="A2"):
    ws.freeze_panes = cell

def add_dv(ws, formula1, sqref, allow_blank=True, prompt=""):
    """Add a list dropdown data validation."""
    dv = DataValidation(
        type="list",
        formula1=formula1,
        allow_blank=allow_blank,
        showDropDown=False
    )
    dv.prompt = prompt
    dv.sqref = sqref
    ws.add_data_validation(dv)
    return dv

def add_num_dv(ws, sqref):
    """Non-negative number validation."""
    dv = DataValidation(
        type="decimal",
        operator="greaterThanOrEqual",
        formula1="0",
        allow_blank=True
    )
    dv.error = "Value must be 0 or more"
    dv.errorTitle = "Invalid Input"
    dv.sqref = sqref
    ws.add_data_validation(dv)


# ═══════════════════════════════════════════════════════════════
# SHEET 0 — COVER SHEET
# ═══════════════════════════════════════════════════════════════
def create_cover_sheet(wb):
    ws = wb.active
    ws.title = "Cover"

    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 40

    # Title
    ws.merge_cells("A1:B1")
    t = ws["A1"]
    t.value = "🌸 Indian Wedding Planner 🌸"
    t.font = title_font(22)
    t.fill = fill(GOLD)
    t.alignment = center()
    t.border = thick_border()
    ws.row_dimensions[1].height = 50

    # Sub-title
    ws.merge_cells("A2:B2")
    s = ws["A2"]
    s.value = "Baniya Groom  ×  Maharashtrian Bride"
    s.font = Font(name="Calibri", bold=True, size=13, color=MAROON)
    s.fill = fill(LIGHT_GOLD)
    s.alignment = center()
    s.border = thin_border()
    ws.row_dimensions[2].height = 28

    ws.row_dimensions[3].height = 12  # spacer

    # Input fields
    fields = [
        ("Bride's Name", ""),
        ("Groom's Name", ""),
        ("Wedding Date", ""),
        ("Venue", ""),
        ("Total Budget (₹)", ""),
    ]

    label_fill  = fill(MAROON)
    value_fill  = fill(IVORY)

    for i, (label, val) in enumerate(fields, start=4):
        lc = ws.cell(row=i, column=1, value=label)
        lc.font  = header_font(11, HEADER_TEXT)
        lc.fill  = label_fill
        lc.alignment = left_align()
        lc.border = thin_border()

        vc = ws.cell(row=i, column=2, value=val)
        vc.font  = data_font(11)
        vc.fill  = value_fill
        vc.alignment = left_align()
        vc.border = thin_border()
        ws.row_dimensions[i].height = 22

    # Name the Total Budget cell for use in formulas
    ws["B8"].value = 0   # default 0; user fills in
    wb.defined_names["TotalBudget"] = openpyxl.workbook.defined_name.DefinedName(
        "TotalBudget", attr_text="Cover!$B$8"
    )

    # Spacer
    ws.row_dimensions[9].height = 14

    # Sheet Index
    idx_header = ws.cell(row=10, column=1, value="📋  Sheet Index")
    idx_header.font = header_font(12, MAROON)
    idx_header.fill = fill(GOLD)
    idx_header.alignment = left_align()
    idx_header.border = thin_border()
    ws.merge_cells("A10:B10")

    sheets_index = [
        ("Sheet 1",  "Master Budget Overview"),
        ("Sheet 2",  "Sakarpuda"),
        ("Sheet 3",  "Engagement"),
        ("Sheet 4",  "Haldi"),
        ("Sheet 5",  "Mehendi"),
        ("Sheet 6",  "Sangeet"),
        ("Sheet 7",  "Antrapat"),
        ("Sheet 8",  "Jaimal & Phere"),
        ("Sheet 9",  "Reception"),
        ("Sheet 10", "Miscellaneous"),
        ("Sheet 11", "Guest List"),
        ("Sheet 12", "Vendor Tracker"),
        ("Sheet 13", "Payment Tracker"),
        ("Sheet 14", "Checklist & Timeline"),
        ("Sheet 15", "Trousseau & Shopping"),
    ]

    for i, (sheet_num, sheet_name) in enumerate(sheets_index, start=11):
        rc = ws.cell(row=i, column=1, value=sheet_num)
        rc.font  = data_font(10)
        rc.fill  = fill(LIGHT_GOLD if i % 2 == 0 else IVORY)
        rc.border = thin_border()
        rc.alignment = left_align()

        nc = ws.cell(row=i, column=2, value=sheet_name)
        nc.font      = Font(name="Calibri", size=10, color="0000FF", underline="single")
        nc.fill      = fill(LIGHT_GOLD if i % 2 == 0 else IVORY)
        nc.border    = thin_border()
        nc.alignment = left_align()
        # Hyperlink to the sheet
        nc.hyperlink = f"#'{sheet_name}'!A1"
        ws.row_dimensions[i].height = 18

    freeze(ws, "A3")


# ═══════════════════════════════════════════════════════════════
# SHEET 1 — MASTER BUDGET OVERVIEW
# ═══════════════════════════════════════════════════════════════
CEREMONIES = [
    ("Sakarpuda",      0.04),
    ("Engagement",     0.08),
    ("Haldi",          0.05),
    ("Mehendi",        0.06),
    ("Sangeet",        0.18),
    ("Antrapat",       0.04),
    ("Jaimal & Phere", 0.20),
    ("Reception",      0.22),
    ("Miscellaneous",  0.13),
]

# Map ceremony name → its sheet name (same as ceremony name)
CEREMONY_SHEET_NAMES = [c[0] for c in CEREMONIES]

def create_budget_sheet(wb):
    ws = wb.create_sheet("Budget Overview")

    ws.sheet_view.showGridLines = False
    set_col_widths(ws, [
        (1, 22), (2, 14), (3, 20), (4, 18), (5, 16), (6, 18)
    ])
    ws.row_dimensions[1].height = 40

    # Title
    ws.merge_cells("A1:F1")
    t = ws["A1"]
    t.value = "Master Budget Overview"
    t.font = title_font(16)
    t.fill = fill(GOLD)
    t.alignment = center()
    t.border = thick_border()

    # Headers row 2
    headers = ["Ceremony", "Budget %", "Allocated Amount (₹)",
               "Actual Spent (₹)", "Variance (₹)", "Status"]
    for c, h in enumerate(headers, 1):
        cell = ws.cell(row=2, column=c, value=h)
        cell.fill = fill(MAROON)
        cell.font = header_font()
        cell.alignment = center()
        cell.border = thin_border()
    ws.row_dimensions[2].height = 30

    # Data rows
    for i, (name, pct) in enumerate(CEREMONIES, start=3):
        row = i
        alt = (i % 2 == 0)
        bg  = LIGHT_GOLD if alt else IVORY

        ws.cell(row=row, column=1, value=name).fill = fill(bg)
        ws.cell(row=row, column=1).font = data_font()
        ws.cell(row=row, column=1).border = thin_border()
        ws.cell(row=row, column=1).alignment = left_align()

        pct_cell = ws.cell(row=row, column=2, value=pct)
        pct_cell.number_format = pct_fmt()
        pct_cell.fill = fill(bg)
        pct_cell.font = data_font()
        pct_cell.border = thin_border()
        pct_cell.alignment = center()

        # Allocated = TotalBudget * pct%
        alloc_cell = ws.cell(row=row, column=3,
                             value=f"=Cover!$B$8*B{row}")
        alloc_cell.number_format = currency_fmt()
        alloc_cell.fill = fill(bg)
        alloc_cell.font = data_font()
        alloc_cell.border = thin_border()
        alloc_cell.alignment = center()

        # Actual Spent — user fills or 0
        actual_cell = ws.cell(row=row, column=4, value=0)
        actual_cell.number_format = currency_fmt()
        actual_cell.fill = fill(WHITE)
        actual_cell.font = data_font()
        actual_cell.border = thin_border()
        actual_cell.alignment = center()

        # Variance
        var_cell = ws.cell(row=row, column=5,
                           value=f"=C{row}-D{row}")
        var_cell.number_format = currency_fmt()
        var_cell.fill = fill(bg)
        var_cell.font = data_font()
        var_cell.border = thin_border()
        var_cell.alignment = center()

        # Status
        status_cell = ws.cell(row=row, column=6,
            value=f'=IF(E{row}<0,"Over Budget",IF(E{row}=0,"On Track","Under Budget"))')
        status_cell.fill = fill(bg)
        status_cell.font = data_font()
        status_cell.border = thin_border()
        status_cell.alignment = center()

        ws.row_dimensions[row].height = 20

    # Total row
    total_row = len(CEREMONIES) + 3
    ws.cell(row=total_row, column=1, value="TOTAL").font = header_font(11, MAROON)
    ws.cell(row=total_row, column=1).fill = fill(GOLD)
    ws.cell(row=total_row, column=1).border = thin_border()
    ws.cell(row=total_row, column=1).alignment = center()

    for col in range(2, 7):
        c_letter = get_column_letter(col)
        first = 3
        last  = total_row - 1
        cell  = ws.cell(row=total_row, column=col,
                        value=f"=SUM({c_letter}{first}:{c_letter}{last})")
        if col in (2,):
            cell.number_format = pct_fmt()
        elif col in (3, 4, 5):
            cell.number_format = currency_fmt()
        cell.fill = fill(GOLD)
        cell.font = header_font(11, MAROON)
        cell.border = thin_border()
        cell.alignment = center()
    ws.row_dimensions[total_row].height = 22

    # ── Conditional formatting on Status column (F)
    status_col = "F"
    data_range = f"{status_col}3:{status_col}{total_row-1}"
    green_fill = PatternFill("solid", fgColor=GREEN)
    red_fill   = PatternFill("solid", fgColor=RED)
    ora_fill   = PatternFill("solid", fgColor=ORANGE)

    ws.conditional_formatting.add(data_range,
        CellIsRule(operator="equal", formula=['"Under Budget"'], fill=green_fill))
    ws.conditional_formatting.add(data_range,
        CellIsRule(operator="equal", formula=['"Over Budget"'],  fill=red_fill))
    ws.conditional_formatting.add(data_range,
        CellIsRule(operator="equal", formula=['"On Track"'],     fill=ora_fill))

    # ── Pie Chart
    chart = PieChart()
    chart.title = "Wedding Budget Distribution"
    chart.style = 10

    labels = Reference(ws, min_col=1, min_row=3, max_row=total_row-1)
    data   = Reference(ws, min_col=2, min_row=2, max_row=total_row-1)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(labels)
    chart.dataLabels = openpyxl.chart.label.DataLabelList()
    chart.dataLabels.showPercent = True
    chart.dataLabels.showCatName = True
    chart.width  = 18
    chart.height = 14

    ws.add_chart(chart, f"A{total_row + 2}")

    freeze(ws, "A3")


# ═══════════════════════════════════════════════════════════════
# SHEETS 2–10 — INDIVIDUAL CEREMONY SHEETS
# ═══════════════════════════════════════════════════════════════

CEREMONY_ITEMS = {
    "Sakarpuda": [
        ("Venue & Logistics", "Venue/hall booking"),
        ("Puja",              "Puja samagri & thali"),
        ("Puja",              "Priest (Pandit) fees"),
        ("Food & Sweets",     "Sweets & dry fruits exchange (Baniya tradition)"),
        ("Bride's Attire",    "Nauvari saree for bride (Maharashtrian custom)"),
        ("Gifts",             "Gifts for bride's family"),
        ("Photography",       "Photographer"),
        ("Catering",          "Catering – tea, snacks, refreshments"),
        ("Decor",             "Floral decoration"),
        ("Stationery",        "Invitation cards"),
    ],
    "Engagement": [
        ("Venue & Logistics", "Venue booking & setup"),
        ("Decor",             "Decoration & floral arrangements"),
        ("Catering",          "Full meal catering"),
        ("Jewelry",           "Engagement ring – Bride"),
        ("Jewelry",           "Engagement ring – Groom"),
        ("Attire",            "Bride's outfit"),
        ("Attire",            "Groom's outfit"),
        ("Beauty",            "Mehendi artist (light design)"),
        ("Photography",       "Photographer"),
        ("Photography",       "Videographer"),
        ("Entertainment",     "DJ / Music system"),
        ("Stationery",        "Invitation cards"),
        ("Gifts",             "Return gifts for guests"),
        ("Puja",              "Priest fees"),
        ("Decor",             "Fresh flowers"),
    ],
    "Haldi": [
        ("Venue & Logistics", "Venue (home/lawn)"),
        ("Haldi",             "Haldi ingredients & decorated thali"),
        ("Decor",             "Marigold floral decoration"),
        ("Attire",            "Traditional yellow outfits (bride + groom)"),
        ("Photography",       "Photographer"),
        ("Catering",          "Catering – breakfast/brunch"),
        ("Entertainment",     "Dhol player"),
        ("Beauty",            "Ubtan & beauty supplies"),
    ],
    "Mehendi": [
        ("Beauty",        "Mehendi artist(s) – bridal"),
        ("Beauty",        "Mehendi artist(s) – guests"),
        ("Decor",         "Stage & seating decoration"),
        ("Entertainment", "DJ / playlist setup"),
        ("Catering",      "Catering – snacks + dinner"),
        ("Attire",        "Bride's outfit"),
        ("Photography",   "Photographer"),
        ("Gifts",         "Favors / return gifts"),
        ("Lighting",      "Decorative lighting"),
    ],
    "Sangeet": [
        ("Venue & Logistics", "Venue booking"),
        ("Decor",             "Stage & backdrop décor"),
        ("Entertainment",     "Sound system & DJ"),
        ("Entertainment",     "Live band / performers"),
        ("Entertainment",     "Choreographer fees"),
        ("Lighting",          "LED / fairy lights"),
        ("Catering",          "Dinner + drinks catering"),
        ("Attire",            "Bride's outfit (multiple changes)"),
        ("Attire",            "Groom's outfit"),
        ("Photography",       "Photographer"),
        ("Photography",       "Videographer"),
        ("Entertainment",     "Photo booth"),
        ("Stationery",        "Invitation cards"),
        ("Entertainment",     "Anchor / emcee"),
        ("Gifts",             "Return gifts for guests"),
    ],
    "Antrapat": [
        ("Puja",        "Priest (Pandit) fees"),
        ("Ritual",      "Antrapat cloth / silk"),
        ("Puja",        "Puja samagri"),
        ("Decor",       "Flowers & garlands"),
        ("Photography", "Photographer"),
    ],
    "Jaimal & Phere": [
        ("Decor",         "Mandap decoration"),
        ("Puja",          "Priest (Pandit) fees"),
        ("Puja",          "Sacred fire / havan samagri"),
        ("Attire",        "Bride's Nauvari saree + blouse"),
        ("Attire",        "Groom's sherwani + accessories"),
        ("Jewelry",       "Wedding jewelry (bride)"),
        ("Logistics",     "Doli / vidai arrangements"),
        ("Photography",   "Photographer"),
        ("Photography",   "Videographer"),
        ("Catering",      "Full day catering (breakfast + lunch + dinner)"),
        ("Entertainment", "Band / Baraat music"),
        ("Logistics",     "Horse / decorated car for groom"),
        ("Decor",         "Flower garlands (Jaimal)"),
        ("Beauty",        "Bridal makeup artist"),
        ("Beauty",        "Hair stylist"),
        ("Beauty",        "Mehendi touch-up artist"),
    ],
    "Reception": [
        ("Venue & Logistics", "Venue (banquet hall / hotel)"),
        ("Catering",          "Dinner catering – veg"),
        ("Catering",          "Dinner catering – non-veg"),
        ("Catering",          "Bar / beverages"),
        ("Decor",             "Stage & floral décor"),
        ("Entertainment",     "DJ + dance floor setup"),
        ("Entertainment",     "Live music / band"),
        ("Lighting",          "Lighting & AV setup"),
        ("Photography",       "Photographer"),
        ("Photography",       "Videographer"),
        ("Gifts",             "Return gifts / favors"),
        ("Attire",            "Bride's reception outfit"),
        ("Attire",            "Groom's reception outfit"),
        ("Stationery",        "Invitation cards"),
        ("Entertainment",     "Anchor / emcee"),
        ("Entertainment",     "Photo booth"),
        ("Cake",              "Wedding / reception cake"),
    ],
    "Miscellaneous": [
        ("Travel",       "Travel & accommodation – outstation guests"),
        ("Travel",       "Honeymoon booking"),
        ("Stationery",   "Wedding website"),
        ("Buffer",       "Emergency buffer fund"),
        ("Vendor Tips",  "Tips for vendors"),
        ("Attire",       "Dress alterations"),
        ("Beauty",       "Pre-wedding facials & parlour"),
        ("Stationery",   "Stationery & printing"),
    ],
}

def create_ceremony_sheet(wb, ceremony_name, budget_row):
    """budget_row: row number in Budget Overview sheet for this ceremony."""
    ws = wb.create_sheet(ceremony_name)
    ws.sheet_view.showGridLines = False

    set_col_widths(ws, [
        (1, 20), (2, 35), (3, 25), (4, 18), (5, 18),
        (6, 12), (7, 16), (8, 30)
    ])

    # ── Row 1: Ceremony Title
    ws.merge_cells("A1:H1")
    t = ws["A1"]
    t.value = f"✨ {ceremony_name} ✨"
    t.font  = title_font(16)
    t.fill  = fill(GOLD)
    t.alignment = center()
    t.border = thick_border()
    ws.row_dimensions[1].height = 40

    # ── Row 2: Summary info
    summary_labels = [
        ("Allocated Budget (₹)", f"='Budget Overview'!C{budget_row}"),
        ("Total Estimated Cost (₹)", None),   # will be SUM formula
        ("Total Actual Cost (₹)", None),
        ("Variance (₹)", None),
    ]

    # We'll put summary in row 2 spanning cols 1-8 in pairs
    summary_cols = [(1,2), (3,4), (5,6), (7,8)]
    for (lc, vc), (label, formula) in zip(summary_cols, summary_labels):
        lbl = ws.cell(row=2, column=lc, value=label)
        lbl.font = header_font(10, HEADER_TEXT)
        lbl.fill = fill(MAROON)
        lbl.alignment = center()
        lbl.border = thin_border()
        ws.merge_cells(
            start_row=2, start_column=lc,
            end_row=2,   end_column=lc
        )

        val = ws.cell(row=2, column=vc)
        val.fill = fill(LIGHT_GOLD)
        val.font = data_font(10)
        val.alignment = center()
        val.border = thin_border()
        val.number_format = currency_fmt()
        if formula:
            val.value = formula
    ws.row_dimensions[2].height = 22

    # ── Row 3: Column Headers
    col_headers = [
        "Category", "Item / Description", "Vendor Name",
        "Estimated Cost (₹)", "Actual Cost (₹)",
        "Paid", "Payment Date", "Notes"
    ]
    for c, h in enumerate(col_headers, 1):
        cell = ws.cell(row=3, column=c, value=h)
        cell.fill = fill(MAROON)
        cell.font = header_font()
        cell.alignment = center()
        cell.border = thin_border()
    ws.row_dimensions[3].height = 28

    # ── Data Rows
    items = CEREMONY_ITEMS.get(ceremony_name, [])
    first_data_row = 4
    for i, (cat, desc) in enumerate(items):
        row = first_data_row + i
        alt = (i % 2 == 0)
        style_data_row(ws, row, 8, alternate=alt)
        ws.cell(row=row, column=1, value=cat)
        ws.cell(row=row, column=2, value=desc)
        ws.cell(row=row, column=3, value="")   # Vendor Name
        ws.cell(row=row, column=4, value=0).number_format = currency_fmt()
        ws.cell(row=row, column=5, value=0).number_format = currency_fmt()
        ws.cell(row=row, column=6, value="No")
        ws.cell(row=row, column=7, value="").number_format = "DD-MMM-YYYY"
        ws.cell(row=row, column=8, value="")
        ws.row_dimensions[row].height = 18

    last_data_row = first_data_row + len(items) - 1

    # ── Now fill summary SUM formulas
    # Total Estimated = SUM(D4:D_last)
    ws["D2"].value = f"=SUM(D{first_data_row}:D{last_data_row})"
    # Total Actual    = SUM(E4:E_last)
    ws["F2"].value  = f"=SUM(E{first_data_row}:E{last_data_row})"
    # Variance        = Allocated - Actual
    ws["H2"].value  = f"='Budget Overview'!C{budget_row}-F2"

    # ── Dropdowns
    paid_range = f"F{first_data_row}:F{last_data_row}"
    add_dv(ws, '"Yes,No"', paid_range)

    # Non-negative on estimated/actual cost columns
    add_num_dv(ws, f"D{first_data_row}:E{last_data_row}")

    freeze(ws, "A4")


# ═══════════════════════════════════════════════════════════════
# SHEET 11 — GUEST LIST & MEAL COUNT
# ═══════════════════════════════════════════════════════════════
def create_guest_sheet(wb):
    ws = wb.create_sheet("Guest List")
    ws.sheet_view.showGridLines = False

    # ── Title
    ws.merge_cells("A1:Q1")
    t = ws["A1"]
    t.value = "Guest List & Meal Count"
    t.font  = title_font(16)
    t.fill  = fill(GOLD)
    t.alignment = center()
    t.border = thick_border()
    ws.row_dimensions[1].height = 38

    # ── Summary table header (row 2)
    summary_cols = ["Ceremony", "Total Guests", "Confirmed", "Pending", "Declined"]
    for c, h in enumerate(summary_cols, 1):
        cell = ws.cell(row=2, column=c, value=h)
        cell.fill = fill(MAROON); cell.font = header_font()
        cell.alignment = center(); cell.border = thin_border()
    ws.row_dimensions[2].height = 22

    ceremony_names = [c[0] for c in CEREMONIES]
    for i, cname in enumerate(ceremony_names, start=3):
        alt = (i % 2 == 0)
        for c in range(1, 6):
            cell = ws.cell(row=i, column=c)
            cell.fill = fill(LIGHT_GOLD if alt else IVORY)
            cell.font = data_font(); cell.border = thin_border()
            cell.alignment = center()
        ws.cell(row=i, column=1, value=cname)
        # Confirmed/Pending/Declined will be filled by user (or formula referencing main table)
        ws.row_dimensions[i].height = 18

    spacer_row = 3 + len(ceremony_names)
    ws.row_dimensions[spacer_row].height = 10

    # ── Main Guest Table headers
    header_row = spacer_row + 1
    guest_headers = [
        "Guest Name", "Relation", "Side",
        "Phone", "Email",
        "Sakarpuda", "Engagement", "Haldi", "Mehendi",
        "Sangeet", "Antrapat", "Jaimal & Phere", "Reception",
        "RSVP Status", "Dietary Preference",
        "Gift Received", "Thank You Sent"
    ]
    for c, h in enumerate(guest_headers, 1):
        cell = ws.cell(row=header_row, column=c, value=h)
        cell.fill = fill(MAROON); cell.font = header_font()
        cell.alignment = center(); cell.border = thin_border()
    ws.row_dimensions[header_row].height = 28

    set_col_widths(ws, [
        (1,22),(2,18),(3,12),(4,14),(5,24),
        (6,12),(7,12),(8,12),(9,12),
        (10,12),(11,12),(12,16),(13,12),
        (14,16),(15,20),(16,14),(17,16)
    ])

    # ── Sample Guests
    sample_guests = [
        ("Priya Sharma",   "Cousin",       "Bride"),
        ("Rahul Gupta",    "Brother",      "Groom"),
        ("Sunita Joshi",   "Aunt",         "Bride"),
        ("Vikram Agarwal", "Uncle",        "Groom"),
        ("Meera Desai",    "Family Friend","Bride"),
    ]

    first_guest_row = header_row + 1
    for i, (name, rel, side) in enumerate(sample_guests):
        row = first_guest_row + i
        alt = (i % 2 == 0)
        for c in range(1, 18):
            cell = ws.cell(row=row, column=c)
            cell.fill = fill(LIGHT_GOLD if alt else IVORY)
            cell.font = data_font(); cell.border = thin_border()
            cell.alignment = left_align()
        ws.cell(row=row, column=1, value=name)
        ws.cell(row=row, column=2, value=rel)
        ws.cell(row=row, column=3, value=side)
        # ceremony columns: Y/N
        for c in range(6, 14):
            ws.cell(row=row, column=c, value="Y")
        ws.cell(row=row, column=14, value="Confirmed")
        ws.cell(row=row, column=15, value="Veg")
        ws.cell(row=row, column=16, value="No")
        ws.cell(row=row, column=17, value="No")
        ws.row_dimensions[row].height = 18

    last_guest_row = first_guest_row + 49  # space for 50 guests

    # ── Dropdowns
    add_dv(ws, '"Bride,Groom,Both"',
           f"C{first_guest_row}:C{last_guest_row}")
    add_dv(ws, '"Confirmed,Pending,Declined"',
           f"N{first_guest_row}:N{last_guest_row}")
    add_dv(ws, '"Veg,Non-Veg,Jain,Vegan"',
           f"O{first_guest_row}:O{last_guest_row}")
    add_dv(ws, '"Yes,No"',
           f"P{first_guest_row}:P{last_guest_row}")
    add_dv(ws, '"Yes,No"',
           f"Q{first_guest_row}:Q{last_guest_row}")
    for col_letter in ["F","G","H","I","J","K","L","M"]:
        add_dv(ws, '"Y,N"',
               f"{col_letter}{first_guest_row}:{col_letter}{last_guest_row}")

    freeze(ws, f"A{header_row+1}")


# ═══════════════════════════════════════════════════════════════
# SHEET 12 — VENDOR TRACKER
# ═══════════════════════════════════════════════════════════════
VENDOR_CATEGORIES = [
    "Venue", "Catering", "Photographer", "Videographer",
    "Decorator", "Mehendi Artist", "Makeup Artist", "DJ",
    "Live Band", "Dhol Player", "Priest/Pandit",
    "Horse/Baraat", "Transport", "Invitation Printer",
    "Choreographer", "Emcee/Anchor", "Cake",
    "Honeymoon Travel Agent"
]

def create_vendor_sheet(wb):
    ws = wb.create_sheet("Vendor Tracker")
    ws.sheet_view.showGridLines = False

    # Title
    ws.merge_cells("A1:P1")
    t = ws["A1"]
    t.value = "Vendor Tracker"
    t.font  = title_font(16)
    t.fill  = fill(GOLD)
    t.alignment = center()
    t.border = thick_border()
    ws.row_dimensions[1].height = 38

    # Headers
    headers = [
        "Vendor Category", "Vendor Name", "Contact Person",
        "Phone", "Email", "Service Description", "Ceremony",
        "Quote (₹)", "Negotiated (₹)", "Advance Paid (₹)",
        "Balance Due (₹)", "Contract Signed",
        "Payment Due Date", "Status", "Rating (1–5)", "Notes"
    ]
    for c, h in enumerate(headers, 1):
        cell = ws.cell(row=2, column=c, value=h)
        cell.fill = fill(MAROON); cell.font = header_font()
        cell.alignment = center(); cell.border = thin_border()
    ws.row_dimensions[2].height = 28

    set_col_widths(ws, [
        (1,22),(2,22),(3,18),(4,14),(5,24),
        (6,28),(7,18),(8,16),(9,16),(10,16),
        (11,16),(12,16),(13,18),(14,16),(15,14),(16,28)
    ])

    first_data = 3
    for i, cat in enumerate(VENDOR_CATEGORIES):
        row = first_data + i
        alt = (i % 2 == 0)
        for c in range(1, 17):
            cell = ws.cell(row=row, column=c)
            cell.fill = fill(LIGHT_GOLD if alt else IVORY)
            cell.font = data_font(); cell.border = thin_border()
            cell.alignment = left_align()

        ws.cell(row=row, column=1,  value=cat)
        ws.cell(row=row, column=8).number_format  = currency_fmt()
        ws.cell(row=row, column=9).number_format  = currency_fmt()
        ws.cell(row=row, column=10).number_format = currency_fmt()

        # Balance Due = Negotiated - Advance Paid
        bd = ws.cell(row=row, column=11,
                     value=f"=IF(I{row}=\"\",0,I{row}-J{row})")
        bd.number_format = currency_fmt()

        ws.cell(row=row, column=12, value="No")
        ws.cell(row=row, column=13).number_format = "DD-MMM-YYYY"
        ws.cell(row=row, column=14, value="Shortlisted")
        ws.row_dimensions[row].height = 18

    last_data = first_data + len(VENDOR_CATEGORIES) - 1

    # Dropdowns
    cat_list = ",".join(VENDOR_CATEGORIES)
    add_dv(ws, f'"{cat_list}"',
           f"A{first_data}:A{last_data+20}")
    add_dv(ws, '"Booked,Negotiating,Shortlisted,Cancelled"',
           f"N{first_data}:N{last_data+20}")
    add_dv(ws, '"Yes,No"',
           f"L{first_data}:L{last_data+20}")

    ceremony_list = ",".join([c[0] for c in CEREMONIES])
    add_dv(ws, f'"{ceremony_list}"',
           f"G{first_data}:G{last_data+20}")

    add_num_dv(ws, f"H{first_data}:J{last_data+20}")

    freeze(ws, "A3")


# ═══════════════════════════════════════════════════════════════
# SHEET 13 — PAYMENT & CASH FLOW TRACKER
# ═══════════════════════════════════════════════════════════════
def create_payment_sheet(wb):
    ws = wb.create_sheet("Payment Tracker")
    ws.sheet_view.showGridLines = False

    # Title
    ws.merge_cells("A1:I1")
    t = ws["A1"]
    t.value = "Payment & Cash Flow Tracker"
    t.font  = title_font(16)
    t.fill  = fill(GOLD)
    t.alignment = center()
    t.border = thick_border()
    ws.row_dimensions[1].height = 38

    headers = [
        "Date", "Vendor / Payee", "Ceremony", "Category",
        "Amount (₹)", "Payment Mode", "Paid By",
        "Receipt No.", "Notes"
    ]
    for c, h in enumerate(headers, 1):
        cell = ws.cell(row=2, column=c, value=h)
        cell.fill = fill(MAROON); cell.font = header_font()
        cell.alignment = center(); cell.border = thin_border()
    ws.row_dimensions[2].height = 28

    set_col_widths(ws, [
        (1,14),(2,26),(3,18),(4,18),
        (5,16),(6,18),(7,20),(8,14),(9,28)
    ])

    # 50 blank data rows
    first_data = 3
    last_data  = first_data + 49
    for i in range(first_data, last_data + 1):
        alt = (i % 2 == 0)
        for c in range(1, 10):
            cell = ws.cell(row=i, column=c)
            cell.fill = fill(LIGHT_GOLD if alt else IVORY)
            cell.font = data_font(); cell.border = thin_border()
            cell.alignment = left_align()
        ws.cell(row=i, column=1).number_format = "DD-MMM-YYYY"
        ws.cell(row=i, column=5).number_format = currency_fmt()
        ws.row_dimensions[i].height = 18

    # Running Total
    total_row = last_data + 1
    ws.cell(row=total_row, column=4, value="TOTAL").font = header_font(11, MAROON)
    ws.cell(row=total_row, column=4).fill = fill(GOLD)
    ws.cell(row=total_row, column=4).border = thin_border()
    ws.cell(row=total_row, column=4).alignment = center()

    tot_cell = ws.cell(row=total_row, column=5,
                       value=f"=SUM(E{first_data}:E{last_data})")
    tot_cell.number_format = currency_fmt()
    tot_cell.fill = fill(GOLD)
    tot_cell.font = header_font(11, MAROON)
    tot_cell.border = thin_border()
    tot_cell.alignment = center()
    ws.row_dimensions[total_row].height = 22

    # ── Monthly Summary
    monthly_start = total_row + 3
    ws.merge_cells(
        start_row=monthly_start-1, start_column=1,
        end_row=monthly_start-1,   end_column=3
    )
    ms_title = ws.cell(row=monthly_start-1, column=1,
                       value="Monthly Summary")
    ms_title.font = header_font(12, MAROON)
    ms_title.fill = fill(GOLD)
    ms_title.alignment = center()
    ms_title.border = thin_border()
    ws.row_dimensions[monthly_start-1].height = 24

    for c, h in enumerate(["Month","Total Spent (₹)","No. of Transactions"], 1):
        cell = ws.cell(row=monthly_start, column=c, value=h)
        cell.fill = fill(MAROON); cell.font = header_font()
        cell.alignment = center(); cell.border = thin_border()

    import datetime
    months = ["January","February","March","April","May","June",
              "July","August","September","October","November","December"]
    for i, month in enumerate(months, start=monthly_start+1):
        alt = (i % 2 ==
