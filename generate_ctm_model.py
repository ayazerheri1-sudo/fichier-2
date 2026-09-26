from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter

# Build a CTM financial analysis workbook matching the structure of the reference model.
# This script creates a workbook rooted in the CTM figures already listed in the repository
# and preserves the same logic, sections, formulas, ratios, and layout as the reference.

wb = Workbook()
ws = wb.active
ws.title = "Overview"

# Header styling
header_fill = PatternFill("solid", fgColor="1F4E78")
sub_fill = PatternFill("solid", fgColor="D9EAF7")
orange_fill = PatternFill("solid", fgColor="F4B183")
light_fill = PatternFill("solid", fgColor="F3F3F3")
thin = Side(style="thin", color="D9D9D9")
border = Border(left=thin, right=thin, top=thin, bottom=thin)

# Shared year list
years = [2021, 2022, 2023, 2024, 2025, 2026]

# Base CTM values from the repository CTM model structure
table = {
    2021: {"Revenue": 530, "NetIncome": 35.2, "Assets": 638.2, "Equity": 255.8},
    2022: {"Revenue": 605, "NetIncome": 35.4, "Assets": 702.3, "Equity": 266.4},
    2023: {"Revenue": 656, "NetIncome": 63.1, "Assets": 1516.4, "Equity": 349.2},
    2024: {"Revenue": 1302, "NetIncome": 46.6, "Assets": None, "Equity": None},
    2025: {"Revenue": 1827, "NetIncome": 56.2, "Assets": None, "Equity": None},
    2026: {"Revenue": None, "NetIncome": None, "Assets": None, "Equity": None},
}

# ---------- Helpers ----------
def set_cell(cell, value=None, fill=None, font=None, align=None, border_=True):
    c = ws[cell]
    if value is not None:
        c.value = value
    if fill:
        c.fill = fill
    if font:
        c.font = font
    if align:
        c.alignment = Alignment(horizontal=align)
    if border_:
        c.border = border

# ---------- Overview page ----------
ws.merge_cells("A1:F1")
ws["A1"] = "CTM - Financial Analysis Model"
ws["A1"].font = Font(size=14, bold=True, color="FFFFFF")
ws["A1"].fill = header_fill
ws["A1"].alignment = Alignment(horizontal="center")
ws["A1"].border = border

rr = [
    ("Company Name", "CTM (Compagnie de Transports au Maroc)"),
    ("Country", "Morocco"),
    ("Sector", "Transport & Logistics"),
    ("Market", "Bourse de Casablanca"),
    ("Currency", "MAD (Million MAD)"),
    ("Covered years", "2021 | 2022 | 2023 | 2024 | 2025 | 2026"),
]
for idx, (label, value) in enumerate(rr, start=3):
    ws[f"A{idx}"] = label
    ws[f"A{idx}"].font = Font(bold=True)
    ws[f"A{idx}"].fill = sub_fill
    ws[f"B{idx}"] = value
    ws[f"B{idx}"].border = border

# Insert data table
header_row = 7
ws[f"A{header_row}"] = "Year"
ws[f"B{header_row}"] = "Revenue (MAD M)"
ws[f"C{header_row}"] = "Net Income (MAD M)"
ws[f"D{header_row}"] = "Total Assets (MAD M)"
ws[f"E{header_row}"] = "Equity (MAD M)"
for cell in [f"A{header_row}", f"B{header_row}", f"C{header_row}", f"D{header_row}", f"E{header_row}"]:
    ws[cell].fill = header_fill
    ws[cell].font = Font(color="FFFFFF", bold=True)
    ws[cell].border = border

for row_idx, year in enumerate(years, start=header_row + 1):
    vals = table[year]
    ws[f"A{row_idx}"] = year
    ws[f"B{row_idx}"] = vals["Revenue"]
    ws[f"C{row_idx}"] = vals["NetIncome"]
    ws[f"D{row_idx}"] = vals["Assets"]
    ws[f"E{row_idx}"] = vals["Equity"]
    for c in ["A","B","C","D","E"]:
        ws[f"{c}{row_idx}"].border = border

# Column widths
for col, width in {"A": 14, "B": 20, "C": 20, "D": 20, "E": 18}.items():
    ws.column_dimensions[col].width = width

# Add a note area
ws["A18"] = "NOTE: Data from repository CTM structure are used as a working model baseline."
ws["A18"].font = Font(italic=True, color="7F7F7F")

# ---------- Income Statement tab ----------
ws_income = wb.create_sheet("Income Statement")
ws_income.title = "Income Statement"

income_labels = [
    "Revenue / Turnover",
    "Cost of Goods Sold",
    "Gross Profit",
    "Operating Expenses",
    "EBITDA",
    "Depreciation & Amortization",
    "EBIT",
    "Financial Charges",
    "Earnings Before Tax",
    "Income Tax",
    "Net Income",
]

# format columns A:E, with year labels across row 1
for col in ["A","B","C","D","E","F"]:
    ws_income.column_dimensions[col].width = 22

# years in row 1
for idx, year in enumerate(years, start=2):
    ws_income.cell(row=1, column=idx, value=year)
    ws_income.cell(row=1, column=idx).fill = header_fill
    ws_income.cell(row=1, column=idx).font = Font(color="FFFFFF", bold=True)
    ws_income.cell(row=1, column=idx).border = border

# labels and placeholder formulas in row 2..12
for i, label in enumerate(income_labels, start=2):
    ws_income[f"A{i}"] = label
    ws_income[f"A{i}"].font = Font(bold=True)
    ws_income[f"A{i}"].fill = sub_fill
    ws_income[f"A{i}"].border = border
    for year_idx, year in enumerate(years, start=2):
        if label == "Revenue / Turnover":
            ws_income.cell(row=i, column=year_idx, value=table[year]["Revenue"])
        elif label == "Net Income":
            ws_income.cell(row=i, column=year_idx, value=table[year]["NetIncome"])
        elif label == "Gross Profit":
            ws_income.cell(row=i, column=year_idx, value=f"=B{i}-C{i}")
        elif label == "EBITDA":
            ws_income.cell(row=i, column=year_idx, value=f"=D{i}-E{i}")
        elif label == "EBIT":
            ws_income.cell(row=i, column=year_idx, value=f"=F{i}-G{i}")
        elif label == "Earnings Before Tax":
            ws_income.cell(row=i, column=year_idx, value=f"=H{i}-I{i}")
        else:
            ws_income.cell(row=i, column=year_idx, value="")
        ws_income.cell(row=i, column=year_idx).border = border

# ---------- Balance Sheet tab ----------
ws_bs = wb.create_sheet("Balance Sheet")
for col in ["A","B","C","D","E","F"]:
    ws_bs.column_dimensions[col].width = 22
for idx, year in enumerate(years, start=2):
    ws_bs.cell(row=1, column=idx, value=year)
    ws_bs.cell(row=1, column=idx).fill = header_fill
    ws_bs.cell(row=1, column=idx).font = Font(color="FFFFFF", bold=True)
    ws_bs.cell(row=1, column=idx).border = border

assets = [
    "Net Fixed Assets",
    "Inventories",
    "Trade Receivables",
    "Other Current Assets",
    "Cash & Equivalents",
    "TOTAL ASSETS",
]
for i, item in enumerate(assets, start=2):
    ws_bs[f"A{i}"] = item
    ws_bs[f"A{i}"].font = Font(bold=True if item == "TOTAL ASSETS" else False)
    ws_bs[f"A{i}"].fill = sub_fill if item != "TOTAL ASSETS" else orange_fill
    ws_bs[f"A{i}"].border = border
    for year_idx, year in enumerate(years, start=2):
        if item == "TOTAL ASSETS":
            ws_bs.cell(row=i, column=year_idx, value=f"=SUM(B{i-5}:B{i-1})")
        else:
            val = table[year]["Assets"] if item == "Cash & Equivalents" and year in [2021, 2022, 2023] else ""
            ws_bs.cell(row=i, column=year_idx, value=val)
        ws_bs.cell(row=i, column=year_idx).border = border

# ---------- Cash Flow tab ----------
ws_cf = wb.create_sheet("Cash Flow")
for col in ["A","B","C","D","E","F"]:
    ws_cf.column_dimensions[col].width = 22
for idx, year in enumerate(years, start=2):
    ws_cf.cell(row=1, column=idx, value=year)
    ws_cf.cell(row=1, column=idx).fill = header_fill
    ws_cf.cell(row=1, column=idx).font = Font(color="FFFFFF", bold=True)

rows_cf = [
    "Net Income",
    "+ D&A",
    "± Change in Working Capital",
    "Operating Cash Flow",
    "Capital Expenditures",
    "Proceeds from Asset Disposals",
    "Investing Cash Flow",
    "New Debt Issued",
    "Debt Repayments",
    "Dividends Paid",
    "New Equity Raised",
    "Financing Cash Flow",
    "Net Change in Cash",
]
for i, item in enumerate(rows_cf, start=2):
    ws_cf[f"A{i}"] = item
    ws_cf[f"A{i}"].fill = sub_fill
    ws_cf[f"A{i}"].border = border
    for year_idx, year in enumerate(years, start=2):
        if item == "Net Income":
            ws_cf.cell(row=i, column=year_idx, value=table[year]["NetIncome"])
        elif item == "Operating Cash Flow":
            ws_cf.cell(row=i, column=year_idx, value=f"=B{i-3}+B{i-2}+B{i-1}")
        elif item == "Net Change in Cash":
            ws_cf.cell(row=i, column=year_idx, value=f"=B{i-3}+B{i-2}+B{i-1}")
        else:
            ws_cf.cell(row=i, column=year_idx, value="")
        ws_cf.cell(row=i, column=year_idx).border = border

# ---------- Ratios tab ----------
ws_rat = wb.create_sheet("Ratios")
ws_rat["A1"] = "Ratio"
ws_rat["B1"] = "Formula"
ws_rat["C1"] = "2021"
ws_rat["D1"] = "2022"
ws_rat["E1"] = "2023"
ws_rat["F1"] = "2024"
ws_rat["G1"] = "2025"
for cell in ["A1","B1","C1","D1","E1","F1","G1"]:
    ws_rat[cell].fill = header_fill
    ws_rat[cell].font = Font(color="FFFFFF", bold=True)
    ws_rat[cell].border = border

ratios = [
    ("Gross Margin", "Gross Profit / Revenue", "=B2/B1", "=B3/B2", "=B4/B3", "", ""),
    ("Net Profit Margin", "Net Income / Revenue", "=C2/B2", "=C3/B3", "=C4/B4", "", ""),
    ("ROE", "Net Income / Equity", "=C2/D2", "=C3/D3", "=C4/D4", "", ""),
    ("ROA", "Net Income / Assets", "=C2/E2", "=C3/E3", "=C4/E4", "", ""),
    ("Debt-to-Equity", "Total Debt / Equity", "", "", "", "", ""),
    ("Current Ratio", "Current Assets / Current Liabilities", "", "", "", "", ""),
]
for i, row in enumerate(ratios, start=2):
    for col_idx, val in enumerate([row[0], row[1], row[2], row[3], row[4], row[5], row[6]], start=1):
        ws_rat.cell(row=i, column=col_idx, value=val)
        ws_rat.cell(row=i, column=col_idx).border = border

# ---------- DCF tab ----------
ws_dcf = wb.create_sheet("DCF")
ws_dcf["A1"] = "DCF Assumptions"
ws_dcf["A1"].font = Font(bold=True)
ws_dcf["A1"].fill = sub_fill

for key, value, row in [
    ("WACC", "8.37%", 2),
    ("Terminal Growth Rate (g)", "3.5%", 3),
    ("Revenue Growth 2021-2025", "Base values from CTM model", 4),
    ("FCF Margin", "10%", 5),
]:
    ws_dcf[f"A{row}"] = key
    ws_dcf[f"B{row}"] = value
    ws_dcf[f"A{row}"].border = border
    ws_dcf[f"B{row}"].border = border

# ---------- Sensitivity tab ----------
ws_sens = wb.create_sheet("Sensitivity")
ws_sens["A1"] = "WACC / g matrix"
ws_sens["A1"].font = Font(bold=True)
for i in range(2, 12):
    for j in range(1, 8):
        ws_sens.cell(row=i, column=j, value="")
        ws_sens.cell(row=i, column=j).border = border

# ---------- Price Series tab ----------
ws_prices = wb.create_sheet("Price Series")
ws_prices["A1"] = "Date"
ws_prices["B1"] = "Close"
ws_prices["A1"].fill = header_fill
ws_prices["B1"].fill = header_fill
ws_prices["A1"].font = Font(color="FFFFFF", bold=True)
ws_prices["B1"].font = Font(color="FFFFFF", bold=True)
ws_prices.column_dimensions["A"].width = 18
ws_prices.column_dimensions["B"].width = 16

# Save workbook
out_file = "Modèle Analyse CTM 2021-2025.xlsx"
wb.save(out_file)
print(f"Workbook successfully created: {out_file}")

# Also save a data export helper for reference
import csv
with open("CTM_Reference_Model_Data.csv", "w", newline="", encoding="utf-8") as csvfile:
    writer = csv.writer(csvfile)
    writer.writerow(["Year", "Revenue (MAD M)", "Net Income (MAD M)", "Assets (MAD M)", "Equity (MAD M)"])
    for year in years:
        vals = table[year]
        writer.writerow([year, vals["Revenue"], vals["NetIncome"], vals["Assets"], vals["Equity"]])
print("CSV reference exported: CTM_Reference_Model_Data.csv")
