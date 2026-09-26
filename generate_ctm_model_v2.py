#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
CTM Financial Analysis Model v2.0
Reproduces the Akdital reference model structure with enhanced formulas, 
ratios, DCF analysis, and professional formatting.

Generates: Modèle Analyse CTM 2021-2025_FINAL.xlsx
"""

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment, numbers
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

# ==================== CONFIGURATION ====================

# Color scheme (matching Akdital reference)
HEADER_COLOR = "1F4E78"        # Dark blue
SUB_HEADER_COLOR = "D9EAF7"    # Light blue
HIGHLIGHT_COLOR = "F4B183"     # Orange
LIGHT_GRAY = "F3F3F3"
WHITE = "FFFFFF"

# Border style
thin_border = Border(
    left=Side(style="thin", color="D9D9D9"),
    right=Side(style="thin", color="D9D9D9"),
    top=Side(style="thin", color="D9D9D9"),
    bottom=Side(style="thin", color="D9D9D9")
)

# Font styles
font_header = Font(name="Arial", size=11, bold=True, color=WHITE)
font_title = Font(name="Arial", size=14, bold=True, color=WHITE)
font_section = Font(name="Arial", size=11, bold=True)
font_normal = Font(name="Arial", size=10)
font_label = Font(name="Arial", size=10, bold=True)

# Years covered
years = [2021, 2022, 2023, 2024, 2025, 2026]

# CTM base data (from repository CTM model structure)
# All values in MAD Millions
ctm_data = {
    2021: {
        "Revenue": 530,
        "COGS": 265,  # Estimated ~50%
        "OpEx": 106,  # Estimated ~20%
        "DA": 5.3,
        "FinCharges": 8,
        "Tax": 28,
        "NetIncome": 35.2,
        "Assets": 638.2,
        "Equity": 255.8,
        "Debt": 150,
        "Cash": 50,
    },
    2022: {
        "Revenue": 605,
        "COGS": 302.5,
        "OpEx": 121,
        "DA": 6,
        "FinCharges": 10,
        "Tax": 32,
        "NetIncome": 35.4,
        "Assets": 702.3,
        "Equity": 266.4,
        "Debt": 165,
        "Cash": 55,
    },
    2023: {
        "Revenue": 656,
        "COGS": 328,
        "OpEx": 131,
        "DA": 6.5,
        "FinCharges": 12,
        "Tax": 45,
        "NetIncome": 63.1,
        "Assets": 1516.4,
        "Equity": 349.2,
        "Debt": 320,
        "Cash": 180,
    },
    2024: {
        "Revenue": 1302,
        "COGS": 651,
        "OpEx": 260,
        "DA": 13,
        "FinCharges": 25,
        "Tax": 120,
        "NetIncome": 46.6,
        "Assets": None,
        "Equity": None,
        "Debt": 400,
        "Cash": 200,
    },
    2025: {
        "Revenue": 1827,
        "COGS": 913,
        "OpEx": 365,
        "DA": 18,
        "FinCharges": 35,
        "Tax": 175,
        "NetIncome": 56.2,
        "Assets": None,
        "Equity": None,
        "Debt": 450,
        "Cash": 250,
    },
    2026: {
        "Revenue": None,
        "COGS": None,
        "OpEx": None,
        "DA": None,
        "FinCharges": None,
        "Tax": None,
        "NetIncome": None,
        "Assets": None,
        "Equity": None,
        "Debt": None,
        "Cash": None,
    },
}

# ==================== HELPER FUNCTIONS ====================

def set_cell_style(cell, value=None, fill=None, font=None, alignment_h="left", 
                   alignment_v="center", number_format=None, border=True):
    """Apply styling to a cell."""
    if value is not None:
        cell.value = value
    if fill:
        cell.fill = PatternFill(start_color=fill, end_color=fill, fill_type="solid")
    if font:
        cell.font = font
    cell.alignment = Alignment(horizontal=alignment_h, vertical=alignment_v, wrap_text=False)
    if number_format:
        cell.number_format = number_format
    if border:
        cell.border = thin_border

def add_section_title(ws, row, title, start_col="A", end_col="G"):
    """Add a section title spanning multiple columns."""
    ws.merge_cells(f"{start_col}{row}:{end_col}{row}")
    cell = ws[f"{start_col}{row}"]
    set_cell_style(cell, value=title, fill=HEADER_COLOR, font=font_title, alignment_h="center")
    ws.row_dimensions[row].height = 25

def add_header_row(ws, row, headers, start_col="A"):
    """Add a header row with styling."""
    for idx, header in enumerate(headers):
        col = get_column_letter(ord(start_col) - ord("A") + 1 + idx)
        cell = ws[f"{col}{row}"]
        set_cell_style(cell, value=header, fill=HEADER_COLOR, font=font_header, alignment_h="center")

def format_number(value, decimals=0):
    """Format number with proper locale."""
    if value is None:
        return ""
    if isinstance(value, str):
        return value
    if decimals == 0:
        return f"{int(value):,}"
    return f"{value:,.{decimals}f}"

# ==================== CREATE WORKBOOK ====================

wb = Workbook()
wb.remove(wb.active)  # Remove default sheet

# ==================== SHEET 1: ACCUEIL (OVERVIEW) ====================

ws_accueil = wb.create_sheet("ACCUEIL", 0)
ws_accueil.column_dimensions["A"].width = 30
for col in ["B", "C", "D", "E", "F", "G"]:
    ws_accueil.column_dimensions[col].width = 18

# Title
add_section_title(ws_accueil, 1, "CTM STOCK FINANCIAL ANALYSIS", "A", "G")

# Company info
info_data = [
    ("Ticker / Company Name", "CTM - Compagnie de Transports au Maroc"),
    ("Country", "Morocco"),
    ("Sector", "Transport & Logistics"),
    ("Market", "Bourse de Casablanca (BVC)"),
    ("Currency", "MAD (Million Dirhams)"),
    ("Covered Period", "2021 | 2022 | 2023 | 2024 | 2025 | 2026"),
]

for idx, (label, value) in enumerate(info_data, start=3):
    ws_accueil[f"A{idx}"] = label
    set_cell_style(ws_accueil[f"A{idx}"], font=font_section, fill=SUB_HEADER_COLOR)
    ws_accueil[f"B{idx}"] = value
    set_cell_style(ws_accueil[f"B{idx}"], font=font_normal)

# Key figures table
key_row = 11
add_section_title(ws_accueil, key_row, "KEY FIGURES", "A", "G")

header_row = key_row + 1
headers = ["Year"] + [str(y) for y in years]
add_header_row(ws_accueil, header_row, headers)

metrics = [
    ("Revenue (MAD M)", "Revenue"),
    ("Net Income (MAD M)", "NetIncome"),
    ("Total Assets (MAD M)", "Assets"),
    ("Equity (MAD M)", "Equity"),
    ("Debt (MAD M)", "Debt"),
    ("Cash (MAD M)", "Cash"),
]

for metric_idx, (label, key) in enumerate(metrics, start=header_row + 1):
    ws_accueil[f"A{metric_idx}"] = label
    set_cell_style(ws_accueil[f"A{metric_idx}"], font=font_label, fill=SUB_HEADER_COLOR)
    
    for year_idx, year in enumerate(years, start=2):
        col = get_column_letter(year_idx)
        val = ctm_data[year].get(key)
        if val:
            cell = ws_accueil[f"{col}{metric_idx}"]
            set_cell_style(cell, value=val, number_format="#,##0.0", font=font_normal, alignment_h="right")

# ==================== SHEET 2: INCOME STATEMENT ====================

ws_income = wb.create_sheet("INCOME STATEMENT", 1)
ws_income.column_dimensions["A"].width = 35
for col in ["B", "C", "D", "E", "F", "G"]:
    ws_income.column_dimensions[col].width = 16

add_section_title(ws_income, 1, "INCOME STATEMENT (Compte de Résultat)", "A", "G")

# Headers
header_row = 3
headers = ["Metric (Métrique)"] + [str(y) for y in years]
for idx, header in enumerate(headers):
    col = get_column_letter(idx + 1)
    cell = ws_income[f"{col}{header_row}"]
    set_cell_style(cell, value=header, fill=HEADER_COLOR, font=font_header, alignment_h="center")

# Income statement rows
income_items = [
    ("Revenue / Turnover (CA HT)", "Revenue"),
    ("Cost of Goods Sold (Coût des ventes)", "COGS"),
    ("Gross Profit (Marge brute)", None),  # Formula
    ("Operating Expenses (Charges d'exploitation)", "OpEx"),
    ("EBITDA (EBE)", None),  # Formula
    ("Depreciation & Amortization (DAP)", "DA"),
    ("EBIT (Résultat d'exploitation)", None),  # Formula
    ("Financial Charges (Intérêts)", "FinCharges"),
    ("Earnings Before Tax (RCAI)", None),  # Formula
    ("Income Tax (IS)", "Tax"),
    ("Net Income (Résultat net)", "NetIncome"),
]

for item_idx, (label, key) in enumerate(income_items, start=header_row + 1):
    ws_income[f"A{item_idx}"] = label
    
    # Color alternating rows
    if item_idx % 2 == 0:
        set_cell_style(ws_income[f"A{item_idx}"], font=font_label, fill=LIGHT_GRAY)
    else:
        set_cell_style(ws_income[f"A{item_idx}"], font=font_label)
    
    # Total rows highlight
    if "Gross Profit" in label or "EBITDA" in label or "EBIT" in label or "Net Income" in label:
        set_cell_style(ws_income[f"A{item_idx}"], fill=HIGHLIGHT_COLOR, font=Font(name="Arial", size=10, bold=True))
    
    # Values
    for year_idx, year in enumerate(years, start=2):
        col = get_column_letter(year_idx)
        cell = ws_income[f"{col}{item_idx}"]
        
        if key:
            val = ctm_data[year].get(key)
            if val:
                set_cell_style(cell, value=val, number_format="#,##0.0", alignment_h="right")
        else:
            # Formulas for calculated rows
            if "Gross Profit" in label:
                # Revenue - COGS
                cell.value = f"={get_column_letter(year_idx)}{header_row+1}-{get_column_letter(year_idx)}{header_row+2}"
            elif "EBITDA" in label:
                # Gross Profit - OpEx
                cell.value = f"={get_column_letter(year_idx)}{header_row+3}-{get_column_letter(year_idx)}{header_row+4}"
            elif "EBIT" in label and "Earnings" not in label:
                # EBITDA - DA
                cell.value = f"={get_column_letter(year_idx)}{header_row+5}-{get_column_letter(year_idx)}{header_row+6}"
            elif "Earnings Before Tax" in label:
                # EBIT - FinCharges
                cell.value = f"={get_column_letter(year_idx)}{header_row+7}-{get_column_letter(year_idx)}{header_row+8}"
            elif "Net Income" in label:
                # EBT - Tax
                ebt_row = header_row + 9
                tax_row = header_row + 10
                cell.value = f"={get_column_letter(year_idx)}{ebt_row}-{get_column_letter(year_idx)}{tax_row}"
            
            set_cell_style(cell, number_format="#,##0.0", alignment_h="right")

# ==================== SHEET 3: BALANCE SHEET ====================

ws_bs = wb.create_sheet("BALANCE SHEET", 2)
ws_bs.column_dimensions["A"].width = 35
for col in ["B", "C", "D", "E", "F", "G"]:
    ws_bs.column_dimensions[col].width = 16

add_section_title(ws_bs, 1, "BALANCE SHEET (Bilan)", "A", "G")

# Headers
header_row = 3
headers = ["Metric (Métrique)"] + [str(y) for y in years]
add_header_row(ws_bs, header_row, headers)

# Assets
bs_items = [
    ("ASSETS (ACTIF)", None, True),
    ("Net Fixed Assets (Immobilisations nettes)", None, False),
    ("Current Assets (Actif circulant)", None, False),
    ("Cash & Equivalents (Trésorerie)", "Cash", False),
    ("TOTAL ASSETS (Total Actif)", "Assets", True),
    ("", None, False),
    ("LIABILITIES & EQUITY (PASSIF & CAPITAUX PROPRES)", None, True),
    ("Total Equity (Capitaux propres)", "Equity", False),
    ("Long-term Debt (Dettes LT)", None, False),
    ("Short-term Debt (Dettes CT)", None, False),
    ("Total Debt (Total Dettes)", "Debt", False),
    ("TOTAL LIABILITIES & EQUITY", None, True),
]

for item_idx, (label, key, is_section) in enumerate(bs_items, start=header_row + 1):
    ws_bs[f"A{item_idx}"] = label
    
    if is_section:
        set_cell_style(ws_bs[f"A{item_idx}"], font=Font(name="Arial", size=10, bold=True), fill=HIGHLIGHT_COLOR)
    else:
        set_cell_style(ws_bs[f"A{item_idx}"], font=font_label, fill=SUB_HEADER_COLOR if not label else LIGHT_GRAY)
    
    for year_idx, year in enumerate(years, start=2):
        col = get_column_letter(year_idx)
        cell = ws_bs[f"{col}{item_idx}"]
        
        if key:
            val = ctm_data[year].get(key)
            if val:
                set_cell_style(cell, value=val, number_format="#,##0.0", alignment_h="right")
        else:
            set_cell_style(cell, alignment_h="right")

# ==================== SHEET 4: CASH FLOW ====================

ws_cf = wb.create_sheet("CASH FLOW", 3)
ws_cf.column_dimensions["A"].width = 35
for col in ["B", "C", "D", "E", "F", "G"]:
    ws_cf.column_dimensions[col].width = 16

add_section_title(ws_cf, 1, "CASH FLOW STATEMENT (Tableau de Trésorerie)", "A", "G")

header_row = 3
headers = ["Metric (Métrique)"] + [str(y) for y in years]
add_header_row(ws_cf, header_row, headers)

cf_items = [
    ("OPERATING ACTIVITIES (Exploitation)", None, True),
    ("Net Income (Résultat net)", "NetIncome", False),
    ("+ D&A (+ Amortissements)", "DA", False),
    ("± Change in Working Capital (± ΔBFR)", None, False),
    ("Operating Cash Flow (CF Exploitation)", None, True),
    ("", None, False),
    ("INVESTING ACTIVITIES (Investissement)", None, True),
    ("Capital Expenditures (CAPEX)", None, False),
    ("Investing Cash Flow (CF Investissement)", None, True),
    ("", None, False),
    ("FINANCING ACTIVITIES (Financement)", None, True),
    ("New Debt (Nouveaux emprunts)", None, False),
    ("Debt Repayments (Remboursements)", None, False),
    ("Dividends Paid (Dividendes versés)", None, False),
    ("Financing Cash Flow (CF Financement)", None, True),
    ("", None, False),
    ("NET CHANGE IN CASH (Variation nette)", None, True),
]

for item_idx, (label, key, is_section) in enumerate(cf_items, start=header_row + 1):
    ws_cf[f"A{item_idx}"] = label
    
    if is_section:
        set_cell_style(ws_cf[f"A{item_idx}"], font=Font(name="Arial", size=10, bold=True), fill=HIGHLIGHT_COLOR)
    else:
        set_cell_style(ws_cf[f"A{item_idx}"], font=font_label, fill=SUB_HEADER_COLOR if not label else LIGHT_GRAY)
    
    for year_idx, year in enumerate(years, start=2):
        col = get_column_letter(year_idx)
        cell = ws_cf[f"{col}{item_idx}"]
        
        if key:
            val = ctm_data[year].get(key)
            if val:
                set_cell_style(cell, value=val, number_format="#,##0.0", alignment_h="right")
        else:
            set_cell_style(cell, alignment_h="right")

# ==================== SHEET 5: RATIOS & INDICATORS ====================

ws_ratios = wb.create_sheet("RATIOS", 4)
ws_ratios.column_dimensions["A"].width = 40
ws_ratios.column_dimensions["B"].width = 45
for col in ["C", "D", "E", "F", "G", "H"]:
    ws_ratios.column_dimensions[col].width = 14

add_section_title(ws_ratios, 1, "RATIOS & FINANCIAL INDICATORS", "A", "H")

# Profitability Ratios
ratios_section = [
    ("PROFITABILITY RATIOS (Rentabilité)", True),
    ("Gross Margin", "Marge brute = (Gross Profit / Revenue) × 100%", [
        f"=IF(B4=0,0,(B4-B5)/B4*100)" if i == 0 else "" for i in range(6)
    ]),
    ("EBITDA Margin", "Marge EBITDA = (EBITDA / Revenue) × 100%", [
        f"=IF(B4=0,0,B7/B4*100)" if i == 0 else "" for i in range(6)
    ]),
    ("EBIT Margin", "Marge d'exploitation = (EBIT / Revenue) × 100%", [
        f"=IF(B4=0,0,B9/B4*100)" if i == 0 else "" for i in range(6)
    ]),
    ("Net Profit Margin", "Marge nette = (Net Income / Revenue) × 100%", [
        f"=IF(B4=0,0,B13/B4*100)" if i == 0 else "" for i in range(6)
    ]),
    ("ROE", "Rendement des fonds propres = (Net Income / Equity) × 100%", [
        f"=IF(B18=0,0,B13/B18*100)" if i == 0 else "" for i in range(6)
    ]),
    ("ROA", "Rentabilité de l'actif = (Net Income / Total Assets) × 100%", [
        f"=IF(B22=0,0,B13/B22*100)" if i == 0 else "" for i in range(6)
    ]),
]

row = 3
for item in ratios_section:
    if len(item) == 2:  # Section header
        ws_ratios[f"A{row}"] = item[0]
        set_cell_style(ws_ratios[f"A{row}"], font=Font(name="Arial", size=11, bold=True), fill=SUB_HEADER_COLOR)
    else:  # Ratio
        ws_ratios[f"A{row}"] = item[0]
        set_cell_style(ws_ratios[f"A{row}"], font=font_normal)
        ws_ratios[f"B{row}"] = item[1]
        set_cell_style(ws_ratios[f"B{row}"], font=font_normal, alignment_h="left")
    row += 1

# Financial Structure Ratios
row += 1
ws_ratios[f"A{row}"] = "FINANCIAL STRUCTURE RATIOS (Structure financière)"
set_cell_style(ws_ratios[f"A{row}"], font=Font(name="Arial", size=11, bold=True), fill=SUB_HEADER_COLOR)
row += 1

structure_ratios = [
    ("Debt-to-Equity Ratio", "Ratio d'endettement = (Total Debt / Equity)"),
    ("Debt-to-Assets Ratio", "Levier financier = (Total Debt / Total Assets)"),
    ("Equity Ratio", "Ratio d'autonomie = (Equity / Total Assets)"),
    ("Interest Coverage", "Couverture des intérêts = (EBIT / Financial Charges)"),
    ("Net Debt", "Dette nette = (Total Debt - Cash)"),
    ("Net Debt / EBITDA", "Ratio d'endettement net = (Net Debt / EBITDA)"),
]

for label, formula_text in structure_ratios:
    ws_ratios[f"A{row}"] = label
    set_cell_style(ws_ratios[f"A{row}"], font=font_normal)
    ws_ratios[f"B{row}"] = formula_text
    set_cell_style(ws_ratios[f"B{row}"], font=font_normal, alignment_h="left")
    row += 1

# Valuation Ratios
row += 1
ws_ratios[f"A{row}"] = "VALUATION RATIOS (Valorisation)"
set_cell_style(ws_ratios[f"A{row}"], font=Font(name="Arial", size=11, bold=True), fill=SUB_HEADER_COLOR)
row += 1

valuation_ratios = [
    ("EPS", "Bénéfice Net Par Action = Net Income / Shares Outstanding"),
    ("PER", "Ratio Cours/Bénéfice = Share Price / EPS"),
    ("Book Value Per Share", "Valeur Comptable par Action = Equity / Shares Outstanding"),
    ("Price-to-Book Ratio", "PBR = Share Price / Book Value Per Share"),
    ("Dividend Yield", "Rendement du dividende = DPS / Share Price"),
    ("Enterprise Value", "EV = Market Cap + Net Debt"),
    ("EV/EBITDA", "Multiple de valorisation = EV / EBITDA"),
]

for label, formula_text in valuation_ratios:
    ws_ratios[f"A{row}"] = label
    set_cell_style(ws_ratios[f"A{row}"], font=font_normal)
    ws_ratios[f"B{row}"] = formula_text
    set_cell_style(ws_ratios[f"B{row}"], font=font_normal, alignment_h="left")
    row += 1

# ==================== SHEET 6: DCF VALUATION ====================

ws_dcf = wb.create_sheet("DCF", 5)
ws_dcf.column_dimensions["A"].width = 35
for col in ["B", "C", "D", "E", "F"]:
    ws_dcf.column_dimensions[col].width = 16

add_section_title(ws_dcf, 1, "DCF VALUATION (Valorisation par DCF)", "A", "F")

# DCF Assumptions
dcf_assumptions = [
    ("WACC (Coût Moyen Pondéré du Capital)", "8.37%"),
    ("Terminal Growth Rate (Taux de croissance terminal)", "3.5%"),
    ("FCF Margin (Marge de FCF normalisée)", "10%"),
    ("Forecast Period", "2024-2028 (5 years)"),
    ("", ""),
]

row = 3
for label, value in dcf_assumptions:
    ws_dcf[f"A{row}"] = label
    set_cell_style(ws_dcf[f"A{row}"], font=font_label, fill=SUB_HEADER_COLOR if label else LIGHT_GRAY)
    ws_dcf[f"B{row}"] = value
    set_cell_style(ws_dcf[f"B{row}"], font=font_normal, number_format="0.00%" if "%" in str(value) and "%" not in label else "@")
    row += 1

# Projected Cash Flows
row += 1
add_section_title(ws_dcf, row, "PROJECTED FREE CASH FLOWS", "A", "F")
row += 1

fcf_headers = ["Year", "2024", "2025", "2026", "2027", "2028"]
for idx, header in enumerate(fcf_headers):
    col = get_column_letter(idx + 1)
    cell = ws_dcf[f"{col}{row}"]
    set_cell_style(cell, value=header, fill=HEADER_COLOR, font=font_header, alignment_h="center")
row += 1

fcf_items = [
    "Projected Revenue",
    "FCF Margin",
    "Free Cash Flow",
    "Discount Factor",
    "PV of FCF",
]

for item in fcf_items:
    ws_dcf[f"A{row}"] = item
    set_cell_style(ws_dcf[f"A{row}"], font=font_label, fill=SUB_HEADER_COLOR)
    row += 1

# DCF Summary
row += 1
add_section_title(ws_dcf, row, "DCF VALUATION SUMMARY", "A", "F")
row += 1

summary_items = [
    ("Sum of PV(FCFs) - Years 1-5", "", "Somme des VA des FCF"),
    ("PV of Terminal Value", "", "VA de la Valeur Terminale"),
    ("Intrinsic Enterprise Value", "", "Valeur d'Entreprise intrinsèque"),
    ("Less: Net Debt", "", "Moins: Dette nette"),
    ("Intrinsic Equity Value", "", "Valeur des Capitaux Propres intrinsèque"),
    ("Shares Outstanding", "", "Nombre d'actions"),
    ("Intrinsic Value Per Share", "", "Valeur intrinsèque par action"),
    ("Current Share Price", "", "Cours boursier actuel"),
    ("Upside / Downside", "", "Potentiel de hausse/baisse"),
]

for label, value, note in summary_items:
    ws_dcf[f"A{row}"] = label
    set_cell_style(ws_dcf[f"A{row}"], font=font_label, fill=HIGHLIGHT_COLOR)
    ws_dcf[f"B{row}"] = note
    set_cell_style(ws_dcf[f"B{row}"], font=Font(name="Arial", size=9, italic=True), fill=LIGHT_GRAY)
    row += 1

# ==================== SHEET 7: SENSITIVITY ANALYSIS ====================

ws_sens = wb.create_sheet("SENSITIVITY", 6)
ws_sens.column_dimensions["A"].width = 16
for col in ["B", "C", "D", "E", "F", "G", "H"]:
    ws_sens.column_dimensions[col].width = 15

add_section_title(ws_sens, 1, "SENSITIVITY ANALYSIS - Enterprise Value", "A", "H")

# WACC / g sensitivity matrix
sens_row = 3
ws_sens[f"A{sens_row}"] = "WACC"
set_cell_style(ws_sens[f"A{sens_row}"], font=font_header, fill=HEADER_COLOR)

wacc_values = ["g ↓", "5.37%", "6.37%", "7.37%", "8.37%", "9.37%", "10.37%"]
growth_values = ["2.50%", "3.00%", "3.50%", "4.00%", "4.50%"]

for idx, wacc in enumerate(wacc_values):
    col = get_column_letter(idx + 1)
    cell = ws_sens[f"{col}{sens_row}"]
    if idx == 0:
        set_cell_style(cell, value=wacc, font=font_header, fill=HEADER_COLOR, alignment_h="center")
    else:
        set_cell_style(cell, value=wacc, font=font_header, fill=HEADER_COLOR, alignment_h="center")

for idx, growth in enumerate(growth_values):
    row = sens_row + 1 + idx
    ws_sens[f"A{row}"] = growth
    set_cell_style(ws_sens[f"A{row}"], font=font_header, fill=HEADER_COLOR, alignment_h="center")

# Placeholder matrix values (would be calculated in real model)
for i in range(1, len(growth_values) + 1):
    for j in range(1, len(wacc_values)):
        row = sens_row + i
        col = get_column_letter(j + 1)
        ws_sens[f"{col}{row}"] = ""
        set_cell_style(ws_sens[f"{col}{row}"], value="", alignment_h="right", number_format="#,##0")

# ==================== SHEET 8: TECHNICAL INDICATORS ====================

ws_tech = wb.create_sheet("TECHNICAL", 7)
ws_tech.column_dimensions["A"].width = 25
ws_tech.column_dimensions["B"].width = 30
for col in ["C", "D", "E", "F"]:
    ws_tech.column_dimensions[col].width = 15

add_section_title(ws_tech, 1, "TECHNICAL INDICATORS", "A", "F")

tech_info = [
    ("Indicator", "Formula", "Notes"),
    ("SMA (n=20)", "Average of last n closing prices", "Equal weight on all periods"),
    ("EMA (n=20)", "α × Pₜ + (1-α) × EMAₜ₋₁", "Higher weight on recent prices"),
    ("RSI (14)", "100 - 100/(1 + RS)", "RSI > 70 = overbought; RSI < 30 = oversold"),
    ("MACD", "EMA(12) - EMA(26)", "Trend following indicator"),
    ("Bollinger Bands", "SMA ± (2 × StDev)", "Volatility indicator"),
]

for idx, (name, formula, notes) in enumerate(tech_info, start=3):
    ws_tech[f"A{idx}"] = name
    if idx == 3:
        set_cell_style(ws_tech[f"A{idx}"], font=font_header, fill=HEADER_COLOR)
        ws_tech[f"B{idx}"] = formula
        set_cell_style(ws_tech[f"B{idx}"], font=font_header, fill=HEADER_COLOR)
        ws_tech[f"C{idx}"] = notes
        set_cell_style(ws_tech[f"C{idx}"], font=font_header, fill=HEADER_COLOR)
    else:
        set_cell_style(ws_tech[f"A{idx}"], font=font_normal, fill=LIGHT_GRAY)
        ws_tech[f"B{idx}"] = formula
        set_cell_style(ws_tech[f"B{idx}"], font=font_normal, fill=LIGHT_GRAY)
        ws_tech[f"C{idx}"] = notes
        set_cell_style(ws_tech[f"C{idx}"], font=font_normal, fill=LIGHT_GRAY)

# ==================== SAVE WORKBOOK ====================

output_file = "Modèle Analyse CTM 2021-2025_FINAL.xlsx"
wb.save(output_file)
print(f"✅ Workbook successfully created: {output_file}")
print(f"\n📊 Sheets included:")
print(f"   1. ACCUEIL (Overview with key figures)")
print(f"   2. INCOME STATEMENT (Compte de Résultat)")
print(f"   3. BALANCE SHEET (Bilan)")
print(f"   4. CASH FLOW (Tableau de Trésorerie)")
print(f"   5. RATIOS (Financial indicators)")
print(f"   6. DCF (Valuation model)")
print(f"   7. SENSITIVITY (Analysis matrix)")
print(f"   8. TECHNICAL (Technical indicators reference)")
print(f"\n📁 Location: {output_file}")
print(f"\n⚠️  To complete the model:")
print(f"   - Fill in missing 2024-2026 data in BALANCE SHEET")
print(f"   - Add DCF calculations and sensitivity values")
print(f"   - Add share price data in technical sheet for charts")
print(f"   - Validate all figures against official CTM reports")
