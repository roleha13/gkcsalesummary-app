# processor.py (ENTERPRISE-GRADE RESILIENT VERSION)

import os
import re
import logging
from datetime import datetime

import pdfplumber

from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import (
    Alignment,
    PatternFill,
    Font,
    Border,
    Side
)

from openpyxl.chart import (
    LineChart,
    PieChart,
    BarChart,
    Reference
)

from openpyxl.chart.label import DataLabelList
from openpyxl.chart.series import DataPoint
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.chart.axis import ChartLines


# =========================================================
# LOGGING (ENTERPRISE SAFE)
# =========================================================

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s"
)

logger = logging.getLogger("casino_processor")


# =========================================================
# STYLES
# =========================================================

MONEY_FMT = '#,##0.00;(#,##0.00)'

YELLOW_FILL = PatternFill(
    start_color="FFFFF2CC",
    end_color="FFFFF2CC",
    fill_type="solid"
)

HEADER_FILL = PatternFill(
    fill_type="solid",
    fgColor="1F4E78"
)

thin = Side(style='thin', color='000000')
border = Border(left=thin, right=thin, top=thin, bottom=thin)


# =========================================================
# SAFE PARSING UTILITIES
# =========================================================

def safe_float(value):
    """Convert messy string → float safely"""
    try:
        return float(str(value).replace(",", "").strip())
    except:
        return 0.0


def extract_numbers(line: str):
    nums = re.findall(r"[\d,]+", line)
    return nums


def safe_extract_number(line: str):
    nums = extract_numbers(line)
    if not nums:
        return 0.0
    value = safe_float(nums[-1])
    return -value if "(" in line and ")" in line else value


# =========================================================
# SMART SLOT DETECTOR (NO HARD CODING)
# =========================================================

def extract_slots_dynamic(slots_text: str):
    """
    Dynamically extracts slot categories from ANY PDF layout
    """
    results = {}

    for line in slots_text.splitlines():
        line = line.strip()

        if not line:
            continue

        if re.search(r"\d", line):
            parts = re.split(r"\s{2,}|\t+", line)

            label = parts[0].strip() if parts else "UNKNOWN"

            value = safe_extract_number(line)

            # normalize label (removes weird spacing issues)
            label = re.sub(r"\s+", " ", label)

            results[label] = value

    return results


# =========================================================
# ROBUST BLOCK EXTRACTION
# =========================================================

def safe_block(text, pattern, name):
    try:
        match = re.search(pattern, text, re.S)
        return match.group(1 if match.groups() else 0)
    except Exception as e:
        logger.warning(f"Missing block: {name} | {e}")
        return ""


# =========================================================
# MAIN PDF EXTRACTION
# =========================================================

def extract_values_from_pdf(pdf_path: str):

    try:
        with pdfplumber.open(pdf_path) as pdf:
            text = "\n".join(page.extract_text() or "" for page in pdf.pages)

        # DATE
        date_match = re.search(r"Date\s+(\d{1,2}-[A-Za-z]{3}-\d{2})", text)

        if not date_match:
            logger.error(f"Date missing in {pdf_path}")
            return None

        date_obj = datetime.strptime(date_match.group(1), "%d-%b-%y")

        # SAFE BLOCK EXTRACTION
        ar_text = safe_block(text, r"AMERICAN ROULETTE(.*?CARDS)", "AR")
        cards_text = safe_block(text, r"CARDS(.*?TABLES TOTALS)", "CARDS")
        slots_text = safe_block(text, r"SLOTS\s*\nBANK No\..*", "SLOTS")

        slots_data = extract_slots_dynamic(slots_text)

        return {
            "date": date_obj,

            "TABLE AR": safe_extract_number(ar_text),
            "TABLE CARDS": safe_extract_number(cards_text),

            "SLOTS AC+CT": slots_data.get("SLOTS AC+CT", 0.0),
            "SLOTS EG+AM+NOV": slots_data.get("SLOTS EG+AM+NOV", 0.0),
            "SLOTS TBJ": slots_data.get("SLOTS TBJ", 0.0),
        }

    except Exception as e:
        logger.exception(f"Failed parsing {pdf_path}: {e}")
        return None


# =========================================================
# MAIN PROCESSING PIPELINE
# =========================================================

def process_pdfs_to_excel(pdf_files, output_folder):

    rows = []

    for pdf in pdf_files:
        result = extract_values_from_pdf(pdf)
        if result:
            rows.append(result)
        else:
            logger.warning(f"Skipped file due to extraction failure: {pdf}")

    if not rows:
        raise ValueError("No valid PDF data extracted.")

    rows.sort(key=lambda r: r["date"])

    first_date = rows[0]["date"]
    month_name = first_date.strftime('%B').upper()
    year_full = first_date.year
    month_number = first_date.month
    month_abbrev_year = first_date.strftime('%b-%y')

    wb = Workbook()
    ws = wb.active
    ws.title = "Sales Summary"

    headers = [
        'Date', 'TABLE AR', 'TABLE CARDS', 'SLOTS AC+CT',
        'SLOTS EG+AM+NOV', 'SLOTS TBJ',
        'TOTAL WINNINGS', 'TIPS', 'GROSS INCOME',
        'CREDIT GIVEN', 'CREDIT REPAID', 'NET OF CREDIT & TIPS'
    ]

    num_cols = len(headers)

    # =====================================================
    # HEADER TITLE
    # =====================================================

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=num_cols)
    ws['A1'] = 'GOLDEN KEY CASINO'
    ws['A1'].font = Font(bold=True, size=14)
    ws['A1'].alignment = Alignment(horizontal='center')

    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=num_cols)
    ws['A2'] = 'SALES ANALYSIS FOR'
    ws['A2'].alignment = Alignment(horizontal='center')

    ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=num_cols)
    ws['A3'] = month_abbrev_year
    ws['A3'].fill = YELLOW_FILL
    ws['A3'].alignment = Alignment(horizontal='center')

    # =====================================================
    # HEADERS
    # =====================================================

    ws.append(headers)

    for cell in ws[4]:
        cell.font = Font(bold=True)
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal='center')
        cell.border = border

    # =====================================================
    # DATA
    # =====================================================

    start_row = 5

    for i, row in enumerate(rows, start=start_row):

        ws[f'A{i}'] = row["date"]
        ws[f'A{i}'].number_format = 'd-mmm-yy'

        keys = ['TABLE AR', 'TABLE CARDS', 'SLOTS AC+CT', 'SLOTS EG+AM+NOV', 'SLOTS TBJ']

        for col, key in enumerate(keys, start=2):
            ws.cell(row=i, column=col, value=row.get(key, 0)).number_format = MONEY_FMT

        ws[f'G{i}'] = f'=SUM(B{i}:F{i})'
        ws[f'G{i}'].number_format = MONEY_FMT

        ws[f'H{i}'].number_format = MONEY_FMT
        ws[f'I{i}'] = f'=G{i}+H{i}'
        ws[f'I{i}'].number_format = MONEY_FMT

        ws[f'J{i}'].number_format = MONEY_FMT
        ws[f'K{i}'].number_format = MONEY_FMT
        ws[f'L{i}'] = f'=I{i}+J{i}+K{i}'
        ws[f'L{i}'].number_format = MONEY_FMT

    total_row = start_row + len(rows)

    ws[f'A{total_row}'] = "TOTAL"
    ws[f'A{total_row}'].font = Font(bold=True)

    for col_idx in range(2, num_cols + 1):
        col_letter = get_column_letter(col_idx)

        ws[f'{col_letter}{total_row}'] = (
            f'=SUM({col_letter}{start_row}:{col_letter}{total_row-1})'
        )

        ws[f'{col_letter}{total_row}'].font = Font(bold=True)
        ws[f'{col_letter}{total_row}'].number_format = MONEY_FMT

    # =====================================================
    # DASHBOARD
    # =====================================================

    dashboard = wb.create_sheet("Dashboard")
    dashboard.sheet_view.showGridLines = False

    dashboard.merge_cells("A1:H1")
    dashboard["A1"] = "GOLDEN KEY CASINO – EXECUTIVE DASHBOARD"
    dashboard["A1"].font = Font(color="FFFFFF", bold=True, size=16)
    dashboard["A1"].fill = HEADER_FILL
    dashboard["A1"].alignment = Alignment(horizontal="center")

    metrics = ['TABLE AR', 'TABLE CARDS', 'SLOTS AC+CT', 'SLOTS EG+AM+NOV', 'SLOTS TBJ']

    for idx, metric in enumerate(metrics, start=4):
        dashboard[f'A{idx}'] = metric
        dashboard[f'B{idx}'] = f"='Sales Summary'!{get_column_letter(idx-2)}{total_row}"
        dashboard[f'B{idx}'].number_format = MONEY_FMT

    dashboard["A9"] = "TOTAL WINNINGS"
    dashboard["B9"] = f"='Sales Summary'!G{total_row}"
    dashboard["B9"].number_format = MONEY_FMT

    for idx in range(4, 9):
        dashboard[f'C{idx}'] = f'=B{idx}/$B$9'
        dashboard[f'C{idx}'].number_format = '0%'

    # =====================================================
    # PIE CHART
    # =====================================================

    pie = PieChart()
    pie.add_data(Reference(dashboard, min_col=2, min_row=4, max_row=8))
    pie.set_categories(Reference(dashboard, min_col=1, min_row=4, max_row=8))
    pie.dataLabels = DataLabelList()
    pie.dataLabels.showPercent = True

    dashboard.add_chart(pie, "D3")

    # =====================================================
    # LINE CHART (CLEAN + STABLE)
    # =====================================================

    line_dates = Reference(ws, min_col=1, min_row=5, max_row=total_row-1)

    line = LineChart()
    data = Reference(ws, min_col=7, max_col=7, min_row=4, max_row=total_row-1)

    line.add_data(data, titles_from_data=True)
    line.set_categories(line_dates)

    line.legend = None
    line.x_axis.title = "Day of Month"
    line.y_axis.title = "Amount"

    dashboard.add_chart(line, "D30")

    # =====================================================
    # STACKED BAR (ONLY LEGEND HERE)
    # =====================================================

    stacked = BarChart()
    data = Reference(ws, min_col=2, max_col=6, min_row=4, max_row=total_row-1)

    stacked.add_data(data, titles_from_data=True)
    stacked.set_categories(line_dates)

    stacked.type = "col"
    stacked.grouping = "stacked"
    stacked.legend.position = "t"

    colors = ["C0504D", "4F81BD", "9BBB59", "8064A2", "F79646"]

    for i, series in enumerate(stacked.series):
        series.graphicalProperties = GraphicalProperties(
            solidFill=colors[i % len(colors)]
        )

    dashboard.add_chart(stacked, "D75")

    # =====================================================
    # SAVE
    # =====================================================

    file_name = f"{month_number:02d}. SALES ANALYSIS {month_name} {year_full}.xlsx"
    output_path = os.path.join(output_folder, file_name)

    wb.save(output_path)
    return output_path
