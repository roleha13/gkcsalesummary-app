import os
import re
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
from openpyxl.chart.legend import Legend


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
# PDF EXTRACTION
# =========================================================

def get_last_number_from_line_with_keyword(text_block: str, keyword: str):
    for line in text_block.splitlines():
        if keyword in line:
            nums = re.findall(r"[\d,]+", line)
            if not nums:
                continue
            value = float(nums[-1].replace(",", ""))
            return -value if '(' in line and ')' in line else value
    raise ValueError(f"Could not find numeric value for keyword '{keyword}'.")


def extract_values_from_pdf(pdf_path: str):
    with pdfplumber.open(pdf_path) as pdf:
        text = "\n".join(page.extract_text() or "" for page in pdf.pages)

    date_match = re.search(r"Date\s+(\d{1,2}-[A-Za-z]{3}-\d{2})", text)
    if not date_match:
        raise ValueError(f"Date not found in {pdf_path}")

    date_obj = datetime.strptime(date_match.group(1), "%d-%b-%y")

    ar_text = re.search(r"AMERICAN ROULETTE(.*?CARDS)", text, re.S).group(1)
    cards_text = re.search(r"CARDS(.*?TABLES TOTALS)", text, re.S).group(1)
    slots_text = re.search(r"SLOTS\s*\nBANK No\..*", text, re.S).group(0)

    return {
        "date": date_obj,
        "TABLE AR": get_last_number_from_line_with_keyword(ar_text, "Sub-Totals"),
        "TABLE CARDS": get_last_number_from_line_with_keyword(cards_text, "Sub-Totals"),
        "SLOTS AC+CT": get_last_number_from_line_with_keyword(slots_text, "SLOTS AT+CT"),
        "SLOTS EG+AM+NOV": get_last_number_from_line_with_keyword(slots_text, "SLOTS EG+AM+NOV"),
        "SLOTS TBJ": get_last_number_from_line_with_keyword(slots_text, "SLOTS TBJ"),
    }


# =========================================================
# MAIN PROCESS
# =========================================================

def process_pdfs_to_excel(pdf_files, output_folder):

    rows = [extract_values_from_pdf(pdf) for pdf in pdf_files]
    rows.sort(key=lambda r: r['date'])

    first_date = rows[0]['date']

    month_name = first_date.strftime('%B').upper()
    month_abbrev_year = first_date.strftime('%b-%y')
    year_full = first_date.year
    month_number = first_date.month

    wb = Workbook()
    ws = wb.active
    ws.title = 'Sales Summary'

    headers = [
        'Date','TABLE AR','TABLE CARDS','SLOTS AC+CT',
        'SLOTS EG+AM+NOV','SLOTS TBJ','TOTAL WINNINGS',
        'TIPS','GROSS INCOME','CREDIT GIVEN','CREDIT REPAID',
        'NET OF CREDIT & TIPS'
    ]

    num_cols = len(headers)

    # ================= HEADER =================

    ws.merge_cells('A1:L1')
    ws['A1'] = 'GOLDEN KEY CASINO'
    ws['A1'].font = Font(bold=True, size=14)
    ws['A1'].alignment = Alignment(horizontal='center')

    ws.merge_cells('A2:L2')
    ws['A2'] = 'SALES ANALYSIS FOR'
    ws['A2'].alignment = Alignment(horizontal='center')

    ws.merge_cells('A3:L3')
    ws['A3'] = month_abbrev_year
    ws['A3'].fill = YELLOW_FILL
    ws['A3'].alignment = Alignment(horizontal='center')

    ws.append(headers)

    for cell in ws[4]:
        cell.font = Font(bold=True)
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal='center')
        cell.border = border

    # ================= DATA =================

    start_row = 5

    for i, row in enumerate(rows, start=start_row):

        ws[f'A{i}'] = row['date']
        ws[f'A{i}'].number_format = 'd-mmm-yy'

        keys = [
            'TABLE AR','TABLE CARDS','SLOTS AC+CT',
            'SLOTS EG+AM+NOV','SLOTS TBJ'
        ]

        for col, key in enumerate(keys, start=2):
            ws.cell(i, col, row[key]).number_format = MONEY_FMT

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

    ws[f'A{total_row}'] = 'TOTAL'
    ws[f'A{total_row}'].font = Font(bold=True)

    for col in range(2, num_cols + 1):
        col_letter = get_column_letter(col)
        ws[f'{col_letter}{total_row}'] = (
            f'=SUM({col_letter}{start_row}:{col_letter}{total_row-1})'
        )
        ws[f'{col_letter}{total_row}'].font = Font(bold=True)

    # ================= DASHBOARD =================

    dashboard = wb.create_sheet('Dashboard')
    dashboard.sheet_view.showGridLines = False

    dashboard.merge_cells('A1:H1')
    dashboard['A1'] = 'GOLDEN KEY CASINO – EXECUTIVE DASHBOARD'
    dashboard['A1'].font = Font(color='FFFFFF', bold=True, size=16)
    dashboard['A1'].fill = HEADER_FILL
    dashboard['A1'].alignment = Alignment(horizontal='center')

    # ================= PIE =================

    metrics = ['TABLE AR','TABLE CARDS','SLOTS AC+CT','SLOTS EG+AM+NOV','SLOTS TBJ']

    for idx, m in enumerate(metrics, start=4):
        dashboard[f'A{idx}'] = m
        dashboard[f'B{idx}'] = f"='Sales Summary'!{get_column_letter(idx-2)}{total_row}"

    dashboard['A9'] = 'TOTAL WINNINGS'
    dashboard['B9'] = f"='Sales Summary'!G{total_row}"

    pie = PieChart()
    pie.add_data(Reference(dashboard, min_col=2, min_row=4, max_row=8))
    pie.set_categories(Reference(dashboard, min_col=1, min_row=4, max_row=8))
    pie.title = 'Gaming Results Summary'
    pie.legend.position = 'r'
    dashboard.add_chart(pie, 'D3')

    # ================= LINE CHARTS =================

    line_dates = Reference(ws, min_col=1, min_row=5, max_row=total_row-1)

    def build_line(col, title, pos):

        chart = LineChart()

        data = Reference(ws, min_col=col, max_col=col,
                         min_row=4, max_row=total_row-1)

        chart.add_data(data, titles_from_data=True)
        chart.set_categories(line_dates)

        chart.title = title

        chart.width = 20
        chart.height = 9

        # NO LEGEND
        chart.legend = None

        # AXIS LABELS
        chart.x_axis.title = "Day of Month"
        chart.y_axis.title = "Amount"

        # LIGHT GRIDLINES
        chart.y_axis.majorGridlines = ChartLines()

        # AXIS LINES (light gray)
        chart.x_axis.spPr = GraphicalProperties()
        chart.y_axis.spPr = GraphicalProperties()

        # SERIES STYLE (LIGHT BLUE)
        s = chart.series[0]
        s.graphicalProperties.line.solidFill = "5B9BD5"
        s.graphicalProperties.line.width = 25000
        s.marker.symbol = "circle"
        s.marker.size = 7

        dashboard.add_chart(chart, pos)

    build_line(7, "Monthly Total Winnings", "D30")
    build_line(8, "Monthly Tips", "D50")

    # ================= STACKED CHART =================

    stacked = BarChart()

    data = Reference(ws, min_col=2, max_col=6,
                     min_row=4, max_row=total_row-1)

    stacked.add_data(data, titles_from_data=True)
    stacked.set_categories(line_dates)

    stacked.type = 'col'
    stacked.grouping = 'stacked'
    stacked.overlap = 100

    stacked.title = "Daily Gaming Mix"

    stacked.width = 22
    stacked.height = 11

    # ONLY LEGEND HERE (TOP)
    stacked.legend.position = 't'

    # AXES
    stacked.x_axis.title = "Day of Month"
    stacked.y_axis.title = "Amount"

    stacked.y_axis.majorGridlines = ChartLines()

    colors = ['C0504D','4F81BD','9BBB59','8064A2','F79646']

    for i, s in enumerate(stacked.series):
        s.graphicalProperties.solidFill = colors[i]

    dashboard.add_chart(stacked, 'D75')

    # ================= SAVE =================

    file_name = f"{month_number:02d}. SALES ANALYSIS {month_name} {year_full}.xlsx"

    output_path = os.path.join(output_folder, file_name)

    wb.save(output_path)

    return output_path
