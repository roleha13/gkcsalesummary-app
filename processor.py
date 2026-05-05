# processor.py
import os
import re
import calendar
from datetime import datetime
import pdfplumber
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment, PatternFill, Font
from openpyxl.chart import LineChart, PieChart, BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.series import DataPoint
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.chart.axis import ChartLines
from openpyxl.chart.legend import Legend, Layout, ManualLayout

MONEY_FMT = '#,##0.00;(#,##0.00)'
YELLOW_FILL = PatternFill(start_color="FFFFF2CC", end_color="FFFFF2CC", fill_type="solid")


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
        'Date','TABLE AR','TABLE CARDS','SLOTS AC+CT','SLOTS EG+AM+NOV','SLOTS TBJ',
        'TOTAL WINNINGS','TIPS','GROSS INCOME','CREDIT GIVEN','CREDIT REPAID','NET OF CREDIT & TIPS'
    ]

    num_cols = len(headers)

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

    ws.append(headers)

    start_row = 5
    for i, row in enumerate(rows, start=start_row):
        ws[f'A{i}'] = row['date']
        ws[f'A{i}'].number_format = 'd-mmm-yy'

        vals = ['TABLE AR','TABLE CARDS','SLOTS AC+CT','SLOTS EG+AM+NOV','SLOTS TBJ']
        for col, key in enumerate(vals, start=2):
            ws.cell(row=i, column=col, value=row[key]).number_format = MONEY_FMT

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

    for col_idx in range(2, num_cols + 1):
        col_letter = get_column_letter(col_idx)
        ws[f'{col_letter}{total_row}'] = f'=SUM({col_letter}{start_row}:{col_letter}{total_row-1})'
        ws[f'{col_letter}{total_row}'].number_format = MONEY_FMT

    dashboard = wb.create_sheet('Dashboard')

    # ===== EXECUTIVE DASHBOARD STYLING =====
    dashboard.sheet_view.showGridLines = False
    dashboard['A1'] = 'GOLDEN KEY CASINO – EXECUTIVE DASHBOARD'
    dashboard['A1'].font = Font(color='FFFFFF', bold=True, size=16)
    dashboard.merge_cells('A1:H1')
    dashboard['A1'].alignment = Alignment(horizontal='center')
    dashboard['A1'].fill = PatternFill('solid', fgColor='1F4E78')

    metrics = ['TABLE AR','TABLE CARDS','SLOTS AC+CT','SLOTS EG+AM+NOV','SLOTS TBJ']
    dashboard['A3'] = 'Gaming Results Summary'
    dashboard['A3'].font = Font(bold=True, size=12)

    from openpyxl.styles import Border, Side
    thin = Side(style='thin', color='000000')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for idx, metric in enumerate(metrics, start=4):
        dashboard[f'A{idx}'] = metric
        dashboard[f'B{idx}'] = f"='Sales Summary'!{get_column_letter(idx-2)}{total_row}"
        dashboard[f'B{idx}'].number_format = MONEY_FMT
        dashboard[f'A{idx}'].border = border
        dashboard[f'B{idx}'].border = border
        dashboard[f'C{idx}'].border = border

    dashboard['A9'] = 'TOTAL WINNINGS'
    dashboard['B9'] = f"='Sales Summary'!G{total_row}"
    dashboard['B9'].number_format = MONEY_FMT
    dashboard['A9'].border = border
    dashboard['B9'].border = border
    dashboard['C9'].border = border

    for idx in range(4, 9):
        dashboard[f'C{idx}'] = f'=B{idx}/$B$9'
        dashboard[f'C{idx}'].number_format = '0%'

    # ===== PIE CHART =====
    pie = PieChart()
    pie.add_data(Reference(dashboard, min_col=2, min_row=4, max_row=8))
    pie.set_categories(Reference(dashboard, min_col=1, min_row=4, max_row=8))
    pie.title = 'Gaming Results Summary'
    pie.style = 10
    pie.height = 7
    pie.width = 9
    pie.legend.position = 'r'
    pie.dataLabels = DataLabelList()
    pie.dataLabels.showPercent = True
    pie.dataLabels.showLeaderLines = True

    slice_colors = ['C0504D','4F81BD','9BBB59','8064A2','F79646']
    for i, color in enumerate(slice_colors):
        pt = DataPoint(idx=i)
        pt.graphicalProperties = GraphicalProperties(solidFill=color)
        pie.series[0].data_points.append(pt)

    dashboard.add_chart(pie, 'D3')

    line_dates = Reference(ws, min_col=1, min_row=5, max_row=total_row-1)

    # ===== TOTAL WINNINGS LINE =====
    line = LineChart()
    line.add_data(Reference(ws, min_col=7, min_row=5, max_row=total_row-1), titles_from_data=False)
    line.set_categories(line_dates)
    line.title = 'Monthly Total Winnings'
    line.style = 13
    line.marker = 'circle'
    line.x_axis.title = 'Day of Month'
    line.y_axis.title = 'Amount'
    line.legend = Legend()
    line.y_axis.majorGridlines = Chartlines()
    line.width = 14
    line.height = 7
    dashboard.add_chart(line, 'D30')

    # ===== TIPS LINE =====
    tips = LineChart()
    tips.add_data(Reference(ws, min_col=8, min_row=5, max_row=total_row-1), titles_from_data=False)
    tips.set_categories(line_dates)
    tips.title = 'Monthly Tips'
    tips.style = 13
    tips.marker = 'circle'
    tips.x_axis.title = 'Day of Month'
    tips.y_axis.title = 'Amount'
    tips.legend = Legend()
    tips.y_axis.majorGridlines = Chartlines()
    tips.width = 14
    tips.height = 7
    dashboard.add_chart(tips, 'D50')

    # ===== STACKED DAILY MIX =====
    stacked = BarChart()
    stacked.type = 'col'
    stacked.grouping = 'stacked'
    stacked.overlap = 100
    stacked.add_data(Reference(ws, min_col=2, max_col=6, min_row=5, max_row=total_row-1), titles_from_data=True)
    stacked.set_categories(line_dates)
    stacked.title = 'Daily Gaming Mix'
    stacked.style = 12
    stacked.x_axis.title = 'Day of Month'
    stacked.y_axis.title = 'Amount'
    stacked.legend = Legend()
    stacked.y_axis.majorGridlines = Chartlines()
    stacked.width = 16
    stacked.height = 9
    dashboard.add_chart(stacked, 'D75')

    file_name = f"{month_number:02d}. SALES ANALYSIS {month_name} {year_full}.xlsx"
    output_path = os.path.join(output_folder, file_name)
    wb.save(output_path)

    return output_path



