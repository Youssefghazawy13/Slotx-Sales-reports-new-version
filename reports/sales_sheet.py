# reports/sales_sheet.py

from openpyxl.styles import Font
from utils.excel_helpers import auto_fit_columns, apply_header_style, format_money_cell


def create_sales_sheet(
    wb,
    brand_name: str,
    mode: str,
    brand_sales
):

    ws = wb.create_sheet("Sales")

    headers = [
        "Branch Name",
        "Brand Name",
        "Product Name",
        "Barcode",
        "Quantity",
        "Total Price"
    ]

    ws.append(headers)
    apply_header_style(ws)

    total_qty = 0
    total_money = 0

    for _, row in brand_sales.iterrows():
        qty = row.get("quantity", 0)
        total = row.get("total", 0)

        ws.append([
            mode,
            brand_name,
            row.get("product_name", ""),
            row.get("barcode", ""),
            qty,
            total
        ])

        ws.cell(row=ws.max_row, column=6).number_format = '#,##0.00'

        total_qty += qty
        total_money += total

    ws.append(["", "", "", "TOTAL", total_qty, total_money])

    ws.cell(row=ws.max_row, column=6).number_format = '#,##0.00'

    for cell in ws[ws.max_row]:
        cell.font = Font(bold=True)

    auto_fit_columns(ws)

    return total_qty, total_money
