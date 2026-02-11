# reports/sales_sheet.py

from openpyxl.styles import Font
from utils.excel_helpers import auto_fit_columns


def create_sales_sheet(
    wb,
    brand_name: str,
    mode: str,
    brand_sales
):
    """
    Create Sales sheet.

    Returns:
        total_sales_qty,
        total_sales_money
    """

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

    # Bold header
    for cell in ws[1]:
        cell.font = Font(bold=True)

    total_qty = 0
    total_money = 0

    for _, row in brand_sales.iterrows():
        qty = row.get("quantity", 0)
        total = row.get("total", 0)

        ws.append([
            mode,                  # Branch Name (Merged or single)
            brand_name,
            row.get("product_name", ""),
            row.get("barcode", ""),
            qty,
            total
        ])

        total_qty += qty
        total_money += total

    # Add one TOTAL row only
    ws.append([
        "",
        "",
        "",
        "TOTAL",
        total_qty,
        total_money
    ])

    # Bold total row
    for cell in ws[ws.max_row]:
        cell.font = Font(bold=True)

    auto_fit_columns(ws)

    return total_qty, total_money
