# reports/inventory_sheet.py

from openpyxl.styles import Font
from utils.excel_helpers import auto_fit_columns
from core.kpi_engine import calculate_status


def create_inventory_sheet_single(
    wb,
    brand_name: str,
    mode: str,
    brand_inventory,
    brand_sales,
    has_deal: bool
):
    """
    Create Inventory sheet for single branch mode.

    Returns:
        total_inventory_qty,
        total_inventory_value
    """

    ws = wb.create_sheet("Inventory")

    headers = [
        "Branch Name",
        "Brand Name",
        "Product Name",
        "Barcode",
        "Unit Price",
        "Available Quantity",
        "Status"
    ]

    ws.append(headers)

    for cell in ws[1]:
        cell.font = Font(bold=True)

    total_inventory_qty = 0
    total_inventory_value = 0

    for _, row in brand_inventory.iterrows():
        barcode = row.get("barcode")
        qty = row.get("available_quantity", 0)
        price = row.get("unit_price", 0)

        sales_qty = brand_sales[
            brand_sales["barcode"] == barcode
        ]["quantity"].sum()

        status = calculate_status(
            sales_qty=sales_qty,
            inventory_qty=qty,
            has_deal=has_deal
        )

        ws.append([
            mode,
            brand_name,
            row.get("product_name", ""),
            barcode,
            price,
            qty,
            status
        ])

        total_inventory_qty += qty
        total_inventory_value += qty * price

    auto_fit_columns(ws)

    return total_inventory_qty, total_inventory_value
