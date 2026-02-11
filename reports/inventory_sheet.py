from openpyxl.styles import Font
from utils.excel_helpers import (
    auto_fit_columns,
    apply_header_style,
    get_status_fill,
    format_money_cell
)
from core.kpi_engine import calculate_status


def create_inventory_sheet_single(
    wb,
    brand_name,
    mode,
    brand_inventory,
    brand_sales,
    has_deal
):

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
    apply_header_style(ws)

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
            sales_qty,
            qty,
            has_deal
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

        format_money_cell(ws.cell(row=ws.max_row, column=5))
        ws.cell(row=ws.max_row, column=7).fill = get_status_fill(status)

        total_inventory_qty += qty
        total_inventory_value += qty * price

    auto_fit_columns(ws)

    return total_inventory_qty, total_inventory_value


def create_inventory_sheet_merged(
    wb,
    brand_name,
    inventory_alex,
    inventory_zam,
    brand_sales,
    has_deal
):

    ws = wb.create_sheet("Inventory")

    headers = [
        "Brand Name",
        "Product Name",
        "Barcode",
        "Unit Price",
        "Alexandria Quantity",
        "Zamalek Quantity",
        "Total Stock",
        "Status"
    ]

    ws.append(headers)
    apply_header_style(ws)

    merged_inventory = {}

    for _, row in inventory_alex.iterrows():
        barcode = row.get("barcode")
        merged_inventory.setdefault(barcode, {
            "product_name": row.get("product_name", ""),
            "unit_price": row.get("unit_price", 0),
            "alex_qty": 0,
            "zam_qty": 0
        })
        merged_inventory[barcode]["alex_qty"] += row.get("available_quantity", 0)

    for _, row in inventory_zam.iterrows():
        barcode = row.get("barcode")
        merged_inventory.setdefault(barcode, {
            "product_name": row.get("product_name", ""),
            "unit_price": row.get("unit_price", 0),
            "alex_qty": 0,
            "zam_qty": 0
        })
        merged_inventory[barcode]["zam_qty"] += row.get("available_quantity", 0)

    total_inventory_qty = 0
    total_inventory_value = 0

    for barcode, data in merged_inventory.items():

        total_stock = data["alex_qty"] + data["zam_qty"]
        price = data["unit_price"]

        sales_qty = brand_sales[
            brand_sales["barcode"] == barcode
        ]["quantity"].sum()

        status = calculate_status(
            sales_qty,
            total_stock,
            has_deal
        )

        ws.append([
            brand_name,
            data["product_name"],
            barcode,
            price,
            data["alex_qty"],
            data["zam_qty"],
            total_stock,
            status
        ])

        format_money_cell(ws.cell(row=ws.max_row, column=4))
        ws.cell(row=ws.max_row, column=8).fill = get_status_fill(status)

        total_inventory_qty += total_stock
        total_inventory_value += total_stock * price

    auto_fit_columns(ws)

    return total_inventory_qty, total_inventory_value
