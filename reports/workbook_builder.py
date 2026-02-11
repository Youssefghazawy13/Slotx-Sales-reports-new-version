from openpyxl import Workbook
from io import BytesIO

from reports.sales_sheet import create_sales_sheet
from reports.inventory_sheet import create_inventory_sheet
from reports.report_sheet import create_report_sheet
from reports.metadata_sheet import create_metadata_sheet


def build_brand_workbook(
    brand_name: str,
    mode: str,
    payout_cycle: str,
    brand_sales,
    brand_inventory,
    deals_dict: dict
):

    # Defensive: ensure not None
    if brand_sales is None:
        brand_sales = []

    if brand_inventory is None:
        brand_inventory = []

    # Skip empty brand (no data at all)
    if len(brand_sales) == 0 and len(brand_inventory) == 0:
        return None

    wb = Workbook()

    # Remove default sheet
    default_sheet = wb.active
    wb.remove(default_sheet)

    # ==========================================
    # Calculate Totals
    # ==========================================

    total_sales_qty = 0
    total_sales_money = 0

    if len(brand_sales) > 0:
        total_sales_qty = brand_sales["quantity"].sum()
        total_sales_money = brand_sales["total"].sum()

    total_inventory_qty = 0
    total_inventory_value = 0

    if len(brand_inventory) > 0:

        if mode.lower() == "merged":
            total_inventory_qty = (
                brand_inventory.get("alex_qty", 0).sum()
                + brand_inventory.get("zamalek_qty", 0).sum()
            )
        else:
            total_inventory_qty = brand_inventory.get("quantity", 0).sum()

        total_inventory_value = (
            brand_inventory.get("price", 0) *
            brand_inventory.get("quantity", 0)
        ).sum()

    # ==========================================
    # Create Sheets (Strict Order)
    # ==========================================

    create_sales_sheet(
        wb,
        brand_sales,
        mode
    )

    create_inventory_sheet(
        wb,
        brand_inventory,
        mode
    )

    create_report_sheet(
        wb,
        brand_name,
        mode,
        payout_cycle,
        brand_sales,
        total_inventory_qty,
        total_inventory_value,
        total_sales_qty,
        total_sales_money,
        deals_dict
    )

    create_metadata_sheet(
        wb,
        mode,
        payout_cycle
    )

    # ==========================================
    # Save to Buffer
    # ==========================================

    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    return buffer
