from openpyxl import Workbook
from io import BytesIO
import pandas as pd

from reports.sales_sheet import create_sales_sheet
from reports.inventory_sheet import create_inventory_sheet
from reports.report_sheet import create_report_sheet
from reports.metadata_sheet import create_metadata_sheet


def build_brand_workbook(
    brand_name,
    mode,
    payout_cycle,
    brand_sales,
    brand_inventory,
    deals_dict
):

    if brand_sales is None:
        brand_sales = pd.DataFrame()

    if brand_inventory is None:
        brand_inventory = pd.DataFrame()

    if brand_sales.empty and brand_inventory.empty:
        return None

    wb = Workbook()
    wb.remove(wb.active)

    # ==========================
    # Sales Totals
    # ==========================

    total_sales_qty = (
        brand_sales["quantity"].sum()
        if "quantity" in brand_sales.columns else 0
    )

    total_sales_money = (
        brand_sales["total"].sum()
        if "total" in brand_sales.columns else 0
    )

    # ==========================
    # Inventory Totals
    # ==========================

    total_inventory_qty = 0
    total_inventory_value = 0

    if not brand_inventory.empty:

        if mode.lower() == "merged":

            alex_qty = (
                brand_inventory["alex_qty"].sum()
                if "alex_qty" in brand_inventory.columns else 0
            )

            zam_qty = (
                brand_inventory["zamalek_qty"].sum()
                if "zamalek_qty" in brand_inventory.columns else 0
            )

            total_inventory_qty = alex_qty + zam_qty

        else:
            total_inventory_qty = (
                brand_inventory["quantity"].sum()
                if "quantity" in brand_inventory.columns else 0
            )

        if "price" in brand_inventory.columns and "quantity" in brand_inventory.columns:
            total_inventory_value = (
                brand_inventory["price"]
                * brand_inventory["quantity"]
            ).sum()

    # ==========================
    # Create Sheets (ORDER)
    # ==========================

    create_sales_sheet(wb, brand_sales, mode)
    create_inventory_sheet(wb, brand_inventory, mode)

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

    create_metadata_sheet(wb, mode, payout_cycle)

    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    return buffer
