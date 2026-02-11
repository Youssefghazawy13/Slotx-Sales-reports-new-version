from openpyxl import Workbook
from io import BytesIO

from reports.sales_sheet import create_sales_sheet
from reports.inventory_sheet import (
    create_inventory_sheet_single,
    create_inventory_sheet_merged
)
from reports.report_sheet import create_report_sheet
from reports.metadata_sheet import create_metadata_sheet

from core.deals_engine import has_deal


def build_brand_workbook(
    brand_name: str,
    mode: str,
    payout_cycle: str,
    brand_sales,
    brand_inventory=None,
    inventory_alex=None,
    inventory_zam=None,
    deals_dict=None
):

    wb = Workbook()

    if "Sheet" in wb.sheetnames:
        wb.remove(wb["Sheet"])

    # --- SALES ---
    total_sales_qty, total_sales_money = create_sales_sheet(
        wb,
        brand_name,
        mode,
        brand_sales
    )

    has_brand_deal = has_deal(brand_name, deals_dict or {})

    # --- INVENTORY ---
    if mode == "Merged":

        total_inventory_qty, total_inventory_value = (
            create_inventory_sheet_merged(
                wb,
                brand_name,
                inventory_alex if inventory_alex is not None else brand_sales.iloc[0:0],
                inventory_zam if inventory_zam is not None else brand_sales.iloc[0:0],
                brand_sales,
                has_brand_deal
            )
        )

    else:

        total_inventory_qty, total_inventory_value = (
            create_inventory_sheet_single(
                wb,
                brand_name,
                mode,
                brand_inventory if brand_inventory is not None else brand_sales.iloc[0:0],
                brand_sales,
                has_brand_deal
            )
        )

    # --- REPORT ---
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
        deals_dict or {}
    )

    # --- METADATA ---
    create_metadata_sheet(
        wb,
        mode,
        payout_cycle
    )

    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    wb.close()

    return buffer
