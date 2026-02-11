from openpyxl import Workbook
from io import BytesIO

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

    wb = Workbook()

    # Remove default sheet
    if "Sheet" in wb.sheetnames:
        wb.remove(wb["Sheet"])

    # =========================
    # SALES SHEET
    # =========================
    create_sales_sheet(
        wb=wb,
        brand_sales=brand_sales,
        mode=mode
    )

    # =========================
    # INVENTORY SHEET
    # =========================
    create_inventory_sheet(
        wb=wb,
        brand_inventory=brand_inventory,
        mode=mode
    )

    # =========================
    # REPORT SHEET
    # =========================
    create_report_sheet(
        wb,
        brand_name,
        mode,
        payout_cycle,
        brand_sales,
        brand_inventory,
        deals_dict
    )

    # =========================
    # METADATA SHEET
    # =========================
    create_metadata_sheet(
        wb=wb,
        mode=mode,
        payout_cycle=payout_cycle
    )

    # =========================
    # SAVE TO BUFFER
    # =========================
    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    return buffer
