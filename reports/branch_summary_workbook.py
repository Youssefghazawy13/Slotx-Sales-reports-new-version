from openpyxl import Workbook
from io import BytesIO

from reports.branch_summary_performance import create_branch_performance_sheet
from reports.metadata_sheet import create_metadata_sheet


def build_branch_summary_workbook(
    branch_name,
    payout_cycle,
    sales_df,
    inventory_df,
    deals_dict
):

    wb = Workbook()

    if "Sheet" in wb.sheetnames:
        wb.remove(wb["Sheet"])

    # =========================
    # PERFORMANCE TAB
    # =========================

    create_branch_performance_sheet(
        wb=wb,
        branch_name=branch_name,
        payout_cycle=payout_cycle,
        sales_df=sales_df,
        inventory_df=inventory_df,
        deals_dict=deals_dict
    )

    # =========================
    # METADATA
    # =========================

    create_metadata_sheet(
        wb,
        f"{branch_name} Summary",
        branch_name,
        payout_cycle
    )

    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    return buffer