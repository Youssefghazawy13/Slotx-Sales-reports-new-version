from openpyxl import Workbook
from reports.branch_summary_performance import create_performance_sheet
from reports.metadata_sheet import create_metadata_sheet
from reports.sales_sheet import create_sales_sheet
from reports.inventory_sheet import create_inventory_sheet


def build_branch_summary_workbook(
    branch_name,
    payout_cycle,
    sales_df,
    inventory_df,
    deals_dict
):

    wb = Workbook()

    # Remove default sheet
    if "Sheet" in wb.sheetnames:
        wb.remove(wb["Sheet"])

    # ============================
    # All Sales Details
    # ============================

    create_sales_sheet(
        wb=wb,
        brand_sales=sales_df,
        mode=branch_name
    )

    # ============================
    # All Inventory
    # ============================

    create_inventory_sheet(
        wb=wb,
        brand_inventory=inventory_df,
        mode=branch_name
    )

    # ============================
    # Deals Sheet
    # ============================

    ws_deals = wb.create_sheet("Deals")

    ws_deals.append(["Brand", "Deal %", "Rent (EGP)"])

    for brand, deal in deals_dict.items():
        ws_deals.append([
            brand,
            deal.get("percentage", 0),
            deal.get("rent", 0)
        ])

    ws_deals.column_dimensions["A"].width = 25
    ws_deals.column_dimensions["B"].width = 15
    ws_deals.column_dimensions["C"].width = 18

    # ============================
    # Branch Summary Sheet
    # ============================

    create_performance_sheet(
        wb=wb,
        branch_name=branch_name,
        payout_cycle=payout_cycle,
        sales_df=sales_df,
        inventory_df=inventory_df,
        deals_dict=deals_dict
    )

    # ============================
    # Metadata Sheet
    # ============================

    create_metadata_sheet(
        wb,
        f"{branch_name} Summary",
        branch_name,
        payout_cycle
    )

    return wb
