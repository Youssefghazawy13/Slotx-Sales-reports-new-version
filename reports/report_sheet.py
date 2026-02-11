# reports/report_sheet.py

from openpyxl.styles import Font
from utils.excel_helpers import auto_fit_columns, format_money_cell
from core.kpi_engine import (
    get_best_selling_product,
    get_best_selling_size,
    apply_deal
)
from core.deals_engine import generate_deal_text


def create_report_sheet(
    wb,
    brand_name: str,
    mode: str,
    payout_cycle: str,
    brand_sales,
    total_inventory_qty: float,
    total_inventory_value: float,
    total_sales_qty: float,
    total_sales_money: float,
    deals_dict: dict
):

    ws = wb.create_sheet("Report")

    deal_text = generate_deal_text(brand_name, deals_dict)

    percentage = deals_dict.get(brand_name, {}).get("percentage", 0)
    rent = deals_dict.get(brand_name, {}).get("rent", 0)

    after_percentage, after_rent = apply_deal(
        total_sales_money,
        percentage,
        rent
    )

    best_product = get_best_selling_product(brand_sales)
    best_size = get_best_selling_size(brand_sales)

    report_data = [
        ["Branch Name:", mode],
        ["Brand Name:", brand_name],
        ["Payout Cycle:", payout_cycle],
        ["Brand Deal:", deal_text],
        [""],
        ["Best Selling Product:", best_product],
        ["Best Selling Size:", best_size],
        [""],
        ["Total Inventory Quantity:", total_inventory_qty],
        ["Total Inventory Stock Value:", total_inventory_value],
        [""],
        ["Total Sales Quantity:", total_sales_qty],
        ["Total Sales Money:", total_sales_money],
        ["After Percentage:", after_percentage],
        ["After Rent:", after_rent],
    ]

    for row in report_data:
        ws.append(row)

    # Bold only first column (labels)
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=1):
        for cell in row:
            if cell.value:
                cell.font = Font(bold=True)

    # Format money fields
    money_rows = [
        "Total Inventory Stock Value:",
        "Total Sales Money:",
        "After Percentage:",
        "After Rent:"
    ]

    for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
        if row[0].value in money_rows:
            format_money_cell(row[1])

    auto_fit_columns(ws)
