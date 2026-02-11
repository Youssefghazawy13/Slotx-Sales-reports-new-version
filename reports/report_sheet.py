# reports/report_sheet.py

from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from utils.excel_helpers import auto_fit_columns
from core.kpi_engine import apply_deal


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

    percentage = deals_dict.get(brand_name, {}).get("percentage", 0)
    rent = deals_dict.get(brand_name, {}).get("rent", 0)

    after_percentage, after_rent = apply_deal(
        total_sales_money,
        percentage,
        rent
    )

    # =====================================================
    # SMALL KPI CARDS (SLIM DESIGN)
    # =====================================================

    def create_kpi_card(row, col, title, value):

        ws.merge_cells(start_row=row, start_column=col,
                       end_row=row, end_column=col+2)

        cell = ws.cell(row=row, column=col)
        cell.value = f"{title}: {value}"

        cell.font = Font(size=11, bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center",
                                   vertical="center")

        fill = PatternFill(
            start_color="0A1F5C",
            end_color="0A1F5C",
            fill_type="solid"
        )

        border = Border(
            left=Side(style="thin"),
            right=Side(style="thin"),
            top=Side(style="thin"),
            bottom=Side(style="thin")
        )

        for c in range(col, col+3):
            ws.cell(row=row, column=c).fill = fill
            ws.cell(row=row, column=c).border = border

    create_kpi_card(2, 1, "Total Sales",
                    f"{total_sales_money:,.2f} EGP")

    create_kpi_card(2, 5, "Net After Deal",
                    f"{after_rent:,.2f} EGP")

    create_kpi_card(2, 9, "Inventory Units",
                    total_inventory_qty)

    # =====================================================
    # REPORT DETAILS BELOW KPIs
    # =====================================================

    start_row = 5

    ws.cell(row=start_row, column=1,
            value="Branch Name:").font = Font(bold=True)
    ws.cell(row=start_row, column=2, value=mode)

    ws.cell(row=start_row+1, column=1,
            value="Brand Name:").font = Font(bold=True)
    ws.cell(row=start_row+1, column=2, value=brand_name)

    ws.cell(row=start_row+2, column=1,
            value="Payout Cycle:").font = Font(bold=True)
    ws.cell(row=start_row+2, column=2, value=payout_cycle)

    ws.cell(row=start_row+3, column=1,
            value="Total Sales Quantity:").font = Font(bold=True)
    ws.cell(row=start_row+3, column=2, value=total_sales_qty)

    ws.cell(row=start_row+4, column=1,
            value="Total Sales Money:").font = Font(bold=True)
    ws.cell(row=start_row+4, column=2,
            value=f"{total_sales_money:,.2f} EGP")

    ws.cell(row=start_row+5, column=1,
            value="Total Inventory Quantity:").font = Font(bold=True)
    ws.cell(row=start_row+5, column=2, value=total_inventory_qty)

    ws.cell(row=start_row+6, column=1,
            value="Total Inventory Value:").font = Font(bold=True)
    ws.cell(row=start_row+6, column=2,
            value=f"{total_inventory_value:,.2f} EGP")

    ws.cell(row=start_row+7, column=1,
            value="After Percentage:").font = Font(bold=True)
    ws.cell(row=start_row+7, column=2,
            value=f"{after_percentage:,.2f} EGP")

    ws.cell(row=start_row+8, column=1,
            value="After Rent:").font = Font(bold=True)
    ws.cell(row=start_row+8, column=2,
            value=f"{after_rent:,.2f} EGP")

    # =====================================================
    # TOP PRODUCTS PERFORMANCE
    # =====================================================

    top_start = start_row + 11

    ws.cell(row=top_start, column=1,
            value="Top Products Performance").font = Font(size=13, bold=True)

    ws.cell(row=top_start+1, column=1,
            value="Product").font = Font(bold=True)
    ws.cell(row=top_start+1, column=2,
            value="Quantity").font = Font(bold=True)
    ws.cell(row=top_start+1, column=3,
            value="Sales").font = Font(bold=True)

    product_sales = (
        brand_sales.groupby("product_name")["quantity"]
        .sum()
        .sort_values(ascending=False)
        .head(3)
    )

    row_pointer = top_start + 2

    for product, qty in product_sales.items():

        revenue = brand_sales[
            brand_sales["product_name"] == product
        ]["total"].sum()

        ws.cell(row=row_pointer, column=1, value=product)
        ws.cell(row=row_pointer, column=2, value=qty)
        ws.cell(row=row_pointer, column=3,
                value=f"{revenue:,.2f} EGP")

        row_pointer += 1

    auto_fit_columns(ws)
