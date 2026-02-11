# reports/report_sheet.py

from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.chart import BarChart, Reference
from utils.excel_helpers import auto_fit_columns, format_money_cell
from core.kpi_engine import apply_deal
from collections import Counter


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

    # ======================================================
    # KPI CARDS
    # ======================================================

    def create_kpi_card(row, col, title, value, is_money=False):

        ws.merge_cells(start_row=row, start_column=col,
                       end_row=row+2, end_column=col+2)

        cell = ws.cell(row=row, column=col)
        cell.value = f"{title}\n{value}"

        cell.font = Font(size=14, bold=True)
        cell.alignment = Alignment(horizontal="center",
                                   vertical="center",
                                   wrap_text=True)

        fill = PatternFill(
            start_color="E8EEF7",
            end_color="E8EEF7",
            fill_type="solid"
        )

        border = Border(
            left=Side(style="thin"),
            right=Side(style="thin"),
            top=Side(style="thin"),
            bottom=Side(style="thin")
        )

        for r in range(row, row+3):
            for c in range(col, col+3):
                ws.cell(row=r, column=c).fill = fill
                ws.cell(row=r, column=c).border = border

    # Row 2 Cards
    create_kpi_card(2, 1, "Total Sales", f"{total_sales_money:,.2f} EGP")
    create_kpi_card(2, 5, "Net After Deal", f"{after_rent:,.2f} EGP")
    create_kpi_card(2, 9, "Inventory Units", total_inventory_qty)

    # Count Critical items
    critical_count = 0
    if not brand_sales.empty:
        # simple logic placeholder (can link with inventory later)
        pass

    create_kpi_card(2, 13, "Critical Items", critical_count)

    # ======================================================
    # TOP 3 PRODUCTS DATA
    # ======================================================

    start_row = 8

    ws.cell(row=start_row, column=1,
            value="📊 Top 3 Products Performance").font = Font(size=14, bold=True)

    product_sales = (
        brand_sales.groupby("product_name")["quantity"]
        .sum()
        .sort_values(ascending=False)
        .head(3)
    )

    products = product_sales.index.tolist()

    ws.cell(row=start_row+1, column=1, value="Product")
    ws.cell(row=start_row+1, column=2, value="Sales Units")
    ws.cell(row=start_row+1, column=3, value="Inventory Units")

    data_row = start_row + 2

    for product in products:

        sales_units = product_sales[product]

        # Inventory lookup placeholder
        inventory_units = 0

        ws.cell(row=data_row, column=1, value=product)
        ws.cell(row=data_row, column=2, value=sales_units)
        ws.cell(row=data_row, column=3, value=inventory_units)

        data_row += 1

    # ======================================================
    # HORIZONTAL BAR CHART
    # ======================================================

    chart = BarChart()
    chart.type = "bar"   # Horizontal
    chart.style = 10
    chart.title = "Top 3 Products"
    chart.y_axis.title = "Products"
    chart.x_axis.title = "Units"

    data = Reference(ws,
                     min_col=2,
                     min_row=start_row+1,
                     max_row=data_row-1,
                     max_col=3)

    cats = Reference(ws,
                     min_col=1,
                     min_row=start_row+2,
                     max_row=data_row-1)

    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)

    chart.height = 7
    chart.width = 18

    ws.add_chart(chart, f"E{start_row+2}")

    auto_fit_columns(ws)
