from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.worksheet.table import Table, TableStyleInfo
from utils.excel_helpers import auto_fit_columns
import pandas as pd


def create_branch_performance_sheet(
    wb,
    branch_name,
    payout_cycle,
    sales_df,
    inventory_df,
    deals_dict
):

    ws = wb.create_sheet("Performance")

    # =====================================================
    # KPI CARDS (2 CARDS SIDE BY SIDE)
    # =====================================================

    total_sales_money = sales_df["total"].sum()
    total_sales_qty = sales_df["quantity"].sum()

    total_percentage_deduction = 0
    total_rent_deduction = 0

    for brand in sales_df["brand"].unique():
        brand_sales = sales_df[sales_df["brand"] == brand]
        brand_total = brand_sales["total"].sum()

        deal = deals_dict.get(brand, {"percentage": 0, "rent": 0})

        total_percentage_deduction += brand_total * (deal["percentage"] / 100)
        total_rent_deduction += deal["rent"]

    after_all = (
        total_sales_money
        - total_percentage_deduction
        - total_rent_deduction
    )

    # Card Style
    card_fill = PatternFill(start_color="0A1F5C", end_color="0A1F5C", fill_type="solid")

    # Card 1
    ws["A1"] = "Total Branch Sales"
    ws["A2"] = total_sales_money
    ws["A2"].number_format = '#,##0.00 "EGP"'

    # Card 2
    ws["C1"] = "Sales After All Deductions"
    ws["C2"] = after_all
    ws["C2"].number_format = '#,##0.00 "EGP"'

    for col in ["A", "C"]:
        for row in [1, 2]:
            cell = ws[f"{col}{row}"]
            cell.fill = card_fill
            cell.font = Font(color="FFFFFF", bold=True)
            cell.alignment = Alignment(horizontal="center")

    ws.row_dimensions[1].height = 25
    ws.row_dimensions[2].height = 25

    # =====================================================
    # PERFORMANCE TABLE
    # =====================================================

    start_row = 5

    table_data = []

    brands = sales_df["brand"].unique()

    for brand in brands:

        brand_sales = sales_df[sales_df["brand"] == brand]
        brand_inventory = inventory_df[inventory_df["brand"] == brand]

        sales_qty = brand_sales["quantity"].sum()
        sales_money = brand_sales["total"].sum()

        deal = deals_dict.get(brand, {"percentage": 0, "rent": 0})

        after_percentage = sales_money - (sales_money * deal["percentage"] / 100)
        after_rent = after_percentage - deal["rent"]
        after_all = after_rent

        inventory_qty = brand_inventory["available_quantity"].sum()
        inventory_value = (
            brand_inventory["available_quantity"]
            * brand_inventory["sale_price"]
        ).sum()

        table_data.append([
            brand,
            sales_qty,
            sales_money,
            after_percentage,
            after_rent,
            after_all,
            inventory_qty,
            inventory_value
        ])

    df = pd.DataFrame(
        table_data,
        columns=[
            "Brand",
            "Sales Qty",
            "Sales Money",
            "After Percentage",
            "After Rent",
            "After All Deductions",
            "Inventory Qty",
            "Inventory Value"
        ]
    )

    df = df.sort_values(by="Sales Money", ascending=False).reset_index(drop=True)
    df.insert(0, "Rank", df.index + 1)

    headers = [
        "Rank",
        "Brand",
        "Sales Qty",
        "Sales Money",
        "After Percentage",
        "After Rent",
        "After All Deductions",
        "Inventory Qty",
        "Inventory Value"
    ]

    ws.append([])
    ws.append(headers)

    header_row = start_row

    header_fill = PatternFill(start_color="0A1F5C", end_color="0A1F5C", fill_type="solid")

    for col in range(1, len(headers) + 1):
        cell = ws.cell(row=header_row, column=col)
        cell.fill = header_fill
        cell.font = Font(bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center")

    for _, row in df.iterrows():
        ws.append(row.tolist())

    end_row = ws.max_row

    table = Table(
        displayName="BranchPerformance",
        ref=f"A{header_row}:I{end_row}"
    )

    style = TableStyleInfo(
        name="TableStyleMedium2",
        showRowStripes=True
    )

    table.tableStyleInfo = style
    ws.add_table(table)

    # Number Formatting
    money_cols = [4, 5, 6, 7, 9]
    for col in money_cols:
        for row in ws.iter_rows(min_row=header_row + 1, min_col=col, max_col=col):
            for cell in row:
                cell.number_format = '#,##0.00 "EGP"'

    auto_fit_columns(ws)