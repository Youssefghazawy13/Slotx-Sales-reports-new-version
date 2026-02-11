from openpyxl.styles import Font, PatternFill, Alignment
from utils.excel_helpers import auto_fit_columns


def create_performance_sheet(
    wb,
    branch_name,
    payout_cycle,
    sales_df,
    inventory_df,
    deals_dict
):

    ws = wb.create_sheet("Performance")

    # ==============================
    # AGGREGATION PER BRAND
    # ==============================

    sales_grouped = (
        sales_df.groupby("brand")
        .agg({
            "quantity": "sum",
            "total": "sum"
        })
        .reset_index()
    )

    sales_grouped = sales_grouped.sort_values(
        by="total",
        ascending=False
    ).reset_index(drop=True)

    # Ranking
    sales_grouped["Rank"] = sales_grouped.index + 1

    # ==============================
    # INVENTORY CALCULATION
    # ==============================

    inventory_grouped = (
        inventory_df.groupby("name_en")
        .agg({
            "available_quantity": "sum",
            "sale_price": "mean"
        })
        .reset_index()
    )

    inventory_grouped["inventory_value"] = (
        inventory_grouped["available_quantity"] *
        inventory_grouped["sale_price"]
    )

    # ==============================
    # MERGE SALES + INVENTORY
    # ==============================

    summary_df = sales_grouped.merge(
        inventory_grouped,
        left_on="brand",
        right_on="name_en",
        how="left"
    )

    summary_df["inventory_value"] = summary_df["inventory_value"].fillna(0)
    summary_df["available_quantity"] = summary_df["available_quantity"].fillna(0)

    # ==============================
    # APPLY DEALS
    # ==============================

    after_percentage_list = []
    after_rent_list = []
    health_list = []

    for _, row in summary_df.iterrows():

        brand = row["brand"]
        sales_money = row["total"]

        deal = deals_dict.get(brand, {"percentage": 0, "rent": 0})
        percentage = deal["percentage"]
        rent = deal["rent"]

        after_percentage = sales_money - (sales_money * percentage / 100)
        after_rent = after_percentage - rent

        after_percentage_list.append(after_percentage)
        after_rent_list.append(after_rent)

        # Brand Health
        if percentage == 0 and rent == 0:
            health = "No Deal"
        elif row["available_quantity"] <= 2:
            health = "Critical"
        else:
            health = "Healthy"

        health_list.append(health)

    summary_df["after_percentage"] = after_percentage_list
    summary_df["after_rent"] = after_rent_list
    summary_df["health"] = health_list

    # ==============================
    # KPI CARDS
    # ==============================

    total_sales_money = summary_df["total"].sum()
    total_inventory_value = summary_df["inventory_value"].sum()
    total_after_rent = summary_df["after_rent"].sum()

    kpi_fill = PatternFill(
        start_color="0A1F5C",
        end_color="0A1F5C",
        fill_type="solid"
    )

    ws["A1"] = "Total Sales"
    ws["A2"] = total_sales_money

    ws["B1"] = "Inventory Value"
    ws["B2"] = total_inventory_value

    ws["C1"] = "Net After Deductions"
    ws["C2"] = total_after_rent

    for col in ["A", "B", "C"]:
        ws[f"{col}1"].fill = kpi_fill
        ws[f"{col}2"].fill = kpi_fill
        ws[f"{col}1"].font = Font(bold=True, color="FFFFFF")
        ws[f"{col}2"].font = Font(bold=True, color="FFFFFF")
        ws[f"{col}1"].alignment = Alignment(horizontal="center")
        ws[f"{col}2"].alignment = Alignment(horizontal="center")
        ws[f"{col}2"].number_format = '#,##0.00 "EGP"'
        ws.column_dimensions[col].width = 22

    # ==============================
    # TABLE HEADER
    # ==============================

    start_row = 5

    headers = [
        "Rank",
        "Brand",
        "Sales Qty",
        "Sales Money",
        "Inventory Qty",
        "Inventory Value",
        "After %",
        "After Rent",
        "Brand Health"
    ]

    ws.append([])
    ws.append(headers)

    header_row = ws.max_row

    header_fill = PatternFill(
        start_color="0A1F5C",
        end_color="0A1F5C",
        fill_type="solid"
    )

    for col in range(1, len(headers) + 1):
        cell = ws.cell(row=header_row, column=col)
        cell.fill = header_fill
        cell.font = Font(bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center")

    # ==============================
    # TABLE DATA
    # ==============================

    for _, row in summary_df.iterrows():
        ws.append([
            row["Rank"],
            row["brand"],
            row["quantity"],
            row["total"],
            row["available_quantity"],
            row["inventory_value"],
            row["after_percentage"],
            row["after_rent"],
            row["health"]
        ])

    last_row = ws.max_row

    # Zebra
    stripe_fill = PatternFill(
        start_color="E9EEF7",
        end_color="E9EEF7",
        fill_type="solid"
    )

    for r in range(header_row + 1, last_row + 1):
        if r % 2 == 0:
            for c in range(1, 10):
                ws.cell(row=r, column=c).fill = stripe_fill

    # Currency format
    for col in [4, 6, 7, 8]:
        for r in range(header_row + 1, last_row + 1):
            ws.cell(row=r, column=col).number_format = '#,##0.00 "EGP"'

    auto_fit_columns(ws)
