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

    # =====================================================
    # GROUP SALES PER BRAND
    # =====================================================

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

    sales_grouped["Rank"] = sales_grouped.index + 1

    # =====================================================
    # GROUP INVENTORY PER BRAND (FIXED)
    # =====================================================

    inventory_grouped = (
        inventory_df.groupby("brand")
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

    # =====================================================
    # MERGE CORRECTLY ON BRAND
    # =====================================================

    summary_df = sales_grouped.merge(
        inventory_grouped,
        on="brand",
        how="left"
    )

    summary_df["available_quantity"] = summary_df["available_quantity"].fillna(0)
    summary_df["inventory_value"] = summary_df["inventory_value"].fillna(0)

    # =====================================================
    # APPLY DEALS
    # =====================================================

    after_percentage_list = []
    after_rent_list = []
    after_all_list = []

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
        after_all_list.append(after_rent)

    summary_df["after_percentage"] = after_percentage_list
    summary_df["after_rent"] = after_rent_list
    summary_df["after_all"] = after_all_list

    # =====================================================
# KPI CARDS (5 SIDE BY SIDE)
# =====================================================

total_sales_money = summary_df["total"].sum()

total_percentage_deduction = (
    summary_df["total"] - summary_df["after_percentage"]
).sum()

total_rent_deduction = (
    summary_df["after_percentage"] - summary_df["after_rent"]
).sum()

total_after_all = summary_df["after_all"].sum()

total_deductions = (
    total_percentage_deduction + total_rent_deduction
)

kpi_fill = PatternFill(
    start_color="0A1F5C",
    end_color="0A1F5C",
    fill_type="solid"
)

kpi_titles = [
    "Total Branch Sales",
    "Total % Deducted",
    "Total Rent Deducted",
    "Total Deductions",
    "Sales After All Deductions"
]

kpi_values = [
    total_sales_money,
    total_percentage_deduction,
    total_rent_deduction,
    total_deductions,
    total_after_all
]

for i, (title, value) in enumerate(zip(kpi_titles, kpi_values)):

    col_letter = chr(65 + i)  # A, B, C, D, E

    ws[f"{col_letter}1"] = title
    ws[f"{col_letter}2"] = value

    ws[f"{col_letter}1"].fill = kpi_fill
    ws[f"{col_letter}2"].fill = kpi_fill

    ws[f"{col_letter}1"].font = Font(bold=True, color="FFFFFF")
    ws[f"{col_letter}2"].font = Font(bold=True, color="FFFFFF")

    ws[f"{col_letter}1"].alignment = Alignment(horizontal="center")
    ws[f"{col_letter}2"].alignment = Alignment(horizontal="center")

    ws[f"{col_letter}2"].number_format = '#,##0.00 "EGP"'

    ws.column_dimensions[col_letter].width = 26

    # =====================================================
    # PERFORMANCE TABLE
    # =====================================================

    ws.append([])
    ws.append([])

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

    for _, row in summary_df.iterrows():
        ws.append([
            row["Rank"],
            row["brand"],
            row["quantity"],
            row["total"],
            row["after_percentage"],
            row["after_rent"],
            row["after_all"],
            row["available_quantity"],
            row["inventory_value"]
        ])

    last_row = ws.max_row

    stripe_fill = PatternFill(
        start_color="E9EEF7",
        end_color="E9EEF7",
        fill_type="solid"
    )

    for r in range(header_row + 1, last_row + 1):
        if r % 2 == 0:
            for c in range(1, 10):
                ws.cell(row=r, column=c).fill = stripe_fill

    # Currency Formatting
    for col in [4, 5, 6, 7, 9]:
        for r in range(header_row + 1, last_row + 1):
            ws.cell(row=r, column=col).number_format = '#,##0.00 "EGP"'

    auto_fit_columns(ws)