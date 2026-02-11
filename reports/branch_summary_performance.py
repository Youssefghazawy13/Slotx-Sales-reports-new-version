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

    ws = wb.create_sheet(f"{branch_name} Summary")

    # =============================
    # AGGREGATE PER BRAND
    # =============================

    grouped = (
        sales_df.groupby("brand")
        .agg({"quantity": "sum", "total": "sum"})
        .reset_index()
    )

    inventory_grouped = (
        inventory_df.groupby("brand")
        .agg({
            "available_quantity": "sum",
            "sale_price": "mean"
        })
        .reset_index()
    )

    inventory_grouped["inventory_value"] = (
        inventory_grouped["available_quantity"]
        * inventory_grouped["sale_price"]
    )

    df = grouped.merge(
        inventory_grouped,
        on="brand",
        how="left"
    ).fillna(0)

    total_percentage_deduction = 0
    total_rent_deduction = 0

    after_all_list = []

    for _, row in df.iterrows():

        brand = row["brand"]
        sales_money = row["total"]

        deal = deals_dict.get(brand, {"percentage": 0, "rent": 0})

        percentage = deal["percentage"]
        rent = deal["rent"]

        percentage_deduction = sales_money * percentage / 100

        after_all = sales_money - percentage_deduction - rent

        total_percentage_deduction += percentage_deduction
        total_rent_deduction += rent

        after_all_list.append(after_all)

    df["after_all"] = after_all_list

    df = df.sort_values(
        by="after_all",
        ascending=False
    ).reset_index(drop=True)

    df["Rank"] = df.index + 1

    total_sales = df["total"].sum()
    total_inventory_value = df["inventory_value"].sum()
    total_after_all = df["after_all"].sum()

    # =============================
    # KPI CARDS
    # =============================

    kpi_fill = PatternFill(
        start_color="0A1F5C",
        end_color="0A1F5C",
        fill_type="solid"
    )

    kpis = [
        ("Total Branch Sales", total_sales),
        ("Total Rent Deductions", total_rent_deduction),
        ("Total Percentage Deductions", total_percentage_deduction),
        ("Sales After All Deductions", total_after_all),
        ("Inventory Value", total_inventory_value),
    ]

    row = 1
    col = 1

    for title, value in kpis:

        ws.cell(row=row, column=col).value = title
        ws.cell(row=row+1, column=col).value = value

        ws.cell(row=row, column=col).font = Font(bold=True, color="FFFFFF")
        ws.cell(row=row+1, column=col).font = Font(bold=True, color="FFFFFF")

        ws.cell(row=row, column=col).fill = kpi_fill
        ws.cell(row=row+1, column=col).fill = kpi_fill

        ws.cell(row=row+1, column=col).number_format = '#,##0.00 "EGP"'

        col += 2

    # =============================
    # TABLE HEADER
    # =============================

    start_row = 5
    headers = [
        "Rank",
        "Brand",
        "Sales Qty",
        "Sales Money",
        "Inventory Qty",
        "Inventory Value",
        "After All Deductions"
    ]

    ws.append([])
    ws.append(headers)

    header_row = ws.max_row

    header_fill = PatternFill(
        start_color="0A1F5C",
        end_color="0A1F5C",
        fill_type="solid"
    )

    for c in range(1, len(headers) + 1):
        cell = ws.cell(row=header_row, column=c)
        cell.fill = header_fill
        cell.font = Font(bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center")

    # =============================
    # TABLE DATA
    # =============================

    for _, row in df.iterrows():
        ws.append([
            row["Rank"],
            row["brand"],
            row["quantity"],
            row["total"],
            row["available_quantity"],
            row["inventory_value"],
            row["after_all"]
        ])

    last_row = ws.max_row

    for r in range(header_row+1, last_row+1):
        ws[f"D{r}"].number_format = '#,##0.00 "EGP"'
        ws[f"F{r}"].number_format = '#,##0.00 "EGP"'
        ws[f"G{r}"].number_format = '#,##0.00 "EGP"'

    auto_fit_columns(ws)