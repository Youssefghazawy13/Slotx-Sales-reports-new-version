from openpyxl.styles import Font, PatternFill, Alignment
from utils.excel_helpers import auto_fit_columns


def extract_best_selling_product(brand_sales):

    if brand_sales.empty:
        return "", 0

    grouped = (
        brand_sales.groupby("name_ar")["quantity"]
        .sum()
        .sort_values(ascending=False)
    )

    if grouped.empty:
        return "", 0

    return grouped.index[0], grouped.iloc[0]


def extract_best_selling_size(brand_sales):

    if brand_sales.empty:
        return ""

    size_sales = {}

    for _, row in brand_sales.iterrows():

        product_name = str(row.get("name_ar", ""))
        qty = row.get("quantity", 0)

        if "-" in product_name:
            size = product_name.split("-")[-1].strip()
            size_sales[size] = size_sales.get(size, 0) + qty

    if not size_sales:
        return ""

    return max(size_sales, key=size_sales.get)


def create_report_sheet(
    wb,
    brand_name,
    mode,
    payout_cycle,
    brand_sales,
    brand_inventory,
    deals_dict
):

    ws = wb.create_sheet("Report")

    # =========================
    # CALCULATIONS
    # =========================

    total_sales_qty = brand_sales["quantity"].sum() if not brand_sales.empty else 0
    total_sales_money = brand_sales["total"].sum() if not brand_sales.empty else 0

    if not brand_inventory.empty:

        brand_inventory["sale_price"] = (
            brand_inventory["sale_price"]
            .fillna(0)
            .astype(float)
        )

        brand_inventory["available_quantity"] = (
            brand_inventory["available_quantity"]
            .fillna(0)
            .astype(float)
        )

        total_inventory_qty = brand_inventory["available_quantity"].sum()

        total_inventory_value = (
            brand_inventory["sale_price"] *
            brand_inventory["available_quantity"]
        ).sum()

    else:
        total_inventory_qty = 0
        total_inventory_value = 0

    deal = deals_dict.get(brand_name, {"percentage": 0, "rent": 0})

    percentage = deal["percentage"]
    rent = deal["rent"]

    after_percentage = total_sales_money - (total_sales_money * percentage / 100)
    after_rent = after_percentage - rent

    best_product, best_qty = extract_best_selling_product(brand_sales)
    best_size = extract_best_selling_size(brand_sales)

    # =========================
    # KPI CARDS (B & C)
    # =========================

    kpi_fill = PatternFill(
        start_color="0A1F5C",
        end_color="0A1F5C",
        fill_type="solid"
    )

    ws["B1"] = "Total Sales"
    ws["B2"] = total_sales_money

    ws["C1"] = "Inventory Value"
    ws["C2"] = total_inventory_value

    for col in ["B", "C"]:
        ws[f"{col}1"].font = Font(color="FFFFFF", bold=True)
        ws[f"{col}2"].font = Font(color="FFFFFF", bold=True)

        ws[f"{col}1"].alignment = Alignment(horizontal="center")
        ws[f"{col}2"].alignment = Alignment(horizontal="center")

        ws[f"{col}1"].fill = kpi_fill
        ws[f"{col}2"].fill = kpi_fill

    ws["B2"].number_format = '#,##0.00'
    ws["C2"].number_format = '#,##0.00'

    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 18

    # =========================
    # DETAILS SECTION
    # =========================

    start_row = 5

    report_data = [
        ("Branch Name:", mode),
        ("Brand Name:", brand_name),
        ("Payout Cycle:", payout_cycle),
        ("", ""),
        ("Total Inventory Quantity:", total_inventory_qty),
        ("Total Inventory Value:", total_inventory_value),
        ("", ""),
        ("Best Selling Product:", f"{best_product} ({best_qty})"),
        ("Best Selling Size:", best_size),
        ("", ""),
        ("Total Sales Quantity:", total_sales_qty),
        ("Total Sales Money:", total_sales_money),
        ("After Percentage:", after_percentage),
        ("After Rent:", after_rent),
    ]

    for label, value in report_data:

        ws[f"B{start_row}"] = label
        ws[f"C{start_row}"] = value

        ws[f"B{start_row}"].font = Font(bold=True)

        if isinstance(value, (int, float)):
            ws[f"C{start_row}"].number_format = '#,##0.00'

        start_row += 1

    auto_fit_columns(ws)
