from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from utils.excel_helpers import auto_fit_columns
from core.kpi_engine import apply_deal


# =====================================================
# Helpers
# =====================================================

def detect_product_column(df):
    possible_cols = [
        "product_name",
        "Product Name",
        "name_ar",
        "Product",
        "product"
    ]

    for col in possible_cols:
        if col in df.columns:
            return col

    return None


def extract_best_selling_product(brand_sales):

    if brand_sales.empty:
        return ""

    product_col = detect_product_column(brand_sales)

    if product_col is None or "quantity" not in brand_sales.columns:
        return ""

    grouped = (
        brand_sales
        .groupby(product_col)["quantity"]
        .sum()
        .sort_values(ascending=False)
    )

    if grouped.empty:
        return ""

    return grouped.index[0]


def extract_best_selling_size(brand_sales):

    if brand_sales.empty:
        return ""

    product_col = detect_product_column(brand_sales)

    if product_col is None or "quantity" not in brand_sales.columns:
        return ""

    size_counter = {}

    for _, row in brand_sales.iterrows():

        name = str(row.get(product_col, ""))
        qty = row.get("quantity", 0)

        if "-" in name:
            size = name.split("-")[-1].strip()
            if size:
                size_counter[size] = size_counter.get(size, 0) + qty

    if not size_counter:
        return ""

    return max(size_counter, key=size_counter.get)


# =====================================================
# Main Report Sheet
# =====================================================

def create_report_sheet(
    wb,
    brand_name,
    mode,
    payout_cycle,
    brand_sales,
    total_inventory_qty,
    total_inventory_value,
    total_sales_qty,
    total_sales_money,
    deals_dict
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
    # KPI CARDS (Row 1)
    # =====================================================

    ws.column_dimensions["B"].width = 25
    ws.column_dimensions["D"].width = 25

    def create_kpi_card(row, col_letter, title, value):

        cell = ws[f"{col_letter}{row}"]
        cell.value = f"{title}\n{value}"

        cell.font = Font(size=12, bold=True, color="FFFFFF")
        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True
        )

        ws.row_dimensions[row].height = 45

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

        cell.fill = fill
        cell.border = border

    create_kpi_card(
        1,
        "B",
        "Total Sales",
        f"{total_sales_money:,.2f} EGP"
    )

    create_kpi_card(
        1,
        "D",
        "Net After Deal",
        f"{after_rent:,.2f} EGP"
    )

    # =====================================================
    # REPORT BODY (OLD ORDER + SPACING)
    # =====================================================

    row_pointer = 4

    def write_row(label, value):
        nonlocal row_pointer

        ws.cell(
            row=row_pointer,
            column=1,
            value=label
        ).font = Font(bold=True)

        ws.cell(
            row=row_pointer,
            column=2,
            value=value
        )

        row_pointer += 1

    best_product = extract_best_selling_product(brand_sales)
    best_size = extract_best_selling_size(brand_sales)

    if percentage and rent:
        brand_deal_text = f"{rent} EGP + {percentage}% Deducted From The Sales"
    elif percentage:
        brand_deal_text = f"{percentage}% Deducted From The Sales"
    elif rent:
        brand_deal_text = f"{rent} EGP"
    else:
        brand_deal_text = "No Deal"

    # ---- Basic Info ----
    write_row("Branch Name:", mode)
    write_row("Brand Name:", brand_name)
    write_row("Brand Deal:", brand_deal_text)
    write_row("Payout Cycle:", payout_cycle)

    row_pointer += 1

    # ---- Performance ----
    write_row("Best Selling Size:", best_size)
    write_row("Best Selling Product:", best_product)

    row_pointer += 2

    # ---- Inventory ----
    write_row(
        "Total Brand Inventory Quantities:",
        total_inventory_qty
    )

    write_row(
        "Total Brand Inventory Stock Price:",
        f"{total_inventory_value:,.2f} EGP"
    )

    row_pointer += 2

    # ---- Sales ----
    write_row(
        "Total Sales Quantity:",
        total_sales_qty
    )

    write_row(
        "Total Sales (Money):",
        f"{total_sales_money:,.2f} EGP"
    )

    row_pointer += 1

    write_row(
        "Total Sales After Percentage:",
        f"{after_percentage:,.2f} EGP"
    )

    write_row(
        "Total Sales After Rent:",
        f"{after_rent:,.2f} EGP"
    )

    auto_fit_columns(ws)

    ws.column_dimensions["B"].width = 25
    ws.column_dimensions["D"].width = 25
