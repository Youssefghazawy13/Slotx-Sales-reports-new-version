from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from utils.excel_helpers import auto_fit_columns
from core.kpi_engine import apply_deal


# ============================================
# Helper Functions
# ============================================

def extract_best_selling_size(brand_sales):

    size_counter = {}

    for _, row in brand_sales.iterrows():
        name = str(row.get("product_name", ""))
        qty = row.get("quantity", 0)

        if "-" in name:
            size = name.split("-")[-1].strip()
            if size:
                size_counter[size] = size_counter.get(size, 0) + qty

    if not size_counter:
        return ""

    return max(size_counter, key=size_counter.get)


def extract_best_selling_product(brand_sales):

    product_sales = (
        brand_sales.groupby("product_name")["quantity"]
        .sum()
        .sort_values(ascending=False)
    )

    if product_sales.empty:
        return ""

    return product_sales.index[0]


# ============================================
# Main Report Sheet
# ============================================

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

    # ============================================
    # FIX KPI COLUMN WIDTHS
    # ============================================

    ws.column_dimensions["A"].width = 5
    ws.column_dimensions["B"].width = 25
    ws.column_dimensions["C"].width = 5
    ws.column_dimensions["D"].width = 25

    # ============================================
    # KPI CARDS (ROW 1)
    # ============================================

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

    # ============================================
    # REPORT BODY (OLD STYLE WITH SPACING)
    # ============================================

    row_pointer = 4

    def write_row(label, value):
        nonlocal row_pointer

        ws.cell(row=row_pointer, column=1,
                value=label).font = Font(bold=True)
        ws.cell(row=row_pointer, column=2,
                value=value)

        row_pointer += 1

    best_size = extract_best_selling_size(brand_sales)
    best_product = extract_best_selling_product(brand_sales)

    # Brand Deal Text
    if percentage and rent:
        brand_deal_text = f"{rent} EGP + {percentage}% Deducted From The Sales"
    elif percentage:
        brand_deal_text = f"{percentage}% Deducted From The Sales"
    elif rent:
        brand_deal_text = f"{rent} EGP"
    else:
        brand_deal_text = "No Deal"

    # ----- Basic Info -----
    write_row("Branch Name:", mode)
    write_row("Brand Name:", brand_name)
    write_row("Brand Deal:", brand_deal_text)
    write_row("Payout Cycle:", payout_cycle)

    row_pointer += 1  # blank row

    # ----- Performance -----
    write_row("Best Selling Size:", best_size)
    write_row("Best Selling Product:", best_product)

    row_pointer += 2  # blank rows

    # ----- Inventory -----
    write_row("Total Brand Inventory Quantities:",
              total_inventory_qty)

    write_row("Total Brand Inventory Stock Price:",
              f"{total_inventory_value:,.2f} EGP")

    row_pointer += 2  # blank rows

    # ----- Sales -----
    write_row("Total Sales Quantity:",
              total_sales_qty)

    write_row("Total Sales (Money):",
              f"{total_sales_money:,.2f} EGP")

    row_pointer += 1  # blank row

    write_row("Total Sales After Percentage:",
              f"{after_percentage:,.2f} EGP")

    write_row("Total Sales After Rent:",
              f"{after_rent:,.2f} EGP")

    # ============================================
    # AUTO FIT
    # ============================================

    auto_fit_columns(ws)

    # Re-fix KPI width after auto-fit
    ws.column_dimensions["B"].width = 25
    ws.column_dimensions["D"].width = 25
