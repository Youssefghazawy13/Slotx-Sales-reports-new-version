from openpyxl.styles import Font, PatternFill, Alignment
from utils.excel_helpers import auto_fit_columns


def create_sales_sheet(wb, brand_sales, mode):

    ws = wb.create_sheet("Sales")

    headers = [
        "Branch",
        "Brand",
        "Product",
        "Barcode",
        "Quantity",
        "Total Price"
    ]

    ws.append(headers)

    # =========================
    # HEADER STYLE
    # =========================

    header_fill = PatternFill(
        start_color="0A1F5C",
        end_color="0A1F5C",
        fill_type="solid"
    )

    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = Font(bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center")

    total_qty = 0
    total_price = 0

    # =========================
    # DATA ROWS
    # =========================

    for _, row in brand_sales.iterrows():

        branch = mode
        brand = row.get("brand", "")
        product = row.get("name_ar", "")
        barcode = row.get("barcode", "")
        qty = float(row.get("quantity", 0) or 0)
        price = float(row.get("total", 0) or 0)

        total_qty += qty
        total_price += price

        ws.append([
            branch,
            brand,
            product,
            barcode,
            qty,
            f"{price:,.2f} EGP"
        ])

    last_row = ws.max_row + 1

    # =========================
    # TOTAL ROW (NO WORD TOTAL)
    # =========================

    ws.append([
        "",
        "",
        "",
        "",
        f"Total={int(total_qty)}",
        f"Total={total_price:,.2f} EGP"
    ])

    for cell in ws[ws.max_row]:
        cell.font = Font(bold=True)

    # =========================
    # ZEBRA STYLE
    # =========================

    stripe_fill = PatternFill(
        start_color="E9EEF7",
        end_color="E9EEF7",
        fill_type="solid"
    )

    for row in range(2, ws.max_row):
        if row % 2 == 0:
            for col in range(1, ws.max_column + 1):
                ws.cell(row=row, column=col).fill = stripe_fill

    auto_fit_columns(ws)
