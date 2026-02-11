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
    # HEADER STYLE (SAME AS INVENTORY)
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

    # =========================
    # DATA ROWS
    # =========================

    total_qty = 0
    total_money = 0

    for _, row in brand_sales.iterrows():

        qty = float(row.get("quantity", 0) or 0)
        total = float(row.get("total", 0) or 0)

        total_qty += qty
        total_money += total

        ws.append([
            mode,
            row.get("brand", ""),
            row.get("product_name", ""),
            row.get("barcode", ""),
            qty,
            total
        ])

    # =========================
    # TOTAL ROW (INSIDE TABLE)
    # =========================

    ws.append([
        "",
        "",
        "",
        "",
        f"Total={int(total_qty)}",
        total_money
    ])

    last_row = ws.max_row

    # =========================
    # ZEBRA STYLE (SAME AS INVENTORY)
    # =========================

    stripe_fill = PatternFill(
        start_color="E9EEF7",
        end_color="E9EEF7",
        fill_type="solid"
    )

    for row in range(2, last_row + 1):
        if row % 2 == 0:
            for col in range(1, ws.max_column + 1):
                ws.cell(row=row, column=col).fill = stripe_fill

    # =========================
    # NUMBER FORMATTING
    # =========================

    # Quantity column
    for row in ws.iter_rows(min_row=2, min_col=5, max_col=5):
        for cell in row:
            if isinstance(cell.value, (int, float)):
                cell.number_format = '#,##0'

    # Total Price column
    for row in ws.iter_rows(min_row=2, min_col=6, max_col=6):
        for cell in row:
            if isinstance(cell.value, (int, float)):
                cell.number_format = '#,##0.00 "EGP"'

    auto_fit_columns(ws)