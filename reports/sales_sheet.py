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

    # ==============================
    # HEADER STYLE
    # ==============================

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
    total_money = 0

    # ==============================
    # DATA ROWS
    # ==============================

    for _, row in brand_sales.iterrows():

        branch_name = mode

        brand = row.get("brand", "")
        product = row.get("name_ar", "")
        barcode = row.get("barcode", "")
        qty = row.get("quantity", 0)
        total = row.get("total", 0)

        total_qty += qty
        total_money += total

        ws.append([
            branch_name,
            brand,
            product,
            barcode,
            qty,
            total
        ])

    # ==============================
    # TOTAL ROW
    # ==============================

    ws.append([
        "",
        "",
        "",
        "TOTAL",
        total_qty,
        total_money
    ])

    total_row = ws.max_row

    for col in range(1, ws.max_column + 1):
        ws.cell(row=total_row, column=col).font = Font(bold=True)

    # ==============================
    # ZEBRA STRIPES
    # ==============================

    stripe_fill = PatternFill(
        start_color="E9EEF7",
        end_color="E9EEF7",
        fill_type="solid"
    )

    for row in range(2, total_row):

        if row % 2 == 0:
            for col in range(1, ws.max_column + 1):
                ws.cell(row=row, column=col).fill = stripe_fill

    # ==============================
    # NUMBER FORMAT
    # ==============================

    for row in ws.iter_rows(min_row=2, min_col=6, max_col=6):
        for cell in row:
            cell.number_format = '#,##0.00'

    # ==============================
    # AUTO FIT
    # ==============================

    auto_fit_columns(ws)
