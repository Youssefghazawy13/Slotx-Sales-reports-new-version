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

    # HEADER STYLE (same as inventory)
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

    # IMPORTANT: use .get() instead of direct access
    for _, row in brand_sales.iterrows():

        product = row.get("Product", "")
        barcode = row.get("Barcode", "")
        qty = float(row.get("Quantity", 0) or 0)
        total = float(row.get("Total Price", 0) or 0)
        brand = row.get("Brand", "")

        total_qty += qty
        total_money += total

        ws.append([
            mode,
            brand,
            product,
            barcode,
            qty,
            total
        ])

    # TOTAL ROW inside table
    ws.append([
        "",
        "",
        "",
        "",
        f"Total={int(total_qty)}",
        f"Total={total_money:,.2f} EGP"
    ])

    last_row = ws.max_row

    # Zebra style (same as inventory)
    stripe_fill = PatternFill(
        start_color="E9EEF7",
        end_color="E9EEF7",
        fill_type="solid"
    )

    for r in range(2, last_row + 1):
        if r % 2 == 0:
            for c in range(1, ws.max_column + 1):
                ws.cell(row=r, column=c).fill = stripe_fill

    # Quantity format
    for row in ws.iter_rows(min_row=2, max_row=last_row-1, min_col=5, max_col=5):
        for cell in row:
            cell.number_format = '#,##0'

    # Price format
    for row in ws.iter_rows(min_row=2, max_row=last_row-1, min_col=6, max_col=6):
        for cell in row:
            cell.number_format = '#,##0.00 "EGP"'

    auto_fit_columns(ws)