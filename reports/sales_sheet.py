from openpyxl.styles import Font
from openpyxl.worksheet.table import Table, TableStyleInfo
from utils.excel_helpers import auto_fit_columns


def create_sales_sheet(
    wb,
    brand_sales,
    mode: str
):

    ws = wb.create_sheet("Sales")

    headers = [
        "Branch Name",
        "Brand Name",
        "Product",
        "Barcode",
        "Quantity",
        "Total Price"
    ]

    ws.append(headers)

    for cell in ws[1]:
        cell.font = Font(bold=True)

    total_qty = 0
    total_money = 0

    for _, row in brand_sales.iterrows():

        branch = mode
        brand = row.get("brand", "")
        product = row.get("product_name", "")
        barcode = row.get("barcode", "")
        qty = row.get("quantity", 0)
        total = row.get("total", 0)

        ws.append([
            branch,
            brand,
            product,
            barcode,
            qty,
            total
        ])

        total_qty += qty
        total_money += total

    # =====================================
    # Total Row (Only One)
    # =====================================

    total_row_index = ws.max_row + 1

    ws.append([
        "",
        "",
        "",
        "TOTAL",
        total_qty,
        total_money
    ])

    for cell in ws[total_row_index]:
        cell.font = Font(bold=True)

    # =====================================
    # Thousand Separator
    # =====================================

    for row in ws.iter_rows(min_row=2, min_col=6, max_col=6):
        for cell in row:
            cell.number_format = '#,##0.00'

    # =====================================
    # Freeze Header
    # =====================================

    ws.freeze_panes = "A2"

    # =====================================
    # Excel Table Style
    # =====================================

    table = Table(
        displayName="SalesTable",
        ref=f"A1:F{ws.max_row}"
    )

    style = TableStyleInfo(
        name="TableStyleMedium9",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False
    )

    table.tableStyleInfo = style
    ws.add_table(table)

    auto_fit_columns(ws)
