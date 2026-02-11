from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


def auto_fit(ws):
    for col in ws.columns:
        max_length = 0
        col_letter = get_column_letter(col[0].column)

        for cell in col:
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))

        ws.column_dimensions[col_letter].width = max_length + 3


def create_sales_sheet(wb, brand_sales, mode):
    ws = wb.create_sheet("Sales")

    header_fill = PatternFill(
        start_color="1F4E78",
        end_color="1F4E78",
        fill_type="solid"
    )

    header_font = Font(
        bold=True,
        color="FFFFFF"
    )

    thin_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin")
    )

    headers = [
        "Branch",
        "Brand",
        "Product",
        "Barcode",
        "Quantity",
        "Total Price"
    ]

    ws.append(headers)

    for col in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border

    total_qty = 0
    total_money = 0

    for _, row in brand_sales.iterrows():
        qty = row.get("quantity", 0)
        price = row.get("total", 0)

        total_qty += qty
        total_money += price

        ws.append([
            mode,  # ده كان branch_name
            row.get("brand", ""),
            row.get("product_name", ""),
            row.get("barcode", ""),
            qty,
            f"{price:,.2f} EGP"
        ])

    total_row_index = ws.max_row + 1

    ws.cell(row=total_row_index, column=5).value = f"Total={int(total_qty)}"
    ws.cell(row=total_row_index, column=6).value = f"Total={total_money:,.2f} EGP"

    for col in range(1, 7):
        cell = ws.cell(row=total_row_index, column=col)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border

    auto_fit(ws)