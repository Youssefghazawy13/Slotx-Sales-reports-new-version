from openpyxl.styles import Font, PatternFill, Alignment
from utils.excel_helpers import auto_fit_columns


def get_status(qty):

    if qty <= 2:
        return "Critical"
    elif qty <= 5:
        return "Low"
    elif qty <= 10:
        return "Medium"
    else:
        return "Good"


def get_note(status):

    if status == "Critical":
        return "Brand requires urgent restocking"

    elif status == "Low":
        return "Stock level is low – restock recommended"

    elif status == "Medium":
        return "Stock level is moderate – monitor movement"

    elif status == "Good":
        return "Stock level is healthy"

    return ""


def create_inventory_sheet(wb, brand_inventory, mode):

    ws = wb.create_sheet("Inventory")

    is_merged = mode.lower() == "merged"

    if is_merged:
        headers = [
            "Product",
            "Barcode",
            "Price",
            "Alexandria Qty",
            "Zamalek Qty",
            "Total Qty",
            "Status",
            "Notes"
        ]
    else:
        headers = [
            "Product",
            "Barcode",
            "Price",
            "Quantity",
            "Status",
            "Notes"
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

    # =========================
    # DATA ROWS
    # =========================

    for _, row in brand_inventory.iterrows():

        product = row.get("name_en", "")
        barcode = row.get("barcodes", "")
        price = float(row.get("sale_price", 0) or 0)

        if is_merged:

            alex = float(row.get("alex_qty", 0) or 0)
            zam = float(row.get("zamalek_qty", 0) or 0)
            total = float(row.get("available_quantity", 0) or 0)

            status = get_status(total)
            note = get_note(status)

            ws.append([
                product,
                barcode,
                price,
                alex,
                zam,
                total,
                status,
                note
            ])

        else:

            qty = float(row.get("available_quantity", 0) or 0)
            status = get_status(qty)
            note = get_note(status)

            ws.append([
                product,
                barcode,
                price,
                qty,
                status,
                note
            ])

    last_row = ws.max_row

    # =========================
    # ZEBRA ROWS
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

    # Price column
    for row in ws.iter_rows(min_row=2, min_col=3, max_col=3):
        for cell in row:
            cell.number_format = '#,##0.00'

    # Quantity columns
    if is_merged:
        qty_cols = [4, 5, 6]
    else:
        qty_cols = [4]

    for col in qty_cols:
        for row in ws.iter_rows(min_row=2, min_col=col, max_col=col):
            for cell in row:
                cell.number_format = '#,##0'

    auto_fit_columns(ws)
