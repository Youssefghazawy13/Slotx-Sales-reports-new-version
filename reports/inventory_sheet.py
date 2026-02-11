from openpyxl.styles import Font
from openpyxl.worksheet.table import Table, TableStyleInfo
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

    # Bold header
    for cell in ws[1]:
        cell.font = Font(bold=True)

    for _, row in brand_inventory.iterrows():

        product = row.get("name_en", "")
        barcode = row.get("barcodes", "")
        price = row.get("sale_price", 0)

        if is_merged:

            alex = row.get("alex_qty", 0)
            zam = row.get("zamalek_qty", 0)
            total = row.get("available_quantity", 0)

            status = get_status(total)

            ws.append([
                product,
                barcode,
                price,
                alex,
                zam,
                total,
                status,
                ""
            ])

        else:

            qty = row.get("available_quantity", 0)
            status = get_status(qty)

            ws.append([
                product,
                barcode,
                price,
                qty,
                status,
                ""
            ])

    # ==============================
    # TABLE STYLE (NO FILTERS)
    # ==============================

    last_row = ws.max_row
    last_col = ws.max_column

    table = Table(
        displayName="InventoryTable",
        ref=f"A1:{chr(64+last_col)}{last_row}"
    )

    style = TableStyleInfo(
        name="TableStyleMedium9",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False
    )

    table.tableStyleInfo = style
    table.autoFilter = None  # 🔥 Remove filter arrows

    ws.add_table(table)

    auto_fit_columns(ws)
