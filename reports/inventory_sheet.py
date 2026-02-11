from openpyxl.styles import Font, PatternFill
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

    if brand_inventory.empty:
        ws.append(["No Inventory Data"])
        return

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

    for cell in ws[1]:
        cell.font = Font(bold=True)

    for _, row in brand_inventory.iterrows():

        product = row.get("product_name", "")
        barcode = row.get("barcode", "")
        price = row.get("price", 0)

        if is_merged:

            alex = row.get("alex_qty", 0)
            zam = row.get("zamalek_qty", 0)
            total = row.get("quantity", 0)

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

            qty = row.get("quantity", 0)
            status = get_status(qty)

            ws.append([
                product,
                barcode,
                price,
                qty,
                status,
                ""
            ])

    # Color status
    for row in ws.iter_rows(min_row=2):

        status_cell = row[-2]

        if status_cell.value == "Critical":
            status_cell.fill = PatternFill("solid", fgColor="FFC7CE")
        elif status_cell.value == "Low":
            status_cell.fill = PatternFill("solid", fgColor="FFEB9C")
        elif status_cell.value == "Medium":
            status_cell.fill = PatternFill("solid", fgColor="C6EFCE")
        elif status_cell.value == "Good":
            status_cell.fill = PatternFill("solid", fgColor="A9D08E")

    auto_fit_columns(ws)
