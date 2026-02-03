from openpyxl.styles import Font

def build_sales(wb, sales_df, brand, branch):
    ws = wb.create_sheet("Sales")
    ws.append(["Branch", "Brand", "Product", "Barcode", "Quantity", "Price"])

    total_qty = 0
    total_money = 0

    for _, r in sales_df[sales_df["brand"] == brand].iterrows():
        ws.append([
            branch,
            brand,
            r.get("product"),
            r.get("barcode"),
            r.get("quantity"),
            r.get("total")
        ])
        total_qty += r.get("quantity", 0)
        total_money += r.get("total", 0)

    ws.append(["", "", "", "Total", f"Total={total_qty}", f"Total={total_money}"])

    for cell in ws[1]:
        cell.font = Font(bold=True)
# sales logic
