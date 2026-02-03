from openpyxl.styles import Font

def build_inventory(wb, inv_df, brand, branch):
    ws = wb.create_sheet("Inventory")
    ws.append(["Branch", "Brand", "Product", "Barcode", "Price", "Quantity"])

    for _, r in inv_df[inv_df["brand"] == brand].iterrows():
        ws.append([
            branch,
            brand,
            r.get("product"),
            r.get("barcode"),
            r.get("price"),
            r.get("quantity")
        ])

    for cell in ws[1]:
        cell.font = Font(bold=True)
# inventory logic
