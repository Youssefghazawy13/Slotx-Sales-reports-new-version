from openpyxl.styles import Font

def build_inventory(wb, inv_df, brand, branch):
    ws = wb.create_sheet("Inventory")

    headers = ["Branch", "Brand", "Product", "Barcode", "Price", "Quantity"]
    ws.append(headers)

    for c in ws[1]:
        c.font = Font(bold=True)

    brand_inv = inv_df[inv_df["brand"] == brand]

    for _, r in brand_inv.iterrows():
        ws.append([
            branch,
            brand,
            r.get("product"),
            r.get("barcode"),
            r.get("price"),
            r.get("quantity")
        ])
