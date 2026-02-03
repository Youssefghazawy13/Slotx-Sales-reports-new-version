from openpyxl.styles import Font


def build_report(wb, sales_df, inventory_df, deals_df, brand, branch, payout_cycle):
    ws = wb.create_sheet("Report")

    # =========================
    # SAFE DEAL DETECTION
    # =========================
    brand_col = deals_df.columns[0]

    pct_col = next(
        (c for c in deals_df.columns if "percent" in c.lower()),
        None
    )

    rent_col = next(
        (c for c in deals_df.columns if "rent" in c.lower()),
        None
    )

    deal_row = deals_df[deals_df[brand_col] == brand]

    pct = deal_row[pct_col].iloc[0] if pct_col and not deal_row.empty else 0
    rent = deal_row[rent_col].iloc[0] if rent_col and not deal_row.empty else 0

    # =========================
    # DEAL TEXT FORMAT (FINAL)
    # =========================
    if pct > 0 and rent > 0:
        deal_text = f"{pct}% + {rent} Deducted from the sales"
    elif pct > 0:
        deal_text = f"{pct}% Deducted from the sales"
    elif rent > 0:
        deal_text = f"{rent} Deducted from the sales"
    else:
        deal_text = "No Deal"

    # =========================
    # SALES CALCULATIONS
    # =========================
    brand_sales = sales_df[sales_df["brand"] == brand]

    total_sales_qty = brand_sales["quantity"].sum()
    total_sales_money = brand_sales["total"].sum()

    after_percentage = total_sales_money - (total_sales_money * pct / 100)
    after_rent = after_percentage - rent

    # =========================
    # INVENTORY CALCULATIONS
    # =========================
    brand_inventory = inventory_df[inventory_df["brand"] == brand]

    total_inventory_qty = (
        brand_inventory["quantity"].sum()
        if "quantity" in brand_inventory.columns
        else brand_inventory.iloc[:, -1].sum()
    )

    total_inventory_price = 0
    if "price" in brand_inventory.columns:
        total_inventory_price = (
            brand_inventory["price"] * brand_inventory["quantity"]
        ).sum()

    # =========================
    # BEST SELLING PRODUCT
    # =========================
    best_product = "N/A"
    if not brand_sales.empty:
        best_product = (
            brand_sales.groupby("product")["quantity"]
            .sum()
            .idxmax()
        )

    # =========================
    # WRITE REPORT (ORDERED)
    # =========================
    rows = [
        ("Branch Name", branch),
        ("Brand Name", brand),
        ("Payout Cycle", payout_cycle),
        ("Brand Deal", deal_text),
        ("", ""),

        ("Best Selling Product", best_product),
        ("", ""),

        ("Total Brand Inventory Quantities", total_inventory_qty),
        ("Total Brand Inventory Stock Price", total_inventory_price),
        ("", ""),

        ("Total Sales (Products Quantities)", total_sales_qty),
        ("Total Sales (Money)", total_sales_money),
        ("Total Sales After Percentage", after_percentage),
        ("Total Sales After Rent", after_rent),
    ]

    for row in rows:
        ws.append(row)

    # =========================
    # STYLING
    # =========================
    for cell in ws["A"]:
        cell.font = Font(bold=True)
