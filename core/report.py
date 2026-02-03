from openpyxl.styles import Font


def _find_col(df, keywords):
    """
    Find first column containing any keyword (case-insensitive)
    """
    for col in df.columns:
        name = col.lower()
        for k in keywords:
            if k in name:
                return col
    return None


def build_report(wb, sales_df, inventory_df, deals_df, brand, branch, payout_cycle):
    ws = wb.create_sheet("Report")

    # =========================
    # COLUMN DETECTION (SAFE)
    # =========================
    sales_brand_col = _find_col(sales_df, ["brand"])
    sales_qty_col = _find_col(sales_df, ["qty", "quantity"])
    sales_total_col = _find_col(sales_df, ["total", "amount", "price"])
    sales_product_col = _find_col(sales_df, ["product", "item", "name"])

    inv_brand_col = _find_col(inventory_df, ["brand"])
    inv_qty_col = _find_col(inventory_df, ["qty", "quantity"])
    inv_price_col = _find_col(inventory_df, ["price", "cost"])

    deal_brand_col = deals_df.columns[0]
    deal_pct_col = _find_col(deals_df, ["percent"])
    deal_rent_col = _find_col(deals_df, ["rent"])

    # =========================
    # FILTER DATA
    # =========================
    brand_sales = sales_df[sales_df[sales_brand_col] == brand]
    brand_inventory = inventory_df[inventory_df[inv_brand_col] == brand]
    deal_row = deals_df[deals_df[deal_brand_col] == brand]

    pct = deal_row[deal_pct_col].iloc[0] if deal_pct_col and not deal_row.empty else 0
    rent = deal_row[deal_rent_col].iloc[0] if deal_rent_col and not deal_row.empty else 0

    # =========================
    # DEAL TEXT
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
    # CALCULATIONS
    # =========================
    total_sales_qty = brand_sales[sales_qty_col].sum() if sales_qty_col else 0
    total_sales_money = brand_sales[sales_total_col].sum() if sales_total_col else 0

    after_percentage = total_sales_money - (total_sales_money * pct / 100)
    after_rent = after_percentage - rent

    total_inventory_qty = brand_inventory[inv_qty_col].sum() if inv_qty_col else 0
    total_inventory_price = (
        (brand_inventory[inv_qty_col] * brand_inventory[inv_price_col]).sum()
        if inv_qty_col and inv_price_col else 0
    )

    # =========================
    # BEST SELLING PRODUCT
    # =========================
    best_product = "N/A"
    if sales_product_col and sales_qty_col and not brand_sales.empty:
        grouped = (
            brand_sales
            .groupby(sales_product_col)[sales_qty_col]
            .sum()
        )
        if not grouped.empty:
            best_product = grouped.idxmax()

    # =========================
    # WRITE REPORT
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

    for r in rows:
        ws.append(r)

    for cell in ws["A"]:
        cell.font = Font(bold=True)
