from openpyxl.styles import Font

def build_report(wb, sales, inventory, deals, brand, branch, payout_cycle):
    ws = wb.create_sheet("Report")

    # =========================
    # SAFE DEAL COLUMN DETECTION
    # =========================
    deal_cols = [c.lower() for c in deals.columns]

    # brand column
    brand_col = deals.columns[0]

    # percentage column
    pct_col = next(
        (c for c in deals.columns if "percent" in c.lower()),
        None
    )

    # rent column
    rent_col = next(
        (c for c in deals.columns if "rent" in c.lower()),
        None
    )

    deal_row = deals[deals[brand_col] == brand]

    pct = deal_row[pct_col].iloc[0] if pct_col and not deal_row.empty else 0
    rent = deal_row[rent_col].iloc[0] if rent_col and not deal_row.empty else 0

    # =========================
    # SALES TOTALS
    # =========================
    brand_sales = sales[sales["brand"] == brand]

    total_qty = brand_sales["quantity"].sum()
    total_money = brand_sales["total"].sum()

    after_pct = total_money - (total_money * pct / 100)
    after_rent = after_pct - rent

    # =========================
    # INVENTORY TOTAL
    # =========================
    brand_inventory = inventory[inventory["brand"] == brand]
    inv_qty = brand_inventory.iloc[:, -1].sum()

    # =========================
    # WRITE REPORT
    # =========================
    rows = [
        ("Branch Name", branch),
        ("Brand Name", brand),
        ("Payout Cycle", payout_cycle),
        ("Brand Deal", f"{pct}% + {rent}"),
        ("", ""),
        ("Total Brand Inventory Quantities", inv_qty),
        ("", ""),
        ("Total Sales (Products Quantities)", total_qty),
        ("Total Sales (Money)", total_money),
        ("Total Sales After Percentage", after_pct),
        ("Total Sales After Rent", after_rent),
    ]

    for r in rows:
        ws.append(r)

    for cell in ws["A"]:
        cell.font = Font(bold=True)
