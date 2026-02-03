from openpyxl.styles import Font

def build_report(wb, sales, inventory, deals, brand, branch, payout_cycle):
    ws = wb.create_sheet("Report")

    deal_row = deals[deals["brand"] == brand]
    pct = deal_row["percentage"].iloc[0] if not deal_row.empty else 0
    rent = deal_row["rent"].iloc[0] if not deal_row.empty else 0

    total_qty = sales[sales["brand"] == brand]["quantity"].sum()
    total_money = sales[sales["brand"] == brand]["total"].sum()

    after_pct = total_money - (total_money * pct / 100)
    after_rent = after_pct - rent

    rows = [
        ("Branch Name", branch),
        ("Brand Name", brand),
        ("Payout Cycle", payout_cycle),
        ("Brand Deal", f"{pct}% + {rent}"),
        ("", ""),
        ("Total Sales Quantity", total_qty),
        ("Total Sales Money", total_money),
        ("After Percentage", after_pct),
        ("After Rent", after_rent),
    ]

    for r in rows:
        ws.append(r)

    for cell in ws["A"]:
        cell.font = Font(bold=True)
# report logic
