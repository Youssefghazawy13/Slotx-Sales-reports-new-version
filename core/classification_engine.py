# core/classification_engine.py

def classify_brand(
    total_sales_qty: float,
    total_inventory_qty: float,
    has_deal: bool
) -> str:
    """
    Determine folder classification for a brand.

    Returns:
        "Reports"
        "No_Deal"
        "Empty_Brand_Guard"
        or None (if brand should not generate report)
    """

    # No sales AND no inventory → ignore
    if total_sales_qty == 0 and total_inventory_qty == 0:
        return None

    # Priority 1: Empty Brand Guard
    if total_sales_qty == 0 and total_inventory_qty > 0 and not has_deal:
        return "Empty_Brand_Guard"

    # Priority 2: No Deal
    if not has_deal:
        return "No_Deal"

    # Priority 3: Normal report
    return "Reports"
