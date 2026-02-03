def empty_brands(inventory_df):
    """
    Return brands that exist but have zero quantity
    """
    result = []
    grouped = inventory_df.groupby("brand")["quantity"].sum()

    for brand, qty in grouped.items():
        if qty == 0:
            result.append(brand)

    return result
# summary
