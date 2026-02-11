# core/refund_engine.py

import pandas as pd


def clean_refunds(sales_df: pd.DataFrame):
    """
    Remove refund transactions (quantity < 0)
    AND their corresponding original sales.

    Returns:
        cleaned_df,
        refund_count,
        removed_total_count
    """

    if "quantity" not in sales_df.columns or "barcode" not in sales_df.columns:
        return sales_df.copy(), 0, 0

    sales_df = sales_df.copy()

    original_count = len(sales_df)

    # Identify refunds
    refunds = sales_df[sales_df["quantity"] < 0]
    refund_count = len(refunds)

    if refund_count == 0:
        return sales_df, 0, 0

    indices_to_remove = set(refunds.index.tolist())

    for idx, refund_row in refunds.iterrows():
        barcode = refund_row["barcode"]
        refund_qty = abs(refund_row["quantity"])
        brand = refund_row.get("brand")

        # Match original sale
        matching_sales = sales_df[
            (sales_df["barcode"] == barcode) &
            (sales_df["quantity"] == refund_qty) &
            (sales_df["brand"] == brand) &
            (~sales_df.index.isin(indices_to_remove))
        ]

        if not matching_sales.empty:
            indices_to_remove.add(matching_sales.index[0])

    cleaned_df = sales_df[~sales_df.index.isin(indices_to_remove)].copy()

    removed_total_count = original_count - len(cleaned_df)

    return cleaned_df, refund_count, removed_total_count
