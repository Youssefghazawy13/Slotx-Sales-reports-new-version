# core/brand_detector.py

import pandas as pd


def normalize_brand_name(name: str) -> str:
    return str(name).strip().title()


def detect_brands(
    sales_df: pd.DataFrame,
    inventory_df: pd.DataFrame
) -> set:
    """
    Detect unique brands from sales ∪ inventory.
    """

    sales_brands = set()
    inventory_brands = set()

    if "brand" in sales_df.columns:
        sales_brands = {
            normalize_brand_name(b)
            for b in sales_df["brand"].dropna().unique()
        }

    if "brand" in inventory_df.columns:
        inventory_brands = {
            normalize_brand_name(b)
            for b in inventory_df["brand"].dropna().unique()
        }

    return sales_brands.union(inventory_brands)
