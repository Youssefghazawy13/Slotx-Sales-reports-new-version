# core/kpi_engine.py

import pandas as pd


def calculate_sales_totals(brand_sales: pd.DataFrame):
    """
    Returns:
        total_sales_qty,
        total_sales_money
    """

    if brand_sales.empty:
        return 0, 0

    total_qty = brand_sales["quantity"].sum()
    total_money = brand_sales["total"].sum()

    return total_qty, total_money


def calculate_inventory_totals(brand_inventory: pd.DataFrame):
    """
    Returns:
        total_inventory_qty,
        total_inventory_value
    """

    if brand_inventory.empty:
        return 0, 0

    total_qty = brand_inventory["available_quantity"].sum()

    total_value = (
        brand_inventory["available_quantity"] *
        brand_inventory["unit_price"]
    ).sum()

    return total_qty, total_value


def apply_deal(total_sales_money, percentage, rent):
    """
    Apply percentage first, then rent.
    """

    after_percentage = total_sales_money - (
        total_sales_money * percentage / 100
    )

    after_rent = after_percentage - rent

    return after_percentage, after_rent


def get_best_selling_product(brand_sales: pd.DataFrame):
    """
    Return product with highest total quantity.
    """

    if brand_sales.empty:
        return ""

    grouped = (
        brand_sales
        .groupby("product_name")["quantity"]
        .sum()
        .sort_values(ascending=False)
    )

    if grouped.empty:
        return ""

    return grouped.index[0]


def get_best_selling_size(brand_sales: pd.DataFrame):
    """
    Extract size from product name using '-' split.
    Example: 'Tshirt - L'
    """

    if brand_sales.empty:
        return ""

    size_sales = {}

    for _, row in brand_sales.iterrows():
        name = str(row.get("product_name", ""))
        qty = row.get("quantity", 0)

        if "-" in name:
            size = name.split("-")[-1].strip()
            if size:
                size_sales[size] = size_sales.get(size, 0) + qty

    if not size_sales:
        return ""

    return max(size_sales, key=size_sales.get)


def calculate_status(
    sales_qty,
    inventory_qty,
    has_deal
):
    """
    Determine product-level status.
    """

    # Empty brand handled elsewhere
    if inventory_qty < 5:
        return "Low Stock"

    if sales_qty <= 3:
        return "Slow Moving"

    if not has_deal:
        return "No Deal"

    return "Healthy"
