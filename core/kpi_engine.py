# core/kpi_engine.py

import pandas as pd


def calculate_sales_totals(brand_sales: pd.DataFrame):
    if brand_sales.empty:
        return 0, 0

    total_qty = brand_sales["quantity"].sum()
    total_money = brand_sales["total"].sum()

    return total_qty, total_money


def calculate_inventory_totals(brand_inventory: pd.DataFrame):
    if brand_inventory.empty:
        return 0, 0

    total_qty = brand_inventory["available_quantity"].sum()

    total_value = (
        brand_inventory["available_quantity"] *
        brand_inventory["unit_price"]
    ).sum()

    return total_qty, total_value


def apply_deal(total_sales_money, percentage, rent):
    after_percentage = total_sales_money - (
        total_sales_money * percentage / 100
    )

    after_rent = after_percentage - rent

    return after_percentage, after_rent


def get_best_selling_product(brand_sales: pd.DataFrame):
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


# ✅ Status based ONLY on Inventory
def calculate_status(
    sales_qty,
    inventory_qty,
    has_deal
):

    if inventory_qty <= 2:
        return "Critical"

    if 3 <= inventory_qty <= 5:
        return "Low"

    if 6 <= inventory_qty <= 15:
        return "Medium"

    return "Good"
