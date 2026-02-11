# core/deals_engine.py

import pandas as pd


def normalize_brand_name(name: str) -> str:
    return str(name).strip().title()


def load_deals_by_mode(deals_file, mode: str) -> dict:
    """
    Load deals from a specific sheet (mode).
    mode must match sheet name exactly:
    - 'Zamalek'
    - 'Alexandria'
    - 'Merged'
    """

    try:
        deals_df = pd.read_excel(deals_file, sheet_name=mode)
    except Exception as e:
        raise Exception(f"Could not read deals sheet '{mode}': {str(e)}")

    deals_df.columns = deals_df.columns.str.strip()

    required_columns = [
        "Brand Name",
        "Deal Percentage (%)",
        "Rent Amount (EGP)"
    ]

    for col in required_columns:
        if col not in deals_df.columns:
            raise ValueError(
                f"Missing required column '{col}' in deals sheet '{mode}'"
            )

    deals_dict = {}

    for _, row in deals_df.iterrows():
        brand = normalize_brand_name(row["Brand Name"])

        if not brand or brand.lower() == "nan":
            continue

        percentage = row.get("Deal Percentage (%)", 0)
        rent = row.get("Rent Amount (EGP)", 0)

        percentage = float(percentage) if pd.notna(percentage) else 0.0
        rent = float(rent) if pd.notna(rent) else 0.0

        deals_dict[brand] = {
            "percentage": percentage,
            "rent": rent
        }

    return deals_dict


def has_deal(brand: str, deals_dict: dict) -> bool:
    """
    Check if brand has any deal (percentage or rent).
    """
    brand = normalize_brand_name(brand)

    if brand not in deals_dict:
        return False

    deal = deals_dict[brand]
    return deal["percentage"] > 0 or deal["rent"] > 0


def generate_deal_text(brand: str, deals_dict: dict) -> str:
    """
    Generate formatted deal text for report sheet.
    """

    brand = normalize_brand_name(brand)

    if brand not in deals_dict:
        return "No Deal"

    percentage = deals_dict[brand]["percentage"]
    rent = deals_dict[brand]["rent"]

    if percentage > 0 and rent > 0:
        return f"{percentage}% + {rent} EGP Deducted from the sales"

    elif percentage > 0:
        return f"{percentage}% Deducted from the sales"

    elif rent > 0:
        return f"{rent} EGP Deducted from the sales"

    else:
        return "No Deal"
