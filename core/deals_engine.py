import pandas as pd
import re


# =====================================================
# SMART BRAND NORMALIZATION
# =====================================================

def normalize_brand_name(name: str) -> str:
    """
    Smart normalization:
    - Remove leading/trailing spaces
    - Lowercase everything
    - Remove spaces inside
    - Remove dashes
    - Remove underscores
    - Remove any special characters
    """

    if not name:
        return ""

    name = str(name).strip().lower()

    # remove spaces, dashes, underscores
    name = re.sub(r"[\s\-_]+", "", name)

    # remove any non-alphanumeric characters
    name = re.sub(r"[^a-z0-9]", "", name)

    return name


# =====================================================
# LOAD DEALS BY MODE (Zamalek / Alexandria / Merged)
# =====================================================

def load_deals_by_mode(deals_file, mode: str) -> dict:

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

        raw_brand = row["Brand Name"]
        normalized_brand = normalize_brand_name(raw_brand)

        if not normalized_brand:
            continue

        percentage = row.get("Deal Percentage (%)", 0)
        rent = row.get("Rent Amount (EGP)", 0)

        percentage = float(percentage) if pd.notna(percentage) else 0.0
        rent = float(rent) if pd.notna(rent) else 0.0

        deals_dict[normalized_brand] = {
            "percentage": percentage,
            "rent": rent
        }

    return deals_dict


# =====================================================
# CHECK IF BRAND HAS DEAL
# =====================================================

def has_deal(brand: str, deals_dict: dict) -> bool:

    normalized_brand = normalize_brand_name(brand)

    if normalized_brand not in deals_dict:
        return False

    deal = deals_dict[normalized_brand]
    return deal["percentage"] > 0 or deal["rent"] > 0


# =====================================================
# GENERATE DEAL TEXT
# =====================================================

def generate_deal_text(brand: str, deals_dict: dict) -> str:

    normalized_brand = normalize_brand_name(brand)

    if normalized_brand not in deals_dict:
        return "No Deal"

    percentage = deals_dict[normalized_brand]["percentage"]
    rent = deals_dict[normalized_brand]["rent"]

    if percentage > 0 and rent > 0:
        return f"{percentage}% + {rent} EGP Deducted from the sales"

    elif percentage > 0:
        return f"{percentage}% Deducted from the sales"

    elif rent > 0:
        return f"{rent} EGP Deducted from the sales"

    else:
        return "No Deal"
