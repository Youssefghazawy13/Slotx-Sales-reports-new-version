# utils/column_detector.py

import re

def normalize_column_name(col_name: str) -> str:
    """
    Normalize column name for comparison.
    """
    col = str(col_name).strip().lower()
    col = re.sub(r"[^a-z0-9]+", "_", col)
    return col


def detect_columns(df, file_type: str):
    """
    Detect and standardize required columns dynamically.
    
    file_type: "sales" or "inventory"
    """

    normalized_map = {
        normalize_column_name(col): col
        for col in df.columns
    }

    # Define expected aliases
    expected_columns = get_expected_columns(file_type)

    final_mapping = {}

    for standard_name, aliases in expected_columns.items():
        for alias in aliases:
            normalized_alias = normalize_column_name(alias)
            if normalized_alias in normalized_map:
                final_mapping[normalized_map[normalized_alias]] = standard_name
                break

    # Rename detected columns
    df = df.rename(columns=final_mapping)

    # Validate required columns
    missing = [
        col for col in expected_columns.keys()
        if col not in df.columns
    ]

    if missing:
        raise ValueError(
            f"Missing required columns in {file_type} file: {missing}"
        )

    return df


def get_expected_columns(file_type: str):
    """
    Define expected column aliases.
    """

    if file_type == "sales":
        return {
            "brand": ["brand", "brand_name"],
            "product_name": ["product_name", "name_ar", "product"],
            "barcode": ["barcode", "barcodes"],
            "quantity": ["quantity", "qty"],
            "total": ["total", "total_price", "amount"],
        }

    elif file_type == "inventory":
        return {
            "brand": ["brand", "brand_name"],
            "product_name": ["product_name", "name_en", "product"],
            "barcode": ["barcode", "barcodes"],
            "unit_price": ["unit_price", "sale_price", "price"],
            "available_quantity": ["available_quantity", "qty", "stock"],
        }

    else:
        raise ValueError("Invalid file type for column detection.")
