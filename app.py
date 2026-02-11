import streamlit as st
import pandas as pd
from io import BytesIO
import zipfile

from reports.workbook_builder import build_brand_workbook


st.set_page_config(
    page_title="Slot-X Sales & Inventory Reports",
    layout="centered"
)

st.title("Slot-X Sales & Inventory Reports")


# ============================================
# DEALS LOADER (LOCAL - NO CORE IMPORT)
# ============================================

def load_brand_deals(deals_file, mode: str):

    try:
        deals_df = pd.read_excel(deals_file, sheet_name=mode)

        deals_df.columns = deals_df.columns.str.strip()

        deals_df["Brand Name"] = (
            deals_df["Brand Name"]
            .astype(str)
            .str.strip()
        )

        deals_dict = {}

        for _, row in deals_df.iterrows():

            brand = row.get("Brand Name")

            if not brand:
                continue

            deals_dict[brand] = {
                "percentage": float(row.get("Deal Percentage (%)", 0) or 0),
                "rent": float(row.get("Rent Amount (EGP)", 0) or 0)
            }

        return deals_dict, None

    except Exception as e:
        return None, str(e)


# ============================================
# ZIP BUILDER
# ============================================

def build_reports_zip(brand_workbooks):

    zip_buffer = BytesIO()

    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:

        for brand, data in brand_workbooks.items():

            buffer = data["buffer"]
            has_sales = data["has_sales"]

            if buffer is None:
                continue

            safe_name = (
                str(brand)
                .replace("/", "-")
                .replace("\\", "-")
                .replace(":", "-")
            )

            if has_sales:
                path = f"Reports/{safe_name}.xlsx"
            else:
                path = f"Reports/Empty Brand Guard/{safe_name}.xlsx"

            zip_file.writestr(path, buffer.getvalue())

    zip_buffer.seek(0)
    return zip_buffer


# ============================================
# UI
# ============================================

mode = st.selectbox(
    "Select Mode",
    ["Alexandria", "Zamalek", "Merged"]
)

payout_cycle = st.selectbox(
    "Select Payout Cycle",
    ["Cycle 1", "Cycle 2"]
)

st.divider()

if mode == "Merged":

    col1, col2 = st.columns(2)

    with col1:
        sales_alex = st.file_uploader("Sales - Alexandria", type=["xlsx"])
        inventory_alex = st.file_uploader("Inventory - Alexandria", type=["xlsx"])

    with col2:
        sales_zam = st.file_uploader("Sales - Zamalek", type=["xlsx"])
        inventory_zam = st.file_uploader("Inventory - Zamalek", type=["xlsx"])

else:
    sales_file = st.file_uploader("Sales File", type=["xlsx"])
    inventory_file = st.file_uploader("Inventory File", type=["xlsx"])

deals_file = st.file_uploader("Deals File (Multi-tab)", type=["xlsx"])

st.divider()


# ============================================
# GENERATE
# ============================================

if st.button("Generate Reports"):

    if not deals_file:
        st.error("Upload deals file")
        st.stop()

    deals_dict, error = load_brand_deals(deals_file, mode)

    if error:
        st.error(error)
        st.stop()

    brand_workbooks = {}

    # ---------------- SINGLE ----------------

    if mode != "Merged":

        if not sales_file or not inventory_file:
            st.error("Upload sales & inventory")
            st.stop()

        sales_df = pd.read_excel(sales_file)
        inventory_df = pd.read_excel(inventory_file)

        brands = pd.concat([
            sales_df["brand"],
            inventory_df["brand"]
        ]).dropna().unique()

        for brand in brands:

            brand_sales = sales_df[sales_df["brand"] == brand]
            brand_inventory = inventory_df[inventory_df["brand"] == brand]

            workbook_buffer = build_brand_workbook(
                brand_name=brand,
                mode=mode,
                payout_cycle=payout_cycle,
                brand_sales=brand_sales,
                brand_inventory=brand_inventory,
                deals_dict=deals_dict
            )

            if workbook_buffer is None:
                continue

            brand_workbooks[brand] = {
                "buffer": workbook_buffer,
                "has_sales": not brand_sales.empty
            }

    # ---------------- MERGED ----------------

    else:

        if not all([sales_alex, inventory_alex, sales_zam, inventory_zam]):
            st.error("Upload all 4 files")
            st.stop()

        sales_alex_df = pd.read_excel(sales_alex)
        sales_zam_df = pd.read_excel(sales_zam)

        inventory_alex_df = pd.read_excel(inventory_alex)
        inventory_zam_df = pd.read_excel(inventory_zam)

        brands = pd.concat([
            sales_alex_df["brand"],
            sales_zam_df["brand"],
            inventory_alex_df["brand"],
            inventory_zam_df["brand"]
        ]).dropna().unique()

        for brand in brands:

            brand_sales = pd.concat([
                sales_alex_df[sales_alex_df["brand"] == brand],
                sales_zam_df[sales_zam_df["brand"] == brand]
            ])

            brand_inventory = pd.concat([
                inventory_alex_df[inventory_alex_df["brand"] == brand],
                inventory_zam_df[inventory_zam_df["brand"] == brand]
            ])

            workbook_buffer = build_brand_workbook(
                brand_name=brand,
                mode=mode,
                payout_cycle=payout_cycle,
                brand_sales=brand_sales,
                brand_inventory=brand_inventory,
                deals_dict=deals_dict
            )

            if workbook_buffer is None:
                continue

            brand_workbooks[brand] = {
                "buffer": workbook_buffer,
                "has_sales": not brand_sales.empty
            }

    zip_buffer = build_reports_zip(brand_workbooks)

    st.success("Reports generated successfully!")

    st.download_button(
        label="Download ZIP",
        data=zip_buffer,
        file_name=f"SlotX_Reports_{mode}_{payout_cycle}.zip",
        mime="application/zip"
    )
