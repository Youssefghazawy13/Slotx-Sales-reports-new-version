import streamlit as st
import pandas as pd

from reports.workbook_builder import build_brand_workbook
from core.zip_builder import build_reports_zip
from core.deals_loader import load_brand_deals


st.set_page_config(
    page_title="Slot-X Sales & Inventory Reports",
    layout="centered"
)

st.title("Slot-X Sales & Inventory Reports")

# ============================================
# MODE SELECTOR
# ============================================

mode = st.selectbox(
    "Select Report Mode",
    ["Alexandria", "Zamalek", "Merged"]
)

payout_cycle = st.selectbox(
    "Select Payout Cycle",
    ["Cycle 1", "Cycle 2"]
)

st.divider()

# ============================================
# FILE UPLOADS
# ============================================

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

    # ============================================
    # SINGLE MODE
    # ============================================

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

    # ============================================
    # MERGED MODE
    # ============================================

    else:

        if not all([sales_alex, inventory_alex, sales_zam, inventory_zam]):
            st.error("Upload all 4 files for merged mode")
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

            alex_inv = inventory_alex_df[inventory_alex_df["brand"] == brand]
            zam_inv = inventory_zam_df[inventory_zam_df["brand"] == brand]

            if not alex_inv.empty:
                alex_inv = alex_inv.rename(columns={"quantity": "alex_qty"})
            if not zam_inv.empty:
                zam_inv = zam_inv.rename(columns={"quantity": "zamalek_qty"})

            brand_inventory = pd.merge(
                alex_inv,
                zam_inv,
                on=["product_name", "barcode", "price"],
                how="outer"
            ).fillna(0)

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

    # ============================================
    # ZIP
    # ============================================

    zip_buffer = build_reports_zip(brand_workbooks)

    st.success("Reports generated successfully!")

    st.download_button(
        label="Download ZIP",
        data=zip_buffer,
        file_name=f"SlotX_Reports_{mode}_{payout_cycle}.zip",
        mime="application/zip"
    )
