import streamlit as st
import pandas as pd
import zipfile
from io import BytesIO
from datetime import datetime

from reports.workbook_builder import build_brand_workbook
from reports.all_brands_summary import build_all_brands_summary
from core.deals_loader import load_branch_deals


st.set_page_config(
    page_title="Slot-X Sales & Inventory Reports",
    page_icon="📊",
    layout="centered"
)

st.title("Slot-X Sales & Inventory Reports")

# ==============================
# MODE SELECTION
# ==============================

mode = st.selectbox(
    "Select Mode",
    ["Zamalek", "Alexandria", "Merged"]
)

payout_cycle = st.selectbox(
    "Select Payout Cycle",
    ["Cycle 1", "Cycle 2"]
)

st.divider()

# ==============================
# FILE UPLOADS
# ==============================

if mode == "Merged":

    sales_zam_file = st.file_uploader("Upload Zamalek Sales")
    inventory_zam_file = st.file_uploader("Upload Zamalek Inventory")

    sales_alex_file = st.file_uploader("Upload Alexandria Sales")
    inventory_alex_file = st.file_uploader("Upload Alexandria Inventory")

    deals_file = st.file_uploader("Upload Deals File")

else:

    sales_file = st.file_uploader("Upload Sales File")
    inventory_file = st.file_uploader("Upload Inventory File")
    deals_file = st.file_uploader("Upload Deals File")

# ==============================
# GENERATE BUTTON
# ==============================

if st.button("Generate Reports"):

    zip_buffer = BytesIO()

    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:

        # ==========================================
        # MERGED MODE
        # ==========================================

        if mode == "Merged":

            sales_zam = pd.read_excel(sales_zam_file)
            inv_zam = pd.read_excel(inventory_zam_file)

            sales_alex = pd.read_excel(sales_alex_file)
            inv_alex = pd.read_excel(inventory_alex_file)

            deals_merged = load_branch_deals(deals_file, "Merged")
            deals_zam = load_branch_deals(deals_file, "Zamalek")
            deals_alex = load_branch_deals(deals_file, "Alexandria")

            all_brands = set(
                list(inv_zam["name_en"].unique()) +
                list(inv_alex["name_en"].unique())
            )

            for brand in all_brands:

                zam_inv_brand = inv_zam[inv_zam["name_en"] == brand]
                alex_inv_brand = inv_alex[inv_alex["name_en"] == brand]

                zam_qty = zam_inv_brand["available_quantity"].sum()
                alex_qty = alex_inv_brand["available_quantity"].sum()

                # --------------------------
                # DETERMINE BRANCH TYPE
                # --------------------------

                if alex_qty > 0 and zam_qty > 0:
                    branch_type = "Merged"
                    deals_dict = deals_merged

                    brand_inventory = pd.merge(
                        alex_inv_brand,
                        zam_inv_brand,
                        on=["name_en", "barcodes", "sale_price"],
                        how="outer",
                        suffixes=("_alex", "_zamalek")
                    ).fillna(0)

                    brand_inventory["available_quantity"] = (
                        brand_inventory["available_quantity_alex"] +
                        brand_inventory["available_quantity_zamalek"]
                    )

                elif zam_qty > 0:
                    branch_type = "Zamalek"
                    deals_dict = deals_zam
                    brand_inventory = zam_inv_brand.copy()

                elif alex_qty > 0:
                    branch_type = "Alexandria"
                    deals_dict = deals_alex
                    brand_inventory = alex_inv_brand.copy()

                else:
                    continue

                # --------------------------
                # SALES FILTER
                # --------------------------

                brand_sales = pd.concat([
                    sales_zam[sales_zam["name_ar"] == brand],
                    sales_alex[sales_alex["name_ar"] == brand]
                ])

                total_sales_qty = brand_sales["quantity"].sum()

                deal = deals_dict.get(brand, {"percentage": 0, "rent": 0})
                percentage = deal["percentage"]
                rent = deal["rent"]

                # --------------------------
                # DETERMINE SUBFOLDER
                # --------------------------

                if total_sales_qty == 0:
                    subfolder = "Empty Brand Guard"

                elif percentage == 0 and rent == 0:
                    subfolder = "No Deal"

                else:
                    subfolder = None

                # --------------------------
                # BUILD REPORT
                # --------------------------

                workbook_buffer = build_brand_workbook(
                    brand_name=brand,
                    mode=branch_type,
                    payout_cycle=payout_cycle,
                    brand_sales=brand_sales,
                    brand_inventory=brand_inventory,
                    deals_dict=deals_dict
                )

                base_path = f"Reports/{branch_type}"

                if subfolder:
                    file_path = f"{base_path}/{subfolder}/{brand}.xlsx"
                else:
                    file_path = f"{base_path}/{brand}.xlsx"

                zip_file.writestr(file_path, workbook_buffer.getvalue())

        # ==========================================
        # SINGLE BRANCH MODE
        # ==========================================

        else:

            sales_df = pd.read_excel(sales_file)
            inventory_df = pd.read_excel(inventory_file)
            deals_dict = load_branch_deals(deals_file, mode)

            brands = inventory_df["name_en"].unique()

            for brand in brands:

                brand_inventory = inventory_df[inventory_df["name_en"] == brand]
                brand_sales = sales_df[sales_df["name_ar"] == brand]

                total_sales_qty = brand_sales["quantity"].sum()

                deal = deals_dict.get(brand, {"percentage": 0, "rent": 0})
                percentage = deal["percentage"]
                rent = deal["rent"]

                if total_sales_qty == 0:
                    subfolder = "Empty Brand Guard"

                elif percentage == 0 and rent == 0:
                    subfolder = "No Deal"

                else:
                    subfolder = None

                workbook_buffer = build_brand_workbook(
                    brand_name=brand,
                    mode=mode,
                    payout_cycle=payout_cycle,
                    brand_sales=brand_sales,
                    brand_inventory=brand_inventory,
                    deals_dict=deals_dict
                )

                base_path = f"Reports/{mode}"

                if subfolder:
                    file_path = f"{base_path}/{subfolder}/{brand}.xlsx"
                else:
                    file_path = f"{base_path}/{brand}.xlsx"

                zip_file.writestr(file_path, workbook_buffer.getvalue())

            # SUMMARY ONLY FOR SINGLE MODE
            summary_buffer = build_all_brands_summary(
                branch_name=mode,
                payout_cycle=payout_cycle,
                sales_df=sales_df,
                inventory_df=inventory_df,
                deals_dict=deals_dict
            )

            zip_file.writestr(
                f"Reports/{mode}/All_Brands_Summary.xlsx",
                summary_buffer.getvalue()
            )

    zip_buffer.seek(0)

    st.download_button(
        "Download Reports ZIP",
        data=zip_buffer,
        file_name=f"SlotX_Reports_{mode}.zip",
        mime="application/zip"
    )
