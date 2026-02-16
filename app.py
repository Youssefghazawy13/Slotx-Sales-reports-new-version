import streamlit as st
import pandas as pd
import zipfile
from io import BytesIO

from reports.workbook_builder import build_brand_workbook
from reports.branch_summary_workbook import build_branch_summary_workbook
from core.deals_engine import load_deals_by_mode, normalize_brand_name


st.set_page_config(
    page_title="Slot-X Sales & Inventory Reports",
    page_icon="📊",
    layout="centered"
)

st.title("Slot-X Sales & Inventory Reports")


# =========================================================
# REFUND REMOVAL (OLD LOGIC EXACTLY)
# =========================================================

def remove_refunds_and_original_sales(sales_df):

    if "quantity" not in sales_df.columns or "barcode" not in sales_df.columns:
        return sales_df

    sales_df = sales_df.copy()

    refunds = sales_df[sales_df["quantity"] < 0]

    if refunds.empty:
        return sales_df

    indices_to_remove = set(refunds.index.tolist())

    for idx, refund_row in refunds.iterrows():
        barcode = refund_row["barcode"]
        refund_qty = abs(refund_row["quantity"])
        brand = refund_row["brand"]

        matching_sales = sales_df[
            (sales_df["barcode"] == barcode) &
            (sales_df["brand"] == brand) &
            (sales_df["quantity"] == refund_qty)
        ]

        if not matching_sales.empty:
            indices_to_remove.add(matching_sales.index[0])

    return sales_df.drop(index=indices_to_remove)


# =========================================================
# MODE
# =========================================================

mode = st.selectbox("Select Mode", ["Zamalek", "Alexandria", "Merged"])
payout_cycle = st.selectbox("Select Payout Cycle", ["Cycle 1", "Cycle 2"])

st.divider()


# =========================================================
# FILE UPLOAD
# =========================================================

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


# =========================================================
# GENERATE
# =========================================================

if st.button("Generate Reports"):

    zip_buffer = BytesIO()

    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:

        # =====================================================
        # MERGED MODE
        # =====================================================

        if mode == "Merged":

            sales_zam = remove_refunds_and_original_sales(
                pd.read_excel(sales_zam_file)
            )
            inv_zam = pd.read_excel(inventory_zam_file)

            sales_alex = remove_refunds_and_original_sales(
                pd.read_excel(sales_alex_file)
            )
            inv_alex = pd.read_excel(inventory_alex_file)

            # 🔥 USE SAME NORMALIZATION EVERYWHERE
            sales_zam["brand"] = sales_zam["brand"].apply(normalize_brand_name)
            sales_alex["brand"] = sales_alex["brand"].apply(normalize_brand_name)

            inv_zam["brand"] = inv_zam["brand"].apply(normalize_brand_name)
            inv_alex["brand"] = inv_alex["brand"].apply(normalize_brand_name)

            # Ensure numeric
            inv_zam["available_quantity"] = pd.to_numeric(
                inv_zam["available_quantity"], errors="coerce"
            ).fillna(0)

            inv_alex["available_quantity"] = pd.to_numeric(
                inv_alex["available_quantity"], errors="coerce"
            ).fillna(0)

            # Load deals
            deals_merged = load_deals_by_mode(deals_file, "Merged")
            deals_zam = load_deals_by_mode(deals_file, "Zamalek")
            deals_alex = load_deals_by_mode(deals_file, "Alexandria")

            # Normalize deal keys
            deals_merged = {
                normalize_brand_name(k): v for k, v in deals_merged.items()
            }
            deals_zam = {
                normalize_brand_name(k): v for k, v in deals_zam.items()
            }
            deals_alex = {
                normalize_brand_name(k): v for k, v in deals_alex.items()
            }

            all_brands = set(
                list(inv_zam["brand"].unique()) +
                list(inv_alex["brand"].unique())
            )

            for brand in all_brands:

                zam_inv_brand = inv_zam[inv_zam["brand"] == brand]
                alex_inv_brand = inv_alex[inv_alex["brand"] == brand]

                zam_qty = zam_inv_brand["available_quantity"].sum()
                alex_qty = alex_inv_brand["available_quantity"].sum()

                if zam_qty == 0 and alex_qty == 0:
                    continue

                if alex_qty > 0 and zam_qty > 0:
                    branch_type = "Merged"
                    deals_dict = deals_merged
                    brand_inventory = pd.concat(
                        [alex_inv_brand, zam_inv_brand],
                        ignore_index=True
                    )
                elif zam_qty > 0:
                    branch_type = "Zamalek"
                    deals_dict = deals_zam
                    brand_inventory = zam_inv_brand.copy()
                else:
                    branch_type = "Alexandria"
                    deals_dict = deals_alex
                    brand_inventory = alex_inv_brand.copy()

                brand_sales = pd.concat([
                    sales_zam[sales_zam["brand"] == brand],
                    sales_alex[sales_alex["brand"] == brand]
                ])

                total_sales_qty = brand_sales["quantity"].sum()
                total_inventory_qty = brand_inventory["available_quantity"].sum()

                deal = deals_dict.get(brand, {"percentage": 0, "rent": 0})

                if total_sales_qty == 0 and total_inventory_qty > 0:
                    subfolder = "Empty Brand Guard"
                elif deal["percentage"] == 0 and deal["rent"] == 0:
                    subfolder = "No Deal"
                else:
                    subfolder = None

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

        # =====================================================
        # SINGLE MODE
        # =====================================================

        else:

            sales_df = remove_refunds_and_original_sales(
                pd.read_excel(sales_file)
            )
            inventory_df = pd.read_excel(inventory_file)

            sales_df["brand"] = sales_df["brand"].apply(normalize_brand_name)
            inventory_df["brand"] = inventory_df["brand"].apply(normalize_brand_name)

            deals_dict = load_deals_by_mode(deals_file, mode)
            deals_dict = {
                normalize_brand_name(k): v for k, v in deals_dict.items()
            }

            brands = inventory_df["brand"].unique()

            for brand in brands:

                brand_inventory = inventory_df[
                    inventory_df["brand"] == brand
                ]

                brand_sales = sales_df[
                    sales_df["brand"] == brand
                ]

                total_sales_qty = brand_sales["quantity"].sum()
                total_inventory_qty = brand_inventory["available_quantity"].sum()

                if total_inventory_qty == 0:
                    continue

                deal = deals_dict.get(brand, {"percentage": 0, "rent": 0})

                if total_sales_qty == 0 and total_inventory_qty > 0:
                    subfolder = "Empty Brand Guard"
                elif deal["percentage"] == 0 and deal["rent"] == 0:
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

            summary_wb = build_branch_summary_workbook(
                branch_name=mode,
                payout_cycle=payout_cycle,
                sales_df=sales_df,
                inventory_df=inventory_df,
                deals_dict=deals_dict
            )

            summary_buffer = BytesIO()
            summary_wb.save(summary_buffer)

            zip_file.writestr(
                f"Reports/{mode}/{mode}_Summary.xlsx",
                summary_buffer.getvalue()
            )

    zip_buffer.seek(0)

    st.download_button(
        "Download Reports ZIP",
        data=zip_buffer,
        file_name=f"SlotX_Reports_{mode}.zip",
        mime="application/zip"
    )
