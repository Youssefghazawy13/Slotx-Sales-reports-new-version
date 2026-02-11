import streamlit as st
import pandas as pd
import zipfile
from io import BytesIO

from utils.column_detector import detect_columns
from core.refund_engine import clean_refunds
from core.deals_engine import load_deals_by_mode
from core.brand_detector import detect_brands
from core.classification_engine import classify_brand
from core.kpi_engine import calculate_sales_totals, calculate_inventory_totals
from reports.workbook_builder import build_brand_workbook


st.set_page_config(
    page_title="Slot-X Sales & Inventory Reports",
    layout="wide"
)

st.title("📊 Slot-X Sales & Inventory Reports")

# -----------------------
# MODE SELECTION
# -----------------------

mode = st.selectbox(
    "Select Report Mode",
    ["Alexandria", "Zamalek", "Merged"]
)

payout_cycle = st.selectbox(
    "Select Payout Cycle",
    ["Cycle 1", "Cycle 2"]
)

st.divider()

# -----------------------
# FILE UPLOADS
# -----------------------

if mode == "Merged":
    st.subheader("Zamalek Files")
    sales_zam_file = st.file_uploader("Zamalek Sales", type=["xlsx"])
    inventory_zam_file = st.file_uploader("Zamalek Inventory", type=["xlsx"])

    st.subheader("Alexandria Files")
    sales_alex_file = st.file_uploader("Alexandria Sales", type=["xlsx"])
    inventory_alex_file = st.file_uploader("Alexandria Inventory", type=["xlsx"])

else:
    sales_file = st.file_uploader("Sales File", type=["xlsx"])
    inventory_file = st.file_uploader("Inventory File", type=["xlsx"])

st.subheader("Deals File (3 Tabs Required)")
deals_file = st.file_uploader("Deals File", type=["xlsx"])

st.divider()

# -----------------------
# GENERATE BUTTON
# -----------------------

if st.button("🚀 Generate Reports", use_container_width=True):

    if not deals_file:
        st.error("Deals file is required.")
        st.stop()

    zip_buffer = BytesIO()

    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:

        # -----------------------
        # SINGLE MODE
        # -----------------------

        if mode in ["Alexandria", "Zamalek"]:

            if not sales_file or not inventory_file:
                st.error("Sales and Inventory files are required.")
                st.stop()

            sales_df = detect_columns(
                pd.read_excel(sales_file),
                "sales"
            )

            inventory_df = detect_columns(
                pd.read_excel(inventory_file),
                "inventory"
            )

            sales_df, _, _ = clean_refunds(sales_df)

            deals_dict = load_deals_by_mode(deals_file, mode)

            brands = detect_brands(sales_df, inventory_df)

            for brand in brands:

                brand_sales = sales_df[sales_df["brand"] == brand]
                brand_inventory = inventory_df[inventory_df["brand"] == brand]

                total_sales_qty, _ = calculate_sales_totals(brand_sales)
                total_inventory_qty, _ = calculate_inventory_totals(brand_inventory)

                has_brand_deal = brand in deals_dict

                classification = classify_brand(
                    total_sales_qty,
                    total_inventory_qty,
                    has_brand_deal
                )

                if not classification:
                    continue

                workbook_buffer = build_brand_workbook(
                    brand,
                    mode,
                    payout_cycle,
                    brand_sales,
                    brand_inventory=brand_inventory,
                    deals_dict=deals_dict
                )

                folder_path = f"{mode}/{classification}/{brand}.xlsx"

                zip_file.writestr(folder_path, workbook_buffer.getvalue())

        # -----------------------
        # MERGED MODE
        # -----------------------

        else:

            if not all([
                sales_zam_file,
                inventory_zam_file,
                sales_alex_file,
                inventory_alex_file
            ]):
                st.error("All branch files are required.")
                st.stop()

            # Load & detect columns
            sales_zam = detect_columns(pd.read_excel(sales_zam_file), "sales")
            inventory_zam = detect_columns(pd.read_excel(inventory_zam_file), "inventory")

            sales_alex = detect_columns(pd.read_excel(sales_alex_file), "sales")
            inventory_alex = detect_columns(pd.read_excel(inventory_alex_file), "inventory")

            # Clean refunds per branch
            sales_zam, _, _ = clean_refunds(sales_zam)
            sales_alex, _, _ = clean_refunds(sales_alex)

            # Load deals tabs
            deals_zam = load_deals_by_mode(deals_file, "Zamalek")
            deals_alex = load_deals_by_mode(deals_file, "Alexandria")
            deals_merged = load_deals_by_mode(deals_file, "Merged")

            # Process each branch separately
            for branch_name, sales_df, inventory_df, deals_dict in [
                ("Zamalek", sales_zam, inventory_zam, deals_zam),
                ("Alexandria", sales_alex, inventory_alex, deals_alex)
            ]:

                brands = detect_brands(sales_df, inventory_df)

                for brand in brands:

                    brand_sales = sales_df[sales_df["brand"] == brand]
                    brand_inventory = inventory_df[inventory_df["brand"] == brand]

                    total_sales_qty, _ = calculate_sales_totals(brand_sales)
                    total_inventory_qty, _ = calculate_inventory_totals(brand_inventory)

                    has_brand_deal = brand in deals_dict

                    classification = classify_brand(
                        total_sales_qty,
                        total_inventory_qty,
                        has_brand_deal
                    )

                    if not classification:
                        continue

                    workbook_buffer = build_brand_workbook(
                        brand,
                        branch_name,
                        payout_cycle,
                        brand_sales,
                        brand_inventory=brand_inventory,
                        deals_dict=deals_dict
                    )

                    folder_path = f"{branch_name}/{classification}/{brand}.xlsx"
                    zip_file.writestr(folder_path, workbook_buffer.getvalue())

            # --- MERGED PROCESSING ---

            merged_sales = pd.concat([sales_zam, sales_alex])
            merged_inventory = pd.concat([inventory_zam, inventory_alex])

            brands_merged = detect_brands(merged_sales, merged_inventory)

            for brand in brands_merged:

                brand_sales = merged_sales[merged_sales["brand"] == brand]

                inv_alex = inventory_alex[inventory_alex["brand"] == brand]
                inv_zam = inventory_zam[inventory_zam["brand"] == brand]

                total_sales_qty, _ = calculate_sales_totals(brand_sales)

                total_inventory_qty = (
                    inv_alex["available_quantity"].sum()
                    + inv_zam["available_quantity"].sum()
                )

                has_brand_deal = brand in deals_merged

                classification = classify_brand(
                    total_sales_qty,
                    total_inventory_qty,
                    has_brand_deal
                )

                if not classification:
                    continue

                workbook_buffer = build_brand_workbook(
                    brand,
                    "Merged",
                    payout_cycle,
                    brand_sales,
                    inventory_alex=inv_alex,
                    inventory_zam=inv_zam,
                    deals_dict=deals_merged
                )

                folder_path = f"Merged/{classification}/{brand}.xlsx"
                zip_file.writestr(folder_path, workbook_buffer.getvalue())

    zip_buffer.seek(0)

    st.success("Reports generated successfully.")

    st.download_button(
        "📥 Download ZIP",
        data=zip_buffer,
        file_name=f"SlotX_Reports_{mode}.zip",
        mime="application/zip",
        use_container_width=True
    )
