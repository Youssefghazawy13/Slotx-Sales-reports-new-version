import streamlit as st
from core.zip_builder import generate_reports_zip

st.set_page_config(
    page_title="Slot-X Reports v1.0",
    layout="centered"
)

st.title("📊 Slot-X Sales & Inventory Reports")

# =========================
# UI CONTROLS
# =========================
report_type = st.selectbox(
    "Select Report Type",
    ["Alexandria Reports", "Zamalek Reports", "Merged Reports"]
)

payout_cycle = st.radio(
    "Payout Period",
    ["Cycle 1", "Cycle 2"],
    horizontal=True
)

uploaded = {}

# =========================
# FILE UPLOADS
# =========================
if report_type == "Alexandria Reports":
    uploaded["sales"] = st.file_uploader("Alexandria – Sales", type=["xlsx"])
    uploaded["inventory"] = st.file_uploader("Alexandria – Inventory", type=["xlsx"])
    uploaded["deals"] = st.file_uploader("Alexandria – Brand Deals", type=["xlsx"])

elif report_type == "Zamalek Reports":
    uploaded["sales"] = st.file_uploader("Zamalek – Sales", type=["xlsx"])
    uploaded["inventory"] = st.file_uploader("Zamalek – Inventory", type=["xlsx"])
    uploaded["deals"] = st.file_uploader("Zamalek – Brand Deals", type=["xlsx"])

else:  # Merged
    uploaded["alex_sales"] = st.file_uploader("Alexandria – Sales", type=["xlsx"])
    uploaded["alex_inventory"] = st.file_uploader("Alexandria – Inventory", type=["xlsx"])
    uploaded["zam_sales"] = st.file_uploader("Zamalek – Sales", type=["xlsx"])
    uploaded["zam_inventory"] = st.file_uploader("Zamalek – Inventory", type=["xlsx"])
    uploaded["deals"] = st.file_uploader("Merged – Brand Deals", type=["xlsx"])

# =========================
# GENERATE
# =========================
if st.button("🚀 Generate Reports"):
    if any(v is None for v in uploaded.values()):
        st.error("Please upload all required files.")
    else:
        with st.spinner("Generating reports…"):
            zip_buffer = generate_reports_zip(
                report_type=report_type,
                uploaded=uploaded,
                payout_cycle=payout_cycle  # IMPORTANT
            )

        st.success("Reports generated successfully!")

        st.download_button(
            label="📥 Download ZIP",
            data=zip_buffer,
            file_name="slotx_reports_v1_0.zip",
            mime="application/zip"
        )
# UI entry point – final wired version
