import streamlit as st
from core.zip_builder import generate_reports_zip

st.set_page_config(page_title="Slot-X Reports", layout="centered")
st.title("Slot-X Sales Reports")

report_type = st.selectbox(
    "Report Type",
    ["Alexandria Reports", "Zamalek Reports", "Merged Reports"]
)

payout_cycle = st.radio(
    "Payout Cycle",
    ["Cycle 1", "Cycle 2"],
    horizontal=True
)

uploaded = {}

if report_type == "Alexandria Reports":
    uploaded["sales"] = st.file_uploader("Alexandria Sales", type="xlsx")
    uploaded["inventory"] = st.file_uploader("Alexandria Inventory", type="xlsx")
    uploaded["deals"] = st.file_uploader("Brand Deals", type="xlsx")

elif report_type == "Zamalek Reports":
    uploaded["sales"] = st.file_uploader("Zamalek Sales", type="xlsx")
    uploaded["inventory"] = st.file_uploader("Zamalek Inventory", type="xlsx")
    uploaded["deals"] = st.file_uploader("Brand Deals", type="xlsx")

else:
    uploaded["alex_sales"] = st.file_uploader("Alexandria Sales", type="xlsx")
    uploaded["alex_inventory"] = st.file_uploader("Alexandria Inventory", type="xlsx")
    uploaded["zam_sales"] = st.file_uploader("Zamalek Sales", type="xlsx")
    uploaded["zam_inventory"] = st.file_uploader("Zamalek Inventory", type="xlsx")
    uploaded["deals"] = st.file_uploader("Brand Deals", type="xlsx")

if st.button("Generate Reports"):
    if any(v is None for v in uploaded.values()):
        st.error("Upload all required files")
    else:
        zip_buffer = generate_reports_zip(
            report_type=report_type,
            uploaded=uploaded,
            payout_cycle=payout_cycle
        )

        st.download_button(
            "Download ZIP",
            zip_buffer,
            "slotx_reports.zip",
            mime="application/zip"
        )
