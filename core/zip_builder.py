from io import BytesIO
import zipfile
import pandas as pd
from openpyxl import Workbook
from core.report import build_report
from core.sales import build_sales
from core.inventory import build_inventory
from core.branding import add_metadata

def generate_reports_zip(report_type, uploaded, payout_cycle):
    buffer = BytesIO()

    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
        if report_type != "Merged Reports":
            branch = report_type.split()[0]
            sales = pd.read_excel(uploaded["sales"])
            inventory = pd.read_excel(uploaded["inventory"])
            deals = pd.read_excel(uploaded["deals"])

            for brand in inventory["brand"].unique():
                wb = Workbook()
                wb.remove(wb.active)

                build_sales(wb, sales, brand, branch)
                build_inventory(wb, inventory, brand, branch)
                build_report(wb, sales, inventory, deals, brand, branch, payout_cycle)
                add_metadata(wb)

                bio = BytesIO()
                wb.save(bio)
                zipf.writestr(f"{branch}/{brand}_Report.xlsx", bio.getvalue())

        else:
            sales = pd.concat([
                pd.read_excel(uploaded["alex_sales"]),
                pd.read_excel(uploaded["zam_sales"])
            ])
            inventory = pd.concat([
                pd.read_excel(uploaded["alex_inventory"]),
                pd.read_excel(uploaded["zam_inventory"])
            ])
            deals = pd.read_excel(uploaded["deals"])

            for brand in inventory["brand"].unique():
                wb = Workbook()
                wb.remove(wb.active)

                build_sales(wb, sales, brand, "Merged")
                build_inventory(wb, inventory, brand, "Merged")
                build_report(wb, sales, inventory, deals, brand, "Merged", payout_cycle)
                add_metadata(wb)

                bio = BytesIO()
                wb.save(bio)
                zipf.writestr(f"Merged/{brand}_Report.xlsx", bio.getvalue())

    buffer.seek(0)
    return buffer
# zip builder
