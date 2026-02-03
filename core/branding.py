from openpyxl.styles import Font
from datetime import datetime

def add_metadata(wb):
    ws = wb.create_sheet("Metadata")
    ws.append(["Powered By", "Slot-X Solutions"])
    ws.append(["Version", "v1.0"])
    ws.append(["Generated At", datetime.now().strftime("%Y-%m-%d %H:%M")])

    for cell in ws["A"]:
        cell.font = Font(bold=True)
# metadata + logo
