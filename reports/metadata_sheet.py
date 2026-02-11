# reports/metadata_sheet.py

from openpyxl.styles import Font
from openpyxl.drawing.image import Image
from datetime import datetime
from utils.excel_helpers import auto_fit_columns
import os


def create_metadata_sheet(
    wb,
    mode: str,
    payout_cycle: str
):

    ws = wb.create_sheet("Metadata")

    dark_blue_font = Font(color="1F4E78")
    bold_dark_blue_font = Font(bold=True, color="1F4E78")

    # -------------------------
    # Create header space for logo
    # -------------------------

    ws.merge_cells("A1:D6")

    # Increase row height
    for i in range(1, 7):
        ws.row_dimensions[i].height = 25

    logo_path = os.path.join("assets", "logo.png")

    if os.path.exists(logo_path):
        try:
            img = Image(logo_path)
            img.width = 220
            img.height = 110
            ws.add_image(img, "A1")
        except Exception as e:
            print("Logo load error:", e)

    # -------------------------
    # Metadata content below logo
    # -------------------------

    start_row = 8

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    metadata_rows = [
        ["Powered by:", "Slot-X Solutions"],
        ["Version:", "v1.0"],
        ["Report Type:", mode],
        ["Payout Cycle:", payout_cycle],
        ["Generated At:", now],
    ]

    current_row = start_row

    for row in metadata_rows:
        ws.cell(row=current_row, column=1, value=row[0]).font = bold_dark_blue_font
        ws.cell(row=current_row, column=2, value=row[1]).font = dark_blue_font
        current_row += 1

    auto_fit_columns(ws)
