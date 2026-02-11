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
    # Add Logo (Top-Left)
    # -------------------------

    logo_path = os.path.join("assets", "logo.png")

    logo_height_rows = 6  # عدد الصفوف اللي هنحجزهم للصورة

    if os.path.exists(logo_path):
        try:
            img = Image(logo_path)
            img.width = 160
            img.height = 90
            ws.add_image(img, "A1")

            # نزود ارتفاع الصفوف الأولى
            for i in range(1, logo_height_rows + 1):
                ws.row_dimensions[i].height = 22

        except Exception as e:
            print("Logo load error:", e)

    # -------------------------
    # Start writing metadata AFTER logo area
    # -------------------------

    start_row = logo_height_rows + 2

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
