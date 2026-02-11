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

    # 🔥 Get absolute path safely
    current_dir = os.path.dirname(os.path.abspath(__file__))
    project_root = os.path.dirname(current_dir)
    logo_path = os.path.join(project_root, "assets", "logo.png")

    if os.path.exists(logo_path):
        try:
            img = Image(logo_path)
            img.width = 150
            img.height = 80
            ws.add_image(img, "A1")
        except Exception:
            pass

    ws.append([""])
    ws.append([""])
    ws.append([""])

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    metadata_rows = [
        ["Powered by:", "Slot-X Solutions"],
        ["Version:", "v1.0"],
        ["Report Type:", mode],
        ["Payout Cycle:", payout_cycle],
        ["Generated At:", now],
    ]

    for row in metadata_rows:
        ws.append(row)

    start_row = ws.max_row - len(metadata_rows) + 1

    for row in ws.iter_rows(min_row=start_row, max_row=ws.max_row):
        row[0].font = bold_dark_blue_font
        row[1].font = dark_blue_font

    auto_fit_columns(ws)
