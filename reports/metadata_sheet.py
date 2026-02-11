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
    """
    Create Metadata sheet (always last).
    Logo appears at top-left (A1).
    """

    ws = wb.create_sheet("Metadata")

    # -------------------------
    # Add Logo (Top-Left)
    # -------------------------

    logo_path = "assets/logo.png"

    if os.path.exists(logo_path):
        try:
            img = Image(logo_path)
            img.width = 150
            img.height = 80
            ws.add_image(img, "A1")
        except Exception:
            pass  # Prevent crash if image fails

    # Add spacing below logo
    ws.append([""])
    ws.append([""])
    ws.append([""])

    # -------------------------
    # Metadata Content
    # -------------------------

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

    # Bold first column
    for row in ws.iter_rows(
        min_row=ws.max_row - len(metadata_rows) + 1,
        max_row=ws.max_row,
        min_col=1,
        max_col=1
    ):
        for cell in row:
            if cell.value:
                cell.font = Font(bold=True)

    auto_fit_columns(ws)
