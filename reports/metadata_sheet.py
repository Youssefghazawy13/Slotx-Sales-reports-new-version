# reports/metadata_sheet.py

from openpyxl.styles import Font, Alignment
from datetime import datetime
from utils.excel_helpers import auto_fit_columns


def create_metadata_sheet(
    wb,
    mode: str,
    payout_cycle: str
):

    ws = wb.create_sheet("Metadata")

    # -------------------------
    # SLOT-X TEXT BANNER
    # -------------------------

    # Merge wide area for title
    ws.merge_cells("A1:H4")

    title_cell = ws["A1"]
    title_cell.value = "SLOT-X"

    # Metallic-like text styling
    title_cell.font = Font(
        name="Impact",       # لو مش متاح هي fallback
        size=42,
        bold=True,
        color="BFBFBF"       # Metallic gray
    )

    title_cell.alignment = Alignment(
        horizontal="center",
        vertical="center"
    )

    # Increase height for visual weight
    for i in range(1, 5):
        ws.row_dimensions[i].height = 40

    # -------------------------
    # Metadata Section
    # -------------------------

    dark_blue_font = Font(color="1F4E78")
    bold_dark_blue_font = Font(bold=True, color="1F4E78")

    start_row = 6

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
