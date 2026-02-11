# reports/metadata_sheet.py

from openpyxl.styles import Font
from datetime import datetime
from utils.excel_helpers import auto_fit_columns


def create_metadata_sheet(
    wb,
    mode: str,
    payout_cycle: str
):
    """
    Create Metadata sheet (always last).
    """

    ws = wb.create_sheet("Metadata")

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
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=1):
        for cell in row:
            if cell.value:
                cell.font = Font(bold=True)

    auto_fit_columns(ws)
