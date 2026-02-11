from openpyxl.styles import Font, PatternFill, Alignment
from datetime import datetime


def create_metadata_sheet(
    wb,
    report_title,
    branch_name,
    payout_cycle
):

    ws = wb.create_sheet("Metadata")

    # Header styling
    header_fill = PatternFill(
        start_color="0A1F5C",
        end_color="0A1F5C",
        fill_type="solid"
    )

    header_font = Font(
        bold=True,
        color="FFFFFF"
    )

    # Title
    ws["A1"] = "Slot-X Sales & Inventory Reports"
    ws["A1"].font = Font(size=14, bold=True)

    # Spacing
    ws["A3"] = "Report Title:"
    ws["B3"] = report_title

    ws["A4"] = "Branch:"
    ws["B4"] = branch_name

    ws["A5"] = "Payout Cycle:"
    ws["B5"] = payout_cycle

    ws["A6"] = "Generated On:"
    ws["B6"] = datetime.now().strftime("%Y-%m-%d %H:%M")

    # Make labels bold
    for row in range(3, 7):
        ws[f"A{row}"].font = Font(bold=True)

    # Adjust column width
    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 30

    return ws
