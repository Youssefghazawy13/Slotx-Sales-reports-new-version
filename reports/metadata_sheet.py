from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter
from datetime import datetime


def create_metadata_sheet(wb, report_title, branch_name, payout_cycle):

    ws = wb.create_sheet("Metadata")

    # =========================================
    # HEADER BACKGROUND (Dark Blue)
    # =========================================

    dark_blue = PatternFill(
        start_color="0F2A66",
        end_color="0F2A66",
        fill_type="solid"
    )

    for row in range(1, 6):
        for col in range(1, 10):
            ws.cell(row=row, column=col).fill = dark_blue

    # =========================================
    # SLOT-X TEXT (Centered & Large)
    # =========================================

    ws.merge_cells("A1:I5")

    title_cell = ws["A1"]
    title_cell.value = "SLOT-X"

    title_cell.font = Font(
        name="Arial Black",   # أقرب شكل للفونت اللي كان ظاهر
        size=48,
        bold=True,
        color="FFFFFF"
    )

    title_cell.alignment = Alignment(
        horizontal="center",
        vertical="center"
    )

    # =========================================
    # METADATA CONTENT
    # =========================================

    start_row = 7

    metadata = [
        ("Powered by:", "Slot-X Solutions"),
        ("Version:", "v2.0"),
        ("Report Type:", branch_name),
        ("Payout Cycle:", payout_cycle),
        ("Generated At:", datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
    ]

    for i, (label, value) in enumerate(metadata):
        row = start_row + i

        ws[f"A{row}"] = label
        ws[f"A{row}"].font = Font(bold=True)

        ws[f"B{row}"] = value

    # Auto width
    for col in range(1, 5):
        ws.column_dimensions[get_column_letter(col)].width = 22