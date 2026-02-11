from openpyxl.styles import Font, PatternFill, Alignment
from datetime import datetime


def create_metadata_sheet(wb, report_title, branch_name, payout_cycle):

    ws = wb.create_sheet("Metadata")

    # =============================
    # FULL WIDTH DARK BLUE HEADER
    # =============================

    header_fill = PatternFill(
        start_color="0A1F5C",
        end_color="0A1F5C",
        fill_type="solid"
    )

    for col in range(1, 8):
        ws.cell(row=1, column=col).fill = header_fill
        ws.cell(row=2, column=col).fill = header_fill
        ws.cell(row=3, column=col).fill = header_fill

    # =============================
    # SLOT-X CHROME STYLE
    # =============================

    ws.merge_cells("A1:G3")

    ws["A1"] = "SLOT-X"

    ws["A1"].font = Font(
        size=28,
        bold=True,
        color="C0C0C0"  # chrome silver color
    )

    ws["A1"].alignment = Alignment(
        horizontal="center",
        vertical="center"
    )

    # =============================
    # METADATA INFO
    # =============================

    start_row = 5

    metadata_rows = [
        ("Powered by:", "Slot-X Solutions"),
        ("Version:", "v2.0"),
        ("Report Type:", report_title),
        ("Branch:", branch_name),
        ("Payout Cycle:", payout_cycle),
        ("Generated At:", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    ]

    row = start_row

    for label, value in metadata_rows:

        ws[f"A{row}"] = label
        ws[f"A{row}"].font = Font(bold=True)

        ws[f"B{row}"] = value

        row += 1

    # =============================
    # COLUMN WIDTH
    # =============================

    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 35

    return ws