from openpyxl.styles import Font, PatternFill


def create_metadata_sheet(wb, report_title, branch_name, payout_cycle):

    ws = wb.create_sheet("Metadata")

    fill = PatternFill(
        start_color="0A1F5C",
        end_color="0A1F5C",
        fill_type="solid"
    )

    ws["A1"] = "SLOT-X"
    ws["A1"].font = Font(size=18, bold=True, color="FFFFFF")

    for col in range(1, 6):
        ws.cell(row=1, column=col).fill = fill

    ws["A3"] = "Report:"
    ws["B3"] = report_title

    ws["A4"] = "Branch:"
    ws["B4"] = branch_name

    ws["A5"] = "Payout:"
    ws["B5"] = payout_cycle

    ws["A6"] = "Version:"
    ws["B6"] = "v2.0"

    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 28