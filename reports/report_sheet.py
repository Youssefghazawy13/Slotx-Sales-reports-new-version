    # =====================================================
    # KPI CARDS (Fixed Width Columns B, D, G)
    # =====================================================

    # Fix column widths first
    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["D"].width = 22
    ws.column_dimensions["G"].width = 22

    # Reduce neighboring columns so layout stays clean
    ws.column_dimensions["A"].width = 5
    ws.column_dimensions["C"].width = 5
    ws.column_dimensions["E"].width = 5
    ws.column_dimensions["F"].width = 5

    def create_kpi_card(row, col_letter, title, value):

        cell = ws[f"{col_letter}{row}"]
        cell.value = f"{title}\n{value}"

        cell.font = Font(size=12, bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center",
                                   vertical="center",
                                   wrap_text=True)

        fill = PatternFill(
            start_color="0A1F5C",
            end_color="0A1F5C",
            fill_type="solid"
        )

        border = Border(
            left=Side(style="thin"),
            right=Side(style="thin"),
            top=Side(style="thin"),
            bottom=Side(style="thin")
        )

        # Height أطول شوية
        ws.row_dimensions[row].height = 45

        cell.fill = fill
        cell.border = border

    create_kpi_card(1, "B", "Total Sales",
                    f"{total_sales_money:,.2f} EGP")

    create_kpi_card(1, "D", "Net After Deal",
                    f"{after_rent:,.2f} EGP")

    create_kpi_card(1, "G", "Inventory Units",
                    total_inventory_qty)
