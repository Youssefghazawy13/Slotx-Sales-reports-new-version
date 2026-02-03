
---

## 🧰 `core/excel_utils.py` (كامل)

```python
from openpyxl.styles import Font

def bold_headers(ws):
    """Make first row (headers) bold"""
    for cell in ws[1]:
        cell.font = Font(bold=True)


def auto_fit_columns(ws):
    """Auto fit column widths based on content"""
    for column_cells in ws.columns:
        length = 0
        col_letter = column_cells[0].column_letter
        for cell in column_cells:
            if cell.value:
                length = max(length, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = length + 3
# styling helpers
