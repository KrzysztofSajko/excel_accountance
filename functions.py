from typing import Optional

from openpyxl.cell import Cell
from openpyxl.worksheet.worksheet import Worksheet


def find_cell_by_value(worksheet: Worksheet, value: str, **kwargs) -> Optional[Cell]:
    for row in worksheet.iter_rows(**kwargs):
        for cell in row:
            if cell.value == value:
                return cell
