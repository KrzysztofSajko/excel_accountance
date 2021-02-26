from __future__ import annotations

import os

from dataclasses import dataclass, field
from typing import Optional

from openpyxl.cell import Cell
from openpyxl.worksheet.worksheet import Worksheet

from gender import Gender
from summary import MonthlySummary
from functions import find_cell_by_value


@dataclass
class Employee:
    name: str
    last_name: str
    gender: Gender
    monthly_summaries: dict[str, MonthlySummary] = field(default_factory=dict)

    def find_summary(self, directory: str) -> Optional[str]:
        for filename in os.listdir(directory):
            if filename.endswith(".xlsx"):
                *last_name, name = tuple(map(lambda part: part.capitalize(), filename.split(".")[0].split()))
                last_name = " ".join(last_name)
                if last_name == self.last_name and name == self.name:
                    return os.path.join(directory, filename)

    def add_monthly_summary(self, key: str, worksheet: Worksheet):
        start_cell: Cell = find_cell_by_value(worksheet, "Lp.", max_row=10, max_col=10)
        self.monthly_summaries[key] = MonthlySummary.parse(worksheet, start_cell.row + 1, start_cell.column)
