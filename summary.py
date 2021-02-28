from __future__ import annotations

from dataclasses import dataclass, field

from openpyxl.cell import Cell
from openpyxl.worksheet.worksheet import Worksheet

from functions import find_cell_by_value

MONTHLY_SUMMARY_WIDTH: int = 31
MONTHLY_SUMMARY_HEIGHT: int = 16


@dataclass
class MonthlyEntry:
    days: int = 0
    hours: list[int] = field(default_factory=list)

    def __add__(self, other):
        if type(self) is not type(other):
            raise AttributeError(f"'{type(other)}' can't be added to '{type(self)}'.")

        result = MonthlyEntry(self.days + other.days)
        result.hours = [sum(hours) for hours in zip(self.hours, other.hours)]
        return result

    @classmethod
    def empty(cls) -> MonthlyEntry:
        return MonthlyEntry(0, [0] * MONTHLY_SUMMARY_WIDTH)

    @classmethod
    def parse(cls, row: tuple) -> MonthlyEntry:
        entry: MonthlyEntry = cls()
        entry.hours = [int(cell) if cell else 0 for cell in row]
        # pad with zeros to full length
        if len(entry.hours) < MONTHLY_SUMMARY_WIDTH:
            entry.hours = [*entry.hours, *[0 for _ in range(MONTHLY_SUMMARY_WIDTH - len(entry.hours))]]
        entry.days = entry.calc_days()
        return entry

    @property
    def total_hours(self) -> int:
        return sum(self.hours)

    def calc_days(self) -> int:
        return sum(map(lambda x: 1 if x > 0 else 0, self.hours))


@dataclass
class MonthlySummary:
    entries: list[MonthlyEntry] = field(default_factory=list)

    @classmethod
    def empty(cls) -> MonthlySummary:
        return cls([MonthlyEntry.empty() for _ in range(MONTHLY_SUMMARY_HEIGHT)])

    @classmethod
    def parse(cls, worksheet: Worksheet, start_row: int, start_column: int, width: int) -> MonthlySummary:
        monthly_summary = cls()
        for row in worksheet.iter_rows(min_row=start_row,
                                       max_row=start_row + MONTHLY_SUMMARY_HEIGHT - 1,
                                       min_col=start_column,
                                       max_col=start_column + width - 1,
                                       values_only=True):
            monthly_summary.entries.append(MonthlyEntry.parse(row))
        return monthly_summary

    @classmethod
    def get_from_worksheet(cls, worksheet: Worksheet, width: int = MONTHLY_SUMMARY_WIDTH,
                           offset_row: int = 0, offset_col: int = 0) -> MonthlySummary:
        start_cell: Cell = find_cell_by_value(worksheet, "Lp.",
                                              min_row=offset_row, max_row=offset_row + MONTHLY_SUMMARY_HEIGHT - 1,
                                              min_col=offset_col, max_col=offset_col + MONTHLY_SUMMARY_WIDTH - 1)
        if not Cell:
            raise AttributeError(f"Can't find summary in worksheet {worksheet.title}"
                                 f" under position ({offset_row}, {offset_col})")

        return MonthlySummary.parse(worksheet, start_cell.row + 1, start_cell.column + 2, width)

    @classmethod
    def agregate(cls, summaries: list[MonthlySummary]) -> MonthlySummary:
        return MonthlySummary([sum(entries, start=MonthlyEntry.empty())
                               for entries
                               in zip(*(summary.entries for summary in summaries))])
