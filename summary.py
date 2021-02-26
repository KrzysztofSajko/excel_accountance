from __future__ import annotations

from dataclasses import dataclass, field
from unidecode import unidecode

from openpyxl.worksheet.worksheet import Worksheet


MONTHLY_SUMMARY_WIDTH: int = 33
MONTHLY_SUMMARY_HEIGHT: int = 16


@dataclass
class MonthlyEntry:
    number: str
    title: str
    hours: list[int] = field(init=False)

    def __post_init__(self):
        self.hours = [0] * 31

    def __add__(self, other):
        if type(self) is not type(other):
            raise AttributeError(f"'{type(other)}' can't be added to '{type(self)}'.")
        # unidecode is a temporary fix for typos in data
        # if unidecode(str(self.number)) != unidecode(str(other.number)) or unidecode(str(self.title)) != unidecode(str(other.title)):
        #     raise AttributeError(f"Can't add entries with non matching descriptions.")

        result = MonthlyEntry(self.number, self.title)
        result.hours = [h1 + h2 for h1, h2 in zip(self.hours, other.hours)]
        return result

    @classmethod
    def parse(cls, row: tuple) -> MonthlyEntry:
        entry: MonthlyEntry = cls(row[0], row[1])
        for i, cell in enumerate(row[2:]):
            entry.hours[i] = int(cell) if cell else 0
        return entry

    @property
    def total_hours(self) -> int:
        return sum(self.hours)

    @property
    def days(self) -> int:
        return sum(map(lambda x: 1 if x > 0 else 0, self.hours))


@dataclass
class MonthlySummary:
    entries: list[MonthlyEntry] = field(default_factory=list)

    @classmethod
    def parse(cls, worksheet: Worksheet, start_row: int, start_column: int) -> MonthlySummary:
        monthly_summary = cls()
        for row in worksheet.iter_rows(min_row=start_row,
                                       max_row=start_row + MONTHLY_SUMMARY_HEIGHT - 1,
                                       min_col=start_column,
                                       max_col=start_column+MONTHLY_SUMMARY_WIDTH - 1,
                                       values_only=True):
            monthly_summary.entries.append(MonthlyEntry.parse(row))
        return monthly_summary

    @classmethod
    def agregate(cls, summaries: list[MonthlySummary]) -> MonthlySummary:
        agregation = MonthlySummary()
        agregation.entries = [MonthlyEntry(entry.number, entry.title) for entry in summaries[0].entries]
        for summary in summaries:
            agregation.entries = [entry1 + entry2 for entry1, entry2 in zip(agregation.entries, summary.entries)]

        return agregation
