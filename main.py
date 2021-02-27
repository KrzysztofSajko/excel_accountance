from __future__ import annotations
import os

from typing import Optional
from calendar import monthrange

from openpyxl import load_workbook
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.cell.cell import Cell

from employee import Employee
from functions import find_first
from gender import Gender
from summary import MonthlySummary


def get_employees(workbook: Workbook) -> list[Employee]:
    return [Employee(row[1], row[0], Gender.from_string(row[2]))
            for row
            in workbook.active.iter_rows(min_row=2, max_col=3, values_only=True)]


WORKING_DIRECTORY: str = "C:\\Users\\User\\Desktop\\ROK 2021"
TESTING_FILE: str = "I STYCZEŃ\\BANCERZ ELŻBIETA.xlsx"
DIR: str = "I STYCZEŃ"
EMPLOYEES_FILE: str = "Pracownicy.xlsx"

months: list[str] = ["Styczeń", "Luty", "Marzec",
                     "Kwiecień", "Maj", "Czerwiec",
                     "Lipiec", "Sierpień", "Wrzesień",
                     "Październik", "Listopad", "Grudzień"]

employees: list[Employee] = get_employees(load_workbook(os.path.join(WORKING_DIRECTORY, EMPLOYEES_FILE)))
dirs = [file for file in os.listdir(WORKING_DIRECTORY) if os.path.isdir(os.path.join(WORKING_DIRECTORY, file))]

for month_no, month in enumerate(months):
    # print(month)
    dir_index = find_first(dirs, lambda x: month.lower() in x.lower())
    directory = dirs[dir_index] if dir_index != -1 else None
    for i, employee in enumerate(employees):
        # print(f"  {i}. {employee.name} {employee.last_name}")
        if directory:
            summary_file = employee.find_summary(os.path.join(WORKING_DIRECTORY, directory))
            if summary_file:
                try:
                    employee.monthly_summaries[month] = MonthlySummary.get_from_worksheet(
                        load_workbook(summary_file)
                        .active,
                        width=monthrange(2021, month_no + 1)[1]
                    )
                    continue
                except Exception as e:
                    employee.monthly_summaries[month] = MonthlySummary.empty()
        employee.monthly_summaries[month] = MonthlySummary.empty()

for k, v in employees[7].monthly_summaries.items():
    print(k)
    for sv in v.entries:
        print(" ", sv.total_hours, sv.hours)

# agregation: MonthlySummary = MonthlySummary.agregate([employee.monthly_summaries["styczeń"]
#                                                       for employee
#                                                       in employees])
# for entry in agregation.entries:
#     print(f"Godziny: {entry.total_hours}, dni: {entry.days}, wykaz: {entry.hours}")
