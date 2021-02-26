from __future__ import annotations
import os

from typing import Optional

from openpyxl import load_workbook
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.cell.cell import Cell

from employee import Employee
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

employees: list[Employee] = get_employees(load_workbook(os.path.join(WORKING_DIRECTORY, EMPLOYEES_FILE)))
s = 0
for employee in employees:
    # print(employee.name, employee.last_name, employee.gender)
    # print()
    employee.add_monthly_summary("styczeń",
                                 load_workbook(employee
                                               .find_summary(os.path.join(WORKING_DIRECTORY, DIR)))
                                 .active)
    s += employee.monthly_summaries["styczeń"].entries[7].total_hours
    print(employee.monthly_summaries["styczeń"].entries[7].total_hours)
print(s)
agregation: MonthlySummary = MonthlySummary.agregate([employee.monthly_summaries["styczeń"] for employee in employees])
for entry in agregation.entries:
    print(entry.total_hours, entry.hours)
# workbook: Workbook = load_workbook(os.path.join(WORKING_DIRECTORY, TESTING_FILE))
# worksheet: Worksheet = workbook.active
#
# start_cell: Cell = find_cell_by_value(worksheet, "Lp.", max_row=20, max_col=40)
# if start_cell:
#     summary: MonthlySummary = MonthlySummary.parse(worksheet, start_cell.row + 1, start_cell.column)
#     for entry in summary.entries:
#         print(f"{entry.number} {entry.title} suma godzin: {entry.total_hours} ilość dni: {entry.days}")


# workbooks = []
#
# for dir in os.listdir(WORKING_DIRECTORY):
#     number, name = dir.split(' ', 1)
#     t = "Zestawienie miesieczne" if number != "XIII" else "Zestawienie roczne", name
#     print(t)


# for path, dirs, files in os.walk(WORKING_DIRECTORY):
#     print(path, dirs, files)
#     workbooks: list[Workbook] = [*workbooks, *[load_workbook(os.path.join(path, file)) for file in files if file.endswith("xlsx")]]

# cells: list[Cell] = []
# for workbook in workbooks:
#     sheet: Worksheet = workbook[workbook.sheetnames[0]]
#     cell = find_cell_by_value(sheet, "Lp.", max_col=40, max_row=20)
#     if cell:
#         cells.append(cell)
#
# for i, cell in enumerate(cells):
#     print(i, cell)