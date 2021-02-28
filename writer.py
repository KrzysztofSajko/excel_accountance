from collections import Callable
from dataclasses import dataclass, field
from datetime import date
from calendar import monthrange

from workalendar.core import CoreCalendar
from xlsxwriter.exceptions import FileCreateError
from xlsxwriter.worksheet import Worksheet
from xlsxwriter.workbook import Workbook
from xlsxwriter.format import Format

from employee import Employee
from summary import MonthlyEntry


TABLE_ROW_LABELS = [
    ("1", "Czas przepracowany ogółem, w tym:"),
    ("a)", "w niedziele i święta"),
    ("b)", "w porze nocnej"),
    ("c)", "w godz. nadliczbowych 50%"),
    ("d)", "w godz. nadliczbowych 100%"),
    ("e)", "w dni wolne od pracy"),
    ("2", "Dyżury"),
    ("3", "Urlopy"),
    ("4", "Choroba w tym:"),
    ("a)", "płatne ESM"),
    ("b)", "płatne ZUS"),
    ("5", "Inne zasiłkowe"),
    ("6", "Nieobecności usprawiedliwione"),
    ("a)", "płatne"),
    ("b)", "niepłatne"),
    ("7", "Nieobecności nieusprawiedliwione")
]


@dataclass
class Writer:
    filename: str
    year: int
    employees: list[Employee]
    calendar: CoreCalendar
    extra_holidays: list[date]
    months: list[str]
    workbook: Workbook = field(init=False)
    formats: dict[str, Format] = field(init=False)
    bg_colors: dict[str, str] = field(init=False)

    def __post_init__(self):
        self.workbook = Workbook(self.filename)
        self.formats = {
            "name": self.workbook.add_format({'font_size': 15, 'bold': True}),
            "header": self.workbook.add_format({'font_size': 12, 'bold': True, 'align': 'center', 'border': 1}),
            "label": self.workbook.add_format({'font_size': 8, 'bold': True, 'text_wrap': True, 'border': 1}),
            "content": self.workbook.add_format({'font_size': 10, 'align': 'center', 'border': 1}),
            "summary": self.workbook.add_format({'font_size': 12, 'align': 'center', 'bold': True, 'border': 1}),
        }
        self.bg_colors = {
            "holiday": "#b3ecff",
            "overday": "#ffb3b3"  # days that go over the range of month
        }

    def setup_sheet(self, title: str) -> Worksheet:
        sheet: Worksheet = self.workbook.add_worksheet(title)
        sheet.set_column(0, 0, width=3)
        sheet.set_column(1, 1, width=25)
        sheet.set_column(2, 32, width=3)
        sheet.set_column(33, 34, width=15)
        sheet.hide_zero()
        return sheet

    def copy_format(self, base_format: Format) -> Format:
        properties = [function[4:] for function in dir(base_format) if function.startswith("set_")]
        default_format = self.workbook.add_format()
        return self.workbook.add_format(
            {k: v for k, v in base_format.__dict__.items() if k in properties and default_format.__dict__[k] != v})

    def set_format_bg_color(self, base_format: Format, month, day) -> Format:
        if day <= monthrange(self.year, month)[1]:
            if self.calendar.is_working_day(date(self.year, month, day), extra_holidays=self.extra_holidays):
                return base_format
            bg_color = self.bg_colors["holiday"]
        else:
            bg_color = self.bg_colors["overday"]
        new_format = self.copy_format(base_format)
        new_format.set_bg_color(bg_color)
        return new_format

    def write_day_colored_row(self, sheet: Worksheet, base_format: Format, value_function: Callable, month: int,
                              start_row: int = 0, start_col: int = 0):
        for day in range(31):
            cell_format = self.set_format_bg_color(base_format, month, day + 1)
            sheet.write(start_row, start_col + day, value_function(day), cell_format)

    def setup_table(self, sheet: Worksheet, employee: Employee, row_no: int, month: int):
        def setup_header():
            sheet.write_string(row_no + 1, 0, "Lp.", self.formats["header"])
            sheet.write_blank(row_no + 1, 1,  None, self.formats["header"])
            self.write_day_colored_row(sheet, self.formats["header"], lambda x: f"{x + 1}",
                                       month, row_no + 1, 2)
            sheet.write_string(row_no + 1, 33, "Suma godzin", self.formats["header"])
            sheet.write_string(row_no + 1, 34, "Suma dni", self.formats["header"])
        # name
        sheet.merge_range(row_no, 0, row_no, 34, f"{employee.last_name} {employee.name}", self.formats["name"])
        # header
        setup_header()
        # labels
        sheet.write_column(row_no + 2, 0,
                           [label[0] for label in TABLE_ROW_LABELS],
                           self.formats["label"])
        # summaries
        sheet.write_column(row_no + 2, 1,
                           [label[1] for label in TABLE_ROW_LABELS],
                           self.formats["label"])

    def fill_table_content(self, sheet: Worksheet, entries: list[MonthlyEntry], row_no: int, month: int):
        for entry_no, entry in enumerate(entries):
            self.write_day_colored_row(sheet, self.formats["content"], lambda x: entry.hours[x],
                                       month, row_no + entry_no, 2)
            sheet.write_formula(f"AH{row_no + 1 + entry_no}",
                                f"=SUM(C{row_no + 1 + entry_no}:AG{row_no + 1 + entry_no})",
                                self.formats["summary"])
            sheet.write_formula(f"AI{row_no + 1 + entry_no}",
                                f'=COUNTIF(C{row_no + 1 + entry_no}:AG{row_no + 1 + entry_no}, ">0")',
                                self.formats["summary"])

    def save_and_exit(self):
        while True:
            try:
                self.workbook.close()
            except FileCreateError as exc:
                decision = input(f"Wyjątek przy próbie zapisania pliku: {exc}\n"
                                 f"Zamknij plik {self.filename} jeśli jest otwarty.\n"
                                 f"Czy chcesz ponownie spróbować zapisać? [T/n]: ")
                if decision != 'n':
                    continue
            break

    def create(self):
        for month_no, month in enumerate(self.months):
            month_sheet: Worksheet = self.setup_sheet(month)
            for employee_no, employee in enumerate(self.employees):
                row_no = employee_no * 20
                self.setup_table(month_sheet, employee, row_no, month_no + 1)
                self.fill_table_content(month_sheet, employee.monthly_summaries[month].entries,
                                        row_no + 2, month_no + 1)
        self.save_and_exit()
