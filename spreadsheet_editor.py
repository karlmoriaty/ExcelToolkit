import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter, column_index_from_string
import config


class SpreadsheetEditor:
    def __init__(self, file_path=None):
        self.file_path = file_path
        self.workbook = None
        self.current_sheet = None
        self.load_workbook()

    def load_workbook(self):
        if self.file_path:
            try:
                self.workbook = openpyxl.load_workbook(self.file_path)
            except FileNotFoundError:
                self.workbook = openpyxl.Workbook()
        else:
            self.workbook = openpyxl.Workbook()

    def save_workbook(self, file_path=None):
        save_path = file_path or self.file_path
        if save_path:
            self.workbook.save(save_path)
            self.file_path = save_path

    def get_sheet_names(self):
        return self.workbook.sheetnames

    def select_sheet(self, sheet_name):
        if sheet_name in self.workbook.sheetnames:
            self.current_sheet = self.workbook[sheet_name]
            return True
        return False

    def create_sheet(self, sheet_name):
        if sheet_name not in self.workbook.sheetnames:
            self.current_sheet = self.workbook.create_sheet(sheet_name)
            return True
        return False

    def delete_sheet(self, sheet_name):
        if sheet_name in self.workbook.sheetnames and len(self.workbook.sheetnames) > 1:
            self.workbook.remove(self.workbook[sheet_name])
            if self.current_sheet and self.current_sheet.title == sheet_name:
                self.current_sheet = self.workbook.active
            return True
        return False

    def get_cell_value(self, row, col):
        if self.current_sheet:
            return self.current_sheet.cell(row=row, column=col).value
        return None

    def set_cell_value(self, row, col, value):
        if self.current_sheet:
            self.current_sheet.cell(row=row, column=col).value = value

    def get_range_data(self, start_row, start_col, end_row, end_col):
        data = []
        if self.current_sheet:
            for row in range(start_row, end_row + 1):
                row_data = []
                for col in range(start_col, end_col + 1):
                    row_data.append(self.get_cell_value(row, col))
                data.append(row_data)
        return data

    def set_range_data(self, start_row, start_col, data):
        if self.current_sheet:
            for i, row_data in enumerate(data):
                for j, value in enumerate(row_data):
                    self.set_cell_value(start_row + i, start_col + j, value)

    def apply_font_style(self, row, col, font_name=None, font_size=None, bold=False, italic=False, color=None):
        if self.current_sheet:
            cell = self.current_sheet.cell(row=row, column=col)
            font = cell.font
            cell.font = Font(
                name=font_name or font.name,
                size=font_size or font.size,
                bold=bold,
                italic=italic,
                color=color or font.color
            )

    def apply_fill_style(self, row, col, color):
        if self.current_sheet:
            cell = self.current_sheet.cell(row=row, column=col)
            cell.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")

    def apply_alignment_style(self, row, col, horizontal="center", vertical="center", wrap_text=False):
        if self.current_sheet:
            cell = self.current_sheet.cell(row=row, column=col)
            cell.alignment = Alignment(horizontal=horizontal, vertical=vertical, wrap_text=wrap_text)

    def apply_border_style(self, row, col, border_style="thin", apply_top=True, apply_bottom=True, apply_left=True, apply_right=True):
        if self.current_sheet:
            cell = self.current_sheet.cell(row=row, column=col)
            side = Side(border_style=border_style)
            cell.border = Border(
                top=side if apply_top else cell.border.top,
                bottom=side if apply_bottom else cell.border.bottom,
                left=side if apply_left else cell.border.left,
                right=side if apply_right else cell.border.right
            )

    def apply_range_font_style(self, start_row, start_col, end_row, end_col, **kwargs):
        for row in range(start_row, end_row + 1):
            for col in range(start_col, end_col + 1):
                self.apply_font_style(row, col, **kwargs)

    def apply_range_fill_style(self, start_row, start_col, end_row, end_col, color):
        for row in range(start_row, end_row + 1):
            for col in range(start_col, end_col + 1):
                self.apply_fill_style(row, col, color)

    def apply_range_alignment_style(self, start_row, start_col, end_row, end_col, **kwargs):
        for row in range(start_row, end_row + 1):
            for col in range(start_col, end_col + 1):
                self.apply_alignment_style(row, col, **kwargs)

    def apply_range_border_style(self, start_row, start_col, end_row, end_col, **kwargs):
        for row in range(start_row, end_row + 1):
            for col in range(start_col, end_col + 1):
                self.apply_border_style(row, col, **kwargs)

    def apply_table_border(self, start_row, start_col, end_row, end_col, border_style="thin", **kwargs):
        for row in range(start_row, end_row + 1):
            for col in range(start_col, end_col + 1):
                apply_top = kwargs.get("apply_top", (row == start_row))
                apply_bottom = kwargs.get("apply_bottom", (row == end_row))
                apply_left = kwargs.get("apply_left", (col == start_col))
                apply_right = kwargs.get("apply_right", (col == end_col))
                self.apply_border_style(row, col, border_style, apply_top, apply_bottom, apply_left, apply_right)

    def auto_fit_columns(self):
        if self.current_sheet:
            for col in self.current_sheet.columns:
                max_length = 0
                column = col[0].column_letter
                for cell in col:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 50)
                self.current_sheet.column_dimensions[column].width = adjusted_width

    def get_used_range(self):
        if self.current_sheet:
            return {
                "min_row": self.current_sheet.min_row,
                "max_row": self.current_sheet.max_row,
                "min_col": self.current_sheet.min_column,
                "max_col": self.current_sheet.max_column,
            }
        return None
