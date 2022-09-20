from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment, Border, Side


def set_value_to_cell(cell, value, color = None, alignment = None):

    # Borders
    regular = Side(border_style="thin", color="000000")
    border = Border(regular, regular, regular, regular)


    cell.value = value
    cell.border = border

    if color:
        cell.fill = color
    if alignment:
        cell.alignment = alignment
    else:
        cell.alignment = Alignment(horizontal='left')


def get_cell(col, row):
    return get_column_letter(col) + str(row)