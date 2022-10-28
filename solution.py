from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from util_funcs import get_cell
import os


def main():

    solution_wb = load_workbook('quizes/solution.xlsx')
    solution_ws = solution_wb.active
    result = dict()

    for file in os.listdir('quizes'):
        if file != 'solution.xlsx':
            wb = load_workbook(f'quizes/{file}')
            ws = wb.active
    
            name = file.split('.')[0]

            result[name] = compare_tip(ws, solution_ws)

    print(result)


def compare_tip(ws, solution_ws):

    total_points = 0

    total_points += group_points(ws, solution_ws)

    return total_points


def group_points(ws, solution_ws):
    points = 0
    row = 3
    max = 64

    for i in range(3, max):

        if row % 8 == 0:
            row += 3
        else:
            for col in range(3,6):
                cell = get_cell(col, row)
                if ws[cell].value == solution_ws[cell].value:
                    points += 1
            row += 1
        if row > max:
            return points
def round_of_16_points():
    pass

def quarter_points():
    pass

def semi_points():
    pass

def final_points():
    pass
    

if __name__ == '__main__':
    main()