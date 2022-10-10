
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from Group import Group
from Playoff import Playoff
import os


def main():
    path = os.path.abspath('quizes/test.xlsx')
    if os.path.exists(path):
        os.remove(path)

    # Create workbook and activate worksheet
    wb = Workbook()
    ws = wb.active
    ws.title = 'VM Quiz 2022'

    group_countires = [
        ['Qatar', 'Ecuador', 'Senegal', 'Nederländerna'],
        ['England', 'Iran', 'USA', 'Wales'],
        ['Argentina', 'Saudiarabien', 'Mexiko', 'Polen'],
        ['Frankrike', 'Australien', 'Danmark', 'Tunisien'],
        ['Spanien', 'Costa Rica', 'Tyskland', 'Japan'],
        ['Belgien', 'Kanada', 'Marocko', 'Kroatien'],
        ['Brasilien', 'Serbien', 'Schweiz', 'Kamerun'],
        ['Portugal', 'Ghana', 'Uruguay', 'Sydkorea']
    ]

    group_colors = [
        'FF4500',           # Orange Red
        '9ACD32',           # yellow green
        'FF8C00',           # Dark Orange
        '1E90FF',           # Dodger Blue
        'FFD700',           # Gold
        'A9A9A9',           # Dark Grey
        'EE82EE',           # Violet
        '87CEEB',           # Sky Blue
    ]

    playoff_color = '888888'

    group_labels = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H']

    # Group spacing
    col_start = 1
    row_start = 2
    row_offset = 8
 
    groups = []

    # Creating groups
    for i in range(len(group_countires)):
        groups.append(Group(f'Group {group_labels[i]}', group_countires[i],  group_colors[i], row_start +  i*row_offset , col_start, ws))


    Playoff(row_start, 14, groups, playoff_color, ws)
    
 
    # Set width of all the columns in range below
    for i in range(1,40):
        ws.column_dimensions[get_column_letter(i)].width = 15

    wb.save('test.xlsx')


if __name__ == "__main__":
    main()