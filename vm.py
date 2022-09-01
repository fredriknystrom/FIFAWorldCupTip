from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill, Alignment, Border, Side
from itertools import combinations


def main():
    try:
        # Only works for xl with xlsx ext
        wb = load_workbook('VMQuiz.xlsx')
    except Exception:
        wb = Workbook()

    # Borders
    regular = Side(border_style="thin", color="000000")
    border = Border(regular, regular, regular, regular)

    # Worksheet
    ws = wb.active
    ws.title = 'VM Quiz 2022'

    # Group spacing
    group_col_space = 1
    first_groups_row = 2
    second_groups_row = 18

    # __First level groups__

    # Creating group A
    group_A_countries = ['Qatar', 'Ecuador', 'Senegal', 'Nederländerna']
    group_A_color = 'FF4500' # Orange Red
    group_A = Group('Group A', group_A_countries, first_groups_row, group_col_space, group_A_color, ws, border)
    
    # Creating group B
    group_B_countries = ['England', 'Iran', 'USA', 'Wales']
    group_B_color = '9ACD32' # yellow green
    group_B = Group('Group B', group_B_countries, first_groups_row, 2*group_col_space + 5, group_B_color, ws, border)

    # Creating group C
    group_C_countries = ['Argentina', 'Saudiarabien', 'Mexiko', 'Polen']
    group_C_color = 'FF8C00' # Dark Orange
    group_C = Group('Group C', group_C_countries, first_groups_row, 3*group_col_space + 10, group_C_color, ws, border)

    # Creating group D
    group_D_countries = ['Frankrike', 'Australien', 'Danmark', 'Tunisien']
    group_D_color = '1E90FF' # Dodger Blue
    group_D = Group('Group D', group_D_countries, first_groups_row, 4*group_col_space + 15, group_D_color, ws, border)

    # __Second level groups__

    # Creating group E
    group_E_countries = ['Spanien', 'Costa Rica', 'Tyskland', 'Japan']
    group_E_color = 'FFD700' # Gold
    group_E = Group('Group E', group_E_countries, second_groups_row, group_col_space, group_E_color, ws, border)

    # Creating group F
    group_F_countries = ['Belgien', 'Kanada', 'Marocko', 'Kroatien']
    group_F_color = 'A9A9A9' # Dark Grey
    group_F = Group('Group F', group_F_countries, second_groups_row, 2*group_col_space + 5, group_F_color, ws, border)

    # Creating group G
    group_G_countries = ['Brasilien', 'Serbien', 'Schweiz', 'Kamerun']
    group_G_color = 'EE82EE' # Violet
    group_G = Group('Group G', group_G_countries, second_groups_row, 3*group_col_space + 10, group_G_color, ws, border)

    # Creating group H
    group_H_countries = ['Portugal', 'Ghana', 'Uruguay', 'Sydkorea']
    group_H_color = '87CEEB' # Sky Blue
    group_H = Group('Group H', group_H_countries, second_groups_row, 4*group_col_space + 15, group_H_color, ws, border)
 

    # Auto size all the columns
    for i in range(1,30):
        ws.column_dimensions[get_column_letter(i)].auto_size = True

    wb.save('VMQuiz.xlsx')


class Group():

    def __init__(self, group_name, countries, row_start, col_start, color, ws, border):
        self.group_name = group_name
        self.headers = [group_name, 'Points', 'GS', 'GC', 'GD']
        self.countries = countries
        self.row_start = row_start
        self.row = row_start
        self.col_start = col_start
        self.fill_color = PatternFill(patternType = 'solid', fgColor = color)
        self.ws = ws
        self.border = border

        self.addGroup()

    def addGroup(self):
       
        # Add headers
        for i in range(len(self.headers)):
            col_letter = get_column_letter(self.col_start + i)
            cell = self.ws[col_letter + str(self.row)]
            
            cell.value = self.headers[i]
            cell.fill = self.fill_color
            cell.border = self.border
        
        self.row += 1

        # Generate matches for the group 
        def genCentralHeader(self, row, offset, title):
            start = get_column_letter(self.col_start + offset)
            end = get_column_letter(self.col_start + offset + 1)
            self.ws.merge_cells(f"{start}{row}:{end}{row}")
            cell = self.ws[f"{start}{row}"]
            cell.value = title
            cell.alignment = Alignment(horizontal='center')
            cell.border = self.border
            cell.fill = self.fill_color

        row = self.row_start + 6 

        genCentralHeader(self, row, 0, f"Matches for {self.group_name}")
        genCentralHeader(self, row, 2, 'Result')

        row += 1
        
        matches = list(combinations(self.countries, 2))
        for match in matches:
            for col_offset in range(4):
                cell = self.ws[get_column_letter(self.col_start + col_offset) + str(row)]
                if col_offset == 0:
                    value = match[0]
                elif col_offset == 1:
                    value = match[1]
                else:
                    value = ''
                cell.value = value
                cell.fill = self.fill_color
                cell.border = self.border
            row += 1

    
        # Add country 
        for i in range(len(self.countries)):
         
            for j in range(5):
                col_letter = get_column_letter(self.col_start + j)
                pos =col_letter + str(self.row + i)
                cell = self.ws[pos]
                
                # Set country name
                if j == 0:
                    cell.value = self.countries[i]
                # Get points
                elif j == 1:
                    cell.value = '=SUM(1,2,3)'
                # Get goal scored
                elif j == 2:
                    cell.value = 55 # add getter method
                # Get goal conceeded
                elif j == 3:
                    cell.value = 22 # add getter method
                # Get goal difference
                elif j == 4:
                    gs_cell = get_column_letter(self.col_start + j-2) + str(self.row + i)
                    gc_cell = get_column_letter(self.col_start + j-1) + str(self.row + i)
                    cell.value = f"=SUM({gs_cell}, -{gc_cell})"

                cell.fill = self.fill_color
                cell.border = self.border

        def getPoints():
            pass
        


if __name__ == "__main__":
    main()