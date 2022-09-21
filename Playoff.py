from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter
from util_funcs import get_cell, set_value_to_cell

class Playoff():

    def __init__(self, row_start, col_start, groups, color, ws):

        self.row_start = row_start
        self.col_start = col_start
        self.fill_color = PatternFill(patternType = 'solid', fgColor = color)
        self.groups = groups
        self.ws = ws


        self.generate_playoff()

    def generate_playoff(self):

        self.generate_16()
        self.generate_quarterfinals()
        self.generate_semifinals()
        self.generate_final()


    def generate_16(self):

        cell = self.ws[get_cell(self.col_start, self.row_start)]
        set_value_to_cell(cell, 'Round of 16', self.fill_color)

        row = self.row_start + 1

        for r in range(0,7,2):
            teams = [self.groups[r].get_winner(), self.groups[r+1].get_second(), self.groups[r+1].get_winner(), self.groups[r].get_second()]

            text1 = f'Group {get_column_letter(r+1)}'
            text2 = f'Group {get_column_letter(r+2)}'

            for team in range(0,4,2):
                for c in range(4):
                    cell1 = self.ws[get_cell(self.col_start + c, row)]
                    cell2 = self.ws[get_cell(self.col_start + c, row+1)]
                    if c == 0:
                        value1 = f'1st in {text1}'
                        value2 = teams[team]
                    if c == 1:
                        value1 = f'2nd in {text2}'
                        value2 = teams[team+1]
                    if c == 2:
                        value1 = 'Score'
                        value2 = ''
                    if c == 3:
                        value1 = ''
                        value2 = ''
                    set_value_to_cell(cell1, value1, self.fill_color)
                    set_value_to_cell(cell2, value2, self.fill_color)

                tmp = text1
                text1 = text2
                text2 = tmp
                row += 4

    def get_round_of_16_winners(self):
        col = self.col_start

        winners = []

        for i in range(8):
            row = self.row_start + i*4 + 2
            formula = f'=IF({get_cell(col+2, row)} > {get_cell(col+3, row)}, {get_cell(col, row)}, {get_cell(col+1, row)})'
            winners.append(formula)

        return winners

    def generate_quarterfinals(self):
        col = self.col_start + 5
        cell = self.ws[get_cell(col, self.row_start)]
        set_value_to_cell(cell, 'Quarterfinals', self.fill_color)

        row = self.row_start + 1
        teams = self.get_round_of_16_winners()

        for r in range(0,8,2):

            for c in range(4):
                cell1 = self.ws[get_cell(col + c, row)]
                cell2 = self.ws[get_cell(col + c, row+1)]
                if c == 0:
                    value1 = f'Quarter {r/2+1}'
                    value2 = teams[r]
                if c == 1:
                    value1 = f''
                    value2 = teams[r+1]
                if c == 2:
                    value1 = 'Score'
                    value2 = ''
                if c == 3:
                    value1 = ''
                    value2 = ''
                set_value_to_cell(cell1, value1, self.fill_color)
                set_value_to_cell(cell2, value2, self.fill_color)

            row += 4

    def generate_semifinals(self):
        pass

    def generate_final(self):
        pass
        