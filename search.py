import re


class Search:

    def __init__(self):
        pass
        '''
    Finds patterns
    '''

    def value_position_cols(self, ws, col_1, col_2, value):
        '''
        Search a value in all cells in range of cols
        EX: value_position_cols(sheet, 1, 2 'apple')
        output: 'B34'

        Parameters: 
                value = value (int,string)
                ws = sheet (worksheet)
        Returns: 
                position of cell
        '''
        for cols in ws.iter_cols(min_col=col_1, max_col=col_2):
            for cell in cols:
                if (cell.value) == f'{value}':
                    # removes the strings that talk about which sheet it is (openpyxl's own)
                    cell = re.findall('([A-Z]+\d{1,2})', str(cell))
                    return cell.cordinate

    def value_position_rows(self, ws, row_1, row_2, value):
        '''
        Search a value in all cells in range of rows
        EX: value_position_cols(sheet, 1, 2, 'apple')
        output: 'G2'

        Parameters: 
                value = value (int, string)
                ws = sheet (worksheet)
        Returns: 
                position of cell
        '''
        for rows in ws.iter_rows(min_row=row_1, max_row=row_2):
            for cell in rows:
                if (cell.value) == f'{value}':
                    # removes the strings that talk about which sheet it is (openpyxl's own)
                    cell = re.findall('([A-Z]+\d{1,2})', str(cell))
                    return cell.cordinate

    def find_all_sheets_value(self, sheet1, sheet2, sheet3, string):
        '''
        Find a value in all sheets per name
        '''
        Search().value_position_cols(sheet1, 2, 2, f'{string}')
        Search().value_position_cols(sheet2, 2, 2, f'{string}')
        Search().value_position_cols(sheet3, 2, 2, f'{string}')

    def find_all_sheets_value(self, sheet1, sheet2, sheet3, string):
        '''
        Find a value in all sheets per name

        Returns a list with values of sheet1, sheet2, sheet3.
        '''
        self.cord1 = []
        self.cord2 = []
        self.cord3 = []
        self.cord1.append(self.value_position_cols(sheet1, 2, 2, f'{string}'))
        self.cord2.append(self.value_position_cols(sheet2, 2, 2, f'{string}'))
        self.cord3.append(self.value_position_cols(sheet3, 2, 2, f'{string}'))

        return self.cord1

    def get_cords(self, ws, by, to):
        '''
        generate cords

        Parameters:
            ws: var of sheet
            by: coordinate from
            to: coordinate to
        Returns:
        '''

        self.cords = []
        wsheet = ws[f'{by}'.upper():f'{to}'.upper()]
        for r in wsheet:
            for x in r:
                self.cords.append(x.coordinate)
        return self.cords

    def get_values(self, ws, by, to):
        '''
        Generate cords
        Ex: get_values(sheet, 'C1', 'C4')
            output: ['C1', 'C2', 'C3', 'C4']

        Parameters:
                ws: var of sheet (worksheet)
                by: coordinate from (string)
                to: coordinate to (string)
        Returns:
                a list of coordinates in sequence
        '''

        self.vals = []
        self.wsheet = ws[f'{by}'.upper():f'{to}'.upper()]
        for r in self.wsheet:
            for x in r:
                self.vals.append(x.value)
        return self.vals

    def headcount_list_cols(self, sheet):
        '''
        Iterates each row of the headcount column without None values

        Returns: 
                all column b values ​​in list type
        '''
        list_cols = []

        for cols in sheet.iter_cols(min_col=2, max_col=2, min_row=6, max_row=84):
            for cell in cols:
                # filling a list with all the lines of each sheet
                list_cols.append(cell.value)

        while None in list_cols:
            list_cols.remove(None)  # removing None

        return list_cols
