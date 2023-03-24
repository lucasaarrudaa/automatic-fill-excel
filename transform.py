class Transform():
    
    def __init__(self):
        pass
            
    def brl(self, value):
        '''
        Converts to BRL format
        Parameters: 
                value: value to be converted (int)
        '''
        self.real = 'R$          {:,.2f}'.format(value)
        return self.real

    def fill_cell(self, ws, cell, value):
        '''
        Insert a specific value into a single cell
        Parameters: 
                ws: sheet, ex = resume
                cell: cell to be inserted
                value: value to be inserted
        '''
        ws[f'{cell}'.upper()] = f'{value}'

    def complete(self, ws, values, coords):
        '''
        Fill multiple cells
        For it to work, you have to do a for on all lines, which you are sure will match.
        
        Parameters: 
                ws: sheet you want to fill (worksheet)
                values: (list)
                coords: (list)
        '''
        for v, c in zip(values, coords):
            self.fill_cell(ws, c, v)
            
    def convert_to_brl(self, vals_ws_1, vals_ws_2, vals_ws_3, coords_sheet_destiny, ws_to_fill):
        ''''
        This function will convert all the lines of each sheet to brl, and later insert them into a sheet
        
        Parameters: 
                vals_ws_1: sheet 1 values (list)
                vals_ws_2: sheet 2 values (list)
                vals_ws_3: sheet 3 values (list)
                coords_sheet_destiny: (list)
                ws_to_fill: sheet you want to fill (worksheet)
                each parameter is a list of values ​​from each sheet that you want to insert in the final table.
        Returns: 
                list of values ​​converted to real
        '''
        sum_brl = []  # values to add in rows da linha especifica 
        sum_dd = []  #valores de todas as sheets da linha especifica
    
        for a, l, m in zip(vals_ws_1, vals_ws_2, vals_ws_3):
            '''
            Para cada valor nas linhas de cada sheet, 
            soma tudo e coloca numa variavel, por cada mes. 
            '''
            summ = round(a, 2)+round(l, 2)+round(m, 2)
            sum_dd.append(summ)
        
        for x in sum_dd:
            sum_brl.append(Transform().brl(x)) #for each value in each sheet, transform it into brl

        for v, c in zip(sum_brl, coords_sheet_destiny):
            Transform().fill_cell(ws_to_fill, c, v) #for each value in each sheet, put it in the target sheet
            
        return sum_brl