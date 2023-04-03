from utils.search import Search
from utils.read import Read
from woorkbook.woorkbook import saida, resumo

class Resumo:
    '''
    Generates the coordinates of the columns in the Resumo sheet, accompanied by a number, which represents the row.
    '''
    resumo = Read().ws(saida, 'Resumo')
    
    def __init__(self):
        '''
        Searches all rows for each column of the resumo sheet. The attribute "r" stands for row and is accompanied by the row number.
        '''
        self.r_8 = Search().get_cords(resumo, 'c8', 'o8')
        self.r_9 = Search().get_cords(resumo, 'c9', 'o9')
        self.r_10 = Search().get_cords(resumo, 'c10', 'o10')
        self.r_11 = Search().get_cords(resumo, 'c11', 'o11')
        self.r_12 = Search().get_cords(resumo, 'c12', 'o12')
        self.r_13 = Search().get_cords(resumo, 'c13', 'o13')
        self.r_14 = Search().get_cords(resumo, 'c14', 'o14')
        self.r_15 = Search().get_cords(resumo, 'c15', 'o15')
        self.r_16 = Search().get_cords(resumo, 'c16', 'o16')
        self.r_17_centro_de_custo = Search().get_cords(resumo, 'c7', 'o17')
        self.r_18 = Search().get_cords(resumo, 'c18', 'o17')
        self.r_19 = Search().get_cords(resumo, 'c19', 'o19')
        self.r_20 = Search().get_cords(resumo, 'c20', 'o20')
        self.r_21 = Search().get_cords(resumo, 'c21', 'o21')
        self.r_23 = Search().get_cords(resumo, 'c22', 'b23')
        self.r_24 = Search().get_cords(resumo, 'c24', 'o24')

    def despesas_diretas(self):
        '''
        Returns the coordinates of the "Despesas Diretas" cell in the Resumo sheet (row 8).
        '''
        return self.r_8
    
    def despesas_admnistrativas(self):
        '''
        Returns the coordinates of the "Despesas Administrativas" cell in the Resumo sheet (row 9).
        '''
        return self.r_9
    
    def despesas_financeiras(self):
        '''
        Returns the coordinates of the "Despesas Financeiras" cell in the Resumo sheet (row 10).
        '''
        return self.r_10
    
    def impostos_e_taxas(self):
        '''
        Returns the coordinates of the "Impostos e Taxas" cell in the Resumo sheet (row 11).
        '''
        return self.r_11
