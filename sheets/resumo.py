from utils.search import Search
from utils.read import Read
from woorkbook.woorkbook import saida, resumo

class Resumo:
    '''
    Gera coordenadas das colunas B de resumo, acompanhada de um número, que é a linha.
   
    '''
    resumo = Read().ws(saida, 'Resumo')
    
    def __init__(self):
        '''
        Busca todas as linhas por cada coluna da sheet resumo, no atributo, r significa row, e é
        acompanhado pelo numero da linha.
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
        self.r_17_centro_de_custo = Search().get_cords(resumo, 'bc7', 'o17')
        self.r_18 = Search().get_cords(resumo, 'c18', 'o17')
        self.r_19 = Search().get_cords(resumo, 'c19', 'o19')
        self.r_20 = Search().get_cords(resumo, 'c20', 'o20')
        self.r_21 = Search().get_cords(resumo, 'c21', 'o21')
        self.r_23 = Search().get_cords(resumo, 'c22', 'b23')
        self.r_24 = Search().get_cords(resumo, 'c24', 'o24')

    def despesas_diretas(self):

        return self.r_8
    
    def despesas_admnistrativas(self):
        
        return self.r_9
    
    def despesas_financeiras(self):
        
        return self.r_10
    
    def impostos_e_taxas(self):
        
        return self.r_11