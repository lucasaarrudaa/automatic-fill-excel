from search import Search
from read import Read
from woorkbook import saida

class Ccadm:
    '''
    Gera os valores das colunas B de CCadm, acompanhada de um número, que é a linha.
    '''
    def __init__(self):
        '''
        lê a sheet ccadm e busca todas as linhas por cada coluna da sheet ccadm, no atributo, r significa row, e é
        acompanhado pelo numero da linha.
        '''        
        ccadm = Read().ws(saida, 'CCAdm')
        self.ccadm_46 = Search().get_values(ccadm, 'c46', 'o46')
        self.ccadm_55 = Search().get_values(ccadm, 'c55', 'o55')

    def despesas_diretas(self):

        return self.ccadm_55
    
    def despesas_admnistrativas(self):

        return self.ccadm_55
