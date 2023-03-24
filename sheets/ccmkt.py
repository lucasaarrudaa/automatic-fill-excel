from utils.search import Search
from utils.read import Read
from woorkbook.woorkbook import saida

class Ccmkt:
    '''
    Gera os valores das colunas B de CCadm, acompanhada de um número, que é a linha.
    '''
    def __init__(self):
        '''
        lê a sheet ccmkt e busca todas as linhas por cada coluna da sheet ccmkt, no atributo, r significa row, e é
        acompanhado pelo numero da linha.
        '''
        ccmkt = Read().ws(saida, 'CCMkt')
        self.ccmkt_46 = Search().get_values(ccmkt, 'c46', 'o46')
        self.ccmkt_57 = Search().get_values(ccmkt, 'c57', 'o57')

    def despesas_diretas(self):

        return self.ccmkt_46
    
    def despesas_admnistrativas(self):

        return self.ccmkt_57
    
    