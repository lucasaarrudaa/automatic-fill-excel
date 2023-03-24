from search import Search
from read import Read
from woorkbook import saida

class Cclops:
    '''
    Gera os valores das colunas B de cclops, acompanhada de um número, que é a linha.
    '''
    def __init__(self):
        '''
        lê a sheet cclops e busca todas as linhas por cada coluna da sheet cclops, no atributo, r significa row, e é
        acompanhado pelo numero da linha.
        '''
        cclops = Read().ws(saida, 'CCLops')
        self.cclops_46 = Search().get_values(cclops, 'c46', 'o46')
        self.cclops_53 = Search().get_values(cclops, 'c53', 'o53')

    def despesas_diretas(self):

        return self.cclops_46

        
    def despesas_admnistrativas(self):

        return self.cclops_53