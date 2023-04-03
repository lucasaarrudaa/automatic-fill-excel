from utils.search import Search
from utils.read import Read
from woorkbook.woorkbook import saida

class Ccmkt:
    '''
    Generates the values of column B in CCadm, accompanied by a number which represents the row.
    '''
    def __init__(self):
        '''
        Reads the ccmkt sheet and searches all rows for each column of the ccmkt sheet. The attribute "r" stands for row and
        is accompanied by the row number.
        '''
        ccmkt = Read().ws(saida, 'CCMkt')
        self.ccmkt_46 = Search().get_values(ccmkt, 'c46', 'o46')
        self.ccmkt_57 = Search().get_values(ccmkt, 'c57', 'o57')

    def despesas_diretas(self):
        '''
        Returns the values of column B in ccmkt accompanied by their respective row number (46).
        '''
        return self.ccmkt_46
    
    def despesas_admnistrativas(self):
        '''
        Returns the values of column B in ccmkt accompanied by their respective row number (57).
        '''
        return self.ccmkt_57
