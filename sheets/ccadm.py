from utils.search import Search
from utils.read import Read
from woorkbook.woorkbook import saida

class Ccadm:
    '''
    Generates values for the B columns of CCadm, accompanied by a number, which is the row.
    '''
    def __init__(self):
        '''
        Reads the ccadm sheet and searches all rows for each column of the ccadm sheet. 
        In the attribute, "r" stands for row and is followed by the line number.
        '''        
        ccadm = Read().ws(saida, 'CCAdm')
        self.ccadm_46 = Search().get_values(ccadm, 'c46', 'o46')
        self.ccadm_55 = Search().get_values(ccadm, 'c55', 'o55')

    def despesas_diretas(self):

        return self.ccadm_46
    
    def despesas_admnistrativas(self):

        return self.ccadm_55
