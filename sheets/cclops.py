from utils.search import Search
from utils.read import Read
from woorkbook.woorkbook import saida

class Cclops:
    '''
    Generates values for the B columns of Cclops, accompanied by a number, which is the row.
    '''
    def __init__(self):
        '''
        Reads the cclops sheet and searches all rows for each column of the cclops sheet.
        In the attribute, "r" stands for row and is followed by the line number.
        '''
        cclops = Read().ws(saida, 'CCLops')
        self.cclops_46 = Search().get_values(cclops, 'c46', 'o46')
        self.cclops_53 = Search().get_values(cclops, 'c53', 'o53')

    def direct_expenses(self):
        '''
        Returns the values for column B of cclops, which correspond to direct expenses.
        '''
        return self.cclops_46

    def administrative_expenses(self):
        '''
        Returns the values for column B of cclops, which correspond to administrative expenses.
        '''
        return self.cclops_53
