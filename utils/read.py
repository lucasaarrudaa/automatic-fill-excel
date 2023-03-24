import openpyxl

class Read:
    
    def __init__(self):
        pass
    
    def wb(self, path_file):
        '''
        wb = workbook
        '''
        self.wbk = openpyxl.load_workbook(rf'{path_file}', data_only=True)
        return self.wbk
        
    def ws(self, wb, sheet):
        '''
        wb: var attributted to the worksheet
        '''
        self.ws = wb[f'{sheet}']
        return self.ws
