'''
@author: Maurice Koster
'''

from openpyxl import Workbook

class ExcelLibrary(object):
    '''
    Library for manipulating Excel 2007+ workbooks
    '''


    def __init__(self):
        '''
        Constructor
        '''
        self.open_workbooks = {}
        self.active_workbook = None
        self.active_workbook_shortname = None
        self.active_worksheet = None
      
    def create_workbook(self, filename, workbook_shortname='default'):
        wb = Workbook()
        self.open_workbooks[workbook_shortname] = {"workbook": wb, "filename": filename}
        self.active_workbook = wb
        self.active_workbook_shortname = workbook_shortname
        
    def open_workbook(self, filename):
        pass
    
    def close_workbook(self, save=True):
        pass
    
    def activate_workbook(self, index):
        pass
    
    def save_workbook(self, filename=None):
        if filename:
            fn = filename
        else:
            fn = self.open_workbooks[self.active_workbook_shortname]['filename']
            
        try:
            self.active_workbook.save(fn)
        except:
            pass
            
        
        
    
    def get_active_sheet_name(self):
        return self.active_workbook.get_active_sheet().title
    
    def get_sheet_names(self):
        return self.active_workbook.get_sheet_names()
    
    def select_sheet(self, sheet):
        if isnumeric(sheet):
            pass
        else:
            pass
        