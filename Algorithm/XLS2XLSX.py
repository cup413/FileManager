import win32com.client as win32

import os


class XLS2XLSX:
    def tras(self, file):
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        wb = excel.Workbooks.Open(file)
        wb.SaveAs(file+'x', FileFormat = 51)    #FileFormat = 51 is for .xlsx extension
        wb.Close()                               #FileFormat = 56 is for .xls extension
        excel.Application.Quit()