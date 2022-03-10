import win32com.client as win32

file = r'C:\Users\HP\Desktop\abc\顺北5井基本数据表.xls'
excel = win32.gencache.EnsureDispatch('Excel.Application')
wb = excel.Workbooks.Open(file)
wb.SaveAs(file+'x', FileFormat = 51)    #FileFormat = 51 is for .xlsx extension
wb.Close()                               #FileFormat = 56 is for .xls extension
excel.Application.Quit()