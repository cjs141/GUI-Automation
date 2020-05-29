import xlwings as xw

wb = xw.Book()  # this will create a new workbook

wb = xw.book('C:\path\to\file.xlsx') # this will conncect to the exisiting excel file using path given

sht = wb.sheets['Sheet1'] # instantiates a sheet object

sht.range('A1').value = 1 # writes value

sht.range('A1').value # reads value

sht.range('A2').value = 'Hello World'

sht.range('A2').value

sht.range('A3').value is None 
