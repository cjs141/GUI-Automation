#Python file to delete files
import os
import time
import xlwings as xw

#take a file path from the user
wb = xw.Book('C:\\Users\\craig\\Documents\\CS3398-Vulcans-S2020\\example.xlsm')
sheet = wb.sheets['Sheet1']
user_input = sheet.range('V2').value

#checking to see if the file exists
if os.path.exists(user_input):
    os.remove(user_input)
    sheet.range('V3').value = "Locating File"
    time.sleep(0.5)
    sheet.range('V3').value = "File Deleted"
    time.sleep(2)
    sheet.range('V3').value = "Waiting for user input"
else:
    sheet.range('V3').value = "Locating File"
    time.sleep(0.5)
    sheet.range('V3').value = "File Path Does Not Exist"
    time.sleep(2)
    sheet.range('V3').value = "Waiting for user input"