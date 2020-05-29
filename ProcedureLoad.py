import xlwings as xw
import sys

wb = xw.Book('C:\\Users\\craig\\Documents\\CS3398-Vulcans-S2020\\example.xlsm')
#wb = xw.Book(sys.argv[1]))
sheet = wb.sheets['Sheet1']


ProcedureNumber = sheet.range('I3').value

if ProcedureNumber is None:
    sys.exit(0)


if ProcedureNumber == 1:
    StartAddrInt = ProcedureNumber
    EndAddrInt = 50
else:
    StartAddrInt = ((ProcedureNumber-1)*50) + 1
    EndAddrInt = StartAddrInt + 49


my_values = wb.sheets['Sheet3'].range((StartAddrInt,1), (EndAddrInt,3)).value
wb.sheets['Sheet1'].range('B3:D52').options(numbers=int).value = my_values

wb.save('C:\\Users\\craig\\Documents\\CS3398-Vulcans-S2020\\example.xlsm')