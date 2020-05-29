import xlwings as xw

wb = xw.Book('C:\\Users\\craig\\Documents\\CS3398-Vulcans-S2020\\example.xlsm')
#wb = xw.Book(sys.argv[1]))
sheet = wb.sheets['Sheet2']

x = 1
ProcedureNumber = "B" + str(x)
while sheet.range(ProcedureNumber).value is not None:
    x = x + 1
    ProcedureNumber = "B" + str(x)

sheet.range(ProcedureNumber).value = x
sheet1 = wb.sheets['Sheet1']
CreateDropDown = wb.macro('Sheet1.CreateDropDown')
CreateDropDown()

if x == 1:
    StartAddrInt = x
    EndAddrInt = 50
else:
    StartAddrInt = ((x-1)*50) + 1
    EndAddrInt = StartAddrInt + 49

my_values = wb.sheets['Sheet1'].range('B3:D52').options(numbers=int).value
wb.sheets['Sheet3'].range((StartAddrInt,1), (EndAddrInt,3)).value = my_values

wb.save('C:\\Users\\craig\\Documents\\CS3398-Vulcans-S2020\\example.xlsm')