import xlwings as xw
import pyautogui, sys, time


wb = xw.Book('C:\\Users\\craig\\Documents\\CS3398-Vulcans-S2020\\example.xlsm')
#wb = xw.Book(sys.argv[1]))
sheet = wb.sheets['Sheet1']
print('Press Ctrl-C to quit.')
print(str(sys.argv[0]))

try:
    while True:
        x, y = pyautogui.position()
        sheet.range('L3').value = x
        sheet.range('M3').value = y
except KeyboardInterrupt:
    print('\n')

wb.save('C:\\Users\\craig\\Documents\\CS3398-Vulcans-S2020\\example.xlsm')
#wb.save(sys.argv[1])