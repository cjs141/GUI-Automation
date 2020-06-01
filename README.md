# GUI Automation
> Creator: Craig Sandlin

> Project Description: Automation tool designed to relieve the user from repetitive tasks. Works across all applications; from websites to spreadsheets, file explorer, and everything inbetween. A friendly excel UI allows those without programming knowledge to automate any task that includes clicking, typing, select all, backspace, enter, wait, and more! 

## General info
> The motivation: The mother of all invention- necessity. Electrical measurements were being performed on lighting fixtures to characterize the product before approbation listing with UL and/or Intertek. The sample size for this testing was sometimes upward of 70 discrete measurements. The changes that needed to be made to the software were repetitive and regular. If only I could program the inputs ahead of time! But alas, the software was propriatery and no API was available. It is tasks like these where this software really shines. With some upfront setup the entire process could be automated. Even allowing measurements to be taken over night which improved throughput!


## Technologies
* Python 3 - version 3.5
* Excel - version Latest


## Code Examples
>Display Mouse Position:
>import xlwings as xw
import pyautogui, sys, time


wb = xw.Book('C:\\Users\\craig\\Documents\\CS3398-Vulcans-S2020\\example.xlsm')
#wb = xw.Book(sys.argv[1]))
sheet = wb.sheets['Sheet1']
print('Press Ctrl-C to quit.')
print(str(sys.argv[0]))

try:
    while True:
        x, y = pyautogui.position()
        sheet.range('H2').value = x
        sheet.range('I2').value = y
except KeyboardInterrupt:
    print('\n')

wb.save('C:\\Users\\craig\\Documents\\CS3398-Vulcans-S2020\\example.xlsm')
#wb.save(sys.argv[1])

## Features
* Interactivity through excel
* Mouse Pixel Position Display
* Clicking and Typing Automation
* Load and Save Procedures



## Status
Project is: Closed, no further development is planned.


