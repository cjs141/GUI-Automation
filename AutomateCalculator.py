import pyautogui, time
import xlwings as xw
import os
import cv2
time.sleep(1)
#Program used to automate using the calculator
#take screenshot of the buttons we want it to find on the screen
# > tell it to find where buttons reside >
# > click area that corresponds to the button

#locateCenterOnScreen looks for a match to the screenshots (.png)
#must save the screenshots in the same directory as the program
wb = xw.Book('C:\\Users\\craig\\Documents\\CS3398-Vulcans-S2020\\example.xlsm')
sheet = wb.sheets['Sheet1']

screenshot_dir = "C:\\Users\\craig\\PycharmProjects\\CS3398-Vulcans-S2020"
os.chdir(screenshot_dir)
FirstNumber = sheet.range('Q2').value
print(FirstNumber)
SecondNumber = sheet.range('Q3').value
Operand = sheet.range('Q4').value
time.sleep(1)
pyautogui.click(pyautogui.locateCenterOnScreen(FirstNumber, confidence=0.9))
time.sleep(0.5)
pyautogui.click(pyautogui.locateCenterOnScreen(Operand, confidence=0.9))
time.sleep(0.5)
pyautogui.click(pyautogui.locateCenterOnScreen(SecondNumber, confidence=0.9))
time.sleep(0.5)
pyautogui.click(pyautogui.locateCenterOnScreen('EqualsSymbol.png', confidence=0.9))
time.sleep(3)
