import xlwings as xw
import pyautogui, sys, time

wb = xw.Book('C:\\Users\\craig\\Documents\\CS3398-Vulcans-S2020\\example.xlsm')
sheet = wb.sheets['Sheet1']
for x in range(3, 51):
    instrCell = "B" + str(x)

    if sheet.range(instrCell).value is None:
        break
    if sheet.range(instrCell).value == "Click":
        param1 = "C" + str(x)
        XCoordinate = sheet.range(param1).value
        param2 = "D" + str(x)
        YCoordinate = sheet.range(param2).value
        pyautogui.click(XCoordinate, YCoordinate, 1, 1)
    if sheet.range(instrCell).value == "DoubleClick":
        param1 = "C" + str(x)
        XCoordinate = sheet.range(param1).value
        param2 = "D" + str(x)
        YCoordinate = sheet.range(param2).value
        pyautogui.click(XCoordinate, YCoordinate, 2, 0)
    if sheet.range(instrCell).value == "TripleClick":
        param1 = "C" + str(x)
        XCoordinate = sheet.range(param1).value
        param2 = "D" + str(x)
        YCoordinate = sheet.range(param2).value
        pyautogui.click(XCoordinate, YCoordinate, 3, 0)
        pyautogui.click()
    if sheet.range(instrCell).value == "BackSpace":
        pyautogui.press('backspace')
    if sheet.range(instrCell).value == "Enter":
        pyautogui.press('enter')
    if sheet.range(instrCell).value == "Type":
        param1 = "C" + str(x)
        text = sheet.range(param1).value
        pyautogui.typewrite(text, 0.1)
    if sheet.range(instrCell).value == "DeleteContents":
        for y in range(0, 100):
            pyautogui.press('right')
        for z in range(0, 100):
            pyautogui.press('backspace')
    if sheet.range(instrCell).value == "SelectAll":
        pyautogui.hotkey('ctrl', 'a')
    if sheet.range(instrCell).value == "Wait":
        param1 = "C" + str(x)
        wait = sheet.range(param1).value
        time.sleep(wait)
    if sheet.range(instrCell).value == "MinimizeApplications":
        pyautogui.hotkey('win', 'd')
    if sheet.range(instrCell).value == "WiggleMouse":
        pyautogui.scroll(10)  # scroll up 10 "clicks"
        pyautogui.scroll(-10) # scroll down 10 "clicks"
    if sheet.range(instrCell).value == "MaximizeApplications":
        pyautogui.hotkey('win', 'up')
    if sheet.range(instrCell).value == "copy":
        pyautogui.hotkey('win', 'c')
    if sheet.range(instrCell).value == "MoveMouse":
        param1 = "C" + str(x)
        XCoordinate = sheet.range(param1).value
        param2 = "D" + str(x)
        YCoordinate = sheet.range(param2).value
        pyautogui.moveTo(XCoordinate, YCoordinate, 2, pyautogui.easeInQuad)
    if sheet.range(instrCell).value == "DragMouse":
        param1 = "C" + str(x)
        XCoordinate = sheet.range(param1).value
        param2 = "D" + str(x)
        YCoordinate = sheet.range(param2).value
        pyautogui.dragTo(XCoordinate, YCoordinate, 2, pyautogui.easeInQuad)
    if sheet.range(instrCell).value == "VerticalScroll":
        param1 = "C" + str(x)
        ScrollAmount = sheet.range(param1).value
        pyautogui.scroll(ScrollAmount)
    if sheet.range(instrCell).value == "HorizontalScroll":
        param1 = "C" + str(x)
        ScrollAmount = sheet.range(param1).value
        pyautogui.hscroll(ScrollAmount)


    wb.save('C:\\Users\\craig\\Documents\\CS3398-Vulcans-S2020\\example.xlsm')
#wb.save(sys.argv[1])