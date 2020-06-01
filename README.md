# GUI Automation
> Creator: Craig Sandlin

> Project Description: Automation tool designed to relieve the user from repetitive tasks. Works across all applications; from websites to spreadsheets, file explorer, and everything inbetween. A friendly excel UI allows those without programming knowledge to automate any task that includes clicking, typing, select all, backspace, enter, wait, and more! 

## Table of contents
* [General info](#general-info)
* [Screenshots](#screenshots)
* [Technologies](#technologies)
* [Setup](#setup)
* [Features](#features)
* [Status](#status)
* [Inspiration](#inspiration)
* [Contact](#contact)

## General info
> The motivation: The mother of all invention- necessity. Electrical measurements were being performed on lighting fixtures to characterize the product before approbation listing with UL and/or Intertek. The sample size for this testing was sometimes upward of 70 discrete measurements. The changes that needed to be made to the software were repetitive and regular. If only I could program the inputs ahead of time! But alas, the software was propriatery and no API was available. It is tasks like these where this software really shines. With some upfront setup the entire process could be automated. Even allowing measurements to be taken over night which improved throughput!

## Screenshots
file:///.file/id=6571367.125785890
![Example screenshot](./img/screenshot.png)

## Technologies
* Python 3 - version 3.5
* Excel - version Latest

## Setup
Describe how to install / setup your local environement / add link to demo version.
need excel, pycharm for code editing./ local environment needs pyautogui and may need pip to install pyautogui on certain machines. demo version is accessible though the example.xlsm

## Code Examples
Show examples of usage:
`example of the interactivity with excel`
*Sub PythonMousePosition()

Dim objShell As Object
'''initialize the array as a string'''
Dim PythonExe, PythonScript, arg, formattedArg As String

Set objShell = VBA.CreateObject("Wscript.Shell")

PythonExe = """C:\Users\craig\AppData\Local\Programs\Python\Python38\python.exe"""

PythonScript = """C:\Users\craig\PycharmProjects\CS3398-Vulcans-S2020\MousePosition.py"""

objShell.Run PythonExe & PythonScript, 6


End Sub
'''initialize the funtion to allow interaction with excel as a gui'''
'''format to run explicitly thru the file paths on craig's laptop'''
Sub PythonGUIControl()

Dim objShell As Object
'''initialize this array as a string'''
Dim PythonExe, PythonScript, arg, formattedArg As String

Set objShell = VBA.CreateObject("Wscript.Shell")

PythonExe = """C:\Users\craig\AppData\Local\Programs\Python\Python38\python.exe"""

PythonScript = """C:\Users\craig\PycharmProjects\CS3398-Vulcans-S2020\ExcelUI.py"""

'arg = ThisWorkbook.Sheets("Sheet2").Range("A3").Value
'arg = "Hello World!"
'formattedArg = """" & """" & """" & arg & """" & """" & """"

'formattedArg = """C:\\Users\\CraigSandlin\\Documents\\CS3398-Vulcans-S2020\example.xlsm"""

objShell.Run PythonExe & PythonScript, 6


End Sub

Sub test()
MsgBox "Hello World"
*End Sub

## Features
* Interactivity through excel
* Mouse Pixel Position Display
* Clicking and Typing Automation



## Status
Project is: Closed, no further development is planned.


