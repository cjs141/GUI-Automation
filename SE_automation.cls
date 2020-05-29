VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'''
initialize the function that registers the position of the mouse
format to run explicitly thru the file paths on craig's laptop'''
Sub PythonMousePosition()

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
End Sub
