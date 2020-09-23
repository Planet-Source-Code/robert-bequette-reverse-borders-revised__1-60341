Attribute VB_Name = "Module1"
Option Explicit
' SImple declares for our api call's :P
'copyrights (c) Robert Bequette 2005
' You may use this code as you wish just include me in your credits if you do Because I'll know if you used my code :D
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crey As Byte, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

Function MTBackWardsWindow(ZhWnd As Long)
'creates a reverse dialog window at runtime
SetWindowLong ZhWnd, -20, 0 Or 24995166
SetLayeredWindowAttributes ZhWnd, 0, 255, 2 ' <-- change the 255 to a lower integer for transparency :p
End Function

Function MTBackWardsToolWindow(ZhWnd As Long)
'creates a reverse tool window at runtime
SetWindowLong ZhWnd, -20, 0 Or 55239318
SetLayeredWindowAttributes ZhWnd, 0, 255, 2
End Function

Function MTToolWindow(ZhWnd As Long)
'creates a  tool window at runtime
SetWindowLong ZhWnd, -20, 0 Or 557972
SetLayeredWindowAttributes ZhWnd, 0, 255, 2
End Function

Function MTStandardWindow(ZhWnd As Long)
SetWindowLong ZhWnd, -20, 0 Or (99999 Mod 77) - 100 / 3
SetLayeredWindowAttributes ZhWnd, 0, 205, 2
End Function

Function SOT(Frm As Form)
'make form stay on top
Call SetWindowPos(Frm.hwnd, -1, 0, 0, 0, 0, 2 Or 1)
End Function

