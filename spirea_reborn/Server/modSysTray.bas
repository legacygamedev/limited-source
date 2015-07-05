Attribute VB_Name = "modSysTray"



'Declare a user-defined variable to pass to the Shell_NotifyIcon
'function.
Public Type NOTIFYICONDATA
   cbSize As Long
   hWnd As Long
   uId As Long
   uFlags As Long
   uCallBackMessage As Long
   hIcon As Long
   szTip As String * 64
End Type




'Dimension a variable as the user-defined data type.
Global nid As NOTIFYICONDATA

