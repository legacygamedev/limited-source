Attribute VB_Name = "modSysTray"
Option Explicit

' Declare a user-defined variable to pass to the Shell_NotifyIcon function.
Public Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 128
    dwState As Long
    dwStateMask As Long
    szInfo As String * 256
    uTimeout As Long
    szInfoTitle As String * 64
    dwInfoFlags As Long
End Type

' Declare the constants for the API function. These constants can be
' found in the header file Shellapi.h.

' The following constants are the messages sent to the
' Shell_NotifyIcon function to add, modify, or delete an icon from the System Tray
Public Const NIS_HIDDEN = &H1
Public Const NIS_SHAREDICON = &H2
Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const WM_MOUSEMOVE = &H200
Public Const NIF_STATE = &H8
Public Const NIF_INFO = &H10
' The following constant is the message sent when a mouse event occurs
' within the rectangular boundaries of the icon in the System Tray
' area.

Public Const vbNone = 0
Public Const vbInformation = 1
Public Const vbExclamation = 2
Public Const vbCritical = 3


' The following constants are the flags that indicate the valid
' members of the NOTIFYICONDATA data type.
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4

' The following constants are used to determine the mouse input on the
' the icon in the taskbar status area.

' Left-click constants.
Public Const WM_LBUTTONDBLCLK = &H203   'Double-click
Public Const WM_LBUTTONDOWN = &H201     'Button down
Public Const WM_LBUTTONUP = &H202       'Button up

' Right-click constants.
Public Const WM_RBUTTONDBLCLK = &H206   'Double-click
Public Const WM_RBUTTONDOWN = &H204     'Button down
Public Const WM_RBUTTONUP = &H205       'Button up

'Declare the API function call.
Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean


' Dimension a variable as the user-defined data type.
Global nid As NOTIFYICONDATA

