Attribute VB_Name = "modSysTray"
Option Explicit

' ********************************************
' **               rootSource               **
' ********************************************

'Declare the API function call.
Private Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
'Private Declare Sub keybd_event Lib "user32.dll" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

'Private Const KEYEVENTF_KEYUP As Long = &H2
'Private Const VK_LWIN As Long = &H5B

Private Type NOTIFYICONDATA
   cbSize As Long
   hWnd As Long
   uId As Long
   uFlags As Long
   uCallBackMessage As Long
   hIcon As Long
   szTip As String * 64
End Type

'Shell_NotifyIcon function to add, modify, or delete an icon from the System Tray
Private Const NIM_ADD As Long = &H0
Private Const NIM_MODIFY As Long = &H1
Private Const NIM_DELETE As Long = &H2

Private Const WM_MOUSEMOVE As Long = &H200

Private Const NIF_MESSAGE As Long = &H1
Private Const NIF_ICON As Long = &H2
Private Const NIF_TIP As Long = &H4

'Left-click constants.
Public Const WM_LBUTTONDBLCLK As Long = &H203    'Double-click
'Private Const WM_LBUTTONDOWN As Long = &H201      'Button down
'Private Const WM_LBUTTONUP As Long = &H202        'Button up

'Right-click constants.
'Private Const WM_RBUTTONDBLCLK As Long = &H206    'Double-click
'Private Const WM_RBUTTONDOWN As Long = &H204      'Button down
'Private Const WM_RBUTTONUP As Long = &H205        'Button up

Dim nid As NOTIFYICONDATA

Public Sub InitSystemTray()
    nid.cbSize = Len(nid)
    nid.hWnd = frmServer.hWnd
    nid.uId = vbNull
    nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    nid.uCallBackMessage = WM_MOUSEMOVE
    nid.hIcon = frmServer.Icon
    nid.szTip = GAME_NAME & vbNullChar
    Call Shell_NotifyIcon(NIM_ADD, nid)
End Sub

Public Sub DestroySystemTray()
    nid.cbSize = Len(nid)
    nid.hWnd = frmServer.hWnd
    nid.uId = vbNull
    nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    nid.uCallBackMessage = WM_MOUSEMOVE
    nid.hIcon = frmServer.Icon
    nid.szTip = "Server" & vbNullChar
    Call Shell_NotifyIcon(NIM_DELETE, nid) ' Add to the sys tray
End Sub
