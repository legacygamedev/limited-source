Attribute VB_Name = "modSysTray"
'Tray Icon constants
Public Const AddIcon = &H0 'Add to Tray
Public Const ModifyIcon = &H1 'Modify Details
Public Const DeleteIcon = &H2 'Remove From Tray
Public Const MessageFlag = &H1 'Message
Public Const IconFlag = &H2 'Icon
Public Const TipFlag = &H4 'TooTipText

Type NOTIFYICONDATA
    Size As Long
    Handle As Long
    ID As Long
    Flags As Long
    CallBackMessage As Long
    Icon As Long
    Tip As String * 64
    hWnd As Long
End Type

Public TrayIcon As NOTIFYICONDATA 'Tray icon

'Tray icon
Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal Message As Long, Data As NOTIFYICONDATA) As Boolean

Public Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Public Const KEYEVENTF_KEYUP = &H2
Public Const VK_LWIN = &H5B

'Declare a user-defined variable to pass to the Shell_NotifyIcon
'function.

'Declare the constants for the API function. These constants can be
'found in the header file Shellapi.h.

'The following constants are the messages sent to the
'Shell_NotifyIcon function to add, modify, or delete an icon from the System Tray
Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2

'The following constant is the message sent when a mouse event occurs
'within the rectangular boundaries of the icon in the System Tray
'area.
Public Const WM_MOUSEMOVE = &H200

'The following constants are the flags that indicate the valid
'members of the NOTIFYICONDATA data type.
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4

'The following constants are used to determine the mouse input on the
'the icon in the taskbar status area.

'Left-click constants.
Public Const WM_LBUTTONDBLCLK = &H203   'Double-click
Public Const WM_LBUTTONDOWN = &H201     'Button down
Public Const WM_LBUTTONUP = &H202       'Button up

'Right-click constants.
Public Const WM_RBUTTONDBLCLK = &H206   'Double-click
Public Const WM_RBUTTONDOWN = &H204     'Button down
Public Const WM_RBUTTONUP = &H205       'Button up

'Dimension a variable as the user-defined data type.
Global nid As NOTIFYICONDATA



