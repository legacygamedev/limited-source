Attribute VB_Name = "modChatLink"
Option Explicit

Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
    ByVal hwnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    lParam As Any _
) As Long

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
    ByVal hwnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long _
) As Long

Public Const CFE_LINK = &H20
Public Const CFM_LINK = &H20
Public Const CFM_REVAUTHOR = &H8000
Public Const CFM_LCID = &H2000000

Public Const SCF_SELECTION = &H1
Public Const SCF_WORD = &H2

Public Const WM_USER = &H400
Public Const EM_GETCHARFORMAT = (WM_USER + 58)
Public Const EM_SETCHARFORMAT = (WM_USER + 68)
Public Const EM_SETEVENTMASK = (WM_USER + 69)
Public Const EM_GETEVENTMASK = (WM_USER + 59)
Public Const ENM_UPDATE = &H2
Public Const EN_UPDATE = &H400
Public Const EN_CHANGE = &H300
Public Const ENM_CHANGE = &H1

Public Const WM_NOTIFY = &H4E
Public Const WM_COMMAND = &H111

Public Const EN_LINK = &H70B
Public Const ENM_LINK = &H4000000
Public Const EN_SELCHANGE = &H702

Public Const SW_SHOWNORMAL As Long = 1

'mouse Messages passed on by EN_LINK notifications
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_MOUSEMOVE = &H200
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_SETCURSOR = &H20

Public Type CharRange
  cpMin As Long     ' First character of range (0 for start of doc)
  cpMax As Long     ' Last character of range (-1 for end of doc)
End Type

Public Type nmhdr
    hwndFrom As Long
    idfrom As Long
    code As Long
End Type

Public Type ENLINK
    nmhdr As nmhdr
    Msg As Long
    wParam As Long
    lParam As Long
    chrg As CharRange
End Type

Public Const LF_FACESIZE = 32
Public Type CHARFORMAT2
    cbSize As Long
    dwMask As Long
    dwEffects As Long
    yHeight As Long
    YOffset As Long
    crTextColor As Long
    bCharSet As Byte
    bPitchAndFamily As Byte
    szFaceName(LF_FACESIZE - 1) As Byte
    wWeight As Integer
    sSpacing As Integer
    crBackColor As Long
    lcid As Long
    dwReserved As Long
    sStyle As Integer
    wKerning As Integer
    bUnderlineType As Byte
    bAnimation As Byte
    bRevAuthor As Byte
    bReserved1 As Byte
End Type

Private Declare Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long

Private Const GWLP_WNDPROC = (-4)
Private Const GWLP_USERDATA = (-21)

Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hwnd As Long, ByVal lpString As String) As Long

Public Function MainCLSProc(ByVal hwnd As Long, ByVal Msg As Long, _
    ByVal wParam As Long, ByVal lParam As Long) As Long
    
Dim clsRefToCLS As clsSubClass
Dim pUserData As Long
    
    pUserData = GetProp(hwnd, "objptr")
    
    If pUserData Then
        Set clsRefToCLS = ObjFromPtr(pUserData)
        MainCLSProc = clsRefToCLS.CLSProc(hwnd, Msg, wParam, lParam)
        Set clsRefToCLS = Nothing
    End If
End Function

Private Function ObjFromPtr(ByVal lpObject As Long) As Object
Dim objTemp As Object
    CopyMemory objTemp, lpObject, 4&
    Set ObjFromPtr = objTemp
    CopyMemory objTemp, 0&, 4&
End Function


