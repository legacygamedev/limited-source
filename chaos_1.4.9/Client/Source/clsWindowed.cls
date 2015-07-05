VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWindowed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long


' Used to support captionless drag
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2

' Undocumented message constant.
Private Const WM_GETSYSMENU = &H313

' Used to get window style bits.
Private Const GWL_STYLE = (-16)
Private Const GWL_EXSTYLE = (-20)

' Style bits.
Private Const WS_MAXIMIZEBOX = &H10000
Private Const WS_MINIMIZEBOX = &H20000
Private Const WS_THICKFRAME = &H40000
Private Const WS_SYSMENU = &H80000
Private Const WS_CAPTION = &HC00000

' Extended Style bits.
Private Const WS_EX_TOPMOST = &H8
Private Const WS_EX_TOOLWINDOW = &H80
Private Const WS_EX_CONTEXTHELP = &H400

' Force total pRedraw that shows new styles.
Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOZORDER = &H4
Private Const SWP_NOSIZE = &H1

' Used to toggle into topmost layer.
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private WithEvents mClient As Form
Attribute mClient.VB_VarHelpID = -1
Private mhWnd As Long


Public Property Get Client() As Form
   ' Return reference to client form.
   Set Client = mClient
End Property
Public Property Set Client(ByVal obj As Form)
   ' Store reference to client form.
   Set mClient = obj
   ' Cache hWnd as it'll be accessed frequently.
   If mClient Is Nothing Then
      mhWnd = 0
   Else
      mhWnd = mClient.hWnd
   End If
End Property


Public Function fFlipBit(ByVal Bit As Long, ByVal Value As Boolean) As Boolean
Dim lStyle As Long
   
   ' Retrieve current style bits.
   lStyle = GetWindowLong(mhWnd, GWL_STYLE)
   
   ' Set requested bit On or Off and Redraw.
   If Value Then
      lStyle = lStyle Or Bit
   Else
      lStyle = lStyle And Not Bit
   End If
   Call SetWindowLong(mhWnd, GWL_STYLE, lStyle)
   Call pRedraw
   
   ' Return success code.
   fFlipBit = (lStyle = GetWindowLong(mhWnd, GWL_STYLE))
End Function

Public Property Let Titlebar(ByVal Value As Boolean)
   ' Set WS_CAPTION On or Off as requested.
   Call fFlipBit(WS_CAPTION, Value)
End Property
Public Property Get Titlebar() As Boolean
   ' Return value of WS_CAPTION bit.
   Titlebar = CBool(fStyle And WS_CAPTION)
End Property

Public Sub pRedraw()
   ' Redraw window with new style.
   Const swpFlags As Long = _
      SWP_FRAMECHANGED Or SWP_NOMOVE Or _
      SWP_NOZORDER Or SWP_NOSIZE
   Call SetWindowPos(mhWnd, 0, 0, 0, 0, 0, swpFlags)
End Sub

Private Function fStyle(Optional ByVal NewBits As Long = 0) As Long
   '
   ' Set new style bits.
   '
   If NewBits Then
      Call SetWindowLong(mhWnd, GWL_STYLE, NewBits)
   End If
   ' Retrieve current style bits.
   fStyle = GetWindowLong(mhWnd, GWL_STYLE)
End Function

