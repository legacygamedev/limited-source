VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSubClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mvarhWnd As Long
Private mlpfnWndProc As Long

Private Declare Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongA" ( _
    ByVal hwnd As Long, _
    ByVal nIndex As Long, _
    ByVal dwNewLong As Long _
) As Long

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" ( _
    ByVal lpPrevWndFunc As Long, _
    ByVal hwnd As Long, _
    ByVal Msg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long _
) As Long

Private Const GWLP_WNDPROC = (-4)
Private Const GWLP_USERDATA = (-21)

Public Event WMArrival(hwnd As Long, uMsg As Long, _
     wParam As Long, lParam As Long, lRetVal As Long)

'for debug purposes set the folowing to -1
'you also much add a reference to the DBGWPROC.DLL
#Const DEBUGWINDOWPROC = 0
#If DEBUGWINDOWPROC Then
    Private m_SCHook As WindowProcHook
#End If

Private Declare Function SetProp Lib "user32" Alias "SetPropA" ( _
    ByVal hwnd As Long, _
    ByVal lpString As String, _
    ByVal hData As Long _
) As Long

Friend Function CLSProc(ByVal hwnd As Long, ByVal uMsg As Long, _
    ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim lRetVal As Long
    lRetVal = -10& 'default value
    RaiseEvent WMArrival(hwnd, uMsg, wParam, lParam, lRetVal)
    'if user didn't intervene
    If lRetVal = -10& Then
        CLSProc = CallWindowProc(mlpfnWndProc, hwnd, uMsg, wParam, lParam)
    Else 'otherwise return user's value
        CLSProc = lRetVal
    End If
End Function

Public Function callWindProc(mhWnd, uMsg, wParam, ByVal lParam) As Long
    callWindProc = CallWindowProc(mlpfnWndProc, mhWnd, uMsg, wParam, ByVal lParam)
End Function

Public Function Enable(hwnd As Long) As Boolean
    
    If mlpfnWndProc = 0 Then
        Enable = False
        mvarhWnd = hwnd
        #If DEBUGWINDOWPROC Then
            On Error Resume Next
            Set m_SCHook = CreateWindowProcHook
            If Err Then
                MsgBox Err.Description
                Err.Clear
                Exit Function
            End If
            On Error GoTo 0
            With m_SCHook
                .SetMainProc AddressOf MainCLSProc
                mlpfnWndProc = SetWindowLongPtr(hwnd, GWLP_WNDPROC, .ProcAddress)
                If mlpfnWndProc = 0 Then Exit Function
                .SetDebugProc mlpfnWndProc
            End With
        #Else
            mlpfnWndProc = SetWindowLongPtr(hwnd, GWLP_WNDPROC, AddressOf MainCLSProc)
            If mlpfnWndProc = 0 Then Exit Function
        #End If
       'store handle to this class in a property of the subclassed window
        Call SetProp(hwnd, "objptr", ObjPtr(Me))
    End If
    Enable = True
End Function

Public Sub Unsubclass()
    If mlpfnWndProc Then
        SetWindowLongPtr mvarhWnd, GWLP_WNDPROC, mlpfnWndProc
        mlpfnWndProc = 0
    End If
    #If DEBUGWINDOWPROC Then
      Set m_SCHook = Nothing
    #End If
End Sub

Private Sub Class_Terminate()
    Unsubclass
End Sub




