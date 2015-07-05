Attribute VB_Name = "modResize"
Option Explicit

Private Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Sub CopyMemoryToMinMaxInfo Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As MINMAXINFO, ByVal hpvSource As Long, ByVal cbCopy As Long)
Private Declare Sub CopyMemoryFromMinMaxInfo Lib "kernel32" Alias "RtlMoveMemory" (ByVal hpvDest As Long, hpvSource As MINMAXINFO, ByVal cbCopy As Long)

Private Const GWL_WNDPROC = -4
Private Const WM_GETMINMAXINFO = &H24

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type MINMAXINFO
    ptReserved As POINTAPI
    ptMaxSize As POINTAPI
    ptMaxPosition As POINTAPI
    ptMinTrackSize As POINTAPI
    ptMaxTrackSize As POINTAPI
End Type

Global lpPrevWndProc As Long
Global gHW As Long

Public Sub Hook()
    lpPrevWndProc = SetWindowLong(gHW, GWL_WNDPROC, AddressOf WindowProc)
End Sub

Public Function WindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim MinMax As MINMAXINFO
    
    If uMsg = WM_GETMINMAXINFO Then
        CopyMemoryToMinMaxInfo MinMax, lParam, Len(MinMax)
        
        'Minimum window size
        MinMax.ptMinTrackSize.x = 649
        MinMax.ptMinTrackSize.y = 536

        'Maximum window size
        'MinMax.ptMaxTrackSize.x = 500
        'MinMax.ptMaxTrackSize.y = 500
        
        CopyMemoryFromMinMaxInfo lParam, MinMax, Len(MinMax)
        WindowProc = DefWindowProc(hw, uMsg, wParam, lParam)
    Else
        WindowProc = CallWindowProc(lpPrevWndProc, hw, uMsg, wParam, lParam)
    End If
End Function

Public Sub Unhook()
    Dim tmp As Long
    tmp = SetWindowLong(gHW, GWL_WNDPROC, lpPrevWndProc)
End Sub

