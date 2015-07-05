Attribute VB_Name = "mCustomCursor"
Option Explicit

' ===========================================================================
' FileName:    mCustomCursor
' Author:      Steve McMahon
' Date:        1 July 1999
'
' Uses a CallWindowProcRet windows hook to allow a cursor
' to be set for all windows within a VB project.  Like
' Screen.MousePointer, but allows any cursor to be used,
' particularly an animated one.
'
' Requires cCustomCursor.cls
'
' ---------------------------------------------------------------------------
' Visit vbAccelerator at
'     http://vbaccelerator.com
' ===========================================================================


Private Type CWPRETSTRUCT
   lResult As Long
   lParam As Long
   wParam As Long
   Message As Long
   hWnd As Long
End Type
Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpFn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)
Private Const WH_CALLWNDPROCRET = 12
Private Const WM_SETCURSOR = &H20
Private Const HC_ACTION = 0

Private m_hHook As Long
Private m_lPtr As Long                     ' Pointer to consumer object

Private Property Get ObjectFromPtr(ByVal lPtr As Long) As Object
Dim oTemp As Object
   ' Turn the pointer into an illegal, uncounted interface
   CopyMemory oTemp, lPtr, 4
   ' Do NOT hit the End button here! You will crash!
   ' Assign to legal reference
   Set ObjectFromPtr = oTemp
   ' Still do NOT hit the End button here! You will still crash!
   ' Destroy the illegal reference
   CopyMemory oTemp, 0&, 4
   ' OK, hit the End button if you must
End Property
Public Function WaitMousePointer( _
      ByRef cCC As cCustomCursor, _
      ByVal bState As Boolean _
   ) As Boolean
Dim lpFn As Long
      
   ' If Hook not already installed:
   If bState And m_hHook = 0 Then
      lpFn = HookAddress(AddressOf CallWndProc)
      m_hHook = SetWindowsHookEx(WH_CALLWNDPROCRET, lpFn, 0&, GetCurrentThreadId())
      If m_hHook <> 0 Then
         m_lPtr = ObjPtr(cCC)
      End If
   ElseIf Not bState And m_hHook <> 0 Then
      UnhookWindowsHookEx m_hHook
      m_hHook = 0
      m_lPtr = 0
   End If
   
End Function

Private Function HookAddress(ByVal lPtr As Long) As Long
   ' Work around for VB's poor AddressOf implementation:
   HookAddress = lPtr
End Function

Private Function CallWndProc(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim tCWP As CWPRETSTRUCT
Dim cCC As cCustomCursor

   ' This hook allows you to intercept every message sent to every window
   ' in your application
   If nCode = HC_ACTION Then
      If m_lPtr > 0 Then
         CopyMemory tCWP, ByVal lParam, Len(tCWP)
         If tCWP.Message = WM_SETCURSOR Then
            Set cCC = ObjectFromPtr(m_lPtr)
            If cCC.WindowProc(tCWP.hWnd, tCWP.Message, tCWP.wParam, tCWP.lParam) Then
               '
            End If
         End If
      End If
   End If
   CallWndProc = CallNextHookEx(m_hHook, nCode, wParam, lParam)
End Function


