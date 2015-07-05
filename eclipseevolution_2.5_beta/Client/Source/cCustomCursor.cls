VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cCustomCursor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ===========================================================================
' FileName:    cCustomCursor
' Author:      Steve McMahon
' Date:        1 July 1999
'
' Uses subclassing
'
' Requires mCustomCursor.bas to implement the WaitCursor method.
'
' ---------------------------------------------------------------------------
' Visit vbAccelerator at
'     http://vbaccelerator.com
' ===========================================================================

Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Const WM_SETCURSOR = &H20
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Implements ISubclass

Private m_hWnd As Long
Private m_hCur As Long
Private m_hCurOld As Long
Private m_bWaitPointer As Boolean

Public Property Let WaitCursor(ByVal bState As Boolean)
' Set the current cursor as Screen.MousePointer
   If m_bWaitPointer <> bState Then
      m_bWaitPointer = bState
      WaitMousePointer Me, bState
      If m_hCur <> 0 Then
         SetCursor m_hCur
      End If
   End If
End Property
Public Property Get WaitCursor() As Boolean
' Get whether the current cursor is acting as Screen.MousePointer
   WaitCursor = m_bWaitPointer
End Property

Public Property Let hCursor(ByVal hCur As Long)
' Sets the current cursor handle to be used by the class.
' The cursor is maintained outside.
   m_hCur = hCur
   If m_hWnd <> 0 Then
      SendMessage m_hWnd, WM_SETCURSOR, 0, ByVal 0&
   End If
End Property
Public Property Get hCursor() As Long
' Gets the current cursor handle to be maintained by the class.
   hCursor = m_hCur
End Property

Public Sub Attach(ByVal hWnd As Long)
' Attaches the cCustomCursor class to a specific window.
   Destroy
   m_hWnd = hWnd
   AttachMessage Me, m_hWnd, WM_SETCURSOR
End Sub
Public Sub Destroy()
' Stops subclassing
   m_hCur = 0
   If m_hWnd <> 0 Then
      SendMessage m_hWnd, WM_SETCURSOR, 0, ByVal 0&
      DetachMessage Me, m_hWnd, WM_SETCURSOR
      m_hWnd = 0
   End If

End Sub
Friend Function WindowProc(ByVal hWnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Boolean
' Method to be called by WaitCursor hook
   If m_hCur <> 0 Then
      m_hCurOld = SetCursor(m_hCur)
      WindowProc = True
   End If
End Function

Private Sub Class_Terminate()
   Destroy
   WaitCursor = False
End Sub

Private Property Let ISubClass_MsgResponse(ByVal RHS As SSubTimer.EMsgResponse)
   '
End Property

Private Property Get ISubClass_MsgResponse() As SSubTimer.EMsgResponse
   ISubClass_MsgResponse = emrConsume
End Property

Private Function ISubClass_WindowProc(ByVal hWnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
   If Not WindowProc(hWnd, iMsg, wParam, lParam) Then
      ISubClass_WindowProc = CallOldWindowProc(hWnd, iMsg, wParam, lParam)
   End If
End Function
