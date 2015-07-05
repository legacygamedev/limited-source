Attribute VB_Name = "Module1"
Option Explicit

' Listscr5 sample from BlackBeltVB.com
' http://blackbeltvb.com
'
' Written by Matt Hart
' Copyright 1999 by Matt Hart
' 
' This software is FREEWARE. You may use it as you see fit for 
' your own projects but you may not re-sell the original or the 
' source code. Do not copy this sample to a collection, such as
' a CD-ROM archive. You may link directly to the original sample
' using "http://blackbeltvb.com/listscr5.htm"
' 
' No warranty express or implied, is given as to the use of this
' program. Use at your own risk.

Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const GWL_WNDPROC = (-4)
Public Const GWL_USERDATA = (-21)
Public Const WM_HSCROLL = &H114
Public Const WM_MOUSEMOVE = &H200
    
Public Function WndProc(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If Msg = WM_HSCROLL Then SendMessage hwnd, WM_MOUSEMOVE, 1, ByVal &HFFFF
    WndProc = CallWindowProc(GetWindowLong(hwnd, GWL_USERDATA), hwnd, Msg, wParam, lParam)
End Function
