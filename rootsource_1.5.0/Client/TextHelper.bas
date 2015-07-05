Attribute VB_Name = "TextHelper"
Option Explicit

Public Type POINTAPI
    x As Long
    y As Long
End Type

Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hDC As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As POINTAPI) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Public Function TWidth(ByVal Text As String, ByRef oFont As StdFont) As Long
Dim hDC As Long
Dim iFnt       As IFont
Dim TextSize As POINTAPI

  hDC = CreateCompatibleDC(GetDC(0))
  Set iFnt = oFont
  SelectObject hDC, iFnt.hFont
  GetTextExtentPoint32 hDC, Text, Len(Text), TextSize
  TWidth = TextSize.x
  DeleteObject iFnt.hFont
  DeleteDC hDC
End Function
