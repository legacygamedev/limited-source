Attribute VB_Name = "modText"
Option Explicit

Public Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal H As Long, ByVal W As Long, ByVal E As Long, ByVal O As Long, ByVal W As Long, ByVal I As Long, ByVal u As Long, ByVal s As Long, ByVal C As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal f As String) As Long
Public Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Attribute SetBkMode.VB_UserMemId = 1879048224
Public Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Attribute SetTextColor.VB_UserMemId = 1879048256
Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDC As Long, ByVal X As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Attribute TextOut.VB_UserMemId = 1879048292
Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Attribute SelectObject.VB_UserMemId = 1879048324

Public Sub SetFont(ByVal Font As String, ByVal Size As Byte)
    GameFont = CreateFont(Size, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, Font)
End Sub

Public Sub DrawText(ByVal hDC As Long, ByVal X, ByVal y, ByVal Text As String, color As Long)
    Call SelectObject(hDC, GameFont)
    Call SetBkMode(hDC, vbTransparent)
    Call SetTextColor(hDC, RGB(0, 0, 0))
    Call TextOut(hDC, X + 1, y + 0, Text, Len(Text))
    Call TextOut(hDC, X + 0, y + 1, Text, Len(Text))
    Call TextOut(hDC, X - 1, y - 0, Text, Len(Text))
    Call TextOut(hDC, X - 0, y - 1, Text, Len(Text))
    Call SetTextColor(hDC, color)
    Call TextOut(hDC, X, y, Text, Len(Text))
End Sub

Public Sub AddText(ByVal Msg As String, ByVal color As Integer)
    frmStable.txtChat.SelStart = Len(frmStable.txtChat.Text)
    frmStable.txtChat.SelColor = QBColor(color)
    frmStable.txtChat.SelText = vbNewLine & Msg
    frmStable.txtChat.SelStart = Len(frmStable.txtChat.Text) - 1

    If frmStable.chkAutoScroll.Value = Unchecked Then
        frmStable.txtChat.SelStart = frmStable.txtChat.SelStart
    End If
End Sub

Public Sub TextAdd(ByVal Txt As TextBox, Msg As String, NewLine As Boolean)
    If NewLine Then
        Txt.Text = Txt.Text & (Msg & vbNewLine)
    Else
        Txt.Text = Txt.Text & Msg
    End If

    Txt.SelStart = Len(Txt.Text)
End Sub

