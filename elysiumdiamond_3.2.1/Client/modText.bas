Attribute VB_Name = "modText"

' Copyright (c) 2007-2008 Elysium Source. All rights reserved.
' This code is licensed under the Elysium General License.

Option Explicit

Public Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal H As Long, ByVal W As Long, ByVal E As Long, ByVal O As Long, ByVal W As Long, ByVal I As Long, ByVal u As Long, ByVal s As Long, ByVal C As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal f As String) As Long
Public Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Public Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long

Public Const Quote As String = """"

Public Const Black As Byte = 0
Public Const Blue As Byte = 1
Public Const Green As Byte = 2
Public Const Cyan As Byte = 3
Public Const Red As Byte = 4
Public Const Magenta As Byte = 5
Public Const Brown As Byte = 6
Public Const Grey As Byte = 7
Public Const DarkGrey As Byte = 8
Public Const BrightBlue As Byte = 9
Public Const BrightGreen As Byte = 10
Public Const BrightCyan As Byte = 11
Public Const BrightRed As Byte = 12
Public Const Pink As Byte = 13
Public Const Yellow As Byte = 14
Public Const White As Byte = 15

Public Const SayColor As Byte = Grey
Public Const GlobalColor As Byte = Green
Public Const BroadcastColor As Byte = White
Public Const TellColor As Byte = White
Public Const EmoteColor As Byte = White
Public Const AdminColor As Byte = BrightCyan
Public Const HelpColor As Byte = White
Public Const WhoColor As Byte = Grey
Public Const JoinLeftColor As Byte = Grey
Public Const NpcColor As Byte = White
Public Const AlertColor As Byte = White
Public Const NewMapColor As Byte = Grey

Public TexthDC As Long
Public GameFont As Long

Public Sub SetFont(ByVal Font As String, ByVal Size As Byte, ByVal Bold As Long, ByVal Italic As Long, ByVal UnderLine As Long, ByVal Strike As Long)
    GameFont = CreateFont(Size, 0, 0, 0, Bold, Italic, UnderLine, Strike, 0, 0, 0, 0, 0, Font)
End Sub

Public Sub DrawText(ByVal hDC As Long, ByVal x, ByVal y, ByVal Text As String, Color As Long)
    Call SelectObject(hDC, GameFont)
    Call SetBkMode(hDC, vbTransparent)
    Call SetTextColor(hDC, RGB(0, 0, 0))
    Call TextOut(hDC, x + 1, y + 1, Text, Len(Text))
    Call SetTextColor(hDC, Color)
    Call TextOut(hDC, x, y, Text, Len(Text))
End Sub

Public Sub DrawPlayerNameText(ByVal hDC As Long, ByVal x, ByVal y, ByVal Text As String, Color As Long)
    Call SelectObject(hDC, GameFont)
    Call SetBkMode(hDC, vbTransparent)
    Call SetTextColor(hDC, RGB(0, 0, 0))
    Call TextOut(hDC, x + 1, y + 1, Text, Len(Text))
    Call SetTextColor(hDC, Color)
    Call TextOut(hDC, x, y, Text, Len(Text))
End Sub

Public Sub AddText(ByVal Msg As String, ByVal Color As Integer)
Dim s As String
Dim C As Long
  
    s = vbNewLine & Msg
    C = frmMirage.txtChat.SelStart
    frmMirage.txtChat.SelStart = Len(frmMirage.txtChat.Text)
    frmMirage.txtChat.SelColor = QBColor(Color)
    frmMirage.txtChat.SelText = s
    frmMirage.txtChat.SelStart = Len(frmMirage.txtChat.Text) - 1
    If frmMirage.chkAutoScroll.Value = Unchecked Then frmMirage.txtChat.SelStart = C
End Sub

Public Sub TextAdd(ByVal Txt As TextBox, Msg As String, NewLine As Boolean)
    If NewLine Then
        Txt.Text = Txt.Text + Msg + vbCrLf
    Else
        Txt.Text = Txt.Text + Msg
    End If
    
    Txt.SelStart = Len(Txt.Text) - 1
End Sub

Function Parse$(ByVal Num As Long, ByVal Data As String)
Dim I As Long
Dim n As Long
Dim sChar As Long
Dim eChar As Long

    n = 0
    sChar = 1
    
    For I = 1 To Len(Data)
        If Mid(Data, I, 1) = SEP_CHAR Then
            If n = Num Then
                eChar = I
                Parse = Mid(Data, sChar, eChar - sChar)
                Exit For
            End If
            
            sChar = I + 1
            n = n + 1
        End If
    Next I
End Function

