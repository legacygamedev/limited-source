Attribute VB_Name = "modText"
Option Explicit

Public Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal H As Long, ByVal W As Long, ByVal E As Long, ByVal O As Long, ByVal W As Long, ByVal i As Long, ByVal u As Long, ByVal s As Long, ByVal C As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal f As String) As Long
Public Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Public Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long

Public Const Quote = """"

Public Const Black = 0
Public Const Blue = 1
Public Const Green = 2
Public Const Cyan = 3
Public Const Red = 4
Public Const Magenta = 5
Public Const Brown = 6
Public Const Grey = 7
Public Const DarkGrey = 8
Public Const BrightBlue = 9
Public Const BrightGreen = 10
Public Const BrightCyan = 11
Public Const BrightRed = 12
Public Const Pink = 13
Public Const Yellow = 14
Public Const White = 15

Public Const SayColor = Grey
Public Const GlobalColor = Green
Public Const BroadcastColor = White
Public Const TellColor = White
Public Const EmoteColor = White
Public Const AdminColor = BrightCyan
Public Const HelpColor = White
Public Const WhoColor = Grey
Public Const JoinLeftColor = Grey
Public Const NpcColor = White
Public Const AlertColor = White
Public Const NewMapColor = Grey

Public TexthDC As Long
Public GameFont As Long

Public Sub SetFont(ByVal Font As String, ByVal Size As Byte)
    GameFont = CreateFont(Size, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, Font)
End Sub

Public Sub DrawText(ByVal hDC As Long, ByVal X, ByVal Y, ByVal Text As String, Color As Long)
    Call SelectObject(hDC, GameFont)
    Call SetBkMode(hDC, vbTransparent)
    Call SetTextColor(hDC, QBColor(Black))
    Call TextOut(hDC, X - 1, Y - 1, Text, Len(Text))
    Call SetTextColor(hDC, Color)
    Call TextOut(hDC, X, Y, Text, Len(Text))
End Sub
Public Sub DrawPlayerNameText(ByVal hDC As Long, ByVal X, ByVal Y, ByVal Text As String, Color As Long)
    Call SelectObject(hDC, GameFont)
    Call SetBkMode(hDC, vbTransparent)
    Call SetTextColor(hDC, QBColor(Black))
    Call TextOut(hDC, X - 1, Y - 1, Text, Len(Text))
    Call SetTextColor(hDC, Color)
    Call TextOut(hDC, X, Y, Text, Len(Text))
End Sub

Public Sub AddText(ByVal Msg As String, ByVal Color As Long)
Dim s As String
  
    s = vbNewLine & Msg
    frmMirage.txtChat.SelStart = Len(frmMirage.txtChat.Text)
    If Color > 15 Then
        frmMirage.txtChat.SelColor = Color
    Else
        frmMirage.txtChat.SelColor = QBColor(Color)
    End If
    frmMirage.txtChat.SelText = s
    frmMirage.txtChat.SelStart = Len(frmMirage.txtChat.Text) - 1
End Sub

Public Sub AddHelp(ByVal RT As RichTextBox, ByVal Msg As String, ByVal Color As Long)
Dim s As String
  
    s = vbNewLine & Msg
    RT.SelStart = Len(RT.Text)
    If Color > 15 Then
        RT.SelColor = Color
    Else
        RT.SelColor = QBColor(Color)
    End If
    RT.SelText = s
    RT.SelStart = Len(RT.Text) - 1
End Sub

Public Sub TextAdd(ByVal Txt As TextBox, Msg As String, NewLine As Boolean)
    If NewLine Then
        Txt.Text = Txt.Text + Msg + vbCrLf
    Else
        Txt.Text = Txt.Text + Msg
    End If
    
    Txt.SelStart = Len(Txt.Text) - 1
End Sub

Function Parse(ByVal Num As Long, ByVal Data As String)
Dim i As Long
Dim n As Long
Dim sChar As Long
Dim eChar As Long

    n = 0
    sChar = 1
    
    For i = 1 To Len(Data)
        If Mid(Data, i, 1) = SEP_CHAR Then
            If n = Num Then
                eChar = i
                Parse = Mid(Data, sChar, eChar - sChar)
                Exit For
            End If
            
            sChar = i + 1
            n = n + 1
        End If
    Next i
End Function

