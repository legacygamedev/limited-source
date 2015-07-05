Attribute VB_Name = "modText"
Option Explicit

Public Sub SetFont(ByVal Font As String, ByVal Size As Byte)
    GameFont = CreateFont(Size, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, Font)
End Sub

Public Sub DrawText(ByVal hdc As Long, ByVal x, ByVal y, ByVal Text As String, Color As Long)
    Call SelectObject(hdc, GameFont)
    Call SetBkMode(hdc, vbTransparent)
    Call SetTextColor(hdc, RGB(0, 0, 0))
    Call TextOut(hdc, x + 2, y + 2, Text, Len(Text))
    Call TextOut(hdc, x + 1, y + 1, Text, Len(Text))
    Call SetTextColor(hdc, Color)
    Call TextOut(hdc, x, y, Text, Len(Text))
End Sub

Public Sub AddText(ByVal Msg As String, ByVal Color As Integer)
Dim s As String
  
    s = vbNewLine & Msg
    frmMirage.txtChat.SelStart = Len(frmMirage.txtChat.Text)
    frmMirage.txtChat.SelColor = QBColor(Color)
    frmMirage.txtChat.SelText = s
    frmMirage.txtChat.SelStart = Len(frmMirage.txtChat.Text) - 1
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

