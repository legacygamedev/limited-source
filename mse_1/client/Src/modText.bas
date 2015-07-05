Attribute VB_Name = "modText"
Option Explicit

Public Sub SetFont(ByVal Font As String, ByVal Size As Byte)
    GameFont = CreateFont(Size, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, Font)
End Sub

Public Sub DrawText(ByVal hDC As Long, ByVal X, ByVal Y, ByVal Text As String, Color As Long)
    Call SelectObject(hDC, GameFont)
    Call SetBkMode(hDC, vbTransparent)
    Call SetTextColor(hDC, RGB(0, 0, 0))
    Call TextOut(hDC, X + 2, Y + 2, Text, Len(Text))
    Call TextOut(hDC, X + 1, Y + 1, Text, Len(Text))
    Call SetTextColor(hDC, Color)
    Call TextOut(hDC, X, Y, Text, Len(Text))
End Sub

Public Sub DrawSplicedText(ByVal hDC As Long, Text As String, Color As Long)
'****************************************************************
'* WHEN        WHO        WHAT
'* ----        ---        ----
'* 10/22/2006  BigRed     Created procedure.
'****************************************************************
    Dim strWords() As String
    Dim i As Long
    Dim Y As Long, strLine(0 To 3) As String
    Dim Length As Byte, Count As Byte

    strWords = Split(Text, " ")
    
    For i = 0 To UBound(strWords)
        Length = Length + Len(strWords(i)) + 1
        If Length <= 64 Then
            strLine(Count) = strLine(Count) + strWords(i) + " "
        Else
            Count = Count + 1
            If Count = 4 Then
                Count = Count - 1
                Exit For
            End If
            
            strLine(Count) = strWords(i) & " "
            Length = 0
        End If
    Next i
    
    Y = (MAX_MAPY + 1) * PIC_Y - 20
    
    For i = 0 To 3
        If LenB(strLine(i)) <> 0 Then
            DrawText hDC, 0, Y - (Count * 16), strLine(i), Color
            Y = Y + 16
        End If
    Next i
End Sub

Public Sub AddText(ByVal Msg As String, ByVal Color As Integer)
Dim s As String
  
    s = Msg & vbCrLf
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

