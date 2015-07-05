Attribute VB_Name = "modText"
Option Explicit



Public Sub AddText(ByVal rTxt As RichTextBox, ByVal Msg As String, ByVal Color As Integer)
Dim s As String
  
    s = vbCrLf + Msg
    rTxt.SelStart = Len(rTxt.Text)
    rTxt.SelColor = QBColor(Color)
    rTxt.SelText = s
    rTxt.SelStart = Len(rTxt.Text) - 1
End Sub

Public Sub TextAdd(ByVal Txt As TextBox, Msg As String, NewLine As Boolean)
Static NumLines As Long

    If NewLine Then
        Txt.Text = Txt.Text & vbCrLf & Msg
    Else
        Txt.Text = Txt.Text & Msg
    End If
        
    NumLines = NumLines + 1
    If NumLines >= MAX_LINES Then
        Txt.Text = ""
        NumLines = 0
    End If
    
    Txt.SelStart = Len(Txt.Text)
End Sub


