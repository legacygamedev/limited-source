Attribute VB_Name = "modText"
Option Explicit

' ------------------------------------------
' --              Asphodel 6              --
' ------------------------------------------

Public Sub TextAdd(ByVal Txt As TextBox, Msg As String)
    
    NumLines = NumLines + 1
    
    If NumLines >= MAX_LINES Then
        Txt.Text = vbNullString
        NumLines = 0
    End If
    
    Txt.Text = Txt.Text & vbNewLine & Msg
    
    Txt.SelStart = Len(Txt.Text)
    
End Sub



