Attribute VB_Name = "modText"
Option Explicit

Public Sub AddText(ByVal rTxt As TextBox, ByVal Msg As String)
Dim s As String
  
    s = vbCrLf + Msg
    rTxt.SelStart = Len(rTxt.Text)
    rTxt.SelText = s
    rTxt.SelStart = Len(rTxt.Text) - 1
End Sub

Sub SetStatus(ByRef Status As String)
    AddText frmServer.txtText, Status
    DoEvents
End Sub

Sub UpdateCaption()
    frmServer.Caption = GAME_NAME & " Server <IP " & frmServer.Socket(0).LocalIP & " Port " & Str$(frmServer.Socket(0).LocalPort) & "> (" & OnlinePlayersCount & ")"
End Sub


