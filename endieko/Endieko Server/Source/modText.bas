Attribute VB_Name = "modText"
Option Explicit

Public Const Quote = """"

Public Const MAX_LINES = 500

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
Public Const BroadcastColor = BrightBlue
Public Const TellColor = BrightGreen
Public Const EmoteColor = BrightCyan
Public Const AdminColor = Magenta
Public Const HelpColor = White
Public Const WhoColor = Grey
Public Const JoinLeftColor = Grey
Public Const NpcColor = White
Public Const AlertColor = BrightRed
Public Const NewMapColor = Grey

'Public Sub AddText(ByVal rTxt As RichTextBox, ByVal Msg As String, ByVal Color As Integer)
'Dim s As String
'
'    s = vbCrLf + Msg
'    rTxt.SelStart = Len(rTxt.Text)
'    rTxt.SelColor = QBColor(Color)
'    rTxt.SelText = s
'    rTxt.SelStart = Len(rTxt.Text) - 1
'End Sub

Public Sub Addtext(Text As String)
    'Sub to add text to the console on the next line
    frmServer.txtText.Text = frmServer.txtText.Text & Chr(13) & Chr(10)
    frmServer.txtText.Text = frmServer.txtText.Text & Text
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
        Txt.Text = vbNullString
        NumLines = 0
    End If
    
    Txt.SelStart = Len(Txt.Text)
End Sub


