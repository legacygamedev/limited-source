Attribute VB_Name = "modText"
Option Explicit

Public Const Quote = """"

Public Const MAX_LINES = 500

'Public Const Black = 0
'Public Const Blue = 1
'Public Const Green = 2
'Public Const Cyan = 3
'Public Const Red = 4
'Public Const Magenta = 5
'Public Const Brown = 6
'Public Const Grey = 7
'Public Const DarkGrey = 8
'Public Const BrightBlue = 9
'Public Const BrightGreen = 10
'Public Const BrightCyan = 11
'Public Const BrightRed = 12
'Public Const Pink = 13
'Public Const Yellow = 14
'Public Const White = 0
Public Const Black = 15
Public Const Blue = 9
Public Const Green = 2
Public Const Cyan = 2
Public Const Red = 4
Public Const Magenta = 5
Public Const Brown = 6
Public Const Grey = 7
Public Const DarkGrey = 7
Public Const BrightBlue = 9
Public Const BrightGreen = 10
Public Const BrightCyan = 10
Public Const BrightRed = 12
Public Const Pink = 13
Public Const Yellow = 5
Public Const White = 15

Public Const SayColor = White
Public Const GlobalColor = Green
Public Const BroadcastColor = Pink
Public Const TellColor = Grey
Public Const EmoteColor = Blue
Public Const AdminColor = Blue
Public Const HelpColor = BrightBlue
Public Const WhoColor = BrightRed
Public Const JoinLeftColor = White
Public Const NpcColor = Brown
Public Const AlertColor = Red
Public Const NewMapColor = Pink

Public RGB_SayColor
Public RGB_GlobalColor
Public RGB_BroadcastColor
Public RGB_TellColor
Public RGB_EmoteColor
Public RGB_AdminColor
Public RGB_HelpColor
Public RGB_WhoColor
Public RGB_JoinLeftColor
Public RGB_NpcColor
Public RGB_AlertColor
Public RGB_NewMapColor
Public RGB_WHITE
Public RGB_LIGHTGREY
Public RGB_GuildSay


Public Sub initColours()
 RGB_SayColor = RGB(255, 255, 255)
 RGB_GlobalColor = RGB(115, 171, 255)
 RGB_BroadcastColor = RGB(141, 212, 113)
RGB_TellColor = RGB(250, 162, 27)
RGB_EmoteColor = RGB(255, 119, 119)
 RGB_AdminColor = RGB(228, 220, 0)
 RGB_HelpColor = RGB(113, 208, 120)
 RGB_WhoColor = RGB(255, 0, 0)
 RGB_JoinLeftColor = RGB(0, 224, 11)
 RGB_NpcColor = RGB(199, 199, 199)
RGB_AlertColor = RGB(102, 102, 104)
 RGB_NewMapColor = RGB(255, 255, 255)
 RGB_WHITE = RGB(255, 255, 255)
 RGB_GuildSay = RGB(0, 255, 255)
 RGB_LIGHTGREY = RGB(211, 211, 211)
End Sub


Public Sub AddText(ByVal rTxt As RichTextBox, ByVal msg As String, ByVal Color As Integer)
Dim s As String
  
    s = vbCrLf + msg
    rTxt.SelStart = Len(rTxt.Text)
    rTxt.SelColor = QBColor(Color)
    rTxt.SelText = s
    rTxt.SelStart = Len(rTxt.Text) - 1
End Sub

Public Sub RGB_AddText(ByVal rTxt As RichTextBox, ByVal msg As String, ByVal Color As Long)
Dim s As String
  
    s = vbCrLf + msg
    rTxt.SelStart = Len(rTxt.Text)
    rTxt.SelColor = Color
    rTxt.SelText = s
    rTxt.SelStart = Len(rTxt.Text) - 1
End Sub

Public Sub TextAdd(ByVal Txt As TextBox, msg As String, NewLine As Boolean)
Static NumLines As Long

    If NewLine Then
        Txt.Text = Txt.Text & vbCrLf & msg
    Else
        Txt.Text = Txt.Text & msg
    End If
        
    NumLines = NumLines + 1
    If NumLines >= MAX_LINES Then
        Txt.Text = ""
        NumLines = 0
    End If
    
    Txt.SelStart = Len(Txt.Text)
End Sub


