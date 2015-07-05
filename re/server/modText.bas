Attribute VB_Name = "modText"
Option Explicit

Public Const Quote = """"

Public Const MAX_LINES = 2000

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

Public SayColor As Double
Public GlobalColor As Double
Public BroadcastColor As Double
Public TellColor As Double
Public EmoteColor As Double
Public AdminColor As Double
Public HelpColor As Double
Public WhoColor As Double
Public JoinLeftColor As Double
Public NpcColor As Double
Public AlertColor As Double
Public NewMapColor As Double
Public AdminJoinLeaveColor As Double
Public GuildColor As Double

Public NormalColor As Double
Public MoniterColor As Double
Public MapperColor As Double
Public DeveloperColor As Double
Public OwnerColor As Double
Public PKColor As Double


Public Sub AddText(ByVal rTxt As RichTextBox, ByVal Msg As String, ByVal Color As Integer)
Dim s As String
  
    s = vbCrLf + Msg
    rTxt.SelStart = Len(rTxt.text)
    rTxt.SelColor = QBColor(Color)
    rTxt.SelText = s
    rTxt.SelStart = Len(rTxt.text) - 1
End Sub

Public Sub TextAdd(ByVal Txt As TextBox, Msg As String, NewLine As Boolean)
Static NumLines As Long

    If NewLine Then
        Txt.text = Txt.text & vbCrLf & Msg
    Else
        Txt.text = Txt.text & Msg
    End If
        
    NumLines = NumLines + 1
    If NumLines >= MAX_LINES Then
        Txt.text = ""
        NumLines = 0
    End If
    
    Txt.SelStart = Len(Txt.text)
End Sub


