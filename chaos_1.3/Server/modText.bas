Attribute VB_Name = "modText"

' Copyright (c) 2006 Chaos Engine Source. All rights reserved.
' This code is licensed under the Chaos Engine General License.

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
Public Const SayColor = White
Public Const GlobalColor = Green
Public Const BroadcastColor = Blue
Public Const TellColor = White
Public Const EmoteColor = White
Public Const AdminColor = BrightCyan
Public Const HelpColor = White
Public Const WhoColor = Grey
Public Const JoinLeftColor = Grey
Public Const NpcColor = White
Public Const AlertColor = White
Public Const NewMapColor = Grey
Public Const GuildColor = Yellow
Public Const PartyColor = Green

Public Sub TextAdd(ByVal Txt As TextBox, _
   Msg As String, _
   NewLine As Boolean)
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
