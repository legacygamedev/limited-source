Attribute VB_Name = "modText"

' Copyright (c) 2007-2008 Elysium Source. All rights reserved.
' This code is licensed under the Elysium General License.
Option Explicit

Public Const Quote As String = """"
Public Const MAX_LINES As Long = 2000
Public Const Black As Byte = 0
Public Const Blue As Byte = 1
Public Const Green As Byte = 2
Public Const Cyan As Byte = 3
Public Const Red As Byte = 4
Public Const Magenta As Byte = 5
Public Const Brown As Byte = 6
Public Const Grey As Byte = 7
Public Const DarkGrey As Byte = 8
Public Const BrightBlue As Byte = 9
Public Const BrightGreen As Byte = 10
Public Const BrightCyan As Byte = 11
Public Const BrightRed As Byte = 12
Public Const Pink As Byte = 13
Public Const Yellow As Byte = 14
Public Const White As Byte = 15
Public Const SayColor As Byte = White
Public Const GlobalColor As Byte = Green
Public Const BroadcastColor As Byte = Blue
Public Const TellColor As Byte = White
Public Const EmoteColor As Byte = White
Public Const AdminColor As Byte = BrightCyan
Public Const HelpColor As Byte = White
Public Const WhoColor As Byte = Grey
Public Const JoinLeftColor As Byte = Grey
Public Const NpcColor As Byte = White
Public Const AlertColor As Byte = White
Public Const NewMapColor As Byte = Grey
Public Const GuildColor As Byte = DarkGrey
Public Const PartyColor As Byte = Pink

'Public Sub TextAdd(ByVal Txt As TextBox, _
'   Msg As String, _
'   NewLine As Boolean)
'    Static NumLines As Long

'    If NewLine Then
'        Txt.text = Txt.text & vbCrLf & Msg
'    Else
'        Txt.text = Txt.text & Msg
'    End If

'    NumLines = NumLines + 1

'    If NumLines >= MAX_LINES Then
'        Txt.text = vbNullString
'        NumLines = 0
'    End If

'    Txt.SelStart = Len(Txt.text)
'End Sub
