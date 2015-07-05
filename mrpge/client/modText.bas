Attribute VB_Name = "modText"
Option Explicit

Public Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal H As Long, ByVal w As Long, ByVal E As Long, ByVal O As Long, ByVal w As Long, ByVal i As Long, ByVal u As Long, ByVal s As Long, ByVal C As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal f As String) As Long
Public Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Public Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long

Public Const Quote = """"

Public Const Black = 15
Public Const Blue = 9
Public Const Green = 10
Public Const Cyan = 11
Public Const Red = 12
Public Const Magenta = 13
Public Const Brown = 14
Public Const Grey = 7
Public Const DarkGrey = 7
Public Const BrightBlue = 9
Public Const BrightGreen = 10
Public Const BrightCyan = 11
Public Const BrightRed = 12
Public Const Pink = 13
Public Const Yellow = 14
Public Const White = 15

Public Const SayColor = White
Public Const GlobalColor = Cyan
Public Const TellColor = Grey
Public Const EmoteColor = BrightCyan
Public Const HelpColor = BrightBlue
Public Const WhoColor = BrightRed
Public Const JoinLeftColor = DarkGrey
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
Public RGB_GuildSay
Public RGB_MassMsg



Public TexthDC As Long
Public GameFont As Long

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
 RGB_MassMsg = RGB(173, 216, 230)
End Sub

Public Sub SetFont(ByVal Font As String, ByVal Size As Byte)
    GameFont = CreateFont(Size, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, Font)
End Sub

Public Sub DrawText(ByVal hdc As Long, ByVal x, ByVal y, ByVal text As String, Color As Long)
    Call SelectObject(hdc, GameFont)
    Call SetBkMode(hdc, vbTransparent)
    If Color <> -1 Then
        Call SetTextColor(hdc, RGB(0, 0, 0))
        Call TextOut(hdc, x - 1, y, text, Len(text))
        Call TextOut(hdc, x - 2, y, text, Len(text))
        Call TextOut(hdc, x + 1, y, text, Len(text))
        Call TextOut(hdc, x + 2, y, text, Len(text))
        Call TextOut(hdc, x, y + 1, text, Len(text))
        Call TextOut(hdc, x, y - 1, text, Len(text))
        Call TextOut(hdc, x, y + 2, text, Len(text))
        Call TextOut(hdc, x, y - 2, text, Len(text))
        Call SetTextColor(hdc, Color)
        Call TextOut(hdc, x, y, text, Len(text))
    End If
End Sub

Public Sub AddText(ByVal Msg As String, ByVal Color As Long)
'Dim s As String
'
'    s = vbNewLine & Msg
'    frmMirage.txtChat.SelStart = Len(frmMirage.txtChat.text)
'    If Color > 15 Then
'        frmMirage.txtChat.SelColor = Color
'    Else
'        If Color = 15 Then Color = 0
'        If Color = 10 Then Color = 2
'        frmMirage.txtChat.SelColor = QBColor(Color)
'    End If
'    frmMirage.txtChat.SelBold = False
'    frmMirage.txtChat.SelText = s
'    frmMirage.txtChat.SelStart = Len(frmMirage.txtChat.text) - 1
Call AddTextNew(frmMirage.txtChannelGeneral, Msg, White) 'Color)
Call AddTextNew(frmMirage.txtChannelAll, Msg, White) 'Color)
End Sub

Public Sub AddGlobalText(ByVal Msg As String, ByVal Color As Long)
'Dim s As String
'
'    s = vbNewLine & Msg
'    frmMirage.txtGlobalChat.SelStart = Len(frmMirage.txtGlobalChat.text)
'    If Color > 15 Then
'        frmMirage.txtGlobalChat.SelColor = Color
'    Else
'        If Color = 15 Then Color = 0
'        frmMirage.txtGlobalChat.SelColor = QBColor(Color)
'    End If
'    frmMirage.txtGlobalChat.SelText = s
'    frmMirage.txtGlobalChat.SelStart = Len(frmMirage.txtGlobalChat.text) - 1
'Call AddTextNew(frmMirage.txtChannelGlobal, Msg, Color) 'Color)
'Call AddTextNew(frmMirage.txtChannelAll, Msg, Color) 'Color)

Dim s As String
Dim txtBox As RichTextBox
Set txtBox = frmMirage.txtChannelAll
    s = vbNewLine & Msg
    txtBox.SelStart = Len(txtBox.text)

    txtBox.SelStart = Len(txtBox.text)
    If Color > 15 Then
        txtBox.SelColor = Color
    Else
        If Color = -1 Then Color = 1
        txtBox.SelColor = QBColor(Color)
    End If
    txtBox.SelBold = True
    txtBox.SelText = s
    txtBox.SelStart = Len(txtBox.text) - 1
    txtBox.SelBold = False
    txtBox.SelText = ""

End Sub

Public Sub AddTextNew(ByRef txtBox As RichTextBox, ByVal Msg As String, ByVal Color As Long)
Dim s As String

    s = vbNewLine & Msg
    txtBox.SelStart = Len(txtBox.text)

    txtBox.SelStart = Len(txtBox.text)
    If Color > 15 Then
        txtBox.SelColor = Color
    Else
        If Color = -1 Then Color = 1
        txtBox.SelColor = QBColor(Color)
    End If
    txtBox.SelBold = False
    txtBox.SelText = s
    txtBox.SelStart = Len(txtBox.text) - 1
    txtBox.SelBold = False
    txtBox.SelText = ""

End Sub

Public Sub AddGenerText(ByVal Msg As String, ByVal Color As Long)
Dim s As String

    s = vbNewLine & Msg
    frmMirage.txtGlobalChat.SelStart = Len(frmMirage.txtGlobalChat.text)
    If Color > 15 Then
        frmMirage.txtGlobalChat.SelColor = Color
    Else
        'If Color = 15 Then Color = 0
        frmMirage.txtGlobalChat.SelColor = QBColor(Color)
    End If
    frmMirage.txtGlobalChat.SelText = s
    frmMirage.txtGlobalChat.SelStart = Len(frmMirage.txtGlobalChat.text) - 1
End Sub


Public Sub addServerText(ByVal Msg As String)
Dim s As String
Dim start As String
'    start = vbNewLine & "Server Message: "
'    s = Msg
'    frmMirage.txtChannelGlobal.SelStart = Len(frmMirage.txtChat.Text)
'    frmMirage.txtChannelGlobal.SelColor = QBColor(10)
'    frmMirage.txtChannelGlobal.SelBold = True
'    frmMirage.txtChannelGlobal.SelText = start
'
'    frmMirage.txtChannelGlobal.SelStart = Len(frmMirage.txtChat.Text)
'    frmMirage.txtChannelGlobal.SelColor = vbWhite 'vbBlack
'    frmMirage.txtChannelGlobal.SelBold = True
'    frmMirage.txtChannelGlobal.SelText = s
'    frmMirage.txtChannelGlobal.SelStart = Len(frmMirage.txtChat.Text) - 1
'    frmMirage.txtChannelGlobal.SelBold = False
    
    start = vbNewLine & "Server Message: "
    s = Msg
    frmMirage.txtChannelAll.SelStart = Len(frmMirage.txtChannelAll.text)
    frmMirage.txtChannelAll.SelColor = QBColor(10)
    frmMirage.txtChannelAll.SelBold = True
    frmMirage.txtChannelAll.SelText = start
    
    frmMirage.txtChannelAll.SelStart = Len(frmMirage.txtChannelAll.text)
    frmMirage.txtChannelAll.SelColor = vbWhite 'vbBlack
    frmMirage.txtChannelAll.SelBold = True
    frmMirage.txtChannelAll.SelText = s
    frmMirage.txtChannelAll.SelStart = Len(frmMirage.txtChannelAll.text) - 1
    frmMirage.txtChannelAll.SelBold = False
    frmMirage.txtChannelAll.SelText = ""
End Sub



Public Sub TextAdd(ByVal Txt As TextBox, Msg As String, NewLine As Boolean)
    If NewLine Then
        Txt.text = Txt.text + Msg + vbCrLf
    Else
        Txt.text = Txt.text + Msg
    End If
    
    Txt.SelStart = Len(Txt.text) - 1
End Sub

Function Parse(ByVal num As Long, ByVal data As String)
Dim i As Long
Dim n As Long
Dim sChar As Long
Dim eChar As Long

    n = 0
    sChar = 1
    
    For i = 1 To Len(data)
        If Mid(data, i, 1) = SEP_CHAR Then
            If n = num Then
                eChar = i
                Parse = Mid(data, sChar, eChar - sChar)
                Exit For
            End If
            
            sChar = i + 1
            n = n + 1
        End If
    Next i
End Function


'SIGNS
Sub ShowSign(ByVal header As String, ByVal Msg As String)
    frmSign.Caption = header
    frmSign.txtMessage = Msg
    frmSign.Show
    
End Sub

