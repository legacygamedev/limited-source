VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHyperText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents FormSubClass As clsSubClass
Attribute FormSubClass.VB_VarHelpID = -1
Private mRTB As RichTextBox

'useful marker characters (&h00 - &h1F)
Private tagHL As clsTags
Private tagURL As clsTags
Private tagINS As clsTags
Private tagTGL As clsTags
Private tagMSG As clsTags

Private mEnabled As Boolean
Private mInitialized As Boolean

Public Event Clicked(Button As Integer, URL As String, UnderlyingText As String)
Public Event MouseMove(URL As String)

Private oldRTF As String
Private Sub Class_Initialize()
    Set tagHL = New clsTags
    tagHL.Initialize &H10, &H11
    Set tagINS = New clsTags
    tagINS.Initialize &H12, &H13
    Set tagTGL = New clsTags
    tagTGL.Initialize &H14, &H15
    Set tagURL = New clsTags
    tagURL.Initialize &H16, &H17
    Set tagMSG = New clsTags
    tagMSG.Initialize tagURL.Ending.Value, tagHL.Ending.Value
    
    Set FormSubClass = New clsSubClass
End Sub

Private Sub Class_Terminate()
    DisableLinkNotifications
End Sub

Public Sub Initialize(RTB As RichTextBox)
    If mInitialized Then Exit Sub
    Set mRTB = RTB
    mEnabled = False
    mInitialized = True
    
    EnableLinkNotifications
End Sub

Private Sub EnableLinkNotifications()
Dim EventMask As Long

    If mEnabled Then Exit Sub
    If Not mInitialized Then Exit Sub
    
    'subclass the parent of the RTB to receive EN_LINK notifications
    FormSubClass.Enable mRTB.Parent.hwnd
    
    'get current event mask flags
    EventMask = SendMessage(mRTB.hwnd, EM_GETEVENTMASK, 0&, 0&)
    'turn on Link Notification flag
    EventMask = EventMask Or ENM_LINK Or ENM_CHANGE
    'set RTB to notify parent when user has clicked hyperlink
    SendMessage mRTB.hwnd, EM_SETEVENTMASK, 0&, ByVal EventMask
    mEnabled = True
End Sub

Private Sub DisableLinkNotifications()
Dim EventMask As Long
    If Not mEnabled Then Exit Sub
    
    'turn off subclassing
    FormSubClass.Unsubclass

    'get current event mask flags
    EventMask = SendMessage(mRTB.hwnd, EM_GETEVENTMASK, 0&, 0&)
    'turn off Link Notification flag
    EventMask = EventMask - (EventMask And ENM_LINK)
    SendMessage mRTB.hwnd, EM_SETEVENTMASK, 0&, ByVal EventMask

    mEnabled = False
End Sub

Public Sub RefreshHyperlinks()
    'refresh all hyperlink effects
    Dim start As Integer
    'EnableLinkNotifications
    
    Dim SelStart As Integer, SelLength As Integer
    SelStart = mRTB.SelStart
    SelLength = mRTB.SelLength
    LockWindowUpdate mRTB.hwnd

    While Len(SelectText(start, tagHL)) > 0
        setHyperlink
    Wend
    
    LockWindowUpdate 0&
    mRTB.SelStart = SelStart
    mRTB.SelLength = SelLength
End Sub

Public Sub InsertHyperlink(strURL As String, Optional strText As String = "")
'if strText is supplied then the text is inserted and marked as a hyperlink
'otherwise the selected text is marked
Dim strNew As String
Dim URL As String
    'Note: backslashes have to be escaped
    URL = Replace(strURL, "\", "\\")
    
    With mRTB
        If strText = "" Then strText = .SelText
        If Len(strText) = 0 Then Exit Sub
        .SelText = tagINS.Starting.char
        strNew = "\v" & tagHL.Starting.RTF & tagURL.Starting.RTF & URL & tagURL.Ending.RTF & "\v0 " & _
                strText & "\v" & tagHL.Ending.RTF & "\v0 "
        .TextRTF = Replace(.TextRTF, tagINS.Starting.RTF, strNew)
    End With
    RefreshHyperlinks
    oldRTF = mRTB.TextRTF
End Sub

Private Sub FormSubClass_WMArrival(hwnd As Long, uMsg As Long, _
    wParam As Long, lParam As Long, lRetVal As Long)
'this is where link notifications will happen
Dim notifyCode As nmhdr
Dim LinkData As ENLINK
Dim URL As String, Message As String
Dim Button As Integer


Dim FormatData As CHARFORMAT2
    Select Case uMsg
    Case WM_NOTIFY

        CopyMemory notifyCode, ByVal lParam, LenB(notifyCode)
        If notifyCode.code = EN_LINK Then
        'A RTB sends EN_LINK notifications when it receives certain mouse Messages
        'while the mouse pointer is over text that has the CFE_LINK effect:
        
        'To receive EN_LINK notifications, specify the ENM_LINK flag in the mask
        'sent with the EM_SETEVENTMASK Message.
        
            CopyMemory LinkData, ByVal lParam, Len(LinkData)
            Select Case LinkData.Msg
            Case WM_LBUTTONUP, WM_RBUTTONUP
                'user clicked on a hyperlink
                URL = getURL(LinkData.chrg.cpMin, LinkData.chrg.cpMax, Message)
                Button = (LinkData.Msg And 3) Xor 3
                RaiseEvent Clicked(Button, URL, Message)
            Case WM_MOUSEMOVE, WM_SETCURSOR
                URL = getURL(LinkData.chrg.cpMin, LinkData.chrg.cpMax, Message)
                RaiseEvent MouseMove(URL)
            End Select

        End If
    End Select
        
        lRetVal = FormSubClass.callWindProc(hwnd, uMsg, wParam, lParam)

End Sub

Private Function HighWord(wParam As Long) As Long
    HighWord = Int(wParam / &H10000)
End Function

Private Function getURL(Pos1 As Long, Pos2 As Long, Message As String)
    'parses URL info from the RTF text given a starting position
    'format of RTF should be:
    'HLTAG.1 URLTAG.1 URL URLTAG.2 Message HLTAG.2
    Dim fulltext As String
    fulltext = Mid$(mRTB.Text, Pos1, Pos2 - Pos1 + 1)
    
    getURL = ExtractText(fulltext, tagURL)
    Message = ExtractText(fulltext, tagMSG)
End Function

Private Function ExtractText(fulltext As String, tags As clsTags)
    'extracts and returns the text between the tags
    Dim P1 As Integer, P2 As Integer
    P1 = InStr(fulltext, tags.Starting.char)
    P2 = InStr(P1 + 1, fulltext, tags.Ending.char)
    ExtractText = Mid$(fulltext, P1 + 1, P2 - P1 - 1)
End Function

Private Sub MarkSelection(mytags As clsTags)
'put the hidden tags around the selected text
Dim strSel As String
Dim strRTF As String
    With mRTB
        strSel = .SelText
      'mark the current selection with insertion characters
        .SelText = tagINS.Starting.char
      'replace insertion tag with selected text surrounded by hidden tags
        .TextRTF = Replace(.TextRTF, tagINS.Starting.RTF, _
        "\v" & mytags.Starting.RTF & "\v0 " & strSel & "\v" & _
        mytags.Ending.RTF & "\v0 ")
   End With
End Sub

Private Function SelectText(start As Integer, tags As clsTags) As String
    'select the text that resides between the tags
    Dim P1 As Integer, P2 As Integer
    If start < 0 Then Exit Function
    With mRTB
        P1 = .Find(tags.Starting.char, start)
        .SelLength = 0
        P2 = .Find(tags.Ending.char, P1)
        If P1 < 0 Or P2 < 0 Then
            SelectText = ""
            start = -1
            Exit Function
        End If
        .SelStart = P1
        .SelLength = P2 - P1 + 1
        SelectText = .SelText
        start = P2 + 1
    End With
End Function

Private Sub setHyperlink()
    'mark selected text as hyperlink
    Dim myFormat As CHARFORMAT2
    With myFormat
        .cbSize = LenB(myFormat)
        .dwEffects = CFE_LINK
        .dwMask = CFM_LINK
    End With
    SendMessage mRTB.hwnd, EM_SETCHARFORMAT, SCF_SELECTION, myFormat
End Sub

Private Sub tagSelection(tag1 As String, tag2 As String)
    'Surround the selection with the specified tags
    Dim strRTF As String
    Dim strSel As String
    With mRTB
        strSel = .SelText
        .SelText = tagINS.Starting.char
        strRTF = .TextRTF
        .TextRTF = Replace(strRTF, tagINS.Starting.RTF, tag1 & " " & _
        strSel & tag2 & " ")
    End With
End Sub



