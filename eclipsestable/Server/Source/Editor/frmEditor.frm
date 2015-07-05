VERSION 5.00
Object = "{665BF2B8-F41F-4EF4-A8D0-303FBFFC475E}#2.0#0"; "cmcs21.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.ocx"
Begin VB.Form frmEditor 
   Caption         =   "Eclipse Script Editor"
   ClientHeight    =   9135
   ClientLeft      =   195
   ClientTop       =   615
   ClientWidth     =   12000
   ClipControls    =   0   'False
   Icon            =   "frmEditor.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9135
   ScaleWidth      =   12000
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cdlOpenFile 
      Left            =   11280
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Open Script"
      Filter          =   "All Files (*.*)|*.*|"
   End
   Begin CodeSenseCtl.CodeSense RT 
      Height          =   9135
      Left            =   0
      OleObjectBlob   =   "frmEditor.frx":628A
      TabIndex        =   0
      Top             =   0
      Width           =   12015
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuNew 
         Caption         =   "New Script"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuLoad 
         Caption         =   "Load Script"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save Script"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnuUndo 
         Caption         =   "Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuRedo 
         Caption         =   "Redo"
         Shortcut        =   ^Y
      End
      Begin VB.Menu mnuEditSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuCut 
         Caption         =   "Cut"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEditSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFind 
         Caption         =   "Find"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuFindNext 
         Caption         =   "Find Next"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuReplace 
         Caption         =   "Replace"
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuGoLine 
         Caption         =   "Go To Line"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuEditSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelAll 
         Caption         =   "Select All"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuSelLine 
         Caption         =   "Select Current Line"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuWebsite 
         Caption         =   "Visit Website"
      End
      Begin VB.Menu mnuForums 
         Caption         =   "Visit Forums"
      End
      Begin VB.Menu mnuHelpSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About..."
      End
   End
End
Attribute VB_Name = "frmEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim FileID As Integer

    ResetAllEditVals
    GetEditColors
    EditorSetVals

    If LenB(Command) <> 0 Then
        OpenFile = Command
        If FileExists(OpenFile) Then
            FileID = FreeFile
            Open OpenFile For Input As #FileID
                frmEditor.RT.Text = Input$(LOF(FileID), FileID)
            Close #FileID
        End If
    End If
End Sub

Private Sub EditorSetVals()
    RT.Language = "Basic"

    RT.SetColor cmClrBookmark, ClrData(0).frClr
    RT.SetColor cmClrBookmarkBk, ClrData(0).bgClr
    RT.SetColor cmClrCommentBk, ClrData(1).bgClr
    RT.SetColor cmClrComment, ClrData(1).frClr
    RT.SetColor cmClrHDividerLines, ClrData(2).frClr
    RT.SetColor cmClrVDividerLines, ClrData(3).frClr
    RT.SetColor cmClrHighlightedLine, ClrData(4).frClr
    RT.SetColor cmClrKeyword, ClrData(5).frClr
    RT.SetColor cmClrKeywordBk, ClrData(5).bgClr
    RT.SetColor cmClrLeftMargin, ClrData(6).frClr
    RT.SetColor cmClrLineNumber, ClrData(7).frClr
    RT.SetColor cmClrLineNumberBk, ClrData(7).bgClr
    RT.SetColor cmClrNumber, ClrData(8).frClr
    RT.SetColor cmClrNumberBk, ClrData(8).bgClr
    RT.SetColor cmClrOperator, ClrData(9).frClr
    RT.SetColor cmClrOperatorBk, ClrData(9).bgClr
    RT.SetColor cmClrScopeKeyword, ClrData(10).frClr
    RT.SetColor cmClrScopeKeywordBk, ClrData(10).bgClr
    RT.SetColor cmClrString, ClrData(11).frClr
    RT.SetColor cmClrStringBk, ClrData(11).bgClr
    RT.SetColor cmClrTagElementName, ClrData(12).frClr
    RT.SetColor cmClrTagElementNameBk, ClrData(12).bgClr
    RT.SetColor cmClrTagEntity, ClrData(13).frClr
    RT.SetColor cmClrTagEntityBk, ClrData(13).bgClr
    RT.SetColor cmClrTagAttributeName, ClrData(14).frClr
    RT.SetColor cmClrTagAttributeNameBk, ClrData(14).bgClr
    RT.SetColor cmClrTagText, ClrData(15).frClr
    RT.SetColor cmClrTagTextBk, ClrData(15).bgClr
    RT.SetColor cmClrText, ClrData(16).frClr
    RT.SetColor cmClrTextBk, ClrData(16).bgClr
    RT.SetColor cmClrWindow, ClrData(17).frClr

    If CInt(GetSetting(App.EXEName, "EditOptions", "Highlight", "1")) = 0 Then
        RT.HighlightedLine = -1
    End If

    RT.LineNumbering = CBool(GetSetting(App.EXEName, "EditOptions", "linenumber", "1"))
    RT.DisplayLeftMargin = CBool(GetSetting(App.EXEName, "EditOptions", "leftmargin", "1"))
    RT.DisplayWhitespace = CBool(GetSetting(App.EXEName, "EditOptions", "whitespace", "0"))
    RT.SmoothScrolling = CBool(GetSetting(App.EXEName, "EditOptions", "smoothscroll", "1"))

    RT.LineNumberStart = 1
    RT.EnableDragDrop = True
    RT.ExpandTabs = True
    RT_SelChange RT
End Sub

Public Sub DoHighLight()
    Dim r As CodeSenseCtl.Range

    Set r = RT.GetSel(True)

    If CInt(GetSetting(App.EXEName, "EditOptions", "Highlight", "1")) = 1 Then
        RT.HighlightedLine = r.EndLineNo
    End If
End Sub

Private Sub Form_Resize()
    RT.Width = Me.Width - 120
    RT.Height = (Me.Height) - 700
End Sub

Private Sub mnuCopy_Click()
    Clipboard.Clear
    Clipboard.SetText RT.SelText
End Sub

Private Sub mnuCut_Click()
    Clipboard.Clear
    Clipboard.SetText RT.SelText
    RT.SelText = vbNullString
End Sub

Private Sub mnuDelete_Click()
    RT.SelText = vbNullString
End Sub

Private Sub mnuFind_Click()
    RT.ExecuteCmd cmCmdFind
End Sub

Private Sub mnuFindNext_Click()
    RT.ExecuteCmd cmCmdFindNext
End Sub

Private Sub mnuGoLine_Click()
    RT.ExecuteCmd cmCmdGotoLine, -1
End Sub

Private Sub mnuLoad_Click()
    Dim FileID As Integer
    Dim FileName As String

    cdlOpenFile.ShowOpen

    FileName = cdlOpenFile.FileName
    If FileExists(FileName) Then
        FileID = FreeFile
        Open FileName For Input As #FileID
            frmEditor.RT.Text = Input$(LOF(FileID), FileID)
        Close #FileID
    Else
        cdlOpenFile.FileName = vbNullString
    End If
End Sub

Private Sub mnuPaste_Click()
    RT.Paste
End Sub

Private Sub mnuRedo_Click()
    RT.Redo
End Sub

Private Sub mnuReplace_Click()
    RT.ExecuteCmd cmCmdFindReplace
End Sub

Private Sub mnuSave_Click()
    Dim FileID As Integer
    Dim FileName As String

    FileID = FreeFile
    FileName = cdlOpenFile.FileName
    If FileName = "" Then
        FileName = OpenFile
    End If

    Open FileName For Output As #FileID
        Print #FileID, RT.Text
    Close #FileID
End Sub

Private Sub mnuSelAll_Click()
    RT.ExecuteCmd cmCmdSelectAll
End Sub

Private Sub mnuSelLine_Click()
    RT.ExecuteCmd cmCmdSelectLine
End Sub

Private Sub mnuUndo_Click()
    RT.Undo
End Sub

Private Function RT_KeyPress(ByVal Control As CodeSenseCtl.ICodeSense, ByVal KeyAscii As Long, ByVal Shift As Long) As Boolean
    GetRange
End Function

Private Function RT_KeyUp(ByVal Control As CodeSenseCtl.ICodeSense, ByVal KeyCode As Long, ByVal Shift As Long) As Boolean
    Dim r As CodeSenseCtl.Range
    Dim sLastWord As String

    If KeyCode = 9 Or KeyCode = 13 Then
        AddIntellWord
    End If

    If RT.CurrentWord <> "." Then
        sLastWord = RT.CurrentWord
    End If

    If KeyCode = 190 Then

        Set r = RT.GetSel(False)

        LBoxPos = r.EndColNo
        RT.ExecuteCmd cmCmdCodeList
    End If
End Function
Private Function RT_CodeList(ByVal Control As CodeSenseCtl.ICodeSense, ByVal ListCtrl As CodeSenseCtl.ICodeList) As Boolean
' ListCtrl.hImageList = IMGIntellisence.hImageList
End Function

Private Function RT_CodeListCancel(ByVal Control As CodeSenseCtl.ICodeSense, ByVal ListCtrl As CodeSenseCtl.ICodeList) As Boolean
    AddIntellWord
    RT_CodeListCancel = False
End Function
Private Function RT_CodeListChar(ByVal Control As CodeSenseCtl.ICodeSense, ByVal ListCtrl As CodeSenseCtl.ICodeList, ByVal wChar As Long, ByVal lKeyData As Long) As Boolean
    RT_CodeListChar = False
End Function
Private Function RT_CodeListSelChange(ByVal Control As CodeSenseCtl.ICodeSense, ByVal ListCtrl As CodeSenseCtl.ICodeList, ByVal lItem As Long) As String
    sIntellText = ListCtrl.GetItemText(lItem)
    RT_CodeListSelChange = vbNullString
End Function
Private Function RT_CodeListSelMade(ByVal Control As CodeSenseCtl.ICodeSense, ByVal ListCtrl As CodeSenseCtl.ICodeList) As Boolean
    AddIntellWord
    RT_CodeListSelMade = False
End Function
Private Function RT_CodeListSelWord(ByVal Control As CodeSenseCtl.ICodeSense, ByVal ListCtrl As CodeSenseCtl.ICodeList, ByVal lItem As Long) As Boolean
    RT_CodeListSelWord = True
End Function
Private Function RT_CodeTip(ByVal Control As CodeSenseCtl.ICodeSense) As CodeSenseCtl.cmToolTipType
    RT_CodeTip = cmToolTipTypeNormal
End Function
Private Function RT_MouseDown(ByVal Control As CodeSenseCtl.ICodeSense, ByVal Button As Long, ByVal Shift As Long, ByVal x As Long, ByVal y As Long) As Boolean
    GetRange
    If Button = 2 Then
        Me.PopupMenu Me.mnuEdit
    End If
End Function

Private Function RT_MouseUp(ByVal Control As CodeSenseCtl.ICodeSense, ByVal Button As Long, ByVal Shift As Long, ByVal x As Long, ByVal y As Long) As Boolean
    GetRange
End Function

Private Sub RT_SelChange(ByVal Control As CodeSenseCtl.ICodeSense)
    DoHighLight
End Sub

Private Sub GetRange()
    Dim r As CodeSenseCtl.Range
    Dim LLine As Long
    Dim LCurrent As Long
    Set r = RT.GetSel(False)
    LLine = r.EndLineNo
    LCurrent = r.EndColNo
    LLine = LLine + 1
    LCurrent = LCurrent + 1
End Sub

Private Sub AddIntellWord()
    Dim r As CodeSenseCtl.Range
    If sIntellText <> vbNullString Then
        Set r = RT.GetSel(False)
        r.StartColNo = LBoxPos
        r.EndColNo = r.EndColNo
        RT.SetSel r, False
        r.StartColNo = r.EndColNo + Len(sIntellText)
        r.EndColNo = r.EndColNo + Len(sIntellText)
        RT.SelText = sIntellText
        RT.SetSel r, False

        sIntellText = vbNullString
    End If
End Sub

