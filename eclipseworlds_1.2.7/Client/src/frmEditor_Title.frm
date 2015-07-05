VERSION 5.00
Begin VB.Form frmEditor_Title 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Title Editor"
   ClientHeight    =   4305
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5730
   Icon            =   "frmEditor_Title.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   5730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdChangeDataSize 
      Caption         =   "Change Data Size"
      Height          =   375
      Left            =   180
      TabIndex        =   19
      Top             =   3840
      Width           =   2355
   End
   Begin VB.Frame Frame3 
      Caption         =   "Title List"
      Height          =   3735
      Left            =   120
      TabIndex        =   12
      Top             =   0
      Width           =   2415
      Begin VB.CommandButton cmdPaste 
         Caption         =   "Paste"
         Height          =   315
         Left            =   1680
         TabIndex        =   18
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox txtSearch 
         CausesValidation=   0   'False
         Height          =   270
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdCopy 
         Caption         =   "Copy"
         Height          =   315
         Left            =   960
         TabIndex        =   17
         Top             =   240
         Width           =   615
      End
      Begin VB.ListBox lstIndex 
         Height          =   2985
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   2175
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Properties"
      Height          =   3735
      Left            =   2640
      TabIndex        =   10
      Top             =   0
      Width           =   3015
      Begin VB.HScrollBar scrlColor 
         Height          =   255
         Left            =   120
         Max             =   17
         TabIndex        =   3
         Top             =   1200
         Width           =   2775
      End
      Begin VB.HScrollBar scrlPKReq 
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   2400
         Width           =   2775
      End
      Begin VB.Frame Frame4 
         Caption         =   "Description"
         Height          =   855
         Left            =   120
         TabIndex        =   15
         Top             =   2760
         Width           =   2775
         Begin VB.TextBox txtDesc 
            Height          =   495
            Left            =   120
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   6
            Top             =   240
            Width           =   2535
         End
      End
      Begin VB.HScrollBar scrlLevelReq 
         Height          =   255
         Left            =   120
         Max             =   5
         TabIndex        =   4
         Top             =   1800
         Width           =   2775
      End
      Begin VB.TextBox txtName 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   2775
      End
      Begin VB.Label lblPKReq 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Player Kill Requirement: 0"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   2160
         Width           =   1800
      End
      Begin VB.Label lblLevelReq 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Level Requirement: 0"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   1560
         Width           =   1515
      End
      Begin VB.Label lblColor 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Color: Black"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   960
         Width           =   855
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   465
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   2760
      TabIndex        =   7
      Top             =   3840
      Width           =   855
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   3720
      TabIndex        =   8
      Top             =   3840
      Width           =   855
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   4680
      TabIndex        =   9
      Top             =   3840
      Width           =   855
   End
End
Attribute VB_Name = "frmEditor_Title"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TmpIndex As Long

Private Sub cmdChangeDataSize_Click()
    Dim Res As VbMsgBoxResult, val As String
    Dim dataModified As Boolean, I As Long
    
    If EditorIndex < 1 Or EditorIndex > MAX_TITLES Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    For I = 1 To MAX_TITLES
        If Title_Changed(I) And I <> EditorIndex Then
        
            dataModified = True
            Exit For
        End If
    Next
    
    If dataModified Then
        Res = MsgBox("Do you want to continue and discard the changes you made to your data?", vbYesNo)
        
        If Res = vbNo Then Exit Sub
    End If
    
    val = InputBox("Enter the amount you want the new data size to be.", "Change Data Size", MAX_TITLES)
    
    If Not IsNumeric(val) Then
        Exit Sub
    End If
    
    Res = Abs(val)
    
    If Res = MAX_TITLES Then Exit Sub
    
    Call SendChangeDataSize(Res, EDITOR_TITLE)
    
    Unload frmEditor_Title
    MAX_TITLES = Res
    ReDim title(MAX_TITLES)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdClose_Click", "frmEditor_Title", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdClose_Click()
    If EditorIndex < 1 Or EditorIndex > MAX_TITLES Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    frmAdmin.chkEditor(EDITOR_TITLE).FontBold = False
    frmAdmin.picEye(EDITOR_TITLE).Visible = False
    BringWindowToTop (frmAdmin.hWnd)
    Unload frmEditor_Title
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdClose_Click", "frmEditor_Title", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdDelete_Click()
    Dim TmpIndex As Long
    
    If EditorIndex < 1 Or EditorIndex > MAX_TITLES Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ClearTitle EditorIndex
    
    TmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & title(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = TmpIndex
    
    TitleEditorInit
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdDelete_Click", "frmEditor_Title", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdSave_Click()
    If EditorIndex < 1 Or EditorIndex > MAX_TITLES Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    EditorSave = True
    TitleEditorSave
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdSave_Click", "frmEditor_Title", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub


Private Sub lstIndex_Click()
    If EditorIndex < 1 Or EditorIndex > MAX_TITLES Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    TitleEditorInit
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "1stIndex_Click", "frmEditor_Title", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlLevelReq_Change()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    title(EditorIndex).LevelReq = Trim$(scrlLevelReq.Value)
    lblLevelReq.Caption = "Level Requirement: " & scrlLevelReq.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlLevelReq_Change", "frmEditor_Title", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlPKReq_Change()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    title(EditorIndex).PKReq = Trim$(scrlPKReq.Value)
    lblPKReq.Caption = "Player Kill Requirement: " & scrlPKReq.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlPKReq_Change", "frmEditor_Title", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlColor_Change()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lblColor.Caption = "Color: " & GetColorName(scrlColor.Value)
    title(EditorIndex).Color = scrlColor.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlColor_Change", "frmEditor_Title", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtDesc_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_TITLES Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    title(EditorIndex).Desc = Trim$(txtDesc.text)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "txtDesc_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
    Dim TmpIndex As Long
    
    If EditorIndex < 1 Or EditorIndex > MAX_TITLES Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    TmpIndex = lstIndex.ListIndex
    title(EditorIndex).Name = Trim$(txtName.text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & title(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = TmpIndex
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "txtName_Validate", "frmEditor_Title", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    frmMain.UnsubDaFocus Me.hWnd
    If EditorSave = False Then
        TitleEditorCancel
    Else
        EditorSave = False
    End If
    frmAdmin.chkEditor(EDITOR_TITLE).Value = False
    BringWindowToTop (frmAdmin.hWnd)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "Form_Unload", "frmEditor_Title", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub Form_Load()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    frmMain.SubDaFocus Me.hWnd
    ' Set max values
    scrlLevelReq.max = MAX_LEVEL
    txtDesc.MaxLength = 100
    txtName.MaxLength = NAME_LENGTH
    txtSearch.MaxLength = NAME_LENGTH
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "Form_Load", "frmEditor_Title", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtName_GotFocus()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    txtName.SelStart = Len(txtName)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "txtName_GotFocus", "frmEditor_Title", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtDesc_GotFocus()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    txtDesc.SelStart = Len(txtDesc)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "txtDesc_GotFocus", "frmEditor_Title", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtSearch_Change()
    Dim Find As String, I As Long
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    For I = 0 To lstIndex.ListCount - 1
        Find = Trim$(I + 1 & ": " & txtSearch.text)
        
        ' Make sure we dont try to check a name that's too small
        If Len(lstIndex.List(I)) >= Len(Find) Then
            If UCase$(Mid$(Trim$(lstIndex.List(I)), 1, Len(Find))) = UCase$(Find) Then
                lstIndex.ListIndex = I
                Exit For
            End If
        End If
    Next
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "txtSearch_Change", "frmEditor_Title", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtSearch_GotFocus()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    txtSearch.SelStart = Len(txtSearch)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "txtSearch_GotFocus", "frmEditor_Title", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Dim buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If KeyAscii = vbKeyReturn Then
        cmdSave_Click
        KeyAscii = 0
    ElseIf KeyAscii = vbKeyEscape Then
        cmdClose_Click
        KeyAscii = 0
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "Form_KeyPress", "frmEditor_Title", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdCopy_Click()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    TmpIndex = lstIndex.ListIndex
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdCopy_Click", "frmEditor_Title", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdPaste_Click()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
      
    lstIndex.RemoveItem EditorIndex - 1
    Call CopyMemory(ByVal VarPtr(title(EditorIndex)), ByVal VarPtr(title(TmpIndex + 1)), LenB(title(TmpIndex + 1)))
    lstIndex.AddItem EditorIndex & ": " & Trim$(title(EditorIndex).Name), EditorIndex - 1
    lstIndex.ListIndex = EditorIndex - 1
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdPaste_Click", "frmEditor_Title", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub
