VERSION 5.00
Begin VB.Form frmEditor_Animation 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Animation Editor"
   ClientHeight    =   7560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10095
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEditor_Animation.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   504
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   673
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdChangeDataSize 
      Caption         =   "Change Data Size"
      Height          =   375
      Left            =   120
      TabIndex        =   33
      Top             =   7080
      Width           =   3135
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   5880
      TabIndex        =   13
      Top             =   7080
      Width           =   1455
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   7560
      TabIndex        =   14
      Top             =   7080
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   4200
      TabIndex        =   12
      Top             =   7080
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Properties"
      Height          =   6975
      Left            =   3360
      TabIndex        =   16
      Top             =   0
      Width           =   6615
      Begin VB.ComboBox cmbSound 
         Height          =   300
         Left            =   4080
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   240
         Width           =   2415
      End
      Begin VB.HScrollBar scrlLoopTime 
         Height          =   255
         Index           =   1
         Left            =   3360
         Min             =   1
         TabIndex        =   11
         Top             =   3120
         Value           =   1
         Width           =   3135
      End
      Begin VB.HScrollBar scrlLoopTime 
         Height          =   255
         Index           =   0
         Left            =   120
         Min             =   1
         TabIndex        =   10
         Top             =   3120
         Value           =   1
         Width           =   3135
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   840
         TabIndex        =   2
         Top             =   240
         Width           =   2535
      End
      Begin VB.HScrollBar scrlFrameCount 
         Height          =   255
         Index           =   1
         Left            =   3360
         Min             =   1
         TabIndex        =   9
         Top             =   2520
         Value           =   1
         Width           =   3135
      End
      Begin VB.HScrollBar scrlLoopCount 
         Height          =   255
         Index           =   1
         Left            =   3360
         Min             =   1
         TabIndex        =   7
         Top             =   1920
         Value           =   1
         Width           =   3135
      End
      Begin VB.PictureBox picSprite 
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000012&
         Height          =   3135
         Index           =   1
         Left            =   3360
         ScaleHeight     =   3075
         ScaleWidth      =   3075
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   3600
         Width           =   3135
      End
      Begin VB.HScrollBar scrlSprite 
         Height          =   255
         Index           =   1
         Left            =   3360
         TabIndex        =   5
         Top             =   1320
         Value           =   1
         Width           =   3135
      End
      Begin VB.HScrollBar scrlFrameCount 
         Height          =   255
         Index           =   0
         Left            =   120
         Min             =   1
         TabIndex        =   8
         Top             =   2520
         Value           =   1
         Width           =   3135
      End
      Begin VB.HScrollBar scrlLoopCount 
         Height          =   255
         Index           =   0
         Left            =   120
         Min             =   1
         TabIndex        =   6
         Top             =   1920
         Value           =   1
         Width           =   3135
      End
      Begin VB.PictureBox picSprite 
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000012&
         Height          =   3135
         Index           =   0
         Left            =   120
         ScaleHeight     =   3075
         ScaleWidth      =   3075
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   3600
         Width           =   3135
      End
      Begin VB.HScrollBar scrlSprite 
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   1320
         Value           =   1
         Width           =   3135
      End
      Begin VB.Label Label3 
         Caption         =   "Sound:"
         Height          =   255
         Left            =   3480
         TabIndex        =   30
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblLoopTime 
         Caption         =   "Loop Time: 1"
         Height          =   255
         Index           =   1
         Left            =   3360
         TabIndex        =   29
         Top             =   2880
         Width           =   2655
      End
      Begin VB.Label lblLoopTime 
         Caption         =   "Loop Time: 1"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   28
         Top             =   2880
         Width           =   2655
      End
      Begin VB.Label Label1 
         Caption         =   "Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Layer 1 (Above Player)"
         Height          =   180
         Left            =   3360
         TabIndex        =   26
         Top             =   720
         Width           =   1740
      End
      Begin VB.Label lblFrameCount 
         AutoSize        =   -1  'True
         Caption         =   "Frame Count: 1"
         Height          =   180
         Index           =   1
         Left            =   3360
         TabIndex        =   25
         Top             =   2280
         Width           =   1170
      End
      Begin VB.Label lblLoopCount 
         AutoSize        =   -1  'True
         Caption         =   "Loop Count: 1"
         Height          =   180
         Index           =   1
         Left            =   3360
         TabIndex        =   24
         Top             =   1680
         Width           =   1065
      End
      Begin VB.Label lblSprite 
         AutoSize        =   -1  'True
         Caption         =   "Sprite: 0"
         Height          =   180
         Index           =   1
         Left            =   3360
         TabIndex        =   22
         Top             =   1080
         Width           =   660
      End
      Begin VB.Label lblFrameCount 
         AutoSize        =   -1  'True
         Caption         =   "Frame Count: 1"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   21
         Top             =   2280
         Width           =   1170
      End
      Begin VB.Label lblLoopCount 
         AutoSize        =   -1  'True
         Caption         =   "Loop Count: 1"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   1680
         Width           =   1065
      End
      Begin VB.Label lblSprite 
         AutoSize        =   -1  'True
         Caption         =   "Sprite: 0"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   18
         Top             =   1080
         Width           =   660
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Layer 0 (Below Player)"
         Height          =   180
         Left            =   120
         TabIndex        =   17
         Top             =   720
         Width           =   1695
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Animation List"
      Height          =   6975
      Left            =   120
      TabIndex        =   15
      Top             =   0
      Width           =   3135
      Begin VB.TextBox txtSearch 
         CausesValidation=   0   'False
         Height          =   270
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdCopy 
         Caption         =   "Copy"
         Height          =   315
         Left            =   1680
         TabIndex        =   32
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdPaste 
         Caption         =   "Paste"
         Height          =   315
         Left            =   2400
         TabIndex        =   31
         Top             =   240
         Width           =   615
      End
      Begin VB.ListBox lstIndex 
         Height          =   6105
         ItemData        =   "frmEditor_Animation.frx":038A
         Left            =   120
         List            =   "frmEditor_Animation.frx":038C
         TabIndex        =   1
         Top             =   660
         Width           =   2895
      End
   End
End
Attribute VB_Name = "frmEditor_Animation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private TmpIndex As Long

Private Sub cmbSound_Click()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If cmbSound.ListIndex >= 0 Then
        Audio.StopSounds
        Animation(EditorIndex).Sound = cmbSound.List(cmbSound.ListIndex)
        Audio.PlaySound Animation(EditorIndex).Sound, -1, -1, True
    Else
        Animation(EditorIndex).Sound = vbNullString
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdSound_Click", "frmEditor_Animation", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdChangeDataSize_Click()
    Dim Res As VbMsgBoxResult, val As String
    Dim dataModified As Boolean, I As Long
    
    If EditorIndex < 1 Or EditorIndex > MAX_ANIMATIONS Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    For I = 1 To MAX_ANIMATIONS
        If Animation_Changed(I) And I <> EditorIndex Then
            dataModified = True
            Exit For
        End If
    Next
    
    If dataModified Then
        Res = MsgBox("Do you want to continue and discard the changes you made to your data?", vbYesNo)
        
        If Res = vbNo Then Exit Sub
    End If
    
    val = InputBox("Enter the amount you want the new data size to be.", "Change Data Size", MAX_ANIMATIONS)
    
    If Not IsNumeric(val) Then
        Exit Sub
    End If
    
    Res = Abs(val)
    
    If Res = MAX_ANIMATIONS Then Exit Sub
    
    Call SendChangeDataSize(Res, EDITOR_ANIMATION)
    
    Unload frmEditor_Animation
    MAX_ANIMATIONS = Res
    ReDim Animation(MAX_ANIMATIONS)
    
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdChangeDataSize_Click", "frmEditor_Animation", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdClose_Click()
    If EditorIndex < 1 Or EditorIndex > MAX_ANIMATIONS Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    frmAdmin.chkEditor(EDITOR_ANIMATION).FontBold = False
    frmAdmin.picEye(EDITOR_ANIMATION).Visible = False
    BringWindowToTop (frmAdmin.hWnd)
    Unload frmEditor_Animation
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdClose_Click", "frmEditor_Animation", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdDelete_Click()
    Dim TmpIndex As Long
    
    If EditorIndex < 1 Or EditorIndex > MAX_ANIMATIONS Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ClearAnimation EditorIndex
    
    TmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Animation(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = TmpIndex
    
    AnimationEditorInit
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdDelete_Click", "frmEditor_Animation", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdSave_Click()
    If EditorIndex < 1 Or EditorIndex > MAX_ANIMATIONS Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    EditorSave = True
    AnimationEditorSave
    'frmAdmin.chkEditor(EDITOR_ANIMATION).FontBold = False
    'frmAdmin.picEye(EDITOR_ANIMATION).Visible = False
    'BringWindowToTop (frmAdmin.hWnd)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdSave_Click", "frmEditor_Animation", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub Form_Load()
    Dim I As Long
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    frmMain.SubDaFocus Me.hWnd
    For I = 0 To 1
        scrlSprite(I).max = NumAnimations
        scrlLoopCount(I).max = 100
        scrlFrameCount(I).max = 100
        scrlLoopTime(I).max = 1000
    Next
    
    txtName.MaxLength = NAME_LENGTH
    txtSearch.MaxLength = NAME_LENGTH
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "Form_Load", "frmEditor_Animation", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub lstIndex_Click()
    If EditorIndex < 1 Or EditorIndex > MAX_ANIMATIONS Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    AnimationEditorInit
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "lstIndex_Click", "frmEditor_Animation", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlFrameCount_Change(Index As Integer)
    If EditorIndex < 1 Or EditorIndex > MAX_ANIMATIONS Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lblFrameCount(Index).Caption = "Frame Count: " & scrlFrameCount(Index).Value
    Animation(EditorIndex).Frames(Index) = scrlFrameCount(Index).Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlFrameCount_Change", "frmEditor_Animation", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlFrameCount_Scroll(Index As Integer)
    If EditorIndex < 1 Or EditorIndex > MAX_ANIMATIONS Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    scrlFrameCount_Change Index
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlFrameCount_Scroll", "frmEditor_Animation", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlLoopCount_Change(Index As Integer)
    If EditorIndex < 1 Or EditorIndex > MAX_ANIMATIONS Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lblLoopCount(Index).Caption = "Loop Count: " & scrlLoopCount(Index).Value
    Animation(EditorIndex).LoopCount(Index) = scrlLoopCount(Index).Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlLoopCount_Change", "frmEditor_Animation", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlLoopCount_Scroll(Index As Integer)
    If EditorIndex < 1 Or EditorIndex > MAX_ANIMATIONS Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    scrlLoopCount_Change Index
    Exit Sub
        
' Error handler
ErrorHandler:
    HandleError "scrlLoopCount_Scroll", "frmEditor_Animation", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlLoopTime_Change(Index As Integer)
    If EditorIndex < 1 Or EditorIndex > MAX_ANIMATIONS Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lblLoopTime(Index).Caption = "Loop Time: " & scrlLoopTime(Index).Value
    Animation(EditorIndex).looptime(Index) = scrlLoopTime(Index).Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlLoopTime_Change", "frmEditor_Animation", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlLoopTime_Scroll(Index As Integer)
    If EditorIndex < 1 Or EditorIndex > MAX_ANIMATIONS Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    scrlLoopTime_Change Index
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlLoopTime_Scroll", "frmEditor_Animation", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlSprite_Change(Index As Integer)
    If EditorIndex < 1 Or EditorIndex > MAX_ANIMATIONS Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lblSprite(Index).Caption = "Sprite: " & scrlSprite(Index).Value
    Animation(EditorIndex).Sprite(Index) = scrlSprite(Index).Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlSprite_Change", "frmEditor_Animation", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlSprite_Scroll(Index As Integer)
    If EditorIndex < 1 Or EditorIndex > MAX_ANIMATIONS Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    scrlSprite_Change Index
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlSprite_Scroll", "frmEditor_Animation", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
    Dim TmpIndex As Long

    If EditorIndex < 1 Or EditorIndex > MAX_ANIMATIONS Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    TmpIndex = lstIndex.ListIndex
    Animation(EditorIndex).Name = Trim$(txtName.text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Animation(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = TmpIndex
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "txtName_Validate", "frmEditor_Animation", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtName_GotFocus()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    txtName.SelStart = Len(txtName)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "txtName_GotFocus", "frmEditor_Animation", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    frmMain.UnsubDaFocus Me.hWnd
    If EditorSave = False Then
        AnimationEditorCancel
    Else
        EditorSave = False
    End If
    frmAdmin.chkEditor(EDITOR_ANIMATION).Value = False
    BringWindowToTop (frmAdmin.hWnd)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "Form_Unload", "frmEditor_Animation", Err.Number, Err.Description, Err.Source, Err.HelpContext
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
    HandleError "txtSearch_Change", "frmEditor_Animation", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtSearch_GotFocus()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    txtSearch.SelStart = Len(txtSearch)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "txtSearch_GotFocus", "frmEditor_Animation", Err.Number, Err.Description, Err.Source, Err.HelpContext
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
    HandleError "Form_KeyPress", "frmEditor_Animation", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdCopy_Click()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    TmpIndex = lstIndex.ListIndex
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdCopy_Click", "frmEditor_Animation", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdPaste_Click()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lstIndex.RemoveItem EditorIndex - 1
    Call CopyMemory(ByVal VarPtr(Animation(EditorIndex)), ByVal VarPtr(Animation(TmpIndex + 1)), LenB(Animation(TmpIndex + 1)))
    lstIndex.AddItem EditorIndex & ": " & Trim$(Animation(EditorIndex).Name), EditorIndex - 1
    lstIndex.ListIndex = EditorIndex - 1
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdPaste_Click", "frmEditor_Animation", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub
