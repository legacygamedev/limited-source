VERSION 5.00
Begin VB.Form frmEditor_Resource 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resource Editor"
   ClientHeight    =   8625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8535
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEditor_Resource.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   575
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   569
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton cmdChangeDataSize 
      Caption         =   "Change Data Size"
      Height          =   375
      Left            =   120
      TabIndex        =   47
      Top             =   8160
      Width           =   3135
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   5160
      TabIndex        =   19
      Top             =   8160
      Width           =   1455
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   6840
      TabIndex        =   20
      Top             =   8160
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   3480
      TabIndex        =   18
      Top             =   8160
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Properties"
      Height          =   8055
      Left            =   3360
      TabIndex        =   22
      Top             =   0
      Width           =   5055
      Begin VB.HScrollBar scrlRespawn 
         Height          =   255
         Left            =   2640
         TabIndex        =   45
         Top             =   5280
         Width           =   2295
      End
      Begin VB.HScrollBar scrlLevelReq 
         Height          =   255
         Left            =   2640
         TabIndex        =   43
         Top             =   7080
         Width           =   2295
      End
      Begin VB.HScrollBar scrlExp 
         Height          =   255
         Left            =   2640
         TabIndex        =   10
         Top             =   4680
         Width           =   2295
      End
      Begin VB.HScrollBar scrlSkill 
         Height          =   255
         Left            =   120
         Max             =   3
         Min             =   1
         TabIndex        =   12
         Top             =   7080
         Value           =   1
         Width           =   2295
      End
      Begin VB.ComboBox cmbSound 
         Height          =   300
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1680
         Width           =   3975
      End
      Begin VB.TextBox txtFail 
         Height          =   285
         Left            =   960
         TabIndex        =   4
         Top             =   960
         Width           =   3975
      End
      Begin VB.HScrollBar scrlHighChance 
         Height          =   255
         Left            =   2640
         Max             =   255
         Min             =   2
         TabIndex        =   16
         Top             =   6480
         Value           =   2
         Width           =   2295
      End
      Begin VB.HScrollBar scrlLowChance 
         Height          =   255
         Left            =   120
         Max             =   255
         Min             =   1
         TabIndex        =   15
         Top             =   6480
         Value           =   1
         Width           =   2295
      End
      Begin VB.HScrollBar scrlRewardMax 
         Height          =   255
         Left            =   2640
         Max             =   255
         Min             =   1
         TabIndex        =   14
         Top             =   5880
         Value           =   1
         Width           =   2295
      End
      Begin VB.HScrollBar scrlAnimation 
         Height          =   255
         Left            =   120
         Max             =   6000
         TabIndex        =   17
         Top             =   7680
         Width           =   4815
      End
      Begin VB.HScrollBar scrlExhaustedPic 
         Height          =   255
         Left            =   2640
         TabIndex        =   8
         Top             =   2280
         Width           =   2295
      End
      Begin VB.PictureBox picExhaustedPic 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1680
         Left            =   2640
         ScaleHeight     =   112
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   152
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   2640
         Width           =   2280
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   960
         TabIndex        =   2
         Top             =   240
         Width           =   3975
      End
      Begin VB.HScrollBar scrlNormalPic 
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   2280
         Width           =   2295
      End
      Begin VB.HScrollBar scrlReward 
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   4680
         Width           =   2295
      End
      Begin VB.HScrollBar scrlTool 
         Height          =   255
         Left            =   120
         Max             =   3
         TabIndex        =   11
         Top             =   5280
         Width           =   2295
      End
      Begin VB.HScrollBar scrlRewardMin 
         Height          =   255
         Left            =   120
         Max             =   255
         Min             =   1
         TabIndex        =   13
         Top             =   5880
         Value           =   1
         Width           =   2295
      End
      Begin VB.PictureBox picNormalPic 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1680
         Left            =   120
         ScaleHeight     =   112
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   152
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   2640
         Width           =   2280
      End
      Begin VB.TextBox txtSuccess 
         Height          =   285
         Left            =   960
         TabIndex        =   3
         Top             =   600
         Width           =   3975
      End
      Begin VB.TextBox txtEmpty 
         Height          =   285
         Left            =   960
         TabIndex        =   5
         Top             =   1320
         Width           =   3975
      End
      Begin VB.Label lblRespawn 
         AutoSize        =   -1  'True
         Caption         =   "Respawn Time: 0 s"
         Height          =   180
         Left            =   2640
         TabIndex        =   46
         ToolTipText     =   "In seconds."
         Top             =   5040
         Width           =   2265
      End
      Begin VB.Label lblLevelReq 
         AutoSize        =   -1  'True
         Caption         =   "Skill Level Required: 0"
         Height          =   195
         Left            =   2640
         TabIndex        =   44
         ToolTipText     =   "In seconds."
         Top             =   6840
         Width           =   2295
      End
      Begin VB.Label lblExp 
         AutoSize        =   -1  'True
         Caption         =   "Exp: 0"
         Height          =   195
         Left            =   2640
         TabIndex        =   40
         Top             =   4440
         Width           =   2235
      End
      Begin VB.Label lblSkill 
         AutoSize        =   -1  'True
         Caption         =   "Skill: None"
         Height          =   180
         Left            =   120
         TabIndex        =   39
         Top             =   6840
         Width           =   2145
      End
      Begin VB.Label Label5 
         Caption         =   "Sound:"
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label lblLowChance 
         AutoSize        =   -1  'True
         Caption         =   "Low Chance: 1"
         Height          =   180
         Left            =   120
         TabIndex        =   37
         Top             =   6240
         Width           =   2295
      End
      Begin VB.Label lblHighChance 
         AutoSize        =   -1  'True
         Caption         =   "High Chance: 2"
         Height          =   180
         Left            =   2640
         TabIndex        =   36
         Top             =   6240
         Width           =   2250
      End
      Begin VB.Label lblRewardMax 
         AutoSize        =   -1  'True
         Caption         =   "Maximum Reward: 1"
         Height          =   180
         Left            =   2640
         TabIndex        =   35
         Top             =   5640
         Width           =   2235
      End
      Begin VB.Label lblRewardMin 
         AutoSize        =   -1  'True
         Caption         =   "Minimum Reward: 1"
         DataSource      =   "in"
         Height          =   180
         Left            =   120
         TabIndex        =   34
         Top             =   5640
         Width           =   2295
      End
      Begin VB.Label lblAnim 
         AutoSize        =   -1  'True
         Caption         =   "Animation: None"
         Height          =   180
         Left            =   120
         TabIndex        =   33
         Top             =   7440
         Width           =   4740
      End
      Begin VB.Label lblExhaustedPic 
         AutoSize        =   -1  'True
         Caption         =   "Exhausted Image: 0"
         Height          =   180
         Left            =   2640
         TabIndex        =   32
         Top             =   2040
         Width           =   2130
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         Caption         =   "Name:"
         Height          =   180
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lblFail 
         AutoSize        =   -1  'True
         Caption         =   "Fail:"
         Height          =   180
         Left            =   120
         TabIndex        =   29
         Top             =   960
         Width           =   330
      End
      Begin VB.Label lblNormalPic 
         AutoSize        =   -1  'True
         Caption         =   "Normal Image: 0"
         Height          =   195
         Left            =   120
         TabIndex        =   28
         Top             =   2040
         Width           =   2175
      End
      Begin VB.Label lblReward 
         AutoSize        =   -1  'True
         Caption         =   "Reward: None"
         Height          =   180
         Left            =   120
         TabIndex        =   27
         Top             =   4440
         Width           =   2235
      End
      Begin VB.Label lblTool 
         AutoSize        =   -1  'True
         Caption         =   "Tool Required: None"
         Height          =   180
         Left            =   120
         TabIndex        =   26
         Top             =   5040
         Width           =   2250
      End
      Begin VB.Label lblSuccess 
         AutoSize        =   -1  'True
         Caption         =   "Success:"
         Height          =   180
         Left            =   120
         TabIndex        =   25
         Top             =   600
         Width           =   705
      End
      Begin VB.Label lblEmpty 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Empty:"
         Height          =   180
         Left            =   120
         TabIndex        =   24
         Top             =   1320
         Width           =   540
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Resource List"
      Height          =   8055
      Left            =   120
      TabIndex        =   21
      Top             =   0
      Width           =   3135
      Begin VB.CommandButton cmdPaste 
         Caption         =   "Paste"
         Height          =   315
         Left            =   2400
         TabIndex        =   42
         Top             =   240
         Width           =   615
      End
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
         TabIndex        =   41
         Top             =   240
         Width           =   615
      End
      Begin VB.ListBox lstIndex 
         Height          =   7080
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   2895
      End
   End
End
Attribute VB_Name = "frmEditor_Resource"
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
        Resource(EditorIndex).Sound = cmbSound.List(cmbSound.ListIndex)
        Audio.PlaySound Resource(EditorIndex).Sound, -1, -1, True
    Else
        Resource(EditorIndex).Sound = vbNullString
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdSound_Click", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdChangeDataSize_Click()
    Dim Res As VbMsgBoxResult, val As String
    Dim dataModified As Boolean, I As Long
    
    If EditorIndex < 1 Or EditorIndex > MAX_RESOURCES Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    For I = 1 To MAX_RESOURCES
        If Resource_Changed(I) And I <> EditorIndex Then
        
            dataModified = True
            Exit For
        End If
    Next
    
    If dataModified Then
        Res = MsgBox("Do you want to continue and discard the changes you made to your data?", vbYesNo)
        
        If Res = vbNo Then Exit Sub
    End If
    
    val = InputBox("Enter the amount you want the new data size to be.", "Change Data Size", MAX_RESOURCES)
    
    If Not IsNumeric(val) Then
        Exit Sub
    End If
    
    Res = Abs(val)
    
    If Res = MAX_RESOURCES Then Exit Sub
    
    Call SendChangeDataSize(Res, EDITOR_RESOURCE)
    
    Unload frmEditor_Resource
    MAX_RESOURCES = Res
    ReDim Resource(MAX_RESOURCES)
    
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdChangeDataSize_Click", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdDelete_Click()
    Dim TmpIndex As Long
    
    If EditorIndex < 1 Or EditorIndex > MAX_RESOURCES Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ClearResource EditorIndex
    
    TmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Resource(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = TmpIndex
    
    ResourceEditorInit
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdDelete_Click", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdSave_Click()
    If EditorIndex < 1 Or EditorIndex > MAX_RESOURCES Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If frmEditor_Resource.scrlLowChance.Value >= frmEditor_Resource.scrlHighChance.Value Then
        AlertMsg "The high chance must be greater than the low chance!"
        Exit Sub
    End If
    
    If frmEditor_Resource.scrlRewardMin.Value > frmEditor_Resource.scrlRewardMax.Value Then
        AlertMsg "The maximum reward must be greater than or equal to the minimum reward!"
        Exit Sub
    End If
    
    EditorSave = True
    Call ResourceEditorSave
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdSave_Click", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub Form_Activate()
    hwndLastActiveWnd = hWnd
    
    If FormVisible("frmAdmin") And adminMin Then
        frmAdmin.centerMiniVert Width, Height, Left, Top
    End If
End Sub

Private Sub Form_Load()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    frmMain.SubDaFocus Me.hWnd
    scrlReward.max = MAX_ITEMS
    scrlNormalPic.max = NumResources
    scrlExhaustedPic.max = NumResources
    scrlAnimation.max = MAX_ANIMATIONS
    txtName.MaxLength = NAME_LENGTH
    txtSearch.MaxLength = NAME_LENGTH
    txtSuccess.MaxLength = NAME_LENGTH
    txtFail.MaxLength = NAME_LENGTH
    txtEmpty.MaxLength = NAME_LENGTH
    scrlSkill.max = Skills.Skill_Count - 1
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "Form_Load", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdClose_Click()
    If EditorIndex < 1 Or EditorIndex > MAX_RESOURCES Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    frmAdmin.chkEditor(EDITOR_RESOURCE).FontBold = False
    frmAdmin.picEye(EDITOR_RESOURCE).Visible = False
    BringWindowToTop (frmAdmin.hWnd)
    Unload frmEditor_Resource
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdClose_Click", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    frmMain.UnsubDaFocus Me.hWnd
    If EditorSave = False Then
        ResourceEditorCancel
    Else
        EditorSave = False
    End If
    frmAdmin.chkEditor(EDITOR_RESOURCE).Value = False
    BringWindowToTop (frmAdmin.hWnd)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "Form_Unload", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub lstIndex_Click()
    If EditorIndex < 1 Or EditorIndex > MAX_RESOURCES Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ResourceEditorInit
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "lstIndex_Click", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlAnimation_Change()
    Dim sString As String
    
    If EditorIndex < 1 Or EditorIndex > MAX_RESOURCES Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If scrlAnimation.Value = 0 Then sString = "None" Else sString = Trim$(Animation(scrlAnimation.Value).Name)
    lblAnim.Caption = "Animation: " & sString
    Resource(EditorIndex).Animation = scrlAnimation.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlAnimation_Change", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlExhaustedPic_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_RESOURCES Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lblExhaustedPic.Caption = "Exhausted Image: " & scrlExhaustedPic.Value
    Resource(EditorIndex).ExhaustedImage = scrlExhaustedPic.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlExhaustedPic_Change", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlExp_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_RESOURCES Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lblExp.Caption = "Exp: " & scrlExp.Value
    Resource(EditorIndex).exp = scrlExp.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlExp_Change", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlLevelReq_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_RESOURCES Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lblLevelReq.Caption = "Skill Level Req: " & scrlLevelReq.Value
    Resource(EditorIndex).LevelReq = scrlLevelReq.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlLevelReq_Change", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlRewardMax_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_RESOURCES Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lblRewardMax.Caption = "Maximum Reward: " & scrlRewardMax.Value
    Resource(EditorIndex).Reward_Max = scrlRewardMax.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlRewardMax_Change", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlRewardMin_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_RESOURCES Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lblRewardMin.Caption = "Minimum Reward: " & scrlRewardMin.Value
    Resource(EditorIndex).Reward_Min = scrlRewardMin.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlRewardMin_Change", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlHighChance_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_RESOURCES Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lblHighChance.Caption = "High Chance: " & scrlHighChance.Value
    Resource(EditorIndex).HighChance = scrlHighChance.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlHighChance_Change", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlLowChance_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_RESOURCES Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lblLowChance.Caption = "Low Chance: " & scrlLowChance.Value
    Resource(EditorIndex).LowChance = scrlLowChance.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlLowChance_Change", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlNormalPic_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_RESOURCES Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lblNormalPic.Caption = "Normal Image: " & scrlNormalPic.Value
    Resource(EditorIndex).ResourceImage = scrlNormalPic.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlNormalPic_Change", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlRespawn_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_RESOURCES Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lblRespawn.Caption = "Respawn Time: " & scrlRespawn.Value & " s"
    Resource(EditorIndex).RespawnTime = scrlRespawn.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlRespawn_Change", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlReward_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_RESOURCES Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If scrlReward.Value > 0 Then
        lblReward.Caption = "Reward: " & Trim$(Item(scrlReward.Value).Name)
    Else
        lblReward.Caption = "Reward: None"
    End If
    
    Resource(EditorIndex).ItemReward = scrlReward.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlReward_Change", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlSkill_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_RESOURCES Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lblSkill.Caption = "Skill: " & GetSkillName(scrlSkill.Value)
    Resource(EditorIndex).Skill = scrlSkill.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlSkill_Change", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlTool_Change()
    Dim Name As String

    If EditorIndex < 1 Or EditorIndex > MAX_RESOURCES Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Select Case scrlTool.Value
        Case 0
            Name = "None"
        Case 1
            Name = "Hatchet"
        Case 2
            Name = "Fishing Pole"
        Case 3
            Name = "Pickaxe"
    End Select

    lblTool.Caption = "Tool Required: " & Name
    Resource(EditorIndex).ToolRequired = scrlTool.Value
    Exit Sub

' Error handler
ErrorHandler:
   HandleError "scrlTool_Change", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtFail_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_RESOURCES Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Resource(EditorIndex).FailMessage = Trim$(txtFail.text)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "txtFail_Change", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtSuccess_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_RESOURCES Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Resource(EditorIndex).SuccessMessage = Trim$(txtSuccess.text)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "txtSuccess_Change", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtEmpty_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_RESOURCES Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Resource(EditorIndex).EmptyMessage = Trim$(txtEmpty.text)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "txtEmpty_Change", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
    Dim TmpIndex As Long
    
    If EditorIndex < 1 Or EditorIndex > MAX_RESOURCES Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    TmpIndex = lstIndex.ListIndex
    Resource(EditorIndex).Name = Trim$(txtName.text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Resource(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = TmpIndex
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "txtName_Validate", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtName_GotFocus()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    txtName.SelStart = Len(txtName)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "txtName_GotFocus", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtSuccess_GotFocus()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    txtSuccess.SelStart = Len(txtSuccess)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "txtSuccess_GotFocus", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtFail_GotFocus()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    txtFail.SelStart = Len(txtFail)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "txtFail_GotFocus", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtEmpty_GotFocus()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    txtEmpty.SelStart = Len(txtEmpty)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "txtEmpty_GotFocus", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
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
    HandleError "txtSearch_Change", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtSearch_GotFocus()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    txtSearch.SelStart = Len(txtSearch)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "txtSearch_GotFocus", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
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
    HandleError "Form_KeyPress", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdCopy_Click()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    TmpIndex = lstIndex.ListIndex
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdCopy_Click", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdPaste_Click()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lstIndex.RemoveItem EditorIndex - 1
    Call CopyMemory(ByVal VarPtr(Resource(EditorIndex)), ByVal VarPtr(Resource(TmpIndex + 1)), LenB(Resource(TmpIndex + 1)))
    lstIndex.AddItem EditorIndex & ": " & Trim$(Resource(EditorIndex).Name), EditorIndex - 1
    lstIndex.ListIndex = EditorIndex - 1
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdPaste_Click", "frmEditor_Resource", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub
