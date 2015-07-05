VERSION 5.00
Begin VB.Form frmEditor_NPC 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Npc Editor"
   ClientHeight    =   6390
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9630
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEditor_NPC.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   426
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   642
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame fraSpell 
      Caption         =   "Spell"
      Height          =   1455
      Left            =   3360
      TabIndex        =   55
      Top             =   4320
      Width           =   3015
      Begin VB.HScrollBar scrlSpellNum 
         Height          =   255
         Left            =   1200
         Max             =   255
         TabIndex        =   58
         Top             =   1080
         Width           =   1695
      End
      Begin VB.HScrollBar scrlSpell 
         Height          =   255
         Left            =   120
         Max             =   10
         Min             =   1
         TabIndex        =   56
         Top             =   240
         Value           =   1
         Width           =   2775
      End
      Begin VB.Label lblSpellNum 
         AutoSize        =   -1  'True
         Caption         =   "Num: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   59
         Top             =   1080
         Width           =   555
      End
      Begin VB.Label lblSpellName 
         Caption         =   "Spell: None"
         Height          =   255
         Left            =   120
         TabIndex        =   57
         Top             =   720
         Width           =   2775
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000000&
         X1              =   120
         X2              =   2880
         Y1              =   600
         Y2              =   600
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Info"
      Height          =   4095
      Left            =   3360
      TabIndex        =   35
      Top             =   120
      Width           =   3015
      Begin VB.TextBox txtSpawnSecs 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1080
         TabIndex        =   45
         Text            =   "0"
         Top             =   3600
         Width           =   1815
      End
      Begin VB.HScrollBar scrlConv 
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   3240
         Width           =   2775
      End
      Begin VB.ComboBox cmbSound 
         Height          =   300
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   43
         Top             =   2040
         Width           =   1695
      End
      Begin VB.HScrollBar scrlAnimation 
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   2640
         Width           =   2775
      End
      Begin VB.PictureBox picSprite 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   2400
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   41
         Top             =   960
         Width           =   480
      End
      Begin VB.HScrollBar scrlSprite 
         Height          =   255
         Left            =   1200
         Max             =   255
         TabIndex        =   40
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox txtName 
         Height          =   270
         Left            =   840
         TabIndex        =   39
         Top             =   240
         Width           =   2055
      End
      Begin VB.ComboBox cmbBehaviour 
         Height          =   300
         ItemData        =   "frmEditor_NPC.frx":3332
         Left            =   1200
         List            =   "frmEditor_NPC.frx":3345
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   1680
         Width           =   1695
      End
      Begin VB.HScrollBar scrlRange 
         Height          =   255
         Left            =   1200
         Max             =   255
         TabIndex        =   37
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox txtAttackSay 
         Height          =   255
         Left            =   840
         TabIndex        =   36
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Spawn Rate:"
         Height          =   180
         Left            =   120
         TabIndex        =   54
         Top             =   3600
         UseMnemonic     =   0   'False
         Width           =   930
      End
      Begin VB.Label lblConv 
         Caption         =   "Conv: None"
         Height          =   255
         Left            =   120
         TabIndex        =   53
         Top             =   3000
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Sound:"
         Height          =   255
         Left            =   120
         TabIndex        =   52
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label lblAnimation 
         Caption         =   "Anim: None"
         Height          =   255
         Left            =   120
         TabIndex        =   51
         Top             =   2400
         Width           =   2775
      End
      Begin VB.Label lblSprite 
         AutoSize        =   -1  'True
         Caption         =   "Sprite: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   50
         Top             =   960
         Width           =   660
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         Caption         =   "Name:"
         Height          =   180
         Left            =   120
         TabIndex        =   49
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Behaviour:"
         Height          =   180
         Left            =   120
         TabIndex        =   48
         Top             =   1680
         UseMnemonic     =   0   'False
         Width           =   810
      End
      Begin VB.Label lblRange 
         AutoSize        =   -1  'True
         Caption         =   "Range: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   47
         Top             =   1320
         UseMnemonic     =   0   'False
         Width           =   675
      End
      Begin VB.Label lblSay 
         AutoSize        =   -1  'True
         Caption         =   "Say:"
         Height          =   180
         Left            =   120
         TabIndex        =   46
         Top             =   600
         UseMnemonic     =   0   'False
         Width           =   345
      End
   End
   Begin VB.Frame Fra7 
      Caption         =   "Vitals"
      Height          =   1815
      Left            =   6480
      TabIndex        =   26
      Top             =   3960
      Width           =   3015
      Begin VB.TextBox txtLevel 
         Height          =   285
         Left            =   960
         TabIndex        =   30
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox txtDamage 
         Height          =   285
         Left            =   960
         TabIndex        =   29
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox txtHP 
         Height          =   285
         Left            =   960
         TabIndex        =   28
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox txtEXP 
         Height          =   285
         Left            =   960
         TabIndex        =   27
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Damage:"
         Height          =   180
         Left            =   120
         TabIndex        =   34
         Top             =   600
         Width           =   675
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Level:"
         Height          =   180
         Left            =   120
         TabIndex        =   33
         Top             =   1320
         Width           =   705
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Exp:"
         Height          =   180
         Left            =   120
         TabIndex        =   32
         Top             =   960
         Width           =   585
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Health:"
         Height          =   180
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   555
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Stats"
      Height          =   1455
      Left            =   6480
      TabIndex        =   15
      Top             =   120
      Width           =   3015
      Begin VB.HScrollBar scrlStat 
         Height          =   255
         Index           =   5
         Left            =   1080
         Max             =   255
         TabIndex        =   20
         Top             =   840
         Width           =   855
      End
      Begin VB.HScrollBar scrlStat 
         Height          =   255
         Index           =   4
         Left            =   120
         Max             =   255
         TabIndex        =   19
         Top             =   840
         Width           =   855
      End
      Begin VB.HScrollBar scrlStat 
         Height          =   255
         Index           =   3
         Left            =   2040
         Max             =   255
         TabIndex        =   18
         Top             =   240
         Width           =   855
      End
      Begin VB.HScrollBar scrlStat 
         Height          =   255
         Index           =   2
         Left            =   1080
         Max             =   255
         TabIndex        =   17
         Top             =   240
         Width           =   855
      End
      Begin VB.HScrollBar scrlStat 
         Height          =   255
         Index           =   1
         Left            =   120
         Max             =   255
         TabIndex        =   16
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblStat 
         AutoSize        =   -1  'True
         Caption         =   "Will: 0"
         Height          =   180
         Index           =   5
         Left            =   1080
         TabIndex        =   25
         Top             =   1080
         Width           =   480
      End
      Begin VB.Label lblStat 
         AutoSize        =   -1  'True
         Caption         =   "Agi: 0"
         Height          =   180
         Index           =   4
         Left            =   120
         TabIndex        =   24
         Top             =   1080
         Width           =   465
      End
      Begin VB.Label lblStat 
         AutoSize        =   -1  'True
         Caption         =   "Int: 0"
         Height          =   180
         Index           =   3
         Left            =   2040
         TabIndex        =   23
         Top             =   480
         Width           =   435
      End
      Begin VB.Label lblStat 
         AutoSize        =   -1  'True
         Caption         =   "End: 0"
         Height          =   180
         Index           =   2
         Left            =   1080
         TabIndex        =   22
         Top             =   480
         Width           =   495
      End
      Begin VB.Label lblStat 
         AutoSize        =   -1  'True
         Caption         =   "Str: 0"
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   21
         Top             =   480
         Width           =   435
      End
   End
   Begin VB.Frame fraDrop 
      Caption         =   "Drop"
      Height          =   2175
      Left            =   6480
      TabIndex        =   6
      Top             =   1680
      Width           =   3015
      Begin VB.TextBox txtChance 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   960
         TabIndex        =   10
         Text            =   "0"
         Top             =   720
         Width           =   1935
      End
      Begin VB.HScrollBar scrlNum 
         Height          =   255
         Left            =   1200
         Max             =   255
         TabIndex        =   9
         Top             =   1440
         Width           =   1695
      End
      Begin VB.HScrollBar scrlValue 
         Height          =   255
         Left            =   1200
         Max             =   255
         TabIndex        =   8
         Top             =   1800
         Width           =   1695
      End
      Begin VB.HScrollBar scrlDrop 
         Height          =   255
         Left            =   120
         Min             =   1
         TabIndex        =   7
         Top             =   240
         Value           =   1
         Width           =   2775
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Chance:"
         Height          =   180
         Left            =   120
         TabIndex        =   14
         Top             =   720
         UseMnemonic     =   0   'False
         Width           =   630
      End
      Begin VB.Label lblNum 
         AutoSize        =   -1  'True
         Caption         =   "Num: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   13
         Top             =   1440
         Width           =   555
      End
      Begin VB.Label lblItemName 
         AutoSize        =   -1  'True
         Caption         =   "Item: None"
         Height          =   180
         Left            =   120
         TabIndex        =   12
         Top             =   1080
         Width           =   2775
      End
      Begin VB.Label lblValue 
         AutoSize        =   -1  'True
         Caption         =   "Value: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   11
         Top             =   1800
         UseMnemonic     =   0   'False
         Width           =   645
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000000&
         X1              =   120
         X2              =   2880
         Y1              =   600
         Y2              =   600
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   3960
      TabIndex        =   5
      Top             =   5880
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   7320
      TabIndex        =   4
      Top             =   5880
      Width           =   1455
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   5640
      TabIndex        =   3
      Top             =   5880
      Width           =   1455
   End
   Begin VB.Frame Frame3 
      Caption         =   "NPC List"
      Height          =   5655
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3135
      Begin VB.ListBox lstIndex 
         Height          =   5280
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.CommandButton cmdArray 
      Caption         =   "Change Array Size"
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   5880
      Width           =   2895
   End
End
Attribute VB_Name = "frmEditor_NPC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private DropIndex As Long
Private SpellIndex As Long

Private Sub cmbBehaviour_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Npc(EditorIndex).Behaviour = cmbBehaviour.ListIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmbBehaviour_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdDelete_Click()
Dim tmpIndex As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ClearNPC EditorIndex
    
    tmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Npc(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    NpcEditorInit
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdDelete_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_Load()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    scrlSprite.Max = Count_Char
    scrlAnimation.Max = MAX_ANIMATIONS
    scrlConv.Max = MAX_CONVS
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_Load", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlConv_Change()
    If scrlConv.value > 0 Then
        lblConv.Caption = "Conv: " & Trim$(Conv(scrlConv.value).Name)
    Else
        lblConv.Caption = "Conv: None"
    End If
    Npc(EditorIndex).Conv = scrlConv.value
End Sub

Private Sub cmdSave_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call NpcEditorOk
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdSave_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call NpcEditorCancel
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdCancel_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub lstIndex_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    NpcEditorInit
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lstIndex_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAnimation_Change()
Dim sString As String
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If scrlAnimation.value = 0 Then sString = "None" Else sString = Trim$(Animation(scrlAnimation.value).Name)
    lblAnimation.Caption = "Anim: " & sString
    Npc(EditorIndex).Animation = scrlAnimation.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAnimation_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlDrop_Change()
    DropIndex = scrlDrop.value
    fraDrop.Caption = "Drop - " & DropIndex
    txtChance.Text = Npc(EditorIndex).DropChance(DropIndex)
    scrlNum.value = Npc(EditorIndex).DropItem(DropIndex)
    scrlValue.value = Npc(EditorIndex).DropItemValue(DropIndex)
End Sub

Private Sub scrlSpell_Change()
    SpellIndex = scrlSpell.value
    fraSpell.Caption = "Spell - " & SpellIndex
    scrlSpellNum.value = Npc(EditorIndex).Spell(SpellIndex)
End Sub

Private Sub scrlSpellNum_Change()
    lblSpellNum.Caption = "Num: " & scrlSpellNum.value
    If scrlSpellNum.value > 0 Then
        lblSpellName.Caption = "Spell: " & Trim$(Spell(scrlSpellNum.value).Name)
    Else
        lblSpellName.Caption = "Spell: None"
    End If
    Npc(EditorIndex).Spell(SpellIndex) = scrlSpellNum.value
End Sub

Private Sub scrlSprite_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblSprite.Caption = "Sprite: " & scrlSprite.value
    Npc(EditorIndex).Sprite = scrlSprite.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlSprite_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlRange_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblRange.Caption = "Range: " & scrlRange.value
    Npc(EditorIndex).Range = scrlRange.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlRange_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlNum_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblNum.Caption = "Num: " & scrlNum.value

    If scrlNum.value > 0 Then
        lblItemName.Caption = "Item: " & Trim$(Item(scrlNum.value).Name)
    End If
    
    Npc(EditorIndex).DropItem(DropIndex) = scrlNum.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlNum_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlStat_Change(index As Integer)
Dim prefix As String
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Select Case index
        Case 1
            prefix = "Str: "
        Case 2
            prefix = "End: "
        Case 3
            prefix = "Int: "
        Case 4
            prefix = "Agi: "
        Case 5
            prefix = "Will: "
    End Select
    lblStat(index).Caption = prefix & scrlStat(index).value
    Npc(EditorIndex).Stat(index) = scrlStat(index).value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlStat_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlValue_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblValue.Caption = "Value: " & scrlValue.value
    Npc(EditorIndex).DropItemValue(DropIndex) = scrlValue.value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlValue_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtAttackSay_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Npc(EditorIndex).AttackSay = txtAttackSay.Text
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtAttackSay_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtChance_Validate(Cancel As Boolean)
    On Error GoTo chanceErr
    
    If DropIndex = 0 Then Exit Sub
    
    If Not IsNumeric(txtChance.Text) And Not Right$(txtChance.Text, 1) = "%" And Not InStr(1, txtChance.Text, "/") > 0 And Not InStr(1, txtChance.Text, ".") Then
        txtChance.Text = "0"
        Npc(EditorIndex).DropChance(DropIndex) = 0
        Exit Sub
    End If
    
    If Right$(txtChance.Text, 1) = "%" Then
        txtChance.Text = left(txtChance.Text, Len(txtChance.Text) - 1) / 100
    ElseIf InStr(1, txtChance.Text, "/") > 0 Then
        Dim i() As String
        i = Split(txtChance.Text, "/")
        txtChance.Text = Int(i(0) / i(1) * 1000) / 1000
    End If
    
    If txtChance.Text > 1 Or txtChance.Text < 0 Then
        Err.Description = "Value must be between 0 and 1!"
        GoTo chanceErr
    End If
    
    Npc(EditorIndex).DropChance(DropIndex) = txtChance.Text
    Exit Sub
    
chanceErr:
    txtChance.Text = "0"
    Npc(EditorIndex).DropChance(DropIndex) = 0
End Sub

Private Sub txtDamage_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If Not Len(txtDamage.Text) > 0 Then Exit Sub
    If IsNumeric(txtDamage.Text) Then Npc(EditorIndex).Damage = Val(txtDamage.Text)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtDamage_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtEXP_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If Not Len(txtEXP.Text) > 0 Then Exit Sub
    If IsNumeric(txtEXP.Text) Then Npc(EditorIndex).EXP = Val(txtEXP.Text)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtEXP_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtHP_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If Not Len(txtHP.Text) > 0 Then Exit Sub
    If IsNumeric(txtHP.Text) Then Npc(EditorIndex).HP = Val(txtHP.Text)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtHP_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtLevel_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If Not Len(txtLevel.Text) > 0 Then Exit Sub
    If IsNumeric(txtLevel.Text) Then Npc(EditorIndex).Level = Val(txtLevel.Text)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtlevel_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
Dim tmpIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Then Exit Sub
    tmpIndex = lstIndex.ListIndex
    Npc(EditorIndex).Name = Trim$(txtName.Text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Npc(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtName_Validate", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtSpawnSecs_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If Not Len(txtSpawnSecs.Text) > 0 Then Exit Sub
    Npc(EditorIndex).SpawnSecs = Val(txtSpawnSecs.Text)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtSpawnSecs_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbSound_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If cmbSound.ListIndex >= 0 Then
        Npc(EditorIndex).Sound = cmbSound.List(cmbSound.ListIndex)
    Else
        Npc(EditorIndex).Sound = "None."
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdSound_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
