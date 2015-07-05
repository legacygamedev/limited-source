VERSION 5.00
Begin VB.Form frmEditor_NPC 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NPC Editor"
   ClientHeight    =   9945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8505
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEditor_NPC.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   663
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   567
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CheckBox chkAnimated 
      Caption         =   "Animated"
      Height          =   255
      Left            =   6360
      TabIndex        =   76
      TabStop         =   0   'False
      Top             =   2880
      Width           =   1155
   End
   Begin VB.CommandButton cmdChangeDataSize 
      Caption         =   "Change Data Size"
      Height          =   375
      Left            =   120
      TabIndex        =   75
      Top             =   9480
      Width           =   3135
   End
   Begin VB.CheckBox chkShowOnDeath 
      Caption         =   "Show On Death"
      Height          =   255
      Left            =   6720
      TabIndex        =   63
      TabStop         =   0   'False
      Top             =   4740
      Width           =   1635
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
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
      Height          =   960
      Left            =   6360
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   40
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   360
      Width           =   600
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   3480
      TabIndex        =   26
      Top             =   9480
      Width           =   1455
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   6840
      TabIndex        =   28
      Top             =   9480
      Width           =   1455
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   5160
      TabIndex        =   27
      Top             =   9480
      Width           =   1455
   End
   Begin VB.Frame fraNPC 
      Caption         =   "NPC List"
      Height          =   9255
      Left            =   120
      TabIndex        =   29
      Top             =   120
      Width           =   3135
      Begin VB.CommandButton cmdPaste 
         Caption         =   "Paste"
         Height          =   315
         Left            =   2400
         TabIndex        =   62
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
         TabIndex        =   61
         Top             =   240
         Width           =   615
      End
      Begin VB.ListBox lstIndex 
         Height          =   8445
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   2895
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Properties"
      Height          =   9255
      Left            =   3360
      TabIndex        =   30
      Top             =   120
      Width           =   5055
      Begin VB.ComboBox cmbSound 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   81
         Top             =   8760
         Width           =   3855
      End
      Begin VB.HScrollBar scrlSprite 
         Height          =   255
         Left            =   1080
         Max             =   255
         TabIndex        =   79
         Top             =   2760
         Width           =   1815
      End
      Begin VB.ComboBox cmbMusic 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   77
         Top             =   8400
         Width           =   3855
      End
      Begin VB.TextBox txtMP 
         Height          =   285
         Left            =   3360
         TabIndex        =   10
         Text            =   "0"
         Top             =   3120
         Width           =   1575
      End
      Begin VB.CheckBox chkFactionThreat 
         Caption         =   "Faction Threat"
         Height          =   255
         Left            =   3000
         TabIndex        =   55
         TabStop         =   0   'False
         ToolTipText     =   "Other faction members will defend this NPC"
         Top             =   2400
         Width           =   1575
      End
      Begin VB.TextBox txtAttackSay 
         Height          =   285
         Left            =   1080
         TabIndex        =   8
         Top             =   960
         Width           =   1815
      End
      Begin VB.ComboBox cmbFaction 
         Height          =   315
         ItemData        =   "frmEditor_NPC.frx":038A
         Left            =   1080
         List            =   "frmEditor_NPC.frx":0397
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   2400
         Width           =   1815
      End
      Begin VB.TextBox txtTitle 
         Height          =   270
         Left            =   1080
         TabIndex        =   3
         Top             =   600
         Width           =   1815
      End
      Begin VB.HScrollBar scrlLevel 
         Height          =   255
         Left            =   1080
         Max             =   255
         TabIndex        =   4
         Top             =   1320
         Width           =   1815
      End
      Begin VB.HScrollBar scrlDamage 
         Height          =   255
         Left            =   3000
         TabIndex        =   13
         Top             =   3840
         Width           =   1935
      End
      Begin VB.TextBox txtSpawnSecs 
         Height          =   285
         Left            =   3720
         TabIndex        =   12
         Text            =   "0"
         ToolTipText     =   "In seconds."
         Top             =   3480
         Width           =   1215
      End
      Begin VB.HScrollBar scrlAnimation 
         Height          =   255
         Left            =   3000
         TabIndex        =   14
         Top             =   4200
         Width           =   1935
      End
      Begin VB.TextBox txtName 
         Height          =   270
         Left            =   1080
         TabIndex        =   2
         Top             =   240
         Width           =   1815
      End
      Begin VB.ComboBox cmbBehavior 
         Height          =   315
         ItemData        =   "frmEditor_NPC.frx":03AD
         Left            =   1080
         List            =   "frmEditor_NPC.frx":03BD
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   2040
         Width           =   1815
      End
      Begin VB.HScrollBar scrlRange 
         Height          =   255
         Left            =   1080
         Max             =   255
         TabIndex        =   5
         Top             =   1680
         Width           =   1815
      End
      Begin VB.TextBox txtHP 
         Height          =   285
         Left            =   1080
         TabIndex        =   9
         Text            =   "0"
         Top             =   3120
         Width           =   1815
      End
      Begin VB.TextBox txtEXP 
         Height          =   285
         Left            =   1080
         TabIndex        =   11
         Text            =   "0"
         Top             =   3480
         Width           =   1815
      End
      Begin VB.CheckBox chkSwapVisibility 
         Caption         =   "Show Spells"
         Height          =   255
         Left            =   120
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   4620
         Width           =   1455
      End
      Begin VB.Frame fraStats 
         Caption         =   "Stats"
         Height          =   1455
         Left            =   120
         TabIndex        =   31
         Top             =   4920
         Width           =   4815
         Begin VB.HScrollBar scrlStat 
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   15
            Top             =   480
            Width           =   1455
         End
         Begin VB.HScrollBar scrlStat 
            Height          =   255
            Index           =   2
            Left            =   1680
            TabIndex        =   16
            Top             =   480
            Width           =   1455
         End
         Begin VB.HScrollBar scrlStat 
            Height          =   255
            Index           =   3
            Left            =   3240
            TabIndex        =   17
            Top             =   480
            Width           =   1455
         End
         Begin VB.HScrollBar scrlStat 
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   18
            Top             =   1080
            Width           =   1455
         End
         Begin VB.HScrollBar scrlStat 
            Height          =   255
            Index           =   5
            Left            =   1680
            TabIndex        =   19
            Top             =   1080
            Width           =   1455
         End
         Begin VB.Label lblStat 
            AutoSize        =   -1  'True
            Caption         =   "Str: 0"
            Height          =   180
            Index           =   1
            Left            =   120
            TabIndex        =   36
            Top             =   240
            Width           =   1515
         End
         Begin VB.Label lblStat 
            AutoSize        =   -1  'True
            Caption         =   "End: 0"
            Height          =   180
            Index           =   2
            Left            =   1680
            TabIndex        =   35
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label lblStat 
            AutoSize        =   -1  'True
            Caption         =   "Int: 0"
            Height          =   180
            Index           =   3
            Left            =   3240
            TabIndex        =   34
            Top             =   240
            Width           =   1515
         End
         Begin VB.Label lblStat 
            AutoSize        =   -1  'True
            Caption         =   "Agi: 0"
            Height          =   180
            Index           =   4
            Left            =   120
            TabIndex        =   33
            Top             =   840
            Width           =   1425
         End
         Begin VB.Label lblStat 
            AutoSize        =   -1  'True
            Caption         =   "Spi: 0"
            Height          =   180
            Index           =   5
            Left            =   1680
            TabIndex        =   32
            Top             =   840
            Width           =   450
         End
      End
      Begin VB.Frame fraSpell 
         Caption         =   "Spell: 1"
         Height          =   1455
         Left            =   120
         TabIndex        =   57
         Top             =   4920
         Visible         =   0   'False
         Width           =   4815
         Begin VB.HScrollBar scrlSpellNum 
            Height          =   255
            Left            =   1200
            Max             =   255
            TabIndex        =   21
            Top             =   1080
            Width           =   3495
         End
         Begin VB.HScrollBar scrlSpell 
            Height          =   255
            Left            =   120
            Min             =   1
            TabIndex        =   20
            Top             =   240
            Value           =   1
            Width           =   4575
         End
         Begin VB.Label lblSpellName 
            AutoSize        =   -1  'True
            Caption         =   "Spell: None"
            Height          =   180
            Left            =   120
            TabIndex        =   59
            Top             =   720
            Width           =   870
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00C0C0C0&
            X1              =   120
            X2              =   4920
            Y1              =   600
            Y2              =   600
         End
         Begin VB.Label lblSpellNum 
            AutoSize        =   -1  'True
            Caption         =   "Number: 0"
            Height          =   180
            Left            =   120
            TabIndex        =   58
            Top             =   1080
            Width           =   795
         End
      End
      Begin VB.Frame fraOnDeath 
         Caption         =   "On Death"
         Height          =   1935
         Left            =   120
         TabIndex        =   64
         Top             =   6360
         Visible         =   0   'False
         Width           =   4815
         Begin VB.ComboBox cmbPlayerSwitch 
            Height          =   315
            ItemData        =   "frmEditor_NPC.frx":03F6
            Left            =   1560
            List            =   "frmEditor_NPC.frx":040C
            Style           =   2  'Dropdown List
            TabIndex        =   69
            Top             =   360
            Width           =   3135
         End
         Begin VB.ComboBox cmbPlayerVar 
            Height          =   315
            ItemData        =   "frmEditor_NPC.frx":0462
            Left            =   1560
            List            =   "frmEditor_NPC.frx":0478
            Style           =   2  'Dropdown List
            TabIndex        =   68
            Top             =   1080
            Width           =   3135
         End
         Begin VB.HScrollBar scrlPlayerSwitch 
            Height          =   255
            Left            =   3720
            Max             =   1
            TabIndex        =   67
            Top             =   720
            Width           =   975
         End
         Begin VB.HScrollBar scrlPlayerVar 
            Height          =   255
            Left            =   1560
            Max             =   1000
            TabIndex        =   66
            Top             =   1440
            Width           =   1455
         End
         Begin VB.CheckBox chkPlayerVar 
            Alignment       =   1  'Right Justify
            Caption         =   "Add To Var:"
            Height          =   255
            Left            =   3240
            TabIndex        =   65
            Top             =   1440
            Width           =   1455
         End
         Begin VB.Label Label9 
            Caption         =   "Set Player Switch"
            Height          =   375
            Left            =   120
            TabIndex        =   73
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label8 
            Caption         =   "Set Player Var:"
            Height          =   375
            Left            =   120
            TabIndex        =   72
            Top             =   1080
            Width           =   1455
         End
         Begin VB.Label lblPlayerSwitch 
            Caption         =   "To: False"
            Height          =   255
            Left            =   2640
            TabIndex        =   71
            Top             =   720
            Width           =   975
         End
         Begin VB.Label lblPlayerVar 
            Caption         =   "To: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   70
            Top             =   1440
            Width           =   975
         End
      End
      Begin VB.Frame fraDrop 
         Caption         =   "Drop: 1"
         Height          =   1935
         Left            =   120
         TabIndex        =   45
         Top             =   6360
         Width           =   4815
         Begin VB.CheckBox chkDrop 
            Caption         =   "Random"
            Height          =   255
            Left            =   3720
            TabIndex        =   74
            Top             =   240
            Width           =   1035
         End
         Begin VB.TextBox txtChance 
            Height          =   285
            Left            =   2880
            TabIndex        =   23
            Text            =   "0"
            ToolTipText     =   "Use 0, 1, number%, 1/number, or decimal values."
            Top             =   720
            Width           =   1815
         End
         Begin VB.HScrollBar scrlNum 
            Height          =   255
            Left            =   1200
            Max             =   255
            TabIndex        =   24
            Top             =   1080
            Width           =   3495
         End
         Begin VB.HScrollBar scrlValue 
            Height          =   255
            Left            =   1200
            TabIndex        =   25
            Top             =   1440
            Width           =   3495
         End
         Begin VB.HScrollBar scrlDrop 
            Height          =   255
            Left            =   120
            Min             =   1
            TabIndex        =   22
            Top             =   240
            Value           =   1
            Width           =   3555
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Chance:"
            Height          =   180
            Left            =   2160
            TabIndex        =   49
            Top             =   720
            UseMnemonic     =   0   'False
            Width           =   630
         End
         Begin VB.Label lblNum 
            AutoSize        =   -1  'True
            Caption         =   "Number: 0"
            Height          =   180
            Left            =   120
            TabIndex        =   48
            Top             =   1080
            Width           =   795
         End
         Begin VB.Label lblItemName 
            AutoSize        =   -1  'True
            Caption         =   "Item: None"
            Height          =   180
            Left            =   120
            TabIndex        =   47
            Top             =   720
            Width           =   855
         End
         Begin VB.Label lblValue 
            AutoSize        =   -1  'True
            Caption         =   "Value: 0"
            Height          =   180
            Left            =   120
            TabIndex        =   46
            Top             =   1440
            UseMnemonic     =   0   'False
            Width           =   645
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00C0C0C0&
            X1              =   0
            X2              =   4800
            Y1              =   600
            Y2              =   600
         End
      End
      Begin VB.Label lblSound 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sound:"
         Height          =   255
         Left            =   120
         TabIndex        =   82
         Top             =   8760
         Width           =   615
      End
      Begin VB.Label lblSprite 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sprite: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   80
         Top             =   2760
         Width           =   660
      End
      Begin VB.Label lblMusic 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Music:"
         Height          =   255
         Left            =   120
         TabIndex        =   78
         Top             =   8400
         Width           =   615
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "MP:"
         Height          =   180
         Left            =   3000
         TabIndex        =   56
         Top             =   3120
         Width           =   300
      End
      Begin VB.Label lblAttackSay 
         AutoSize        =   -1  'True
         Caption         =   "Say:"
         Height          =   180
         Left            =   120
         TabIndex        =   54
         Top             =   960
         UseMnemonic     =   0   'False
         Width           =   345
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Faction:"
         Height          =   180
         Left            =   120
         TabIndex        =   53
         Top             =   2400
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Title:"
         Height          =   180
         Left            =   120
         TabIndex        =   52
         Top             =   600
         UseMnemonic     =   0   'False
         Width           =   405
      End
      Begin VB.Label lblLevel 
         Caption         =   "Level: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   51
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label lblDamage 
         Caption         =   "Damage: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   50
         Top             =   3840
         Width           =   2775
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Spawn:"
         Height          =   180
         Left            =   3000
         TabIndex        =   44
         Top             =   3480
         UseMnemonic     =   0   'False
         Width           =   660
      End
      Begin VB.Label lblAnimation 
         Caption         =   "Animation: None"
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   4200
         Width           =   2895
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         Caption         =   "Name:"
         Height          =   180
         Left            =   120
         TabIndex        =   41
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Behavior:"
         Height          =   180
         Left            =   120
         TabIndex        =   40
         Top             =   2040
         UseMnemonic     =   0   'False
         Width           =   840
      End
      Begin VB.Label lblRange 
         AutoSize        =   -1  'True
         Caption         =   "Range: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   39
         Top             =   1680
         UseMnemonic     =   0   'False
         Width           =   675
      End
      Begin VB.Label lblExp 
         AutoSize        =   -1  'True
         Caption         =   "EXP:"
         Height          =   195
         Left            =   120
         TabIndex        =   38
         Top             =   3480
         Width           =   405
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "HP:"
         Height          =   180
         Left            =   120
         TabIndex        =   37
         Top             =   3120
         Width           =   285
      End
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
Private TmpIndex As Long

Private Sub chkAnimated_Click()
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    NPC(EditorIndex).Animated = chkAnimated.Value
    Exit Sub
    
' Error Handler
ErrorHandler:
    HandleError "chkAnimated", "frmEditor_NPC", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub chkDrop_Click()
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    NPC(EditorIndex).DropRandom(DropIndex) = chkDrop.Value
    Exit Sub
    
' Error Handler
ErrorHandler:
    HandleError "chkDrop_Click", "frmEditor_NPC", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub chkFactionThreat_Click()
    If EditorIndex < 1 Or EditorIndex > MAX_NPCS Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If chkFactionThreat.Value = 1 Then
        NPC(EditorIndex).FactionThreat = True
    Else
        NPC(EditorIndex).FactionThreat = False
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "chkFactionThreat_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub chkShowOnDeath_Click()
    If EditorIndex < 1 Or EditorIndex > MAX_NPCS Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    fraDrop.Visible = Not fraDrop.Visible
    fraOnDeath.Visible = Not fraOnDeath.Visible
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "chkShowOnDeath_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub chkSwapVisibility_Click()
    If EditorIndex < 1 Or EditorIndex > MAX_NPCS Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    fraSpell.Visible = Not fraSpell.Visible
    fraStats.Visible = Not fraStats.Visible
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "chkSwapVisibility_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmbBehavior_Click()
    If EditorIndex < 1 Or EditorIndex > MAX_NPCS Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    NPC(EditorIndex).Behavior = cmbBehavior.ListIndex
    
    If NPC(EditorIndex).Behavior = NPC_BEHAVIOR_ATTACKONSIGHT Or NPC(EditorIndex).Behavior = NPC_BEHAVIOR_GUARD Then
        frmEditor_NPC.txtAttackSay.Enabled = True
        frmEditor_NPC.lblAttackSay.Enabled = True
    Else
        frmEditor_NPC.txtAttackSay.Enabled = False
        frmEditor_NPC.lblAttackSay.Enabled = False
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmbBehavior_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmbFaction_Click()
    If EditorIndex < 1 Or EditorIndex > MAX_NPCS Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    NPC(EditorIndex).Faction = cmbFaction.ListIndex
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmbFaction_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmbMusic_Click()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If cmbMusic.ListIndex > 0 Then
        NPC(EditorIndex).Music = cmbMusic.List(cmbMusic.ListIndex)
        Audio.PlayMusic cmbMusic.List(cmbMusic.ListIndex), True
    Else
        NPC(EditorIndex).Music = vbNullString
        Audio.StopMusic
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdMusic_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmbPlayerSwitch_Click()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    NPC(EditorIndex).SwitchNum = cmbPlayerSwitch.ListIndex
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmbPlayerSwitch_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmbPlayerVar_Click()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    NPC(EditorIndex).VariableNum = cmbPlayerVar.ListIndex
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmbPlayerVar_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmbSound_Click()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If cmbSound.ListIndex >= 0 Then
        Audio.StopSounds
        NPC(EditorIndex).Sound = cmbSound.List(cmbSound.ListIndex)
        Audio.PlaySound NPC(EditorIndex).Sound, -1, -1, True
    Else
        NPC(EditorIndex).Sound = vbNullString
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdSound_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdChangeDataSize_Click()
    Dim Res As VbMsgBoxResult, val As String
    Dim dataModified As Boolean, I As Long
    
    If EditorIndex < 1 Or EditorIndex > MAX_NPCS Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    For I = 1 To MAX_NPCS
        If NPC_Changed(I) And I <> EditorIndex Then
        
            dataModified = True
            Exit For
        End If
    Next
    
    If dataModified Then
        Res = MsgBox("Do you want to continue and discard the changes you made to your data?", vbYesNo)
        
        If Res = vbNo Then Exit Sub
    End If
    
    val = InputBox("Enter the amount you want the new data size to be.", "Change Data Size", MAX_NPCS)
    
    If Not IsNumeric(val) Then
        Exit Sub
    End If
    
    Res = Abs(val)
    
    If Res = MAX_NPCS Then Exit Sub
    
    Call SendChangeDataSize(Res, EDITOR_NPC)
    
    Unload frmEditor_NPC
    MAX_NPCS = Res
    ReDim NPC(MAX_NPCS)
    
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdChangeDataSize_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdDelete_Click()
    Dim TmpIndex As Long
    
    If EditorIndex < 1 Or EditorIndex > MAX_NPCS Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ClearNPC EditorIndex
    
    TmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & NPC(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = TmpIndex
    NPCEditorInit
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdDelete_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub Form_Load()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    frmMain.SubDaFocus Me.hWnd
    scrlSprite.max = NumCharacters
    scrlAnimation.max = MAX_ANIMATIONS
    scrlDrop.max = MAX_NPC_DROPS
    scrlLevel.max = MAX_LEVEL
    txtName.MaxLength = NAME_LENGTH
    txtSearch.MaxLength = NAME_LENGTH
    txtTitle.MaxLength = NAME_LENGTH
    scrlNum.max = MAX_ITEMS
    scrlSpell.max = MAX_NPC_SPELLS
    scrlSpellNum.max = MAX_SPELLS
    
    ' Resize the sprite pictures
    If NumCharacters > 0 Then
        frmEditor_NPC.picSprite.Height = Tex_Character(1).Height / 4
        frmEditor_NPC.picSprite.Width = Tex_Character(1).Width / 4
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "Form_Load", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdSave_Click()
    If EditorIndex < 1 Or EditorIndex > MAX_NPCS Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    EditorSave = True
    Call NPCEditorSave
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdSave_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdClose_Click()
    If EditorIndex < 1 Or EditorIndex > MAX_NPCS Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    frmAdmin.chkEditor(EDITOR_NPC).FontBold = False
    frmAdmin.picEye(EDITOR_NPC).Visible = False
    BringWindowToTop (frmAdmin.hWnd)
    Unload frmEditor_NPC
    PlayMapMusic
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdClose_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    frmMain.UnsubDaFocus Me.hWnd
    If EditorSave = False Then
        Call NPCEditorCancel
    Else
        EditorSave = False
    End If
    frmAdmin.chkEditor(EDITOR_NPC).Value = False
    BringWindowToTop (frmAdmin.hWnd)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "Form_Unload", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub lstIndex_Click()
    If EditorIndex < 1 Or EditorIndex > MAX_NPCS Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    NPCEditorInit
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "lstIndex_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlAnimation_Change()
    Dim sString As String
    
    If EditorIndex < 1 Or EditorIndex > MAX_NPCS Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If scrlAnimation.Value = 0 Then sString = "None" Else sString = Trim$(Animation(scrlAnimation.Value).Name)
    lblAnimation.Caption = "Animation: " & sString
    NPC(EditorIndex).Animation = scrlAnimation.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlAnimation_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlDamage_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_NPCS Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lblDamage.Caption = "Damage: " & scrlDamage.Value
    NPC(EditorIndex).Damage = scrlDamage.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlDamage_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlLevel_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_NPCS Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lblLevel.Caption = "Level: " & scrlLevel.Value
    NPC(EditorIndex).Level = scrlLevel.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlLevel_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlPlayerSwitch_Change()
     ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If scrlPlayerSwitch.Value = 0 Then
        lblPlayerSwitch.Caption = "To: False"
    Else
        lblPlayerSwitch.Caption = "To: True"
    End If
    
    NPC(EditorIndex).SwitchVal = scrlPlayerSwitch.Value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlPlayerSwitch_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlPlayerVar_Change()
     ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    lblPlayerVar.Caption = "To: " & scrlPlayerVar.Value
    NPC(EditorIndex).VariableVal = scrlPlayerVar.Value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlPlayerVar_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlSpell_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_NPCS Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    SpellIndex = scrlSpell.Value
    fraSpell.Caption = "Spell: " & SpellIndex
    scrlSpellNum.Value = NPC(EditorIndex).Spell(SpellIndex)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlSpell_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlSpellNum_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_NPCS Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lblSpellNum.Caption = "Number: " & scrlSpellNum.Value

    If scrlSpellNum.Value > 0 Then
        lblSpellName.Caption = "Spell: " & Trim$(Spell(scrlSpellNum.Value).Name)
    Else
        lblSpellName.Caption = "Spell: None"
    End If
    NPC(EditorIndex).Spell(SpellIndex) = scrlSpellNum.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlSpellNum_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlSprite_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_NPCS Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lblSprite.Caption = "Sprite: " & scrlSprite.Value
    NPC(EditorIndex).Sprite = scrlSprite.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlSprite_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlDrop_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_NPCS Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    DropIndex = scrlDrop.Value
    fraDrop.Caption = "Drop: " & DropIndex
    txtChance.text = NPC(EditorIndex).DropChance(DropIndex)
    scrlNum.Value = NPC(EditorIndex).DropItem(DropIndex)
    scrlValue.Value = NPC(EditorIndex).DropValue(DropIndex)
    chkDrop.Value = NPC(EditorIndex).DropRandom(DropIndex)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlDrop_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlRange_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_NPCS Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lblRange.Caption = "Range: " & scrlRange.Value
    NPC(EditorIndex).Range = scrlRange.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlRange_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlNum_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_NPCS Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lblNum.Caption = "Number: " & scrlNum.Value

    If scrlNum.Value > 0 Then
        lblItemName.Caption = "Item: " & Trim$(Item(scrlNum.Value).Name)
    Else
        lblItemName.Caption = "Item: None"
    End If
    NPC(EditorIndex).DropItem(DropIndex) = scrlNum.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlNum_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlStat_Change(Index As Integer)
    Dim prefix As String
    
    If EditorIndex < 1 Or EditorIndex > MAX_NPCS Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Select Case Index
        Case 1
            prefix = "Str: "
        Case 2
            prefix = "End: "
        Case 3
            prefix = "Int: "
        Case 4
            prefix = "Agi: "
        Case 5
            prefix = "Spi: "
    End Select
    lblStat(Index).Caption = prefix & scrlStat(Index).Value
    NPC(EditorIndex).Stat(Index) = scrlStat(Index).Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlStat_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlValue_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_NPCS Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lblValue.Caption = "Value: " & scrlValue.Value
    NPC(EditorIndex).DropValue(DropIndex) = scrlValue.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlValue_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtAttackSay_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_NPCS Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    NPC(EditorIndex).AttackSay = txtAttackSay.text
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "txtAttackSay_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtExp_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_NPCS Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If Not IsNumeric(txtEXP.text) Then txtEXP.text = 0
    If txtEXP.text > MAX_LONG Then txtEXP.text = MAX_LONG
    If txtEXP.text < 0 Then txtEXP.text = 0
    NPC(EditorIndex).exp = txtEXP.text
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "txtExp_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtHP_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_NPCS Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If Not IsNumeric(txtHP.text) Then txtHP.text = 0
    If txtHP.text > MAX_LONG Then txtHP.text = MAX_LONG
    If txtHP.text < 0 Then txtHP.text = 0
    NPC(EditorIndex).HP = txtHP.text
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "txtHP_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtMP_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_NPCS Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If Not IsNumeric(txtMP.text) Then txtMP.text = 0
    If txtMP.text > MAX_LONG Then txtMP.text = MAX_LONG
    If txtMP.text < 0 Then txtMP.text = 0
    NPC(EditorIndex).MP = txtMP.text
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "txtMP_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
    Dim TmpIndex As Long
    
    If EditorIndex < 1 Or EditorIndex > MAX_NPCS Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    TmpIndex = lstIndex.ListIndex
    NPC(EditorIndex).Name = Trim$(txtName.text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & NPC(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = TmpIndex
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "txtName_Validate", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtSpawnSecs_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_NPCS Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If Not IsNumeric(txtSpawnSecs.text) Then txtSpawnSecs.text = 0
    If txtSpawnSecs.text > MAX_LONG Then txtSpawnSecs.text = MAX_LONG
    If txtSpawnSecs.text < 0 Then txtSpawnSecs.text = 0
    NPC(EditorIndex).SpawnSecs = txtSpawnSecs.text
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "txtSpawnSecs_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtChance_Validate(Cancel As Boolean)
    Dim I() As String
    
    If EditorIndex < 1 Or EditorIndex > MAX_NPCS Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Not IsNumeric(txtChance.text) And Not Right$(txtChance.text, 1) = "%" And Not InStr(1, txtChance.text, "/") > 0 And Not InStr(1, txtChance.text, ".") Then
        txtChance.text = "0"
        NPC(EditorIndex).DropChance(DropIndex) = 0
        Exit Sub
    End If
    
    If Right$(txtChance.text, 1) = "%" Then
        txtChance.text = Left$(txtChance.text, Len(txtChance.text) - 1) / 100
    ElseIf InStr(1, txtChance.text, "/") > 0 Then
        I = Split(txtChance.text, "/")
        txtChance.text = Int(I(0) / I(1) * 1000) / 1000
    End If
    
    If txtChance.text > 1 Then
        txtChance.text = "1"
    ElseIf txtChance.text < 0 Then
        txtChance.text = "0"
    End If
    
    NPC(EditorIndex).DropChance(DropIndex) = txtChance.text
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "txtChance_Validate", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtTitle_Validate(Cancel As Boolean)
    If EditorIndex < 1 Or EditorIndex > MAX_NPCS Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    NPC(EditorIndex).title = Trim$(txtTitle.text)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "txtTitle_Validate", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtName_GotFocus()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    txtName.SelStart = Len(txtName)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "txtName_GotFocus", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtTitle_GotFocus()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    txtTitle.SelStart = Len(txtTitle)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "txtTitle_GotFocus", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtAttackSay_GotFocus()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    txtAttackSay.SelStart = Len(txtAttackSay)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "txtAttackSay_GotFocus", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtHP_GotFocus()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    txtHP.SelStart = Len(txtHP)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "txtHP_GotFocus", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtMP_GotFocus()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    txtMP.SelStart = Len(txtMP)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "txtMP_GotFocus", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtSpawnSecs_GotFocus()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    txtSpawnSecs.SelStart = Len(txtSpawnSecs)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "txtSpawnSecs_GotFocus", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtEXP_GotFocus()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    txtEXP.SelStart = Len(txtEXP)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "txtEXP_GotFocus", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtChance_GotFocus()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    txtChance.SelStart = Len(txtChance)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "txtChance_GotFocus", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
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
    HandleError "txtSearch_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtSearch_GotFocus()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    txtSearch.SelStart = Len(txtSearch)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "txtSearch_GotFocus", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
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
    HandleError "Form_KeyPress", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdCopy_Click()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    TmpIndex = lstIndex.ListIndex
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdCopy_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdPaste_Click()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lstIndex.RemoveItem EditorIndex - 1
    Call CopyMemory(ByVal VarPtr(NPC(EditorIndex)), ByVal VarPtr(NPC(TmpIndex + 1)), LenB(NPC(TmpIndex + 1)))
    lstIndex.AddItem EditorIndex & ": " & Trim$(NPC(EditorIndex).Name), EditorIndex - 1
    lstIndex.ListIndex = EditorIndex - 1
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdPaste_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub
