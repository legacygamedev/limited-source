VERSION 5.00
Begin VB.Form frmEditor_Spell 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Spell Editor"
   ClientHeight    =   8250
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10320
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEditor_Spell.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   550
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   688
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdChangeDataSize 
      Caption         =   "Change Data Size"
      Height          =   375
      Left            =   60
      TabIndex        =   68
      Top             =   7800
      Width           =   3195
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   4320
      TabIndex        =   28
      Top             =   7800
      Width           =   1455
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   7680
      TabIndex        =   30
      Top             =   7800
      Width           =   1455
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   6000
      TabIndex        =   29
      Top             =   7800
      Width           =   1455
   End
   Begin VB.Frame Frame3 
      Caption         =   "Spell List"
      Height          =   7695
      Left            =   120
      TabIndex        =   31
      Top             =   0
      Width           =   3135
      Begin VB.CommandButton cmdPaste 
         Caption         =   "Paste"
         Height          =   315
         Left            =   2400
         TabIndex        =   67
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
         TabIndex        =   66
         Top             =   240
         Width           =   615
      End
      Begin VB.ListBox lstIndex 
         Height          =   6885
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   2895
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Properties"
      Height          =   7695
      Left            =   3360
      TabIndex        =   32
      Top             =   0
      Width           =   6855
      Begin VB.Frame Frame4 
         Caption         =   "Description"
         Height          =   1095
         Left            =   120
         TabIndex        =   59
         Top             =   6000
         Width           =   3255
         Begin VB.TextBox txtDesc 
            Height          =   735
            Left            =   120
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   11
            Top             =   240
            Width           =   3015
         End
      End
      Begin VB.ComboBox cmbSound 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   7200
         Width           =   5775
      End
      Begin VB.Frame Frame6 
         Caption         =   "Data"
         Height          =   6855
         Left            =   3480
         TabIndex        =   43
         Top             =   240
         Width           =   3255
         Begin VB.HScrollBar scrlRankUp 
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   6480
            Width           =   3015
         End
         Begin VB.CheckBox chkWeaponDamage 
            Caption         =   "Weapon Damage"
            Height          =   495
            Left            =   1680
            TabIndex        =   60
            TabStop         =   0   'False
            Top             =   1440
            Width           =   1455
         End
         Begin VB.HScrollBar scrlMap 
            Height          =   255
            Left            =   120
            Max             =   100
            TabIndex        =   12
            Top             =   480
            Width           =   1455
         End
         Begin VB.HScrollBar scrlX 
            Height          =   255
            Left            =   120
            Max             =   255
            TabIndex        =   14
            Top             =   1080
            Width           =   1455
         End
         Begin VB.HScrollBar scrlY 
            Height          =   255
            Left            =   1680
            Max             =   255
            TabIndex        =   15
            Top             =   1080
            Width           =   1455
         End
         Begin VB.HScrollBar scrlDir 
            Height          =   255
            Left            =   1680
            Max             =   3
            TabIndex        =   13
            Top             =   480
            Width           =   1455
         End
         Begin VB.HScrollBar scrlVital 
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   1680
            Width           =   1455
         End
         Begin VB.HScrollBar scrlDuration 
            Height          =   255
            Left            =   120
            Max             =   60
            TabIndex        =   17
            Top             =   2280
            Width           =   1455
         End
         Begin VB.HScrollBar scrlInterval 
            Height          =   255
            Left            =   1680
            Max             =   60
            TabIndex        =   18
            Top             =   2280
            Width           =   1455
         End
         Begin VB.HScrollBar scrlRange 
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   2880
            Width           =   1455
         End
         Begin VB.HScrollBar scrlAOE 
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   3480
            Width           =   1455
         End
         Begin VB.HScrollBar scrlAnimCast 
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   4080
            Width           =   3015
         End
         Begin VB.HScrollBar scrlAnim 
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   4680
            Width           =   3015
         End
         Begin VB.HScrollBar scrlStun 
            Height          =   255
            Left            =   120
            Max             =   60
            TabIndex        =   24
            Top             =   5280
            Width           =   3015
         End
         Begin VB.HScrollBar scrlSprite 
            Height          =   255
            Left            =   1680
            TabIndex        =   20
            Top             =   2880
            Width           =   1455
         End
         Begin VB.CheckBox chkAoE 
            Caption         =   "Area of Effect"
            Height          =   375
            Left            =   1680
            TabIndex        =   44
            TabStop         =   0   'False
            Top             =   3425
            Width           =   1335
         End
         Begin VB.HScrollBar scrlCast 
            Height          =   255
            Left            =   120
            Max             =   60
            TabIndex        =   25
            Top             =   5880
            Width           =   3015
         End
         Begin VB.Label lblRankUp 
            Caption         =   "Rank Up Spell: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   61
            Top             =   6240
            Width           =   3015
         End
         Begin VB.Label lblMap 
            Caption         =   "Map: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   57
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label lblX 
            Caption         =   "X: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   56
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label lblY 
            Caption         =   "Y: 0"
            Height          =   255
            Left            =   1680
            TabIndex        =   55
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label lblDir 
            Caption         =   "Direction: Up"
            Height          =   255
            Left            =   1680
            TabIndex        =   54
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label lblVital 
            Caption         =   "Vital: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   53
            Top             =   1440
            Width           =   1455
         End
         Begin VB.Label lblDuration 
            Caption         =   "Duration: 0 s"
            Height          =   255
            Left            =   120
            TabIndex        =   52
            Top             =   2040
            Width           =   1455
         End
         Begin VB.Label lblInterval 
            Caption         =   "Interval: 0 s"
            Height          =   255
            Left            =   1680
            TabIndex        =   51
            Top             =   2040
            Width           =   1455
         End
         Begin VB.Label lblRange 
            Caption         =   "Range: Self-cast"
            Height          =   255
            Left            =   120
            TabIndex        =   50
            Top             =   2640
            Width           =   1575
         End
         Begin VB.Label lblAOE 
            Caption         =   "AoE Range: Self-cast"
            Height          =   255
            Left            =   120
            TabIndex        =   49
            Top             =   3240
            Width           =   2895
         End
         Begin VB.Label lblAnimCast 
            Caption         =   "Cast Animation: None"
            Height          =   255
            Left            =   120
            TabIndex        =   48
            Top             =   3840
            Width           =   2895
         End
         Begin VB.Label lblAnim 
            Caption         =   "Animation: None"
            Height          =   255
            Left            =   120
            TabIndex        =   47
            Top             =   4440
            Width           =   2895
         End
         Begin VB.Label lblStun 
            Caption         =   "Stun Duration: None"
            Height          =   255
            Left            =   120
            TabIndex        =   46
            Top             =   5040
            Width           =   2895
         End
         Begin VB.Label lblSprite 
            Caption         =   "Sprite: 0"
            Height          =   255
            Left            =   1680
            TabIndex        =   45
            Top             =   2640
            Width           =   1455
         End
         Begin VB.Label lblCast 
            Caption         =   "Casting Time: 0 s"
            Height          =   255
            Left            =   120
            TabIndex        =   62
            Top             =   5640
            Width           =   1935
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Information"
         Height          =   5775
         Left            =   120
         TabIndex        =   33
         Top             =   240
         Width           =   3255
         Begin VB.HScrollBar scrlCastRequired 
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   4080
            Width           =   3015
         End
         Begin VB.TextBox txtCool 
            Height          =   270
            Left            =   120
            TabIndex        =   9
            Top             =   4680
            Width           =   3015
         End
         Begin VB.HScrollBar scrlIcon 
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   5280
            Width           =   2415
         End
         Begin VB.ComboBox cmbClass 
            Height          =   300
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   3480
            Width           =   3015
         End
         Begin VB.HScrollBar scrlAccess 
            Height          =   255
            Left            =   120
            Max             =   5
            TabIndex        =   6
            Top             =   2880
            Width           =   3015
         End
         Begin VB.HScrollBar scrlLevel 
            Height          =   255
            Left            =   120
            Max             =   255
            TabIndex        =   5
            Top             =   2280
            Width           =   3015
         End
         Begin VB.HScrollBar scrlMP 
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   1680
            Width           =   3015
         End
         Begin VB.ComboBox cmbType 
            Height          =   300
            ItemData        =   "frmEditor_Spell.frx":038A
            Left            =   120
            List            =   "frmEditor_Spell.frx":03A3
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   1080
            Width           =   3015
         End
         Begin VB.TextBox txtName 
            Height          =   270
            Left            =   120
            TabIndex        =   2
            Top             =   480
            Width           =   3015
         End
         Begin VB.PictureBox picSpriteBorder 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   540
            Left            =   2600
            ScaleHeight     =   34
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   34
            TabIndex        =   42
            TabStop         =   0   'False
            Top             =   5000
            Width           =   540
            Begin VB.PictureBox Picture4 
               BackColor       =   &H00404040&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   480
               Left            =   15
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   64
               TabStop         =   0   'False
               Top             =   15
               Width           =   480
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
                  Height          =   480
                  Left            =   0
                  ScaleHeight     =   32
                  ScaleMode       =   3  'Pixel
                  ScaleWidth      =   32
                  TabIndex        =   65
                  TabStop         =   0   'False
                  Top             =   0
                  Width           =   480
               End
            End
         End
         Begin VB.Label lblCastRequired 
            Caption         =   "Cast Required: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   63
            Top             =   3840
            Width           =   3015
         End
         Begin VB.Label lblIcon 
            Caption         =   "Icon: None"
            Height          =   255
            Left            =   120
            TabIndex        =   41
            Top             =   5040
            Width           =   2415
         End
         Begin VB.Label lblCool 
            Caption         =   "Cooldown Time:"
            Height          =   255
            Left            =   120
            TabIndex        =   40
            Top             =   4440
            Width           =   1935
         End
         Begin VB.Label lblClass 
            Caption         =   "Class Required:"
            Height          =   255
            Left            =   120
            TabIndex        =   39
            Top             =   3240
            Width           =   1815
         End
         Begin VB.Label lblAccess 
            Caption         =   "Access Required: None"
            Height          =   255
            Left            =   120
            TabIndex        =   38
            Top             =   2640
            Width           =   2895
         End
         Begin VB.Label lblLevel 
            Caption         =   "Level Required: None"
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   2040
            Width           =   1815
         End
         Begin VB.Label lblMP 
            Caption         =   "Mana Cost: None"
            Height          =   255
            Left            =   120
            TabIndex        =   36
            Top             =   1440
            Width           =   1815
         End
         Begin VB.Label Label2 
            Caption         =   "Type:"
            Height          =   255
            Left            =   120
            TabIndex        =   35
            Top             =   840
            Width           =   1815
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Name:"
            Height          =   180
            Left            =   120
            TabIndex        =   34
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.Label Label4 
         Caption         =   "Sound:"
         Height          =   255
         Left            =   120
         TabIndex        =   58
         Top             =   7200
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmEditor_Spell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private TmpIndex As Long

Private Sub chkAoE_Click()
    If EditorIndex < 1 Or EditorIndex > MAX_SPELLS Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If frmEditor_Spell.chkAoE = False Then
        frmEditor_Spell.scrlAOE.Enabled = False
    Else
        frmEditor_Spell.scrlAOE.Enabled = True
    End If
    Spell(EditorIndex).IsAoe = frmEditor_Spell.chkAoE.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "chkAoE_Click", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub chkWeaponDamage_Click()
    If EditorIndex < 1 Or EditorIndex > MAX_SPELLS Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If frmEditor_Spell.chkWeaponDamage.Value = 0 Then
        Spell(EditorIndex).WeaponDamage = False
    Else
        Spell(EditorIndex).WeaponDamage = True
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "chkWeaponDamage_Click", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmbClass_Click()
    If EditorIndex < 1 Or EditorIndex > MAX_SPELLS Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Spell(EditorIndex).ClassReq = cmbClass.ListIndex
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmbClass_Click", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmbSound_Click()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If cmbSound.ListIndex >= 0 Then
        Audio.StopSounds
        Spell(EditorIndex).Sound = cmbSound.List(cmbSound.ListIndex)
        Audio.PlaySound Spell(EditorIndex).Sound, -1, -1, True
    Else
        Spell(EditorIndex).Sound = vbNullString
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdSound_Click", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmbType_Click()
    If EditorIndex < 1 Or EditorIndex > MAX_SPELLS Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If Not cmbType.ListIndex = SPELL_TYPE_WARP Then
        scrlX.Enabled = False
        scrlY.Enabled = False
        scrlMap.Enabled = False
        scrlDir.Enabled = False
    Else
        scrlX.Enabled = True
        scrlY.Enabled = True
        scrlMap.Enabled = True
        scrlDir.Enabled = True
    End If
    
    If cmbType.ListIndex = SPELL_TYPE_WARP Or cmbType.ListIndex = SPELL_TYPE_RECALL Then
        scrlRange.Value = 0
        scrlRange.Enabled = False
    Else
        scrlRange.Enabled = True
    End If
        
    Spell(EditorIndex).Type = cmbType.ListIndex
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmbType_Click", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdChangeDataSize_Click()
    Dim Res As VbMsgBoxResult, val As String
    Dim dataModified As Boolean, I As Long
    
    If EditorIndex < 1 Or EditorIndex > MAX_SPELLS Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    For I = 1 To MAX_SPELLS
        If Spell_Changed(I) And I <> EditorIndex Then
        
            dataModified = True
            Exit For
        End If
    Next
    
    If dataModified Then
        Res = MsgBox("Do you want to continue and discard the changes you made to your data?", vbYesNo)
        
        If Res = vbNo Then Exit Sub
    End If
    
    val = InputBox("Enter the amount you want the new data size to be.", "Change Data Size", MAX_SPELLS)
    
    If Not IsNumeric(val) Then
        Exit Sub
    End If
    
    Res = Abs(val)
    
    If Res = MAX_SPELLS Then Exit Sub
    
    Call SendChangeDataSize(Res, EDITOR_SPELL)
    
    Unload frmEditor_Spell
    MAX_SPELLS = Res
    ReDim Spell(MAX_SPELLS)
    
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdChangeDataSize_Click", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdDelete_Click()
    Dim TmpIndex As Long
    
    If EditorIndex < 1 Or EditorIndex > MAX_SPELLS Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    ClearSpell EditorIndex
    
    TmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Spell(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = TmpIndex
    
    SpellEditorInit
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdDelete_Click", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdSave_Click()
    If EditorIndex < 1 Or EditorIndex > MAX_SPELLS Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    EditorSave = True
    SpellEditorSave
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdSave_Click", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub Form_Load()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    frmMain.SubDaFocus Me.hWnd
    
    ' Set max values
    scrlIcon.max = NumSpellIcons
    txtName.MaxLength = NAME_LENGTH
    txtSearch.MaxLength = NAME_LENGTH
    scrlSprite.max = NumCharacters
    scrlAnimCast.max = MAX_ANIMATIONS
    scrlAnim.max = MAX_ANIMATIONS
    scrlAOE.max = MAX_BYTE
    scrlRange.max = MAX_BYTE
    scrlMap.max = MAX_MAPS
    scrlLevel.max = MAX_LEVEL
    txtDesc.MaxLength = 256
    scrlRankUp.max = MAX_SPELLS
    scrlMap.max = MAX_MAPS
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "Form_Load", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    frmMain.UnsubDaFocus Me.hWnd
    If EditorSave = False Then
        SpellEditorCancel
    Else
        EditorSave = False
    End If
    frmAdmin.chkEditor(EDITOR_SPELL).Value = False
    BringWindowToTop (frmAdmin.hWnd)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "Form_Unload", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub lstIndex_Click()
    If EditorIndex < 1 Or EditorIndex > MAX_SPELLS Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    SpellEditorInit
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "lstIndex_Click", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdClose_Click()
    If EditorIndex < 1 Or EditorIndex > MAX_SPELLS Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    frmAdmin.chkEditor(EDITOR_SPELL).FontBold = False
    frmAdmin.picEye(EDITOR_SPELL).Visible = False
    BringWindowToTop (frmAdmin.hWnd)
    Unload frmEditor_Spell
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdClose_Click", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlAccess_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_SPELLS Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If scrlAccess.Value > 0 Then
        lblAccess.Caption = "Access Required: " & scrlAccess.Value
    Else
        lblAccess.Caption = "Access Required: None"
    End If
    Spell(EditorIndex).AccessReq = scrlAccess.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlAccess_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlAnim_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_SPELLS Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If scrlAnim.Value > 0 Then
        lblAnim.Caption = "Animation: " & Trim$(Animation(scrlAnim.Value).Name)
    Else
        lblAnim.Caption = "Animation: None"
    End If
    Spell(EditorIndex).SpellAnim = scrlAnim.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlAnim_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlAnimCast_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_SPELLS Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If scrlAnimCast.Value > 0 Then
        lblAnimCast.Caption = "Cast Animation: " & Trim$(Animation(scrlAnimCast.Value).Name)
    Else
        lblAnimCast.Caption = "Cast Animation: None"
    End If
    Spell(EditorIndex).CastAnim = scrlAnimCast.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlAnimCast_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlAOE_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_SPELLS Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If scrlAOE.Value > 0 Then
        lblAOE.Caption = "AoE Range: " & scrlAOE.Value & " tiles"
    Else
        lblAOE.Caption = "AoE Range: Self-cast"
    End If
    Spell(EditorIndex).AoE = scrlAOE.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlAOE_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlCast_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_SPELLS Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lblCast.Caption = "Casting Time: " & scrlCast.Value & " s"
    Spell(EditorIndex).CastTime = scrlCast.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlCast_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlCastRequired_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_SPELLS Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lblCastRequired.Caption = "Cast Required: " & scrlCastRequired.Value
    Spell(EditorIndex).CastRequired = scrlCastRequired.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlCastRequired_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlDir_Change()
    Dim sDir As String
    
    If EditorIndex < 1 Or EditorIndex > MAX_SPELLS Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Select Case scrlDir.Value
        Case 0
            sDir = "Up"
        Case 1
            sDir = "Down"
        Case 2
            sDir = "Left"
        Case 3
            sDir = "Right"
    End Select
    lblDir.Caption = "Direction: " & sDir
    Spell(EditorIndex).Dir = scrlDir.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlDir_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlDuration_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_SPELLS Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lblDuration.Caption = "Duration: " & scrlDuration.Value & " s"
    Spell(EditorIndex).Duration = scrlDuration.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlDuration_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlIcon_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_SPELLS Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If scrlIcon.Value > 0 Then
        lblIcon.Caption = "Icon: " & scrlIcon.Value
    Else
        lblIcon.Caption = "Icon: None"
    End If
    Spell(EditorIndex).Icon = scrlIcon.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlIcon_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlInterval_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_SPELLS Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lblInterval.Caption = "Interval: " & scrlInterval.Value & " s"
    Spell(EditorIndex).Interval = scrlInterval.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlInterval_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlLevel_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_SPELLS Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If scrlLevel.Value > 0 Then
        lblLevel.Caption = "Level Required: " & scrlLevel.Value
    Else
        lblLevel.Caption = "Level Required: None"
    End If
    Spell(EditorIndex).LevelReq = scrlLevel.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlLevel_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlMap_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_SPELLS Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lblMap.Caption = "Map: " & scrlMap.Value
    Spell(EditorIndex).Map = scrlMap.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlMap_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlMP_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_SPELLS Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If scrlMP.Value > 0 Then
        lblMP.Caption = "Mana Cost: " & scrlMP.Value
    Else
        lblMP.Caption = "Mana Cost: None"
    End If
    Spell(EditorIndex).MPCost = scrlMP.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlMP_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlRange_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_SPELLS Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If scrlRange.Value > 0 Then
        lblRange.Caption = "Range: " & scrlRange.Value & " tiles"
    Else
        lblRange.Caption = "Range: Self-cast"
    End If
    Spell(EditorIndex).Range = scrlRange.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlRange_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlRankUp_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_SPELLS Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lblRankUp.Caption = "Rank Up Spell: " & scrlRankUp.Value
    Spell(EditorIndex).NewSpell = scrlRankUp.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlRankUp_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlSprite_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_SPELLS Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lblSprite.Caption = "Sprite: " & scrlSprite.Value
    Spell(EditorIndex).Sprite = scrlSprite.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlSprite_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlStun_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_SPELLS Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If scrlStun.Value > 0 Then
        lblStun.Caption = "Stun Duration: " & scrlStun.Value & " s"
    Else
        lblStun.Caption = "Stun Duration: None"
    End If
    Spell(EditorIndex).StunDuration = scrlStun.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlStun_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlVital_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_SPELLS Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lblVital.Caption = "Vital: " & scrlVital.Value
    Spell(EditorIndex).Vital = scrlVital.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlVital_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlX_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_SPELLS Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lblX.Caption = "X: " & scrlX.Value
    Spell(EditorIndex).X = scrlX.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlX_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlY_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_SPELLS Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lblY.Caption = "Y: " & scrlY.Value
    Spell(EditorIndex).Y = scrlY.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlY_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtCool_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_SPELLS Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If Not IsNumeric(txtCool.text) Then txtCool.text = 1
    If txtCool.text < 1 Then txtCool = 1
    If txtCool.text > MAX_LONG Then txtCool.text = MAX_LONG
    
    Spell(EditorIndex).CDTime = txtCool.text
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "txtCool_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtDesc_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_SPELLS Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    Spell(EditorIndex).Desc = Trim$(txtDesc.text)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "txtDesc_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
    Dim TmpIndex As Long
    
    If EditorIndex < 1 Or EditorIndex > MAX_SPELLS Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    TmpIndex = lstIndex.ListIndex
    Spell(EditorIndex).Name = Trim$(txtName.text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Spell(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = TmpIndex
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "txtName_Validate", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtName_GotFocus()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    txtName.SelStart = Len(txtName)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "txtName_GotFocus", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtDesc_GotFocus()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    txtDesc.SelStart = Len(txtDesc)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "txtDesc_GotFocus", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtCool_GotFocus()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    txtCool.SelStart = Len(txtCool)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "txtCool_GotFocus", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
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
    HandleError "txtSearch_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtSearch_GotFocus()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    txtSearch.SelStart = Len(txtSearch)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "txtSearch_GotFocus", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
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
    HandleError "Form_KeyPress", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdCopy_Click()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    TmpIndex = lstIndex.ListIndex
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdCopy_Click", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdPaste_Click()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lstIndex.RemoveItem EditorIndex - 1
    Call CopyMemory(ByVal VarPtr(Spell(EditorIndex)), ByVal VarPtr(Spell(TmpIndex + 1)), LenB(Spell(TmpIndex + 1)))
    lstIndex.AddItem EditorIndex & ": " & Trim$(Spell(EditorIndex).Name), EditorIndex - 1
    lstIndex.ListIndex = EditorIndex - 1
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdPaste_Click", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub
