VERSION 5.00
Begin VB.Form frmEditor_Class 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Class Editor"
   ClientHeight    =   8520
   ClientLeft      =   60
   ClientTop       =   480
   ClientWidth     =   7890
   Icon            =   "frmEditor_Class.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   568
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   526
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkAnimated 
      Caption         =   "Animated"
      Height          =   255
      Left            =   5760
      TabIndex        =   58
      TabStop         =   0   'False
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton cmdChangeDataSize 
      Caption         =   "Change Data Size"
      Height          =   375
      Left            =   120
      TabIndex        =   57
      Top             =   8040
      Width           =   2535
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      TabIndex        =   15
      Top             =   8040
      Width           =   1215
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   14
      Top             =   8040
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   13
      Top             =   8040
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      Caption         =   "Class List"
      Height          =   7935
      Left            =   120
      TabIndex        =   16
      Top             =   0
      Width           =   2535
      Begin VB.TextBox txtSearch 
         CausesValidation=   0   'False
         Height          =   270
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdCopy 
         Caption         =   "Copy"
         Height          =   315
         Left            =   1080
         TabIndex        =   54
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdPaste 
         Caption         =   "Paste"
         Height          =   315
         Left            =   1800
         TabIndex        =   53
         Top             =   240
         Width           =   615
      End
      Begin VB.ListBox lstIndex 
         Height          =   7080
         Left            =   120
         TabIndex        =   1
         Top             =   660
         Width           =   2295
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Properties"
      Height          =   7935
      Left            =   2760
      TabIndex        =   17
      Top             =   0
      Width           =   5055
      Begin VB.HScrollBar scrlCombatTree 
         Height          =   255
         Left            =   3720
         Max             =   3
         Min             =   1
         TabIndex        =   55
         Top             =   5640
         Value           =   1
         Width           =   1215
      End
      Begin VB.HScrollBar scrlDir 
         Height          =   255
         Left            =   3720
         Max             =   3
         TabIndex        =   48
         Top             =   4920
         Width           =   1215
      End
      Begin VB.HScrollBar scrlY 
         Height          =   255
         Left            =   3720
         Max             =   255
         TabIndex        =   47
         Top             =   5280
         Width           =   1215
      End
      Begin VB.HScrollBar scrlX 
         Height          =   255
         Left            =   1080
         Max             =   255
         TabIndex        =   46
         Top             =   5280
         Width           =   1215
      End
      Begin VB.HScrollBar scrlMap 
         Height          =   255
         Left            =   1080
         Max             =   100
         Min             =   1
         TabIndex        =   45
         Top             =   4920
         Value           =   1
         Width           =   1215
      End
      Begin VB.ComboBox cmbColor 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         ItemData        =   "frmEditor_Class.frx":038A
         Left            =   1680
         List            =   "frmEditor_Class.frx":03C4
         Style           =   2  'Dropdown List
         TabIndex        =   44
         ToolTipText     =   "Color for login message if not a staff member."
         Top             =   3720
         Width           =   1575
      End
      Begin VB.PictureBox picFace 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   1500
         Left            =   3360
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   3240
         Width           =   1500
      End
      Begin VB.CheckBox chkSwapGender 
         Caption         =   "Swap Gender"
         Height          =   255
         Left            =   1680
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   3240
         Width           =   1455
      End
      Begin VB.CheckBox chkLocked 
         Caption         =   "Locked"
         Height          =   255
         Left            =   4040
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox chkSwapStart 
         Caption         =   "Start Spell"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   5640
         Width           =   1215
      End
      Begin VB.HScrollBar scrlStat 
         Height          =   255
         Index           =   3
         LargeChange     =   10
         Left            =   1080
         Min             =   1
         TabIndex        =   7
         Top             =   2520
         Value           =   1
         Width           =   3810
      End
      Begin VB.HScrollBar scrlStat 
         Height          =   255
         Index           =   5
         LargeChange     =   10
         Left            =   1080
         Min             =   1
         TabIndex        =   6
         Top             =   2160
         Value           =   1
         Width           =   3810
      End
      Begin VB.HScrollBar scrlStat 
         Height          =   255
         Index           =   2
         LargeChange     =   10
         Left            =   1080
         Min             =   1
         TabIndex        =   5
         Top             =   1800
         Value           =   1
         Width           =   3810
      End
      Begin VB.HScrollBar scrlStat 
         Height          =   255
         Index           =   4
         LargeChange     =   10
         Left            =   1080
         Min             =   1
         TabIndex        =   4
         Top             =   1440
         Value           =   1
         Width           =   3810
      End
      Begin VB.HScrollBar scrlStat 
         Height          =   255
         Index           =   1
         LargeChange     =   10
         Left            =   1080
         Min             =   1
         TabIndex        =   3
         Top             =   1080
         Value           =   1
         Width           =   3810
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
         Width           =   4770
      End
      Begin VB.Frame fraStartItem 
         Caption         =   "Start Item: 1"
         Height          =   1935
         Left            =   120
         TabIndex        =   28
         Top             =   5880
         Width           =   4815
         Begin VB.HScrollBar scrlStartItem 
            Height          =   255
            Left            =   120
            Min             =   1
            TabIndex        =   10
            Top             =   240
            Value           =   1
            Width           =   4575
         End
         Begin VB.HScrollBar scrlItemValue 
            Height          =   255
            Left            =   1200
            TabIndex        =   12
            Top             =   1440
            Width           =   3495
         End
         Begin VB.HScrollBar scrlItemNum 
            Height          =   255
            Left            =   1200
            Max             =   255
            TabIndex        =   11
            Top             =   1080
            Width           =   3495
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00C0C0C0&
            X1              =   0
            X2              =   4800
            Y1              =   600
            Y2              =   600
         End
         Begin VB.Label lblItemValue 
            AutoSize        =   -1  'True
            Caption         =   "Value: 0"
            Height          =   180
            Left            =   120
            TabIndex        =   31
            Top             =   1440
            UseMnemonic     =   0   'False
            Width           =   645
         End
         Begin VB.Label lblItemName 
            AutoSize        =   -1  'True
            Caption         =   "Item: None"
            Height          =   180
            Left            =   120
            TabIndex        =   30
            Top             =   720
            Width           =   855
         End
         Begin VB.Label lblItemNum 
            AutoSize        =   -1  'True
            Caption         =   "Number: 0"
            Height          =   180
            Left            =   120
            TabIndex        =   29
            Top             =   1080
            Width           =   795
         End
      End
      Begin VB.Frame fraStartSpell 
         Caption         =   "Start Spell: 1"
         Height          =   1935
         Left            =   120
         TabIndex        =   33
         Top             =   5880
         Visible         =   0   'False
         Width           =   4815
         Begin VB.HScrollBar scrlSpellNum 
            Height          =   255
            Left            =   1200
            Max             =   255
            TabIndex        =   35
            Top             =   1080
            Width           =   3495
         End
         Begin VB.HScrollBar scrlStartSpell 
            Height          =   255
            Left            =   120
            Min             =   1
            TabIndex        =   34
            Top             =   240
            Value           =   1
            Width           =   4575
         End
         Begin VB.Label lblSpellNum 
            AutoSize        =   -1  'True
            Caption         =   "Number: 0"
            Height          =   180
            Left            =   120
            TabIndex        =   37
            Top             =   1080
            Width           =   795
         End
         Begin VB.Label lblSpellName 
            AutoSize        =   -1  'True
            Caption         =   "Spell: None"
            Height          =   195
            Left            =   120
            TabIndex        =   36
            Top             =   720
            Width           =   825
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00C0C0C0&
            X1              =   0
            X2              =   4800
            Y1              =   600
            Y2              =   600
         End
      End
      Begin VB.PictureBox picSprite 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   960
         Left            =   120
         ScaleHeight     =   64
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   40
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   3240
         Width           =   600
      End
      Begin VB.HScrollBar scrlMFace 
         Height          =   255
         Left            =   1680
         TabIndex        =   9
         Top             =   4440
         Width           =   1575
      End
      Begin VB.HScrollBar scrlFFace 
         Height          =   255
         Left            =   1680
         TabIndex        =   42
         Top             =   4440
         Width           =   1575
      End
      Begin VB.HScrollBar scrlMSprite 
         Height          =   255
         Left            =   1630
         TabIndex        =   8
         Top             =   2880
         Width           =   3255
      End
      Begin VB.HScrollBar scrlFSprite 
         Height          =   255
         Left            =   1635
         TabIndex        =   24
         Top             =   2880
         Width           =   3255
      End
      Begin VB.Label lblCombatTree 
         Caption         =   "Combat Tree: Melee"
         Height          =   255
         Left            =   1920
         TabIndex        =   56
         Top             =   5640
         Width           =   1695
      End
      Begin VB.Label lblDir 
         Caption         =   "Direction: Up"
         Height          =   255
         Left            =   2400
         TabIndex        =   52
         Top             =   4920
         Width           =   1335
      End
      Begin VB.Label lblY 
         Caption         =   "Y: 0"
         Height          =   255
         Left            =   2400
         TabIndex        =   51
         Top             =   5280
         Width           =   1215
      End
      Begin VB.Label lblX 
         Caption         =   "X: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   50
         Top             =   5280
         Width           =   975
      End
      Begin VB.Label lblMap 
         Caption         =   "Map: 1"
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   4920
         Width           =   855
      End
      Begin VB.Label lblStat 
         AutoSize        =   -1  'True
         Caption         =   "Int: 1"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   23
         Top             =   2520
         UseMnemonic     =   0   'False
         Width           =   360
      End
      Begin VB.Label lblStat 
         AutoSize        =   -1  'True
         Caption         =   "Spi: 1"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   22
         Top             =   2160
         UseMnemonic     =   0   'False
         Width           =   405
      End
      Begin VB.Label lblStat 
         AutoSize        =   -1  'True
         Caption         =   "End: 1"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   21
         Top             =   1800
         UseMnemonic     =   0   'False
         Width           =   465
      End
      Begin VB.Label lblStat 
         AutoSize        =   -1  'True
         Caption         =   "Agi: 1"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   20
         Top             =   1440
         UseMnemonic     =   0   'False
         Width           =   405
      End
      Begin VB.Label lblStat 
         AutoSize        =   -1  'True
         Caption         =   "Str: 1"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   19
         Top             =   1080
         UseMnemonic     =   0   'False
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   465
      End
      Begin VB.Label lblMSprite 
         AutoSize        =   -1  'True
         Caption         =   "Male Sprite: 0"
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   2880
         UseMnemonic     =   0   'False
         Width           =   975
      End
      Begin VB.Label lblFSprite 
         AutoSize        =   -1  'True
         Caption         =   "Female Sprite: 0"
         Height          =   195
         Left            =   120
         TabIndex        =   26
         Top             =   2880
         UseMnemonic     =   0   'False
         Width           =   1140
      End
      Begin VB.Label lblMFace 
         AutoSize        =   -1  'True
         Caption         =   "Face: 0"
         Height          =   195
         Left            =   1680
         TabIndex        =   41
         Top             =   4200
         UseMnemonic     =   0   'False
         Width           =   540
      End
      Begin VB.Label lblFFace 
         AutoSize        =   -1  'True
         Caption         =   "Face: 0"
         Height          =   195
         Left            =   1680
         TabIndex        =   43
         Top             =   4200
         UseMnemonic     =   0   'False
         Width           =   540
      End
   End
End
Attribute VB_Name = "frmEditor_Class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private ItemIndex As Long
Private SpellIndex As Long
Private TmpIndex As Long

Private Sub chkAnimated_Click()
    If EditorIndex < 1 Or EditorIndex > MAX_CLASSES Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Class(EditorIndex).Animated = chkAnimated.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "chkAnimated_Click", "frmEditor_Class", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub chkLocked_Click()
    If EditorIndex < 1 Or EditorIndex > MAX_CLASSES Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Class(EditorIndex).Locked = chkLocked.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "chkLocked_Click", "frmEditor_Class", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub chkSwapGender_Click()
    If EditorIndex < 1 Or EditorIndex > MAX_CLASSES Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ' Sprites
    scrlMSprite.Visible = Not scrlMSprite.Visible
    scrlFSprite.Visible = Not scrlFSprite.Visible
    lblMSprite.Visible = Not lblMSprite.Visible
    lblFSprite.Visible = Not lblFSprite.Visible

    ' Faces
    lblMFace.Visible = Not lblMFace.Visible
    lblFFace.Visible = Not lblFFace.Visible
    scrlMFace.Visible = Not scrlMFace.Visible
    scrlFFace.Visible = Not scrlFFace.Visible
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "chkSwap_Click", "frmEditor_Class", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub chkSwapStart_Click()
    If EditorIndex < 1 Or EditorIndex > MAX_CLASSES Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    frmEditor_Class.fraStartItem.Visible = Not frmEditor_Class.fraStartItem.Visible
    frmEditor_Class.fraStartSpell.Visible = Not frmEditor_Class.fraStartSpell.Visible
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "chkSwap_Click", "frmEditor_Class", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmbColor_Click()
    If EditorIndex < 1 Or EditorIndex > MAX_CLASSES Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If cmbColor.Visible = False Then Exit Sub
    
    Class(EditorIndex).Color = cmbColor.ListIndex
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmbColor_Click", "frmEditor_Class", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdChangeDataSize_Click()
    Dim Res As VbMsgBoxResult, val As String
    Dim dataModified As Boolean, I As Long
    
    If EditorIndex < 1 Or EditorIndex > MAX_CLASSES Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    For I = 1 To MAX_CLASSES
        If Class_Changed(I) And I <> EditorIndex Then
        
            dataModified = True
            Exit For
        End If
    Next
    
    If dataModified Then
        Res = MsgBox("Do you want to continue and discard the changes you made to your data?", vbYesNo)
        
        If Res = vbNo Then Exit Sub
    End If
    
    val = InputBox("Enter the amount you want the new data size to be.", "Change Data Size", MAX_CLASSES)
    
    If Not IsNumeric(val) Then
        Exit Sub
    End If
    
    Res = Abs(val)
    
    If Res = MAX_CLASSES Then Exit Sub
    
    Call SendChangeDataSize(Res, EDITOR_CLASS)
    
    Unload frmEditor_Class
    MAX_CLASSES = Res
    ReDim Class(MAX_CLASSES)
    
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdChangeDataSize_Click", "frmEditor_Class", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdDelete_Click()
    Dim TmpIndex As Long
    
    If EditorIndex < 1 Or EditorIndex > MAX_CLASSES Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ClearClass EditorIndex
    
    TmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Class(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = TmpIndex
    
    ClassEditorInit
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdDelete_Click", "frmEditor_Class", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdSave_Click()
    If EditorIndex < 1 Or EditorIndex > MAX_CLASSES Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    EditorSave = True
    ClassEditorSave
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdSave_Click", "frmEditor_Class", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdClose_Click()
    If EditorIndex < 1 Or EditorIndex > MAX_CLASSES Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    frmAdmin.chkEditor(EDITOR_CLASS).FontBold = False
    frmAdmin.picEye(EDITOR_CLASS).Visible = False
    BringWindowToTop (frmAdmin.hWnd)
    Unload frmEditor_Class
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdClose_Click", "frmEditor_Class", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    frmMain.UnsubDaFocus Me.hWnd
    
    If EditorSave = False Then
        ClassEditorCancel
    Else
        EditorSave = False
    End If
    
    frmAdmin.chkEditor(EDITOR_CLASS).Value = False
    BringWindowToTop (frmAdmin.hWnd)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "Form_Unload", "frmEditor_Class", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub lstIndex_Click()
    If EditorIndex < 1 Or EditorIndex > MAX_CLASSES Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ClassEditorInit
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "1stIndex_Click", "frmEditor_Class", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub Form_Load()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    frmMain.SubDaFocus Me.hWnd
    txtName.MaxLength = NAME_LENGTH
    txtSearch.MaxLength = NAME_LENGTH
    scrlMSprite.max = NumCharacters
    scrlFSprite.max = NumCharacters
    scrlStartItem.max = MAX_INV
    scrlStartSpell.max = MAX_PLAYER_SPELLS
    scrlItemNum.max = MAX_ITEMS
    scrlSpellNum.max = MAX_SPELLS
    scrlMFace.max = NumFaces
    scrlFFace.max = NumFaces
    scrlMap.max = MAX_MAPS
    
    ' Resize face picture
    If NumFaces > 0 Then
        frmEditor_Class.picFace.Width = Tex_Face(1).Width * Screen.TwipsPerPixelX
        frmEditor_Class.picFace.Height = Tex_Face(1).Height * Screen.TwipsPerPixelY
    End If
    
    ' Resize the sprite pictures
    If NumCharacters > 0 Then
        frmEditor_Class.picSprite.Height = Tex_Character(1).Height / 4 * Screen.TwipsPerPixelY
        frmEditor_Class.picSprite.Width = Tex_Character(1).Width / 4 * Screen.TwipsPerPixelX
        frmEditor_Class.picSprite.Height = Tex_Character(1).Height / 4 * Screen.TwipsPerPixelY
        frmEditor_Class.picSprite.Width = Tex_Character(1).Width / 4 * Screen.TwipsPerPixelX
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "Form_Load", "frmEditor_Class", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlCombatTree_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_CLASSES Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lblCombatTree.Caption = "Combat Tree : " & GetCombatTreeName(scrlCombatTree.Value)
    Class(EditorIndex).CombatTree = scrlCombatTree.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlCombatTree_Change", "frmEditor_Class", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlDir_Change()
    Dim sDir As String
    
    If EditorIndex < 1 Or EditorIndex > MAX_CLASSES Then Exit Sub
    
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
    Class(EditorIndex).Dir = scrlDir.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlDir_Change", "frmEditor_Class", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlMap_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_CLASSES Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lblMap.Caption = "Map : " & scrlMap.Value
    Class(EditorIndex).Map = scrlMap.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlMap_Change", "frmEditor_Class", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlMFace_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_CLASSES Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If scrlMFace.Visible = False Then Exit Sub
    
    lblMFace.Caption = "Face: " & scrlMFace.Value
    Class(EditorIndex).MaleFace = scrlMFace.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlMFace_Change", "frmEditor_Class", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlFFace_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_CLASSES Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If scrlFFace.Visible = False Then Exit Sub
    
    lblFFace.Caption = "Face: " & scrlFFace.Value
    Class(EditorIndex).FemaleFace = scrlFFace.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlFFace_Change", "frmEditor_Class", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlItemNum_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_CLASSES Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lblItemNum.Caption = "Number: " & scrlItemNum.Value

    If scrlItemNum.Value > 0 Then
        lblItemName.Caption = "Item: " & Trim$(Item(scrlItemNum.Value).Name)
    Else
        lblItemName.Caption = "Item: None"
    End If
    
    Class(EditorIndex).StartItem(ItemIndex) = scrlItemNum.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlItemNum_Change", "frmEditor_Class", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlItemValue_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_CLASSES Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lblItemValue.Caption = "Value: " & scrlItemValue.Value
    Class(EditorIndex).StartItemValue(ItemIndex) = scrlItemValue.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlItemValue_Change", "frmEditor_Class", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlSpellNum_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_CLASSES Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lblSpellNum.Caption = "Number: " & scrlSpellNum.Value

    If scrlSpellNum.Value > 0 Then
        lblSpellName.Caption = "Spell: " & Trim$(Spell(scrlSpellNum.Value).Name)
    Else
        lblSpellName.Caption = "Spell: None"
    End If
    Class(EditorIndex).StartSpell(SpellIndex) = scrlSpellNum.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlSpellNum_Change", "frmEditor_Class", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlMSprite_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_CLASSES Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lblMSprite.Caption = "Male Sprite: " & scrlMSprite.Value
    Class(EditorIndex).MaleSprite = scrlMSprite.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlMSprite_Change", "frmEditor_Class", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlFSprite_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_CLASSES Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lblFSprite.Caption = "Female Sprite: " & scrlFSprite.Value
    Class(EditorIndex).FemaleSprite = scrlFSprite.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlFSprite_Change", "frmEditor_Class", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlStartItem_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_CLASSES Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ItemIndex = scrlStartItem.Value
    fraStartItem.Caption = "Start Item: " & ItemIndex
    scrlItemNum.Value = Class(EditorIndex).StartItem(ItemIndex)
    scrlItemValue.Value = Class(EditorIndex).StartItemValue(ItemIndex)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlStartItem_Change", "frmEditor_Class", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlStartSpell_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_CLASSES Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    SpellIndex = scrlStartSpell.Value
    fraStartSpell.Caption = "Start Spell: " & SpellIndex
    scrlSpellNum.Value = Class(EditorIndex).StartSpell(SpellIndex)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlStartspell_Change", "frmEditor_Class", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlStat_Change(Index As Integer)
    Dim text As String
    
    If EditorIndex < 1 Or EditorIndex > MAX_CLASSES Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Select Case Index
        Case 1
            text = "Str: "
        Case 2
            text = "End: "
        Case 3
            text = "Int: "
        Case 4
            text = "Agi: "
        Case 5
            text = "Spi: "
    End Select
    
    lblStat(Index).Caption = text & scrlStat(Index).Value
    Class(EditorIndex).Stat(Index) = scrlStat(Index).Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlStat_Change", "frmEditor_Class", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlX_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_CLASSES Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lblX.Caption = "X : " & scrlX.Value
    Class(EditorIndex).X = scrlX.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlX_Change", "frmEditor_Class", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlY_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_CLASSES Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lblY.Caption = "Y : " & scrlY.Value
    Class(EditorIndex).Y = scrlY.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlY_Change", "frmEditor_Class", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
    Dim TmpIndex As Long

    If EditorIndex < 1 Or EditorIndex > MAX_CLASSES Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    TmpIndex = lstIndex.ListIndex
    Class(EditorIndex).Name = Trim$(txtName.text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Class(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = TmpIndex
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "txtName_Validate", "frmEditor_Class", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtName_GotFocus()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    txtName.SelStart = Len(txtName)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "txtName_GotFocus", "frmEditor_Class", Err.Number, Err.Description, Err.Source, Err.HelpContext
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
    HandleError "frmEditor_Class", "frmEditor_Class", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtSearch_GotFocus()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    txtSearch.SelStart = Len(txtSearch)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "txtSearch_GotFocus", "frmEditor_Class", Err.Number, Err.Description, Err.Source, Err.HelpContext
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
    HandleError "Form_KeyPress", "frmEditor_Class", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdCopy_Click()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    TmpIndex = lstIndex.ListIndex
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdCopy_Click", "frmEditor_Class", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdPaste_Click()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
     
    lstIndex.RemoveItem EditorIndex - 1
    Call CopyMemory(ByVal VarPtr(Class(EditorIndex)), ByVal VarPtr(Class(TmpIndex + 1)), LenB(Class(TmpIndex + 1)))
    lstIndex.AddItem EditorIndex & ": " & Trim$(Class(EditorIndex).Name), EditorIndex - 1
    lstIndex.ListIndex = EditorIndex - 1
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdPaste_Click", "frmEditor_Class", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub
