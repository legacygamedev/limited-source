VERSION 5.00
Begin VB.Form frmItemEditor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item Editor"
   ClientHeight    =   7695
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   10080
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmItemEditor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7695
   ScaleWidth      =   10080
   Begin VB.ComboBox cmbBind 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmItemEditor.frx":0E42
      Left            =   3240
      List            =   "frmItemEditor.frx":0E44
      Style           =   2  'Dropdown List
      TabIndex        =   47
      Top             =   5520
      Width           =   1815
   End
   Begin VB.Frame fraSpell 
      Caption         =   "Spell Data"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      TabIndex        =   38
      Top             =   3960
      Width           =   4815
      Begin VB.HScrollBar scrlSpell 
         Height          =   255
         Left            =   960
         Max             =   1000
         Min             =   1
         TabIndex        =   39
         Top             =   720
         Value           =   1
         Width           =   3015
      End
      Begin VB.Label lblSpellName 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   43
         Top             =   360
         Width           =   3615
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Name:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   42
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Spell"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   720
         Width           =   735
      End
      Begin VB.Label lblSpell 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4080
         TabIndex        =   40
         Top             =   720
         Width           =   495
      End
   End
   Begin VB.HScrollBar scrlPic 
      Height          =   255
      Left            =   840
      Max             =   255
      TabIndex        =   37
      Top             =   840
      Width           =   3135
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   36
      Top             =   6960
      Width           =   1815
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   35
      Top             =   6600
      Width           =   1815
   End
   Begin VB.TextBox txtName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   840
      TabIndex        =   34
      Top             =   240
      Width           =   4215
   End
   Begin VB.ComboBox cmbType 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmItemEditor.frx":0E46
      Left            =   240
      List            =   "frmItemEditor.frx":0E59
      Style           =   2  'Dropdown List
      TabIndex        =   33
      Top             =   1440
      Width           =   4815
   End
   Begin VB.Frame fraStack 
      Caption         =   "Stack Data"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   28
      Top             =   3000
      Width           =   4815
      Begin VB.HScrollBar scrlStackMax 
         Height          =   255
         Left            =   960
         Max             =   9999
         TabIndex        =   30
         Top             =   600
         Value           =   1
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.CheckBox chkStack 
         Caption         =   "Stack?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblStackMax 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4080
         TabIndex        =   32
         Top             =   600
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label lblStack 
         Caption         =   "Stack Max"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   110
         TabIndex        =   31
         Top             =   610
         Visible         =   0   'False
         Width           =   855
      End
   End
   Begin VB.Frame fraEquipmentType 
      Caption         =   "Equipment Type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      TabIndex        =   25
      Top             =   1800
      Width           =   4815
      Begin VB.ComboBox cmbEquipmentType 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmItemEditor.frx":0E82
         Left            =   120
         List            =   "frmItemEditor.frx":0E84
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   360
         Width           =   4455
      End
      Begin VB.CommandButton cmdCalculate 
         Caption         =   "Calculate iLvl"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3360
         TabIndex        =   26
         Top             =   840
         Width           =   1215
      End
   End
   Begin VB.Frame frmAnim 
      Caption         =   "Weapon Animation"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   240
      TabIndex        =   20
      Top             =   5160
      Width           =   2895
      Begin VB.Timer tmrAnimation 
         Interval        =   5
         Left            =   2280
         Top             =   240
      End
      Begin VB.HScrollBar scrlAnim 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   22
         Top             =   1680
         Value           =   1
         Width           =   2655
      End
      Begin VB.PictureBox picSprite 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         CausesValidation=   0   'False
         ClipControls    =   0   'False
         Height          =   480
         Left            =   120
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   21
         Top             =   240
         Width           =   480
      End
      Begin VB.Label Label8 
         Caption         =   "Animation"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label lblAnim 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   23
         Top             =   1440
         Width           =   495
      End
   End
   Begin VB.Frame frmLevelReq 
      Caption         =   "Level Req"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5160
      TabIndex        =   16
      Top             =   120
      Width           =   4815
      Begin VB.HScrollBar scrlLevel 
         Height          =   255
         Left            =   1080
         Max             =   255
         TabIndex        =   17
         Top             =   360
         Value           =   1
         Width           =   2655
      End
      Begin VB.Label lblLevelReq 
         Alignment       =   1  'Right Justify
         Caption         =   "Level Req:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblLevel 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3960
         TabIndex        =   18
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.Frame frmClassReq 
      Caption         =   "Class Reqs."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5160
      TabIndex        =   14
      Top             =   840
      Width           =   4815
      Begin VB.CheckBox chkClass 
         Caption         =   "Class Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   4575
      End
   End
   Begin VB.Frame frmStatReq 
      Caption         =   "Stat Reqs."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5160
      TabIndex        =   10
      Top             =   1560
      Width           =   4815
      Begin VB.HScrollBar scrlStat 
         Height          =   255
         Index           =   0
         Left            =   1080
         Max             =   255
         TabIndex        =   11
         Top             =   360
         Value           =   1
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.Label lblStatName 
         Alignment       =   1  'Right Justify
         Caption         =   "Stat Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lblStat 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   3960
         TabIndex        =   12
         Top             =   360
         Visible         =   0   'False
         Width           =   375
      End
   End
   Begin VB.PictureBox Picture 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   4500
      ScaleHeight     =   14.703
      ScaleMode       =   0  'User
      ScaleWidth      =   7.884
      TabIndex        =   8
      Top             =   600
      Width           =   540
      Begin VB.PictureBox picPic 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   15
         ScaleHeight     =   16
         ScaleMode       =   0  'User
         ScaleWidth      =   10.667
         TabIndex        =   9
         Top             =   15
         Width           =   480
      End
   End
   Begin VB.Frame fraModVitals 
      Caption         =   "Mod Vitals"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5160
      TabIndex        =   4
      Top             =   2280
      Width           =   4815
      Begin VB.TextBox txtModVital 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   3840
         TabIndex        =   6
         Text            =   "1"
         Top             =   360
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.HScrollBar scrlModVital 
         Height          =   255
         Index           =   0
         Left            =   1080
         Max             =   5000
         Min             =   -5000
         TabIndex        =   5
         Top             =   360
         Value           =   1
         Visible         =   0   'False
         Width           =   2640
      End
      Begin VB.Label lblModVitalName 
         Caption         =   "Vital Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   375
         Visible         =   0   'False
         Width           =   855
      End
   End
   Begin VB.Frame fraModStat 
      Caption         =   "Mod Stats"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5160
      TabIndex        =   0
      Top             =   3000
      Width           =   4815
      Begin VB.TextBox txtModStat 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   3840
         LinkItem        =   "txtModStat"
         TabIndex        =   2
         Text            =   "1"
         Top             =   360
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.HScrollBar scrlModStat 
         Height          =   255
         Index           =   0
         Left            =   1080
         Max             =   5000
         Min             =   -5000
         TabIndex        =   1
         Top             =   360
         Value           =   1
         Visible         =   0   'False
         Width           =   2640
      End
      Begin VB.Label lblModStatName 
         Caption         =   "Stat Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Visible         =   0   'False
         Width           =   855
      End
   End
   Begin VB.Label lblItemBind 
      Caption         =   "Item Bind:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   48
      Top             =   5280
      Width           =   1815
   End
   Begin VB.Label lblPic 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4080
      TabIndex        =   46
      Top             =   840
      Width           =   390
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Picture:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   45
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   44
      Top             =   240
      Width           =   615
   End
End
Attribute VB_Name = "frmItemEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private CurrFrame As Byte
Private LastUpdate As Long

Private Sub chkStack_Click()
    
    lblStack.Visible = Not lblStack.Visible
    lblStackMax.Visible = Not lblStackMax.Visible
    scrlStackMax.Visible = Not scrlStackMax.Visible
End Sub

Private Sub cmbEquipmentType_Click()
    If cmbEquipmentType.ListIndex + 1 <> Slots.Weapon Then frmAnim.Visible = False
End Sub

Private Sub cmdCalculate_Click()
' ilvl is just a suggestion for now
' 1.5 Stat points = 1 level
' 3 Vital points = 1 level
Dim iLvl
Dim i As Long
Dim n As Long
Dim Result As Boolean
    
    Result = MsgBox("This will change the level. Do you want to continue?", vbOKCancel)
    If Result Then
        n = 0
        For i = 1 To Stats.Stat_Count
             n = n + scrlModStat(i - 1).Value
        Next
        iLvl = n \ 1.5
    
        n = 0
        For i = 1 To Vitals.Vital_Count
            n = n + scrlModVital(i - 1).Value
        Next
        iLvl = iLvl + (n \ 3)
        
        scrlLevel.Value = Clamp(iLvl, 0, scrlLevel.Max)
    End If
End Sub

Private Sub cmdOk_Click()
    Call ItemEditorOk
End Sub

Private Sub cmdCancel_Click()
    Call ItemEditorCancel
End Sub

Private Sub cmbType_Click()
    
    frmItemEditor.fraModStat.Visible = False
    frmItemEditor.fraModVitals.Visible = False
    frmItemEditor.fraSpell.Visible = False
    frmItemEditor.fraStack.Visible = False
    frmItemEditor.fraEquipmentType.Visible = False
    frmItemEditor.frmAnim.Visible = False
    
    If (cmbType.ListIndex = ITEM_TYPE_EQUIPMENT) Then
        frmItemEditor.fraModStat.Visible = True
        frmItemEditor.fraModVitals.Visible = True
        frmItemEditor.fraEquipmentType.Visible = True
        If chkStack Then
            chkStack.Value = 0
            frmItemEditor.scrlStackMax.Value = 0
        End If
        If cmbEquipmentType.ListIndex + 1 = Slots.Weapon Then frmItemEditor.frmAnim.Visible = True
    End If

    If (cmbType.ListIndex = ITEM_TYPE_POTION) Then
        frmItemEditor.fraModVitals.Visible = True
        fraStack.Visible = True
    End If

    If (cmbType.ListIndex = ITEM_TYPE_KEY) Or (cmbType.ListIndex = ITEM_TYPE_NONE) Then
        fraStack.Visible = True
    End If

    If (cmbType.ListIndex = ITEM_TYPE_SPELL) Then
        fraSpell.Visible = True
    End If
End Sub

Private Sub Form_Load()
    ItemEditorBltItem
End Sub

Private Sub scrlAnim_Change()
    lblAnim.Caption = scrlAnim.Value
End Sub

Private Sub scrlLevel_Change()
    lblLevel.Caption = scrlLevel.Value
End Sub

Private Sub scrlModVital_Change(Index As Integer)
    txtModVital(Index).Text = scrlModVital(Index).Value
End Sub

Private Sub scrlStat_Change(Index As Integer)
    lblStat(Index).Caption = scrlStat(Index).Value
End Sub
Private Sub scrlModStat_Change(Index As Integer)
    txtModStat(Index).Text = scrlModStat(Index).Value
End Sub
Private Sub scrlPic_Change()
    ItemEditorBltItem
    lblPic.Caption = Str$(scrlPic.Value)
End Sub
Private Sub scrlSpell_Change()
    lblSpellName.Caption = Trim$(Spell(scrlSpell.Value).Name)
    lblSpell.Caption = Str$(scrlSpell.Value)
End Sub
Private Sub scrlstackMAX_Change()
    lblStackMax.Caption = Str$(scrlStackMax.Value)
End Sub

Private Sub tmrAnimation_Timer()
Dim sRECT As RECT
Dim dRECT As RECT
Dim Anim As Byte, Frames As Byte, Speed As Byte, Size As Byte
    
    If Not frmItemEditor.frmAnim.Visible Then Exit Sub
    If scrlAnim.Value = 0 Then Exit Sub
    
    Anim = Animation(scrlAnim.Value).Animation
    Frames = Animation(scrlAnim.Value).AnimationFrames
    Speed = Animation(scrlAnim.Value).AnimationSpeed
    Size = Animation(scrlAnim.Value).AnimationSize
        
    If Size = 1 Then
        picSprite.Height = 480
        picSprite.Width = 480
        
         With dRECT
            .Top = 0
            .Bottom = PIC_Y
            .Left = 0
            .Right = PIC_X
        End With
    
        With sRECT
            .Top = Anim * PIC_Y
            .Bottom = .Top + PIC_Y
            .Left = CurrFrame * PIC_X
            .Right = .Left + PIC_X
        End With
        
        DD_AnimationSurf.BltToDC picSprite.hdc, sRECT, dRECT
    ElseIf Size = 2 Then
        picSprite.Height = 960
        picSprite.Width = 960
        
        With dRECT
            .Top = 0
            .Bottom = (PIC_Y * 2)
            .Left = 0
            .Right = (PIC_X * 2)
        End With
        
        With sRECT
            .Top = Anim * (PIC_Y * 2)
            .Bottom = .Top + (PIC_Y * 2)
            .Left = CurrFrame * (PIC_X * 2)
            .Right = .Left + (PIC_X * 2)
        End With
        
        DD_AnimationSurf2.BltToDC picSprite.hdc, sRECT, dRECT
    End If
    
    picSprite.Refresh

    If GetTickCount > LastUpdate Then
    
        CurrFrame = CurrFrame + 1
        If CurrFrame > Frames Then
            CurrFrame = 0
        Else
            LastUpdate = GetTickCount + Speed
        End If
    End If
End Sub

Private Sub ItemEditorBltItem()
Dim rec As RECT
Dim drec As RECT
    
    With rec
        .Top = scrlPic.Value * 32
        .Bottom = .Top + 32
        .Left = 0
        .Right = .Left + 32
    End With
    
    With drec
        .Top = 0
        .Bottom = 32
        .Left = 0
        .Right = 32
    End With
    
    DD_ItemSurf.BltToDC Picpic.hdc, rec, drec
    Picpic.Refresh
End Sub

Private Sub txtModStat_Change(Index As Integer)
    SetTextBox txtModStat(Index), scrlModStat(Index)
End Sub

Private Sub txtModVital_Change(Index As Integer)
    SetTextBox txtModVital(Index), scrlModVital(Index)
End Sub
