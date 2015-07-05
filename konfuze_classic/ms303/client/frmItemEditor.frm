VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Begin VB.Form frmItemEditor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item Editor"
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5280
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
   ScaleHeight     =   289
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   352
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   7223
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   476
      TabMaxWidth     =   2646
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Edit"
      TabPicture(0)   =   "frmItemEditor.frx":2372
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblPic"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label5"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fraEquipment"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "tmrPic"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "picItems"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "picPic"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "scrlPic"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdCancel"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdOk"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "fraVitals"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtName"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cmbType"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "fraSpell"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).ControlCount=   14
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
         Height          =   1455
         Left            =   120
         TabIndex        =   5
         Top             =   2160
         Visible         =   0   'False
         Width           =   4815
         Begin VB.HScrollBar scrlSpell 
            Height          =   255
            Left            =   960
            Max             =   255
            TabIndex        =   6
            Top             =   960
            Value           =   1
            Width           =   3135
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
            Left            =   4200
            TabIndex        =   10
            Top             =   960
            Width           =   495
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "Add HP:"
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
            TabIndex        =   9
            Top             =   960
            Width           =   735
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
            TabIndex        =   8
            Top             =   480
            Width           =   735
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
            TabIndex        =   7
            Top             =   480
            Width           =   3735
         End
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
         ItemData        =   "frmItemEditor.frx":238E
         Left            =   120
         List            =   "frmItemEditor.frx":23BC
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   1680
         Width           =   4815
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
         Left            =   720
         TabIndex        =   20
         Top             =   480
         Width           =   4215
      End
      Begin VB.Frame fraVitals 
         Caption         =   "Vitals Data"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   120
         TabIndex        =   16
         Top             =   2160
         Visible         =   0   'False
         Width           =   4815
         Begin VB.HScrollBar scrlVitalMod 
            Height          =   255
            Left            =   960
            Max             =   255
            TabIndex        =   17
            Top             =   960
            Value           =   1
            Width           =   3135
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Vital Mod:"
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
            TabIndex        =   19
            Top             =   960
            Width           =   735
         End
         Begin VB.Label lblVitalMod 
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
            Height          =   375
            Left            =   4200
            TabIndex        =   18
            Top             =   960
            Width           =   495
         End
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
         Left            =   120
         TabIndex        =   15
         Top             =   3720
         Width           =   1095
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
         Left            =   3840
         TabIndex        =   14
         Top             =   3720
         Width           =   1095
      End
      Begin VB.HScrollBar scrlPic 
         Height          =   255
         Left            =   720
         Max             =   255
         TabIndex        =   13
         Top             =   1080
         Width           =   3135
      End
      Begin VB.PictureBox picPic 
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
         Left            =   4440
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   12
         Top             =   960
         Width           =   480
      End
      Begin VB.PictureBox picItems 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
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
         Left            =   7080
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   11
         Top             =   5640
         Width           =   480
      End
      Begin VB.Timer tmrPic 
         Interval        =   50
         Left            =   7560
         Top             =   5640
      End
      Begin VB.Frame fraEquipment 
         Caption         =   "Equipment Data"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   120
         TabIndex        =   1
         Top             =   2160
         Visible         =   0   'False
         Width           =   4815
         Begin VB.HScrollBar scrlStrength 
            Height          =   255
            Left            =   960
            Max             =   255
            TabIndex        =   2
            Top             =   960
            Value           =   1
            Width           =   3135
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Strength:"
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
            TabIndex        =   4
            Top             =   960
            Width           =   735
         End
         Begin VB.Label lblStrength 
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
            Height          =   375
            Left            =   4200
            TabIndex        =   3
            Top             =   960
            Width           =   495
         End
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
         Left            =   0
         TabIndex        =   24
         Top             =   480
         Width           =   615
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
         Left            =   0
         TabIndex        =   23
         Top             =   1080
         Width           =   615
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
         Left            =   3960
         TabIndex        =   22
         Top             =   1080
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmItemEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOk_Click()
    Call ItemEditorOk
End Sub

Private Sub cmdCancel_Click()
    Call ItemEditorCancel
End Sub

Private Sub cmbType_Click()
    If (cmbType.ListIndex >= ITEM_TYPE_WEAPON) And (cmbType.ListIndex <= ITEM_TYPE_SHIELD) Then
        fraEquipment.Visible = True
    Else
        fraEquipment.Visible = False
    End If
    
    If (cmbType.ListIndex >= ITEM_TYPE_POTIONADDHP) And (cmbType.ListIndex <= ITEM_TYPE_POTIONSUBSP) Then
        fraVitals.Visible = True
    Else
        fraVitals.Visible = False
    End If
    
    If (cmbType.ListIndex = ITEM_TYPE_SPELL) Then
        fraSpell.Visible = True
    Else
        fraSpell.Visible = False
    End If
End Sub

Private Sub scrlPic_Change()
    lblPic.Caption = STR(scrlPic.Value)
End Sub

Private Sub scrlVitalMod_Change()
    lblVitalMod.Caption = STR(scrlVitalMod.Value)
End Sub

Private Sub scrlVitalAdd_Change()
End Sub


Private Sub scrlStrength_Change()
    lblStrength.Caption = STR(scrlStrength.Value)
End Sub

Private Sub scrlSpell_Change()
    lblSpellName.Caption = Trim(Spell(scrlSpell.Value).Name)
    lblSpell.Caption = STR(scrlSpell.Value)
End Sub

Private Sub tmrPic_Timer()
    Call ItemEditorBltItem
End Sub
