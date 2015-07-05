VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmEditEffect 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit Effect"
   ClientHeight    =   4335
   ClientLeft      =   105
   ClientTop       =   420
   ClientWidth     =   4455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   4455
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   7223
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Edit Effect"
      TabPicture(0)   =   "frmEditEffect.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraFortify"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fraDrain"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtName"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmbEffect"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdOK"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdCancel"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   2280
         TabIndex        =   20
         Top             =   3600
         Width           =   1455
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         Height          =   375
         Left            =   480
         TabIndex        =   19
         Top             =   3600
         Width           =   1455
      End
      Begin VB.ComboBox cmbEffect 
         Height          =   315
         ItemData        =   "frmEditEffect.frx":001C
         Left            =   840
         List            =   "frmEditEffect.frx":0029
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   920
         Width           =   3135
      End
      Begin VB.Frame Frame1 
         Caption         =   "Duration of Effect"
         Height          =   735
         Left            =   240
         TabIndex        =   3
         Top             =   1320
         Width           =   3735
         Begin VB.ComboBox cmbTime 
            Height          =   315
            ItemData        =   "frmEditEffect.frx":0045
            Left            =   120
            List            =   "frmEditEffect.frx":005E
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   320
            Width           =   3495
         End
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   840
         MaxLength       =   20
         TabIndex        =   2
         Top             =   560
         Width           =   3135
      End
      Begin VB.Frame fraDrain 
         Caption         =   "Drain"
         Height          =   1095
         Left            =   240
         TabIndex        =   7
         Top             =   2160
         Width           =   3735
         Begin VB.HScrollBar scrlDrainAmount 
            Height          =   255
            Left            =   960
            Max             =   1000
            Min             =   1
            TabIndex        =   11
            Top             =   720
            Value           =   1
            Width           =   2055
         End
         Begin VB.ComboBox cmbDrain 
            Height          =   315
            ItemData        =   "frmEditEffect.frx":00A2
            Left            =   960
            List            =   "frmEditEffect.frx":00BB
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   310
            Width           =   2655
         End
         Begin VB.Label lblDrainAmount 
            Alignment       =   1  'Right Justify
            Caption         =   "1"
            Height          =   255
            Left            =   3050
            TabIndex        =   12
            Top             =   720
            Width           =   615
         End
         Begin VB.Label Label4 
            Caption         =   "Amount:"
            Height          =   255
            Left            =   240
            TabIndex        =   10
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Stat:"
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   360
            Width           =   735
         End
      End
      Begin VB.Frame fraFortify 
         Caption         =   "Fortify"
         Height          =   1095
         Left            =   240
         TabIndex        =   13
         Top             =   2160
         Width           =   3735
         Begin VB.ComboBox cmbFortify 
            Height          =   315
            ItemData        =   "frmEditEffect.frx":00F9
            Left            =   960
            List            =   "frmEditEffect.frx":0112
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   310
            Width           =   2655
         End
         Begin VB.HScrollBar scrlFortifyAmount 
            Height          =   255
            Left            =   960
            Max             =   1000
            Min             =   1
            TabIndex        =   14
            Top             =   720
            Value           =   1
            Width           =   2055
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "Stat:"
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label6 
            Caption         =   "Amount:"
            Height          =   255
            Left            =   240
            TabIndex        =   17
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label lblFortifyAmount 
            Alignment       =   1  'Right Justify
            Caption         =   "1"
            Height          =   255
            Left            =   3050
            TabIndex        =   16
            Top             =   720
            Width           =   615
         End
      End
      Begin VB.Label Label2 
         Caption         =   "Effect:"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Name:"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmEditEffect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbEffect_Click()
    Select Case cmbEffect.ListIndex
        Case EFFECT_TYPE_DRAIN
            fraDrain.Visible = True
            fraFortify.Visible = False
        Case EFFECT_TYPE_FORTIFY
            fraDrain.Visible = False
            fraFortify.Visible = True
        Case EFFECT_TYPE_FREEZE
            fraDrain.Visible = False
            fraFortify.Visible = False
    End Select
End Sub

Private Sub cmdCancel_Click()
    Call EffectEditorCancel
End Sub

Private Sub cmdOk_Click()
    Call EffectEditorOk
End Sub

Private Sub scrlDrainAmount_Change()
    lblDrainAmount.Caption = scrlDrainAmount.Value
End Sub

Private Sub scrlFortifyAmount_Change()
    lblFortifyAmount.Caption = scrlFortifyAmount.Value
End Sub
