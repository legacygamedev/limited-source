VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmEditStatus 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit Status"
   ClientHeight    =   4215
   ClientLeft      =   225
   ClientTop       =   480
   ClientWidth     =   3855
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   3855
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   7011
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Edit Status"
      TabPicture(0)   =   "frmEditStatus.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraFortify"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtName"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmbEffect"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "fraDrain"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "fraTime"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdOK"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdCancel"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   255
         Left            =   1920
         TabIndex        =   16
         Top             =   3600
         Width           =   1455
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   3600
         Width           =   1455
      End
      Begin VB.Frame fraTime 
         Caption         =   "Length of Effect"
         Height          =   735
         Left            =   240
         TabIndex        =   6
         Top             =   1440
         Width           =   3135
         Begin VB.OptionButton optTime 
            Caption         =   "Slow"
            Height          =   255
            Index           =   2
            Left            =   2280
            TabIndex        =   9
            Top             =   360
            Width           =   735
         End
         Begin VB.OptionButton optTime 
            Caption         =   "Medium"
            Height          =   255
            Index           =   1
            Left            =   1200
            TabIndex        =   8
            Top             =   360
            Width           =   1095
         End
         Begin VB.OptionButton optTime 
            Caption         =   "Instant"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   7
            Top             =   360
            Value           =   -1  'True
            Width           =   855
         End
      End
      Begin VB.Frame fraDrain 
         Caption         =   "Drain"
         Height          =   1095
         Left            =   240
         TabIndex        =   5
         Top             =   2280
         Visible         =   0   'False
         Width           =   3135
         Begin VB.HScrollBar scrlDrainAmount 
            Height          =   255
            Left            =   840
            Max             =   1000
            TabIndex        =   12
            Top             =   700
            Width           =   1695
         End
         Begin VB.ComboBox cmbDrainType 
            Height          =   315
            ItemData        =   "frmEditStatus.frx":001C
            Left            =   600
            List            =   "frmEditStatus.frx":0029
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   320
            Width           =   2415
         End
         Begin VB.Label lblDrainAmount 
            Alignment       =   1  'Right Justify
            Caption         =   "1"
            Height          =   255
            Left            =   2490
            TabIndex        =   14
            Top             =   720
            Width           =   615
         End
         Begin VB.Label Label4 
            Caption         =   "Amount:"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   720
            Width           =   615
         End
         Begin VB.Label Label3 
            Caption         =   "Stat:"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.ComboBox cmbEffect 
         Height          =   315
         ItemData        =   "frmEditStatus.frx":0042
         Left            =   840
         List            =   "frmEditStatus.frx":004F
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1040
         Width           =   2535
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   840
         MaxLength       =   20
         TabIndex        =   2
         Top             =   560
         Width           =   2535
      End
      Begin VB.Frame fraFortify 
         Caption         =   "Fortify"
         Height          =   1095
         Left            =   240
         TabIndex        =   17
         Top             =   2280
         Visible         =   0   'False
         Width           =   3135
         Begin VB.ComboBox cmbForifyStat 
            Height          =   315
            ItemData        =   "frmEditStatus.frx":0074
            Left            =   600
            List            =   "frmEditStatus.frx":008D
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   320
            Width           =   2415
         End
         Begin VB.HScrollBar scrlFortifyStat 
            Height          =   255
            Left            =   840
            Max             =   1000
            TabIndex        =   18
            Top             =   700
            Width           =   1695
         End
         Begin VB.Label Label7 
            Caption         =   "Stat:"
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label6 
            Caption         =   "Amount:"
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   720
            Width           =   615
         End
         Begin VB.Label lblFortifyStat 
            Alignment       =   1  'Right Justify
            Caption         =   "1"
            Height          =   255
            Left            =   2490
            TabIndex        =   20
            Top             =   720
            Width           =   615
         End
      End
      Begin VB.Label Label2 
         Caption         =   "Effect:"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1080
         Width           =   1215
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
Attribute VB_Name = "frmEditStatus"
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
    Call StatusEditorCancel
End Sub

Private Sub cmdOK_Click()
    Call StatusEditorOk
End Sub

Private Sub scrlDrainAmount_Change()
    lblDrainAmount.Caption = scrlDrainAmount.Value
End Sub

Private Sub scrlFortifyStat_Change()
    lblFortifyStat.Caption = scrlFortifyStat.Value
End Sub
