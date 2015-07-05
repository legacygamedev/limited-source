VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmTrap 
   Caption         =   "Trap"
   ClientHeight    =   4695
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5055
   LinkTopic       =   "Form1"
   ScaleHeight     =   4695
   ScaleWidth      =   5055
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   7858
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   353
      TabMaxWidth     =   1940
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Set Damage"
      TabPicture(0)   =   "frmTrap.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label5"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label6"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label7"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label8"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label9"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label10"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label11"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cmdCancel"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cmdOk"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Line3"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "HScroll1"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "HScroll2"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "HScroll3"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "HScroll4"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "HScroll5"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).ControlCount=   19
      Begin VB.HScrollBar HScroll5 
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   3000
         Width           =   3855
      End
      Begin VB.HScrollBar HScroll4 
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   2400
         Width           =   3855
      End
      Begin VB.HScrollBar HScroll3 
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   1800
         Width           =   3855
      End
      Begin VB.HScrollBar HScroll2 
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1200
         Width           =   3855
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Width           =   3855
      End
      Begin VB.TextBox Line3 
         Height          =   285
         Left            =   240
         MaxLength       =   100
         TabIndex        =   3
         Top             =   3600
         Width           =   4335
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "Ok"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   3960
         Width           =   1935
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   1
         Top             =   3960
         Width           =   1935
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Text to say when Damage is caused:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   240
         TabIndex        =   19
         Top             =   3360
         Width           =   2310
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   4200
         TabIndex        =   18
         Top             =   3030
         Width           =   315
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Remove Amount of Levels:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   240
         TabIndex        =   16
         Top             =   2760
         Width           =   1710
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   4200
         TabIndex        =   15
         Top             =   2430
         Width           =   315
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Remove Amount of Experience:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   240
         TabIndex        =   13
         Top             =   2160
         Width           =   1980
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   4200
         TabIndex        =   12
         Top             =   1830
         Width           =   315
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Remove Amount of Stamina:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   240
         TabIndex        =   10
         Top             =   1560
         Width           =   1830
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   4200
         TabIndex        =   9
         Top             =   1230
         Width           =   315
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Remove Amount of Mana:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   240
         TabIndex        =   7
         Top             =   960
         Width           =   1665
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   4200
         TabIndex        =   6
         Top             =   630
         Width           =   315
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Remove Amount of Health:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   1725
      End
   End
End
Attribute VB_Name = "frmTrap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
HealthDamage = HScroll1.Value
ManaDamage = HScroll2.Value
StaminaDamage = HScroll3.Value
ExpDamage = HScroll4.Value
LevelDamage = HScroll5.Value
TextDamage = Line3.Text
    Unload Me
End Sub

Private Sub HScroll1_Change()
Label3.Caption = HScroll1.Value
End Sub

Private Sub HScroll2_Change()
Label5.Caption = HScroll2.Value
End Sub

Private Sub HScroll3_Change()
Label6.Caption = HScroll3.Value
End Sub

Private Sub HScroll4_Change()
Label8.Caption = HScroll4.Value
End Sub

Private Sub HScroll5_Change()
Label10.Caption = HScroll5.Value
End Sub
