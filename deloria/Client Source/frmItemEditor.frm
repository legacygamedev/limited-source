VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmItemEditor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item Editor"
   ClientHeight    =   5892
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   10572
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   491
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   881
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraWarp 
      Caption         =   "Scroll/Orb"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   3960
      TabIndex        =   74
      Top             =   480
      Visible         =   0   'False
      Width           =   3135
      Begin VB.TextBox txtMap 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   77
         Top             =   600
         Width           =   2895
      End
      Begin VB.HScrollBar scrlY 
         Height          =   255
         Left            =   120
         Max             =   30
         TabIndex        =   76
         Top             =   1800
         Width           =   2895
      End
      Begin VB.HScrollBar scrlX 
         Height          =   255
         Left            =   120
         Max             =   30
         TabIndex        =   75
         Top             =   1200
         Width           =   2895
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Y: 0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   120
         TabIndex        =   80
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "X: 0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   120
         TabIndex        =   79
         Top             =   960
         Width           =   240
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Map:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   120
         TabIndex        =   78
         Top             =   360
         Width           =   315
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox picSelect 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   1680
      ScaleHeight     =   40
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   40
      TabIndex        =   27
      Top             =   4440
      Width           =   480
   End
   Begin VB.Frame fraEquipment 
      Caption         =   "Equipment Data"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   3960
      TabIndex        =   2
      Top             =   480
      Visible         =   0   'False
      Width           =   3135
      Begin VB.HScrollBar scrlAccessReq 
         Height          =   255
         Left            =   120
         Max             =   4
         TabIndex        =   41
         Top             =   4200
         Width           =   2895
      End
      Begin VB.HScrollBar scrlClassReq 
         Height          =   255
         Left            =   120
         Max             =   1
         Min             =   -1
         TabIndex        =   40
         Top             =   3600
         Value           =   -1
         Width           =   2895
      End
      Begin VB.HScrollBar scrlSpeedReq 
         Height          =   255
         Left            =   120
         Max             =   500
         TabIndex        =   30
         Top             =   3000
         Width           =   2895
      End
      Begin VB.HScrollBar scrlDefReq 
         Height          =   255
         Left            =   120
         Max             =   500
         TabIndex        =   29
         Top             =   2400
         Width           =   2895
      End
      Begin VB.HScrollBar scrlStrReq 
         Height          =   255
         Left            =   120
         Max             =   500
         TabIndex        =   28
         Top             =   1800
         Width           =   2895
      End
      Begin VB.HScrollBar scrlStrength 
         Height          =   255
         Left            =   120
         Max             =   5000
         TabIndex        =   7
         Top             =   1200
         Value           =   1
         Width           =   2895
      End
      Begin VB.HScrollBar scrlDurability 
         Height          =   255
         Left            =   120
         Max             =   5000
         Min             =   -1
         TabIndex        =   5
         Top             =   600
         Value           =   -1
         Width           =   2895
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "0 - Anyone"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   43
         Top             =   3960
         Width           =   1695
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "None"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   1080
         TabIndex        =   42
         Top             =   3360
         Width           =   330
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Access Req :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   39
         Top             =   3960
         Width           =   975
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Class Req :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   38
         Top             =   3360
         Width           =   975
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Speed Req :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   37
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         TabIndex        =   35
         Top             =   2760
         Width           =   495
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         TabIndex        =   34
         Top             =   2160
         Width           =   495
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         TabIndex        =   33
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Defence Req :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   32
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Strength Req :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   31
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label lblStrength 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         TabIndex        =   8
         Top             =   960
         Width           =   495
      End
      Begin VB.Label lblDurability 
         Alignment       =   1  'Right Justify
         Caption         =   "Ind."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         TabIndex        =   6
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Damage :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Durability :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.PictureBox picPic 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2400
      Left            =   360
      ScaleHeight     =   200
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   240
      TabIndex        =   15
      Top             =   1800
      Width           =   2880
      Begin VB.PictureBox picItems 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   0
         ScaleHeight     =   40
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   240
         TabIndex        =   25
         Top             =   0
         Width           =   2880
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7320
      TabIndex        =   14
      Top             =   5280
      Width           =   1455
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8880
      TabIndex        =   13
      Top             =   5280
      Width           =   1455
   End
   Begin VB.TextBox txtName 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   360
      TabIndex        =   1
      Top             =   840
      Width           =   3135
   End
   Begin VB.ComboBox cmbType 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      ItemData        =   "frmItemEditor.frx":0000
      Left            =   360
      List            =   "frmItemEditor.frx":0046
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1200
      Width           =   3135
   End
   Begin VB.Frame fraVitals 
      Caption         =   "Vitals Data"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   3960
      TabIndex        =   9
      Top             =   480
      Visible         =   0   'False
      Width           =   3135
      Begin VB.HScrollBar scrlVitalMod 
         Height          =   255
         Left            =   240
         Max             =   255
         TabIndex        =   11
         Top             =   1080
         Value           =   1
         Width           =   2655
      End
      Begin VB.Label lblVitalMod 
         Alignment       =   1  'Right Justify
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         TabIndex        =   12
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Vital Mod :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   840
         Width           =   735
      End
   End
   Begin VB.Frame fraSpell 
      Caption         =   "Spell Data"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   3960
      TabIndex        =   16
      Top             =   480
      Visible         =   0   'False
      Width           =   3135
      Begin VB.HScrollBar scrlSpell 
         Height          =   255
         Left            =   240
         Max             =   255
         Min             =   1
         TabIndex        =   17
         Top             =   1200
         Value           =   1
         Width           =   2775
      End
      Begin VB.Label lblSpellName 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   21
         Top             =   600
         Width           =   2760
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Spell Name :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Spell Number :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label lblSpell 
         Alignment       =   1  'Right Justify
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1200
         TabIndex        =   18
         Top             =   960
         Width           =   495
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5625
      Left            =   120
      TabIndex        =   22
      Top             =   120
      Width           =   10365
      _ExtentX        =   18288
      _ExtentY        =   9927
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   397
      TabMaxWidth     =   1984
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Edit Item"
      TabPicture(0)   =   "frmItemEditor.frx":012A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label26"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Picture1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "VScroll1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "fraAttributes"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtDesc"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      Begin VB.TextBox txtDesc 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   240
         MaxLength       =   150
         TabIndex        =   69
         Top             =   5160
         Width           =   3135
      End
      Begin VB.Frame fraAttributes 
         Caption         =   "Attributes"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4695
         Left            =   7080
         TabIndex        =   44
         Top             =   360
         Width           =   3135
         Begin VB.HScrollBar scrlAddVit 
            Height          =   230
            Left            =   360
            Max             =   5000
            TabIndex        =   71
            Top             =   2880
            Width           =   2655
         End
         Begin VB.HScrollBar scrlAddEXP 
            Height          =   230
            Left            =   360
            Max             =   100
            TabIndex        =   66
            Top             =   4320
            Width           =   2655
         End
         Begin VB.HScrollBar scrlAddSP 
            Height          =   230
            Left            =   360
            Max             =   5000
            TabIndex        =   64
            Top             =   1440
            Width           =   2655
         End
         Begin VB.HScrollBar scrlAddSpeed 
            Height          =   230
            Left            =   360
            Max             =   5000
            TabIndex        =   56
            Top             =   3840
            Width           =   2655
         End
         Begin VB.HScrollBar scrlAddMagi 
            Height          =   230
            Left            =   360
            Max             =   5000
            TabIndex        =   55
            Top             =   3360
            Width           =   2655
         End
         Begin VB.HScrollBar scrlAddDef 
            Height          =   230
            Left            =   360
            Max             =   5000
            TabIndex        =   54
            Top             =   2400
            Width           =   2655
         End
         Begin VB.HScrollBar scrlAddStr 
            Height          =   230
            Left            =   360
            Max             =   5000
            TabIndex        =   53
            Top             =   1920
            Width           =   2655
         End
         Begin VB.HScrollBar scrlAddMP 
            Height          =   230
            Left            =   360
            Max             =   5000
            TabIndex        =   52
            Top             =   960
            Width           =   2655
         End
         Begin VB.HScrollBar scrlAddHP 
            Height          =   230
            Left            =   360
            Max             =   5000
            TabIndex        =   51
            Top             =   480
            Width           =   2655
         End
         Begin VB.Label lblAddVit 
            Alignment       =   1  'Right Justify
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1080
            TabIndex        =   73
            Top             =   2640
            Width           =   495
         End
         Begin VB.Label Label27 
            Alignment       =   1  'Right Justify
            Caption         =   "Add Vit :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   72
            Top             =   2640
            Width           =   735
         End
         Begin VB.Label lblAddEXP 
            Alignment       =   1  'Right Justify
            Caption         =   "0%"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1080
            TabIndex        =   68
            Top             =   4080
            Width           =   495
         End
         Begin VB.Label Label25 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Add EXP :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   67
            Top             =   4080
            Width           =   855
         End
         Begin VB.Label lblAddSP 
            Alignment       =   1  'Right Justify
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1080
            TabIndex        =   65
            Top             =   1200
            Width           =   495
         End
         Begin VB.Label Label24 
            Alignment       =   1  'Right Justify
            Caption         =   "Add SP :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   63
            Top             =   1200
            Width           =   735
         End
         Begin VB.Label lblAddSpeed 
            Alignment       =   1  'Right Justify
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1080
            TabIndex        =   62
            Top             =   3600
            Width           =   495
         End
         Begin VB.Label lblAddMagi 
            Alignment       =   1  'Right Justify
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1080
            TabIndex        =   61
            Top             =   3120
            Width           =   495
         End
         Begin VB.Label lblAddDef 
            Alignment       =   1  'Right Justify
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1080
            TabIndex        =   60
            Top             =   2160
            Width           =   495
         End
         Begin VB.Label lblAddStr 
            Alignment       =   1  'Right Justify
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1080
            TabIndex        =   59
            Top             =   1680
            Width           =   495
         End
         Begin VB.Label lblAddMP 
            Alignment       =   1  'Right Justify
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1080
            TabIndex        =   58
            Top             =   720
            Width           =   495
         End
         Begin VB.Label lblAddHP 
            Alignment       =   1  'Right Justify
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1080
            TabIndex        =   57
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label23 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Add Speed :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   50
            Top             =   3600
            Width           =   855
         End
         Begin VB.Label Label22 
            Alignment       =   1  'Right Justify
            Caption         =   "Add Magi :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   49
            Top             =   3120
            Width           =   735
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            Caption         =   "Add Def :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   48
            Top             =   2160
            Width           =   735
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            Caption         =   "Add Str :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   47
            Top             =   1680
            Width           =   735
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            Caption         =   "Add MP :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   46
            Top             =   720
            Width           =   735
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            Caption         =   "Add HP :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   45
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   2400
         Left            =   3120
         Max             =   464
         TabIndex        =   26
         Top             =   1680
         Width           =   255
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   540
         Left            =   1530
         ScaleHeight     =   516
         ScaleWidth      =   516
         TabIndex        =   36
         Top             =   4290
         Width           =   540
      End
      Begin VB.Label Label26 
         Caption         =   "Description :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   70
         Top             =   4920
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Item Name :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Item Sprite :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   1440
         Width           =   855
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
    If (cmbType.ListIndex >= ITEM_TYPE_WEAPON) And (cmbType.ListIndex <= ITEM_TYPE_SHIELD) Or (frmItemEditor.cmbType.ListIndex >= ITEM_TYPE_BOOTS) And (frmItemEditor.cmbType.ListIndex <= ITEM_TYPE_GLOVES) Then
        If cmbType.ListIndex = ITEM_TYPE_WEAPON Then
            Label3.Caption = "Damage :"
        Else
            Label3.Caption = "Defence :"
        End If
        fraEquipment.Visible = True
        fraAttributes.Visible = True
    Else
        fraEquipment.Visible = False
        fraAttributes.Visible = False
    End If
    
    If (cmbType.ListIndex >= ITEM_TYPE_POTIONADDHP) And (cmbType.ListIndex <= ITEM_TYPE_POTIONSUBSP) Then
        fraVitals.Visible = True
        fraAttributes.Visible = False
    Else
        fraVitals.Visible = False
    End If
    
    If (cmbType.ListIndex = ITEM_TYPE_SPELL) Then
        fraSpell.Visible = True
        fraAttributes.Visible = False
    Else
        fraSpell.Visible = False
    End If
    
    If (cmbType.ListIndex = ITEM_TYPE_SCROLL) Or (cmbType.ListIndex = ITEM_TYPE_ORB) Then
        fraWarp.Visible = True
        fraAttributes.Visible = False
    Else
        fraWarp.Visible = False
    End If
End Sub

Private Sub Form_Load()
    picItems.Height = 320 * PIC_Y
    Call BitBlt(picSelect.hDC, 0, 0, PIC_X, PIC_Y, picItems.hDC, EditorItemX * PIC_X, EditorItemY * PIC_Y, SRCCOPY)
End Sub

Private Sub Frame1_DragDrop(Source As Control, x As Single, y As Single)

End Sub

Private Sub picItems_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        EditorItemX = Int(x / PIC_X)
        EditorItemY = Int(y / PIC_Y)
    End If
    Call BitBlt(picSelect.hDC, 0, 0, PIC_X, PIC_Y, picItems.hDC, EditorItemX * PIC_X, EditorItemY * PIC_Y, SRCCOPY)
End Sub

Private Sub picItems_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        EditorItemX = Int(x / PIC_X)
        EditorItemY = Int(y / PIC_Y)
    End If
    Call BitBlt(picSelect.hDC, 0, 0, PIC_X, PIC_Y, picItems.hDC, EditorItemX * PIC_X, EditorItemY * PIC_Y, SRCCOPY)
End Sub

Private Sub scrlAccessReq_Change()
    With scrlAccessReq
        Select Case .Value
            Case 0
                Label17.Caption = "0 - Anyone"
            Case 1
                Label17.Caption = "1 - Moniters"
            Case 2
                Label17.Caption = "2 - Mappers"
            Case 3
                Label17.Caption = "3 - Developers"
            Case 4
                Label17.Caption = "4 - Admins"
            End Select
    End With
End Sub

Private Sub scrlAddDef_Change()
    lblAddDef.Caption = scrlAddDef.Value
End Sub

Private Sub scrlAddEXP_Change()
    lblAddEXP.Caption = scrlAddEXP.Value & "%"
End Sub

Private Sub scrlAddHP_Change()
    lblAddHP.Caption = scrlAddHP.Value
End Sub

Private Sub scrlAddMagi_Change()
    lblAddMagi.Caption = scrlAddMagi.Value
End Sub

Private Sub scrlAddMP_Change()
    lblAddMP.Caption = scrlAddMP.Value
End Sub

Private Sub scrlAddSP_Change()
    lblAddSP.Caption = scrlAddSP.Value
End Sub

Private Sub scrlAddSpeed_Change()
    lblAddSpeed.Caption = scrlAddSpeed.Value
End Sub

Private Sub scrlAddStr_Change()
    lblAddStr.Caption = scrlAddStr.Value
End Sub

Private Sub scrlAddVit_Change()
    lblAddVit.Caption = scrlAddVit.Value
End Sub

Private Sub scrlClassReq_Change()
If scrlClassReq.Value = -1 Then
    Label16.Caption = "None"
Else
    Label16.Caption = scrlClassReq.Value & " - " & Trim(Class(scrlClassReq.Value).Name)
End If
End Sub

Private Sub scrlDefReq_Change()
    Label12.Caption = scrlDefReq.Value
End Sub

Private Sub scrlSpeedReq_Change()
    Label13.Caption = scrlSpeedReq.Value
End Sub

Private Sub scrlStrReq_Change()
    Label11.Caption = scrlStrReq.Value
End Sub

Private Sub scrlVitalMod_Change()
    lblVitalMod.Caption = STR(scrlVitalMod.Value)
End Sub

Private Sub scrlDurability_Change()
    lblDurability.Caption = STR(scrlDurability.Value)
    If STR(scrlDurability.Value) = -1 Then
        lblDurability.Caption = "Ind."
    End If
End Sub

Private Sub scrlStrength_Change()
    lblStrength.Caption = STR(scrlStrength.Value)
End Sub

Private Sub scrlSpell_Change()
    lblSpellName.Caption = Trim(Spell(scrlSpell.Value).Name)
    lblSpell.Caption = STR(scrlSpell.Value)
End Sub

Private Sub scrlX_Change()
    Label29.Caption = "X: " & scrlX.Value
End Sub

Private Sub scrlY_Change()
    Label30.Caption = "Y: " & scrlY.Value
End Sub

Private Sub Timer1_Timer()
    Call BitBlt(picSelect.hDC, 0, 0, PIC_X, PIC_Y, picItems.hDC, EditorItemX * PIC_X, EditorItemY * PIC_Y, SRCCOPY)
End Sub

Private Sub VScroll1_Change()
    picItems.Top = (VScroll1.Value * PIC_Y) * -1
End Sub
