VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmMapProperties 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Map Properties"
   ClientHeight    =   7005
   ClientLeft      =   90
   ClientTop       =   390
   ClientWidth     =   7755
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
   Icon            =   "frmMapProperties.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   467
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   517
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   6765
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   11933
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   344
      TabMaxWidth     =   1773
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "frmMapProperties.frx":0FC2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblMapName"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "txtMapName"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraSwitch"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fraSettings"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "fraRespawn"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "fraBGM"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdOk"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdCancel"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "NPC"
      TabPicture(1)   =   "frmMapProperties.frx":0FDE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdSetRand"
      Tab(1).Control(1)=   "cmbNpcX(0)"
      Tab(1).Control(2)=   "cmbNpcY(14)"
      Tab(1).Control(3)=   "cmbNpcX(14)"
      Tab(1).Control(4)=   "cmbNpcY(13)"
      Tab(1).Control(5)=   "cmbNpcX(13)"
      Tab(1).Control(6)=   "cmbNpcY(12)"
      Tab(1).Control(7)=   "cmbNpcX(12)"
      Tab(1).Control(8)=   "cmbNpcY(11)"
      Tab(1).Control(9)=   "cmbNpcX(11)"
      Tab(1).Control(10)=   "cmbNpcY(10)"
      Tab(1).Control(11)=   "cmbNpcX(10)"
      Tab(1).Control(12)=   "cmbNpcY(9)"
      Tab(1).Control(13)=   "cmbNpcX(9)"
      Tab(1).Control(14)=   "cmbNpcY(8)"
      Tab(1).Control(15)=   "cmbNpcX(8)"
      Tab(1).Control(16)=   "cmbNpcY(7)"
      Tab(1).Control(17)=   "cmbNpcX(7)"
      Tab(1).Control(18)=   "cmbNpcY(6)"
      Tab(1).Control(19)=   "cmbNpcX(6)"
      Tab(1).Control(20)=   "cmbNpcY(5)"
      Tab(1).Control(21)=   "cmbNpcX(5)"
      Tab(1).Control(22)=   "cmbNpcY(4)"
      Tab(1).Control(23)=   "cmbNpcX(4)"
      Tab(1).Control(24)=   "cmbNpcY(3)"
      Tab(1).Control(25)=   "cmbNpcX(3)"
      Tab(1).Control(26)=   "cmbNpcY(2)"
      Tab(1).Control(27)=   "cmbNpcX(2)"
      Tab(1).Control(28)=   "cmbNpcY(1)"
      Tab(1).Control(29)=   "cmbNpcX(1)"
      Tab(1).Control(30)=   "cmbNpcY(0)"
      Tab(1).Control(31)=   "cmdCopy(9)"
      Tab(1).Control(32)=   "cmdCopy(13)"
      Tab(1).Control(33)=   "cmdCopy(12)"
      Tab(1).Control(34)=   "cmdCopy(11)"
      Tab(1).Control(35)=   "cmdCopy(10)"
      Tab(1).Control(36)=   "cmdCopy(8)"
      Tab(1).Control(37)=   "cmdCopy(7)"
      Tab(1).Control(38)=   "cmdCopy(6)"
      Tab(1).Control(39)=   "cmdCopy(5)"
      Tab(1).Control(40)=   "cmdCopy(4)"
      Tab(1).Control(41)=   "cmdCopy(3)"
      Tab(1).Control(42)=   "cmdCopy(2)"
      Tab(1).Control(43)=   "cmdCopy(1)"
      Tab(1).Control(44)=   "cmdCopy(0)"
      Tab(1).Control(45)=   "cmdClear"
      Tab(1).Control(46)=   "cmbNpc(0)"
      Tab(1).Control(47)=   "cmbNpc(1)"
      Tab(1).Control(48)=   "cmbNpc(2)"
      Tab(1).Control(49)=   "cmbNpc(3)"
      Tab(1).Control(50)=   "cmbNpc(4)"
      Tab(1).Control(51)=   "cmbNpc(5)"
      Tab(1).Control(52)=   "cmbNpc(6)"
      Tab(1).Control(53)=   "cmbNpc(7)"
      Tab(1).Control(54)=   "cmbNpc(8)"
      Tab(1).Control(55)=   "cmbNpc(9)"
      Tab(1).Control(56)=   "cmbNpc(10)"
      Tab(1).Control(57)=   "cmbNpc(11)"
      Tab(1).Control(58)=   "cmbNpc(12)"
      Tab(1).Control(59)=   "cmbNpc(13)"
      Tab(1).Control(60)=   "cmbNpc(14)"
      Tab(1).Control(61)=   "lblCoordY"
      Tab(1).Control(62)=   "lblMonster"
      Tab(1).Control(63)=   "lblCoordX"
      Tab(1).ControlCount=   64
      Begin VB.CommandButton cmdSetRand 
         Caption         =   "Reset All Coords"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   -74880
         TabIndex        =   97
         Top             =   6240
         Width           =   1575
      End
      Begin VB.ComboBox cmbNpcX 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   -70560
         Style           =   2  'Dropdown List
         TabIndex        =   66
         Top             =   720
         Width           =   735
      End
      Begin VB.ComboBox cmbNpcY 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   14
         Left            =   -69600
         Style           =   2  'Dropdown List
         TabIndex        =   96
         Top             =   5760
         Width           =   735
      End
      Begin VB.ComboBox cmbNpcX 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   14
         Left            =   -70560
         Style           =   2  'Dropdown List
         TabIndex        =   95
         Top             =   5760
         Width           =   735
      End
      Begin VB.ComboBox cmbNpcY 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   13
         Left            =   -69600
         Style           =   2  'Dropdown List
         TabIndex        =   94
         Top             =   5400
         Width           =   735
      End
      Begin VB.ComboBox cmbNpcX 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   13
         Left            =   -70560
         Style           =   2  'Dropdown List
         TabIndex        =   93
         Top             =   5400
         Width           =   735
      End
      Begin VB.ComboBox cmbNpcY 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   12
         Left            =   -69600
         Style           =   2  'Dropdown List
         TabIndex        =   92
         Top             =   5040
         Width           =   735
      End
      Begin VB.ComboBox cmbNpcX 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   12
         Left            =   -70560
         Style           =   2  'Dropdown List
         TabIndex        =   91
         Top             =   5040
         Width           =   735
      End
      Begin VB.ComboBox cmbNpcY 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   11
         Left            =   -69600
         Style           =   2  'Dropdown List
         TabIndex        =   90
         Top             =   4680
         Width           =   735
      End
      Begin VB.ComboBox cmbNpcX 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   11
         Left            =   -70560
         Style           =   2  'Dropdown List
         TabIndex        =   89
         Top             =   4680
         Width           =   735
      End
      Begin VB.ComboBox cmbNpcY 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   10
         Left            =   -69600
         Style           =   2  'Dropdown List
         TabIndex        =   88
         Top             =   4320
         Width           =   735
      End
      Begin VB.ComboBox cmbNpcX 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   10
         Left            =   -70560
         Style           =   2  'Dropdown List
         TabIndex        =   87
         Top             =   4320
         Width           =   735
      End
      Begin VB.ComboBox cmbNpcY 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   9
         Left            =   -69600
         Style           =   2  'Dropdown List
         TabIndex        =   86
         Top             =   3960
         Width           =   735
      End
      Begin VB.ComboBox cmbNpcX 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   9
         Left            =   -70560
         Style           =   2  'Dropdown List
         TabIndex        =   85
         Top             =   3960
         Width           =   735
      End
      Begin VB.ComboBox cmbNpcY 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   8
         Left            =   -69600
         Style           =   2  'Dropdown List
         TabIndex        =   84
         Top             =   3600
         Width           =   735
      End
      Begin VB.ComboBox cmbNpcX 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   8
         Left            =   -70560
         Style           =   2  'Dropdown List
         TabIndex        =   83
         Top             =   3600
         Width           =   735
      End
      Begin VB.ComboBox cmbNpcY 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   7
         Left            =   -69600
         Style           =   2  'Dropdown List
         TabIndex        =   82
         Top             =   3240
         Width           =   735
      End
      Begin VB.ComboBox cmbNpcX 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   7
         Left            =   -70560
         Style           =   2  'Dropdown List
         TabIndex        =   81
         Top             =   3240
         Width           =   735
      End
      Begin VB.ComboBox cmbNpcY 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   6
         Left            =   -69600
         Style           =   2  'Dropdown List
         TabIndex        =   80
         Top             =   2880
         Width           =   735
      End
      Begin VB.ComboBox cmbNpcX 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   6
         Left            =   -70560
         Style           =   2  'Dropdown List
         TabIndex        =   79
         Top             =   2880
         Width           =   735
      End
      Begin VB.ComboBox cmbNpcY 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   5
         Left            =   -69600
         Style           =   2  'Dropdown List
         TabIndex        =   78
         Top             =   2520
         Width           =   735
      End
      Begin VB.ComboBox cmbNpcX 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   5
         Left            =   -70560
         Style           =   2  'Dropdown List
         TabIndex        =   77
         Top             =   2520
         Width           =   735
      End
      Begin VB.ComboBox cmbNpcY 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   -69600
         Style           =   2  'Dropdown List
         TabIndex        =   76
         Top             =   2160
         Width           =   735
      End
      Begin VB.ComboBox cmbNpcX 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   -70560
         Style           =   2  'Dropdown List
         TabIndex        =   75
         Top             =   2160
         Width           =   735
      End
      Begin VB.ComboBox cmbNpcY 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   -69600
         Style           =   2  'Dropdown List
         TabIndex        =   74
         Top             =   1800
         Width           =   735
      End
      Begin VB.ComboBox cmbNpcX 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   -70560
         Style           =   2  'Dropdown List
         TabIndex        =   73
         Top             =   1800
         Width           =   735
      End
      Begin VB.ComboBox cmbNpcY 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   -69600
         Style           =   2  'Dropdown List
         TabIndex        =   72
         Top             =   1440
         Width           =   735
      End
      Begin VB.ComboBox cmbNpcX 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   -70560
         Style           =   2  'Dropdown List
         TabIndex        =   71
         Top             =   1440
         Width           =   735
      End
      Begin VB.ComboBox cmbNpcY 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   -69600
         Style           =   2  'Dropdown List
         TabIndex        =   70
         Top             =   1080
         Width           =   735
      End
      Begin VB.ComboBox cmbNpcX 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   -70560
         Style           =   2  'Dropdown List
         TabIndex        =   69
         Top             =   1080
         Width           =   735
      End
      Begin VB.ComboBox cmbNpcY 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   -69600
         Style           =   2  'Dropdown List
         TabIndex        =   68
         Top             =   720
         Width           =   735
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   53
         Top             =   5400
         Width           =   1800
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "Apply"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   52
         Top             =   5040
         Width           =   1800
      End
      Begin VB.CommandButton cmdCopy 
         Caption         =   "Copy"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   9
         Left            =   -68640
         TabIndex        =   51
         Top             =   3960
         Width           =   615
      End
      Begin VB.CommandButton cmdCopy 
         Caption         =   "Copy"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   13
         Left            =   -68640
         TabIndex        =   50
         Top             =   5400
         Width           =   615
      End
      Begin VB.CommandButton cmdCopy 
         Caption         =   "Copy"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   12
         Left            =   -68640
         TabIndex        =   49
         Top             =   5040
         Width           =   615
      End
      Begin VB.CommandButton cmdCopy 
         Caption         =   "Copy"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   11
         Left            =   -68640
         TabIndex        =   48
         Top             =   4680
         Width           =   615
      End
      Begin VB.CommandButton cmdCopy 
         Caption         =   "Copy"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   10
         Left            =   -68640
         TabIndex        =   47
         Top             =   4320
         Width           =   615
      End
      Begin VB.CommandButton cmdCopy 
         Caption         =   "Copy"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   8
         Left            =   -68640
         TabIndex        =   46
         Top             =   3600
         Width           =   615
      End
      Begin VB.CommandButton cmdCopy 
         Caption         =   "Copy"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   7
         Left            =   -68640
         TabIndex        =   45
         Top             =   3240
         Width           =   615
      End
      Begin VB.CommandButton cmdCopy 
         Caption         =   "Copy"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   6
         Left            =   -68640
         TabIndex        =   44
         Top             =   2880
         Width           =   615
      End
      Begin VB.CommandButton cmdCopy 
         Caption         =   "Copy"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   5
         Left            =   -68640
         TabIndex        =   43
         Top             =   2520
         Width           =   615
      End
      Begin VB.CommandButton cmdCopy 
         Caption         =   "Copy"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   4
         Left            =   -68640
         TabIndex        =   42
         Top             =   2160
         Width           =   615
      End
      Begin VB.CommandButton cmdCopy 
         Caption         =   "Copy"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   3
         Left            =   -68640
         TabIndex        =   41
         Top             =   1800
         Width           =   615
      End
      Begin VB.CommandButton cmdCopy 
         Caption         =   "Copy"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   2
         Left            =   -68640
         TabIndex        =   40
         Top             =   1440
         Width           =   615
      End
      Begin VB.CommandButton cmdCopy 
         Caption         =   "Copy"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   -68640
         TabIndex        =   39
         Top             =   1080
         Width           =   615
      End
      Begin VB.CommandButton cmdCopy 
         Caption         =   "Copy"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   -68640
         TabIndex        =   38
         Top             =   720
         Width           =   615
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear Map NPCs"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   -73080
         TabIndex        =   37
         Top             =   6240
         Width           =   1575
      End
      Begin VB.Frame fraBGM 
         Caption         =   "Background Music"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3375
         Left            =   2160
         TabIndex        =   36
         Top             =   2400
         Width           =   5175
         Begin VB.CheckBox chkURL 
            Caption         =   "Use URL"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   63
            Top             =   2520
            Width           =   1095
         End
         Begin VB.TextBox txtURL 
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
            Left            =   240
            MaxLength       =   100
            TabIndex        =   62
            Top             =   2160
            Width           =   4695
         End
         Begin VB.CommandButton cmdPlay 
            Caption         =   "Play"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   240
            TabIndex        =   60
            Top             =   2880
            Width           =   2280
         End
         Begin VB.CommandButton cmdStop 
            Caption         =   "Stop"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   2640
            TabIndex        =   59
            Top             =   2880
            Width           =   2280
         End
         Begin VB.ListBox lstMusic 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1425
            Left            =   240
            TabIndex        =   58
            Top             =   360
            Width           =   4695
         End
         Begin VB.Label lblURL 
            Caption         =   "URL"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   61
            Top             =   1920
            Width           =   495
         End
      End
      Begin VB.Frame fraRespawn 
         Caption         =   "Respawning"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1485
         Left            =   240
         TabIndex        =   29
         Top             =   3000
         Width           =   1815
         Begin VB.TextBox txtBootMap 
            Alignment       =   2  'Center
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
            Left            =   720
            TabIndex        =   32
            Text            =   "0"
            Top             =   300
            Width           =   855
         End
         Begin VB.TextBox txtBootX 
            Alignment       =   2  'Center
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
            Left            =   720
            TabIndex        =   31
            Text            =   "0"
            Top             =   660
            Width           =   855
         End
         Begin VB.TextBox txtBootY 
            Alignment       =   2  'Center
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
            Left            =   720
            TabIndex        =   30
            Text            =   "0"
            Top             =   1020
            Width           =   855
         End
         Begin VB.Label lblMap 
            Caption         =   "Map"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   240
            TabIndex        =   35
            Top             =   300
            UseMnemonic     =   0   'False
            Width           =   330
         End
         Begin VB.Label lblX 
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   34
            Top             =   660
            Width           =   135
         End
         Begin VB.Label lblY 
            Caption         =   "Y"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   33
            Top             =   1020
            Width           =   120
         End
      End
      Begin VB.Frame fraSettings 
         Caption         =   "Settings"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1620
         Left            =   2160
         TabIndex        =   26
         Top             =   720
         Width           =   5205
         Begin VB.ComboBox cmbWeather 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "frmMapProperties.frx":0FFA
            Left            =   240
            List            =   "frmMapProperties.frx":100A
            Style           =   2  'Dropdown List
            TabIndex        =   57
            Top             =   1080
            Width           =   4695
         End
         Begin VB.ComboBox cmbMoral 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "frmMapProperties.frx":1029
            Left            =   240
            List            =   "frmMapProperties.frx":1039
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   480
            Width           =   4695
         End
         Begin VB.Label lblWeather 
            Caption         =   "Weather"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   240
            TabIndex        =   56
            Top             =   840
            Width           =   720
         End
         Begin VB.Label lblMorality 
            Caption         =   "Morality"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   240
            TabIndex        =   28
            Top             =   240
            Width           =   600
         End
      End
      Begin VB.Frame fraSwitch 
         Caption         =   "Switchover"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2235
         Left            =   240
         TabIndex        =   18
         Top             =   720
         Width           =   1815
         Begin VB.TextBox txtUp 
            Alignment       =   2  'Center
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
            Left            =   720
            TabIndex        =   55
            Text            =   "0"
            Top             =   360
            Width           =   855
         End
         Begin VB.CheckBox chkIndoors 
            Caption         =   "Map Indoors"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   54
            Top             =   1800
            Width           =   1215
         End
         Begin VB.TextBox txtLeft 
            Alignment       =   2  'Center
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
            Left            =   720
            TabIndex        =   24
            Text            =   "0"
            Top             =   1080
            Width           =   855
         End
         Begin VB.TextBox txtDown 
            Alignment       =   2  'Center
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
            Left            =   720
            TabIndex        =   22
            Text            =   "0"
            Top             =   720
            Width           =   855
         End
         Begin VB.TextBox txtRight 
            Alignment       =   2  'Center
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
            Left            =   720
            TabIndex        =   20
            Text            =   "0"
            Top             =   1440
            Width           =   855
         End
         Begin VB.Label lblLeft 
            Caption         =   "Left"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   25
            Top             =   1080
            Width           =   315
         End
         Begin VB.Label lblDown 
            Caption         =   "Down"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   240
            TabIndex        =   23
            Top             =   720
            Width           =   435
         End
         Begin VB.Label lblRight 
            Caption         =   "Right"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   240
            TabIndex        =   21
            Top             =   1440
            Width           =   375
         End
         Begin VB.Label lblUp 
            Caption         =   "Up"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   240
            TabIndex        =   19
            Top             =   360
            Width           =   300
         End
      End
      Begin VB.TextBox txtMapName 
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
         Left            =   1080
         MaxLength       =   20
         TabIndex        =   16
         Top             =   360
         Width           =   6225
      End
      Begin VB.ComboBox cmbNpc 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   -74880
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   720
         Width           =   4095
      End
      Begin VB.ComboBox cmbNpc 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   -74880
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1080
         Width           =   4095
      End
      Begin VB.ComboBox cmbNpc 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   -74880
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1440
         Width           =   4095
      End
      Begin VB.ComboBox cmbNpc 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         ItemData        =   "frmMapProperties.frx":1067
         Left            =   -74880
         List            =   "frmMapProperties.frx":1069
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1800
         Width           =   4095
      End
      Begin VB.ComboBox cmbNpc 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   -74880
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   2160
         Width           =   4095
      End
      Begin VB.ComboBox cmbNpc 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   5
         Left            =   -74880
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   2520
         Width           =   4095
      End
      Begin VB.ComboBox cmbNpc 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   6
         Left            =   -74880
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   2880
         Width           =   4095
      End
      Begin VB.ComboBox cmbNpc 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   7
         ItemData        =   "frmMapProperties.frx":106B
         Left            =   -74880
         List            =   "frmMapProperties.frx":106D
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   3240
         Width           =   4095
      End
      Begin VB.ComboBox cmbNpc 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   8
         Left            =   -74880
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   3600
         Width           =   4095
      End
      Begin VB.ComboBox cmbNpc 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   9
         Left            =   -74880
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   3960
         Width           =   4095
      End
      Begin VB.ComboBox cmbNpc 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   10
         Left            =   -74880
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   4320
         Width           =   4095
      End
      Begin VB.ComboBox cmbNpc 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   11
         Left            =   -74880
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   4680
         Width           =   4095
      End
      Begin VB.ComboBox cmbNpc 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   12
         Left            =   -74880
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   5040
         Width           =   4095
      End
      Begin VB.ComboBox cmbNpc 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   13
         ItemData        =   "frmMapProperties.frx":106F
         Left            =   -74880
         List            =   "frmMapProperties.frx":1071
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   5400
         Width           =   4095
      End
      Begin VB.ComboBox cmbNpc 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   14
         Left            =   -74880
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   5760
         Width           =   4095
      End
      Begin VB.Label lblCoordY 
         Caption         =   "Y Coord"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -69600
         TabIndex        =   67
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblMonster 
         Caption         =   "Monster Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   65
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label lblCoordX 
         Caption         =   "X Coord"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -70560
         TabIndex        =   64
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblMapName 
         Caption         =   "Map Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   240
         TabIndex        =   17
         Top             =   360
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmMapProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClear_Click()
    Dim i As Integer

    For i = 1 To MAX_MAP_NPCS
        cmbNpc(i - 1).ListIndex = 0
        cmbNpcX(i - 1).ListIndex = 0
        cmbNpcY(i - 1).ListIndex = 0
    Next i
End Sub

Private Sub cmdPlay_Click()
    If chkURL.Value = 0 Then
        Call frmMirage.MusicPlayer.PlayMedia(App.Path & "\Music\" & lstMusic.List(lstMusic.ListIndex), False)
    Else
        Call frmMirage.MusicPlayer.PlayMedia(txtURL.Text, False)
    End If
End Sub

Private Sub cmdSetRand_Click()
    Dim X As Long
    
    For X = 1 To 15
        cmbNpcX(X - 1).ListIndex = 0
        cmbNpcY(X - 1).ListIndex = 0
    Next X
End Sub

Private Sub cmdStop_Click()
    Call frmMirage.MusicPlayer.StopMedia
End Sub

Private Sub cmdCopy_Click(Index As Integer)
    cmbNpc(Index + 1).ListIndex = cmbNpc(Index).ListIndex
End Sub

Private Sub Form_Load()
    Dim X As Integer
    Dim y As Integer

    ListMusic (App.Path & "\Music\")
    ListBGS (App.Path & "\BGS\")

    txtMapName.Text = Trim$(Map(GetPlayerMap(MyIndex)).Name)
    txtUp.Text = STR$(Map(GetPlayerMap(MyIndex)).Up)
    txtDown.Text = STR$(Map(GetPlayerMap(MyIndex)).Down)
    txtLeft.Text = STR$(Map(GetPlayerMap(MyIndex)).Left)
    txtRight.Text = STR$(Map(GetPlayerMap(MyIndex)).Right)
    cmbMoral.ListIndex = Map(GetPlayerMap(MyIndex)).Moral
    txtBootMap.Text = STR$(Map(GetPlayerMap(MyIndex)).BootMap)
    txtBootX.Text = STR$(Map(GetPlayerMap(MyIndex)).BootX)
    txtBootY.Text = STR$(Map(GetPlayerMap(MyIndex)).BootY)
    lstMusic = Trim$(Map(GetPlayerMap(MyIndex)).music)
    lstMusic.Text = Trim$(Map(GetPlayerMap(MyIndex)).music)
    chkIndoors.Value = STR$(Map(GetPlayerMap(MyIndex)).Indoors)
    cmbWeather.ListIndex = Map(GetPlayerMap(MyIndex)).Weather

    For X = 1 To 15
        cmbNpc(X - 1).addItem "No NPC"
        cmbNpcX(X - 1).addItem "Rand"
        cmbNpcY(X - 1).addItem "Rand"
    Next X

    For y = 1 To MAX_NPCS
        For X = 1 To 15
            cmbNpc(X - 1).addItem y & ": " & Trim$(Npc(y).Name)
        Next X
    Next y

    For X = 1 To 15
        cmbNpc(X - 1).ListIndex = Map(GetPlayerMap(MyIndex)).Npc(X)
    Next X

    For X = 1 To 15
        For y = 0 To MAX_MAPX
            cmbNpcX(X - 1).addItem y
        Next y
        cmbNpcX(X - 1).ListIndex = Map(GetPlayerMap(MyIndex)).SpawnX(X)
    Next X
    
    For X = 1 To 15
        For y = 0 To MAX_MAPY
            cmbNpcY(X - 1).addItem y
        Next y
        cmbNpcY(X - 1).ListIndex = Map(GetPlayerMap(MyIndex)).SpawnY(X)
    Next X

    Call StopBGM
End Sub

Private Sub cmdOk_Click()
    Dim i As Integer

    Call StopBGM

    Map(GetPlayerMap(MyIndex)).Name = txtMapName.Text
    Map(GetPlayerMap(MyIndex)).Up = Val(txtUp.Text)
    Map(GetPlayerMap(MyIndex)).Down = Val(txtDown.Text)
    Map(GetPlayerMap(MyIndex)).Left = Val(txtLeft.Text)
    Map(GetPlayerMap(MyIndex)).Right = Val(txtRight.Text)
    Map(GetPlayerMap(MyIndex)).Moral = cmbMoral.ListIndex
    Map(GetPlayerMap(MyIndex)).BootMap = Val(txtBootMap.Text)
    Map(GetPlayerMap(MyIndex)).BootX = Val(txtBootX.Text)
    Map(GetPlayerMap(MyIndex)).BootY = Val(txtBootY.Text)
    Map(GetPlayerMap(MyIndex)).Indoors = Val(chkIndoors.Value)
    Map(GetPlayerMap(MyIndex)).Weather = cmbWeather.ListIndex

    For i = 1 To 15
        Map(GetPlayerMap(MyIndex)).Npc(i) = cmbNpc(i - 1).ListIndex
        Map(GetPlayerMap(MyIndex)).SpawnX(i) = cmbNpcX(i - 1).ListIndex
        Map(GetPlayerMap(MyIndex)).SpawnY(i) = cmbNpcY(i - 1).ListIndex
    Next i

    If chkURL.Value = 0 Then
        Map(GetPlayerMap(MyIndex)).music = lstMusic.Text
    Else
        If Not Left$(txtURL.Text, 7) = "http://" Then
            txtURL.Text = "http://" & txtURL.Text
        End If

        Map(GetPlayerMap(MyIndex)).music = txtURL.Text
    End If

    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Call StopBGM

    Unload Me
End Sub
