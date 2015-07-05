VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmMapProperties 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Propiedades de Mapa"
   ClientHeight    =   6285
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10215
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
   ScaleHeight     =   419
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   681
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   6165
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   10874
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   353
      TabMaxWidth     =   1764
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
      TabPicture(0)   =   "frmMapProperties.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label13"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "txtName"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame3"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame4"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdOk"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdCancel"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "NPC'S"
      TabPicture(1)   =   "frmMapProperties.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Copy(28)"
      Tab(1).Control(1)=   "Copy(27)"
      Tab(1).Control(2)=   "Copy(26)"
      Tab(1).Control(3)=   "Copy(25)"
      Tab(1).Control(4)=   "Copy(24)"
      Tab(1).Control(5)=   "cmbNpc(29)"
      Tab(1).Control(6)=   "cmbNpc(28)"
      Tab(1).Control(7)=   "cmbNpc(27)"
      Tab(1).Control(8)=   "cmbNpc(26)"
      Tab(1).Control(9)=   "cmbNpc(25)"
      Tab(1).Control(10)=   "Copy(23)"
      Tab(1).Control(11)=   "Copy(22)"
      Tab(1).Control(12)=   "cmbNpc(24)"
      Tab(1).Control(13)=   "cmbNpc(23)"
      Tab(1).Control(14)=   "Copy(21)"
      Tab(1).Control(15)=   "Copy(20)"
      Tab(1).Control(16)=   "Copy(19)"
      Tab(1).Control(17)=   "cmbNpc(22)"
      Tab(1).Control(18)=   "cmbNpc(21)"
      Tab(1).Control(19)=   "cmbNpc(20)"
      Tab(1).Control(20)=   "Copy(18)"
      Tab(1).Control(21)=   "Copy(17)"
      Tab(1).Control(22)=   "cmbNpc(19)"
      Tab(1).Control(23)=   "Copy(16)"
      Tab(1).Control(24)=   "cmbNpc(18)"
      Tab(1).Control(25)=   "cmbNpc(17)"
      Tab(1).Control(26)=   "Copy(15)"
      Tab(1).Control(27)=   "cmbNpc(16)"
      Tab(1).Control(28)=   "Copy(14)"
      Tab(1).Control(29)=   "cmbNpc(15)"
      Tab(1).Control(30)=   "Copy(9)"
      Tab(1).Control(31)=   "Copy(13)"
      Tab(1).Control(32)=   "Copy(12)"
      Tab(1).Control(33)=   "Copy(11)"
      Tab(1).Control(34)=   "Copy(10)"
      Tab(1).Control(35)=   "Copy(8)"
      Tab(1).Control(36)=   "Copy(7)"
      Tab(1).Control(37)=   "Copy(6)"
      Tab(1).Control(38)=   "Copy(5)"
      Tab(1).Control(39)=   "Copy(4)"
      Tab(1).Control(40)=   "Copy(3)"
      Tab(1).Control(41)=   "Copy(2)"
      Tab(1).Control(42)=   "Copy(1)"
      Tab(1).Control(43)=   "Copy(0)"
      Tab(1).Control(44)=   "Command1"
      Tab(1).Control(45)=   "cmbNpc(0)"
      Tab(1).Control(46)=   "cmbNpc(1)"
      Tab(1).Control(47)=   "cmbNpc(2)"
      Tab(1).Control(48)=   "cmbNpc(3)"
      Tab(1).Control(49)=   "cmbNpc(4)"
      Tab(1).Control(50)=   "cmbNpc(5)"
      Tab(1).Control(51)=   "cmbNpc(6)"
      Tab(1).Control(52)=   "cmbNpc(7)"
      Tab(1).Control(53)=   "cmbNpc(8)"
      Tab(1).Control(54)=   "cmbNpc(9)"
      Tab(1).Control(55)=   "cmbNpc(10)"
      Tab(1).Control(56)=   "cmbNpc(11)"
      Tab(1).Control(57)=   "cmbNpc(12)"
      Tab(1).Control(58)=   "cmbNpc(13)"
      Tab(1).Control(59)=   "cmbNpc(14)"
      Tab(1).ControlCount=   60
      Begin VB.CommandButton Copy 
         Caption         =   "Copiar"
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
         Index           =   28
         Left            =   -65640
         TabIndex        =   88
         Top             =   5280
         Width           =   615
      End
      Begin VB.CommandButton Copy 
         Caption         =   "Copiar"
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
         Index           =   27
         Left            =   -65640
         TabIndex        =   87
         Top             =   4920
         Width           =   615
      End
      Begin VB.CommandButton Copy 
         Caption         =   "Copiar"
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
         Index           =   26
         Left            =   -65640
         TabIndex        =   86
         Top             =   4560
         Width           =   615
      End
      Begin VB.CommandButton Copy 
         Caption         =   "Copiar"
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
         Index           =   25
         Left            =   -65640
         TabIndex        =   85
         Top             =   4200
         Width           =   615
      End
      Begin VB.CommandButton Copy 
         Caption         =   "Copiar"
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
         Index           =   24
         Left            =   -65640
         TabIndex        =   84
         Top             =   3840
         Width           =   615
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
         Index           =   29
         Left            =   -69840
         Style           =   2  'Dropdown List
         TabIndex        =   83
         Top             =   5280
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
         Index           =   28
         Left            =   -69840
         Style           =   2  'Dropdown List
         TabIndex        =   82
         Top             =   4920
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
         Index           =   27
         Left            =   -69840
         Style           =   2  'Dropdown List
         TabIndex        =   81
         Top             =   4560
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
         Index           =   26
         Left            =   -69840
         Style           =   2  'Dropdown List
         TabIndex        =   80
         Top             =   4200
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
         Index           =   25
         Left            =   -69840
         Style           =   2  'Dropdown List
         TabIndex        =   79
         Top             =   3840
         Width           =   4095
      End
      Begin VB.CommandButton Copy 
         Caption         =   "Copiar"
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
         Index           =   23
         Left            =   -65640
         TabIndex        =   78
         Top             =   3480
         Width           =   615
      End
      Begin VB.CommandButton Copy 
         Caption         =   "Copiar"
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
         Index           =   22
         Left            =   -65640
         TabIndex        =   77
         Top             =   3120
         Width           =   615
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
         Index           =   24
         Left            =   -69840
         Style           =   2  'Dropdown List
         TabIndex        =   76
         Top             =   3480
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
         Index           =   23
         Left            =   -69840
         Style           =   2  'Dropdown List
         TabIndex        =   75
         Top             =   3120
         Width           =   4095
      End
      Begin VB.CommandButton Copy 
         Caption         =   "Copiar"
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
         Index           =   21
         Left            =   -65640
         TabIndex        =   74
         Top             =   2760
         Width           =   615
      End
      Begin VB.CommandButton Copy 
         Caption         =   "Copiar"
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
         Index           =   20
         Left            =   -65640
         TabIndex        =   73
         Top             =   2400
         Width           =   615
      End
      Begin VB.CommandButton Copy 
         Caption         =   "Copiar"
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
         Index           =   19
         Left            =   -65640
         TabIndex        =   72
         Top             =   2040
         Width           =   615
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
         Index           =   22
         Left            =   -69840
         Style           =   2  'Dropdown List
         TabIndex        =   71
         Top             =   2760
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
         Index           =   21
         Left            =   -69840
         Style           =   2  'Dropdown List
         TabIndex        =   70
         Top             =   2400
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
         Index           =   20
         Left            =   -69840
         Style           =   2  'Dropdown List
         TabIndex        =   69
         Top             =   2040
         Width           =   4095
      End
      Begin VB.CommandButton Copy 
         Caption         =   "Copiar"
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
         Index           =   18
         Left            =   -65640
         TabIndex        =   68
         Top             =   1680
         Width           =   615
      End
      Begin VB.CommandButton Copy 
         Caption         =   "Copiar"
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
         Index           =   17
         Left            =   -65640
         TabIndex        =   67
         Top             =   1320
         Width           =   615
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
         Index           =   19
         Left            =   -69840
         Style           =   2  'Dropdown List
         TabIndex        =   66
         Top             =   1680
         Width           =   4095
      End
      Begin VB.CommandButton Copy 
         Caption         =   "Copiar"
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
         Index           =   16
         Left            =   -65640
         TabIndex        =   65
         Top             =   960
         Width           =   615
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
         Index           =   18
         Left            =   -69840
         Style           =   2  'Dropdown List
         TabIndex        =   64
         Top             =   1320
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
         Index           =   17
         Left            =   -69840
         Style           =   2  'Dropdown List
         TabIndex        =   63
         Top             =   960
         Width           =   4095
      End
      Begin VB.CommandButton Copy 
         Caption         =   "Copiar"
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
         Index           =   15
         Left            =   -65640
         TabIndex        =   62
         Top             =   600
         Width           =   615
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
         Index           =   16
         Left            =   -69840
         Style           =   2  'Dropdown List
         TabIndex        =   61
         Top             =   600
         Width           =   4095
      End
      Begin VB.CommandButton Copy 
         Caption         =   "Copiar"
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
         Index           =   14
         Left            =   -65640
         TabIndex        =   60
         Top             =   240
         Width           =   615
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
         Index           =   15
         Left            =   -69840
         Style           =   2  'Dropdown List
         TabIndex        =   59
         Top             =   240
         Width           =   4095
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancelar"
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
         Left            =   480
         TabIndex        =   55
         Top             =   5040
         Width           =   1080
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "Ok"
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
         Left            =   480
         TabIndex        =   54
         Top             =   4680
         Width           =   1080
      End
      Begin VB.CommandButton Copy 
         Caption         =   "Copiar"
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
         Left            =   -70680
         TabIndex        =   52
         Top             =   3840
         Width           =   615
      End
      Begin VB.CommandButton Copy 
         Caption         =   "Copiar"
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
         Left            =   -70680
         TabIndex        =   51
         Top             =   5280
         Width           =   615
      End
      Begin VB.CommandButton Copy 
         Caption         =   "Copiar"
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
         Left            =   -70680
         TabIndex        =   50
         Top             =   4920
         Width           =   615
      End
      Begin VB.CommandButton Copy 
         Caption         =   "Copiar"
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
         Left            =   -70680
         TabIndex        =   49
         Top             =   4560
         Width           =   615
      End
      Begin VB.CommandButton Copy 
         Caption         =   "Copiar"
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
         Left            =   -70680
         TabIndex        =   48
         Top             =   4200
         Width           =   615
      End
      Begin VB.CommandButton Copy 
         Caption         =   "Copiar"
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
         Left            =   -70680
         TabIndex        =   47
         Top             =   3480
         Width           =   615
      End
      Begin VB.CommandButton Copy 
         Caption         =   "Copiar"
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
         Left            =   -70680
         TabIndex        =   46
         Top             =   3120
         Width           =   615
      End
      Begin VB.CommandButton Copy 
         Caption         =   "Copiar"
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
         Left            =   -70680
         TabIndex        =   45
         Top             =   2760
         Width           =   615
      End
      Begin VB.CommandButton Copy 
         Caption         =   "Copiar"
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
         Left            =   -70680
         TabIndex        =   44
         Top             =   2400
         Width           =   615
      End
      Begin VB.CommandButton Copy 
         Caption         =   "Copiar"
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
         Left            =   -70680
         TabIndex        =   43
         Top             =   2040
         Width           =   615
      End
      Begin VB.CommandButton Copy 
         Caption         =   "Copiar"
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
         Left            =   -70680
         TabIndex        =   42
         Top             =   1680
         Width           =   615
      End
      Begin VB.CommandButton Copy 
         Caption         =   "Copiar"
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
         Left            =   -70680
         TabIndex        =   41
         Top             =   1320
         Width           =   615
      End
      Begin VB.CommandButton Copy 
         Caption         =   "Copiar"
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
         Left            =   -70680
         TabIndex        =   40
         Top             =   960
         Width           =   615
      End
      Begin VB.CommandButton Copy 
         Caption         =   "Copiar"
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
         Left            =   -70680
         TabIndex        =   39
         Top             =   600
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Limpiar Npcs del Mapa"
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
         Left            =   -66600
         TabIndex        =   38
         Top             =   5760
         Width           =   1575
      End
      Begin VB.Frame Frame4 
         Caption         =   "Musica"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3615
         Left            =   2040
         TabIndex        =   37
         Top             =   1920
         Width           =   5055
         Begin VB.CommandButton Command3 
            Caption         =   "Parar"
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
            TabIndex        =   57
            Top             =   720
            Width           =   1080
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Escuchar"
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
            TabIndex        =   56
            Top             =   360
            Width           =   1080
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
            Height          =   3180
            Left            =   120
            TabIndex        =   53
            Top             =   280
            Width           =   2415
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Revive"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1245
         Left            =   120
         TabIndex        =   30
         Top             =   2880
         Width           =   1815
         Begin VB.TextBox txtBootMap 
            Alignment       =   1  'Right Justify
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
            Left            =   900
            TabIndex        =   33
            Text            =   "0"
            Top             =   300
            Width           =   735
         End
         Begin VB.TextBox txtBootX 
            Alignment       =   1  'Right Justify
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
            Left            =   900
            TabIndex        =   32
            Text            =   "0"
            Top             =   570
            Width           =   735
         End
         Begin VB.TextBox txtBootY 
            Alignment       =   1  'Right Justify
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
            Left            =   900
            TabIndex        =   31
            Text            =   "0"
            Top             =   825
            Width           =   735
         End
         Begin VB.Label Label7 
            Caption         =   "Mapa"
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
            Left            =   120
            TabIndex        =   36
            Top             =   285
            Width           =   690
         End
         Begin VB.Label Label8 
            Caption         =   "   X"
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
            Left            =   120
            TabIndex        =   35
            Top             =   570
            Width           =   1095
         End
         Begin VB.Label Label9 
            Caption         =   "   Y"
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
            Left            =   120
            TabIndex        =   34
            Top             =   840
            Width           =   600
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Global (con el mapa actual)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1140
         Left            =   2070
         TabIndex        =   27
         Top             =   720
         Width           =   5085
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
            ItemData        =   "frmMapProperties.frx":0038
            Left            =   75
            List            =   "frmMapProperties.frx":0045
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   600
            Width           =   4920
         End
         Begin VB.Label Label1 
            Caption         =   "Moralidad del Mapa"
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
            Left            =   120
            TabIndex        =   29
            Top             =   360
            Width           =   1440
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Cambio de Mapa"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1995
         Left            =   120
         TabIndex        =   18
         Top             =   720
         Width           =   1815
         Begin VB.CheckBox chkIndoors 
            Caption         =   "Interior"
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
            TabIndex        =   58
            Top             =   1560
            Width           =   855
         End
         Begin VB.TextBox txtLeft 
            Alignment       =   1  'Right Justify
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
            Left            =   750
            TabIndex        =   25
            Text            =   "0"
            Top             =   1260
            Width           =   735
         End
         Begin VB.TextBox txtDown 
            Alignment       =   1  'Right Justify
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
            Left            =   750
            TabIndex        =   23
            Text            =   "0"
            Top             =   975
            Width           =   735
         End
         Begin VB.TextBox txtRight 
            Alignment       =   1  'Right Justify
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
            Left            =   750
            TabIndex        =   21
            Text            =   "0"
            Top             =   690
            Width           =   735
         End
         Begin VB.TextBox txtUp 
            Alignment       =   1  'Right Justify
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
            Left            =   750
            TabIndex        =   19
            Text            =   "0"
            Top             =   405
            Width           =   735
         End
         Begin VB.Label Label16 
            Caption         =   "Oeste"
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
            TabIndex        =   26
            Top             =   1275
            Width           =   525
         End
         Begin VB.Label Label15 
            Caption         =   "Sur"
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
            Left            =   285
            TabIndex        =   24
            Top             =   1005
            Width           =   435
         End
         Begin VB.Label Label2 
            Caption         =   "Este"
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
            Left            =   285
            TabIndex        =   22
            Top             =   705
            Width           =   375
         End
         Begin VB.Label Label14 
            Caption         =   "Norte"
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
            Left            =   285
            TabIndex        =   20
            Top             =   405
            Width           =   420
         End
      End
      Begin VB.TextBox txtName 
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
         Left            =   165
         MaxLength       =   40
         TabIndex        =   16
         Top             =   315
         Width           =   6105
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
         Top             =   240
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
         Top             =   600
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
         Top             =   960
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
         ItemData        =   "frmMapProperties.frx":006C
         Left            =   -74880
         List            =   "frmMapProperties.frx":006E
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1320
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
         Top             =   1680
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
         Top             =   2040
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
         Top             =   2400
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
         ItemData        =   "frmMapProperties.frx":0070
         Left            =   -74880
         List            =   "frmMapProperties.frx":0072
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   2760
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
         Top             =   3120
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
         Top             =   3480
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
         Top             =   3840
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
         Top             =   4200
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
         Top             =   4560
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
         ItemData        =   "frmMapProperties.frx":0074
         Left            =   -74880
         List            =   "frmMapProperties.frx":0076
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   4920
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
         Top             =   5280
         Width           =   4095
      End
      Begin VB.Label Label13 
         Caption         =   " Nombre   del Mapa"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   6360
         TabIndex        =   17
         Top             =   330
         Width           =   870
      End
   End
End
Attribute VB_Name = "frmMapProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim i As Long
For i = 1 To MAX_MAP_NPCS
        cmbNpc(i - 1).ListIndex = 0
    Next i
End Sub

Private Sub Command2_Click()
    Call StopMidi
    Call PlayMidi(lstMusic.Text)
End Sub

Private Sub Command3_Click()
    Call StopMidi
End Sub

Private Sub Copy_Click(Index As Integer)
    cmbNpc(Index + 1).ListIndex = cmbNpc(Index).ListIndex
End Sub

Private Sub Form_Load()
Dim x As Long, y As Long, i As Long

    txtName.Text = Trim(Map(GetPlayerMap(MyIndex)).Name)
    txtUp.Text = STR(Map(GetPlayerMap(MyIndex)).Up)
    txtDown.Text = STR(Map(GetPlayerMap(MyIndex)).Down)
    txtLeft.Text = STR(Map(GetPlayerMap(MyIndex)).Left)
    txtRight.Text = STR(Map(GetPlayerMap(MyIndex)).Right)
    cmbMoral.ListIndex = Map(GetPlayerMap(MyIndex)).Moral
    txtBootMap.Text = STR(Map(GetPlayerMap(MyIndex)).BootMap)
    txtBootX.Text = STR(Map(GetPlayerMap(MyIndex)).BootX)
    txtBootY.Text = STR(Map(GetPlayerMap(MyIndex)).BootY)
    ListMusic (App.Path & "\Music\")
    lstMusic = Trim(Map(GetPlayerMap(MyIndex)).Music)
    lstMusic.Text = Trim(Map(GetPlayerMap(MyIndex)).Music)
    chkIndoors.Value = STR(Map(GetPlayerMap(MyIndex)).Indoors)
        
    For x = 1 To MAX_MAP_NPCS
        cmbNpc(x - 1).AddItem "No NPC"
    Next x
    
    For y = 1 To MAX_NPCS
        For x = 1 To MAX_MAP_NPCS
            cmbNpc(x - 1).AddItem y & ": " & Trim(Npc(y).Name)
        Next x
    Next y
    
    For i = 1 To MAX_MAP_NPCS
        cmbNpc(i - 1).ListIndex = Map(GetPlayerMap(MyIndex)).Npc(i)
    Next i
    
    Call StopMidi
End Sub

Private Sub cmdOk_Click()
Dim x As Long, y As Long, i As Long

    Map(GetPlayerMap(MyIndex)).Name = txtName.Text
    Map(GetPlayerMap(MyIndex)).Up = Val(txtUp.Text)
    Map(GetPlayerMap(MyIndex)).Down = Val(txtDown.Text)
    Map(GetPlayerMap(MyIndex)).Left = Val(txtLeft.Text)
    Map(GetPlayerMap(MyIndex)).Right = Val(txtRight.Text)
    Map(GetPlayerMap(MyIndex)).Moral = cmbMoral.ListIndex
    Map(GetPlayerMap(MyIndex)).Music = lstMusic.Text
    Map(GetPlayerMap(MyIndex)).BootMap = Val(txtBootMap.Text)
    Map(GetPlayerMap(MyIndex)).BootX = Val(txtBootX.Text)
    Map(GetPlayerMap(MyIndex)).BootY = Val(txtBootY.Text)
    Map(GetPlayerMap(MyIndex)).Indoors = Val(chkIndoors.Value)
    
    For i = 1 To MAX_MAP_NPCS
        Map(GetPlayerMap(MyIndex)).Npc(i) = cmbNpc(i - 1).ListIndex
    Next i
    
    Call StopMidi
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Call StopMidi
    Unload Me
End Sub

