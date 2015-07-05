VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMapProperties 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Map Properties"
   ClientHeight    =   7050
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7350
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
   ScaleHeight     =   470
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   360
      TabIndex        =   88
      Top             =   6600
      Width           =   1080
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
      Left            =   1680
      TabIndex        =   87
      Top             =   6600
      Width           =   1080
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6285
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   11086
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
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "NPC'S"
      TabPicture(1)   =   "frmMapProperties.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Num(0)"
      Tab(1).Control(1)=   "Num(1)"
      Tab(1).Control(2)=   "Num(2)"
      Tab(1).Control(3)=   "Num(3)"
      Tab(1).Control(4)=   "Num(4)"
      Tab(1).Control(5)=   "Num(5)"
      Tab(1).Control(6)=   "Num(6)"
      Tab(1).Control(7)=   "Num(7)"
      Tab(1).Control(8)=   "Num(8)"
      Tab(1).Control(9)=   "Num(9)"
      Tab(1).Control(10)=   "Num(10)"
      Tab(1).Control(11)=   "Num(11)"
      Tab(1).Control(12)=   "Num(12)"
      Tab(1).Control(13)=   "Num(13)"
      Tab(1).Control(14)=   "Num(14)"
      Tab(1).Control(15)=   "cmbNpc(14)"
      Tab(1).Control(16)=   "cmbNpc(13)"
      Tab(1).Control(17)=   "cmbNpc(12)"
      Tab(1).Control(18)=   "cmbNpc(11)"
      Tab(1).Control(19)=   "cmbNpc(10)"
      Tab(1).Control(20)=   "cmbNpc(9)"
      Tab(1).Control(21)=   "cmbNpc(8)"
      Tab(1).Control(22)=   "cmbNpc(7)"
      Tab(1).Control(23)=   "cmbNpc(6)"
      Tab(1).Control(24)=   "cmbNpc(5)"
      Tab(1).Control(25)=   "cmbNpc(4)"
      Tab(1).Control(26)=   "cmbNpc(3)"
      Tab(1).Control(27)=   "cmbNpc(2)"
      Tab(1).Control(28)=   "cmbNpc(1)"
      Tab(1).Control(29)=   "cmbNpc(0)"
      Tab(1).Control(30)=   "Command1"
      Tab(1).Control(31)=   "Copy(0)"
      Tab(1).Control(32)=   "Copy(1)"
      Tab(1).Control(33)=   "Copy(2)"
      Tab(1).Control(34)=   "Copy(3)"
      Tab(1).Control(35)=   "Copy(4)"
      Tab(1).Control(36)=   "Copy(5)"
      Tab(1).Control(37)=   "Copy(6)"
      Tab(1).Control(38)=   "Copy(7)"
      Tab(1).Control(39)=   "Copy(8)"
      Tab(1).Control(40)=   "Copy(10)"
      Tab(1).Control(41)=   "Copy(11)"
      Tab(1).Control(42)=   "Copy(12)"
      Tab(1).Control(43)=   "Copy(13)"
      Tab(1).Control(44)=   "Copy(9)"
      Tab(1).Control(45)=   "Spawn(0)"
      Tab(1).Control(46)=   "Spawn(1)"
      Tab(1).Control(47)=   "Spawn(2)"
      Tab(1).Control(48)=   "Spawn(3)"
      Tab(1).Control(49)=   "Spawn(4)"
      Tab(1).Control(50)=   "Spawn(5)"
      Tab(1).Control(51)=   "Spawn(6)"
      Tab(1).Control(52)=   "Spawn(7)"
      Tab(1).Control(53)=   "Spawn(8)"
      Tab(1).Control(54)=   "Spawn(9)"
      Tab(1).Control(55)=   "Spawn(10)"
      Tab(1).Control(56)=   "Spawn(11)"
      Tab(1).Control(57)=   "Spawn(12)"
      Tab(1).Control(58)=   "Spawn(13)"
      Tab(1).Control(59)=   "Spawn(14)"
      Tab(1).ControlCount=   60
      Begin VB.CommandButton Spawn 
         Caption         =   "Set Spawn Point"
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
         Left            =   -70440
         TabIndex        =   71
         Top             =   5400
         Width           =   1455
      End
      Begin VB.CommandButton Spawn 
         Caption         =   "Set Spawn Point"
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
         Left            =   -70440
         TabIndex        =   70
         Top             =   5040
         Width           =   1455
      End
      Begin VB.CommandButton Spawn 
         Caption         =   "Set Spawn Point"
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
         Left            =   -70440
         TabIndex        =   69
         Top             =   4680
         Width           =   1455
      End
      Begin VB.CommandButton Spawn 
         Caption         =   "Set Spawn Point"
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
         Left            =   -70440
         TabIndex        =   68
         Top             =   4320
         Width           =   1455
      End
      Begin VB.CommandButton Spawn 
         Caption         =   "Set Spawn Point"
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
         Left            =   -70440
         TabIndex        =   67
         Top             =   3960
         Width           =   1455
      End
      Begin VB.CommandButton Spawn 
         Caption         =   "Set Spawn Point"
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
         Left            =   -70440
         TabIndex        =   66
         Top             =   3600
         Width           =   1455
      End
      Begin VB.CommandButton Spawn 
         Caption         =   "Set Spawn Point"
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
         Left            =   -70440
         TabIndex        =   65
         Top             =   3240
         Width           =   1455
      End
      Begin VB.CommandButton Spawn 
         Caption         =   "Set Spawn Point"
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
         Left            =   -70440
         TabIndex        =   64
         Top             =   2880
         Width           =   1455
      End
      Begin VB.CommandButton Spawn 
         Caption         =   "Set Spawn Point"
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
         Left            =   -70440
         TabIndex        =   63
         Top             =   2520
         Width           =   1455
      End
      Begin VB.CommandButton Spawn 
         Caption         =   "Set Spawn Point"
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
         Left            =   -70440
         TabIndex        =   62
         Top             =   2160
         Width           =   1455
      End
      Begin VB.CommandButton Spawn 
         Caption         =   "Set Spawn Point"
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
         Left            =   -70440
         TabIndex        =   61
         Top             =   1800
         Width           =   1455
      End
      Begin VB.CommandButton Spawn 
         Caption         =   "Set Spawn Point"
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
         Left            =   -70440
         TabIndex        =   60
         Top             =   1440
         Width           =   1455
      End
      Begin VB.CommandButton Spawn 
         Caption         =   "Set Spawn Point"
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
         Left            =   -70440
         TabIndex        =   59
         Top             =   1080
         Width           =   1455
      End
      Begin VB.CommandButton Spawn 
         Caption         =   "Set Spawn Point"
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
         Left            =   -70440
         TabIndex        =   58
         Top             =   720
         Width           =   1455
      End
      Begin VB.CommandButton Spawn 
         Caption         =   "Set Spawn Point"
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
         Left            =   -70440
         TabIndex        =   57
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton Copy 
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
         Left            =   -68760
         TabIndex        =   52
         Top             =   3960
         Width           =   615
      End
      Begin VB.CommandButton Copy 
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
         Left            =   -68760
         TabIndex        =   51
         Top             =   5400
         Width           =   615
      End
      Begin VB.CommandButton Copy 
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
         Left            =   -68760
         TabIndex        =   50
         Top             =   5040
         Width           =   615
      End
      Begin VB.CommandButton Copy 
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
         Left            =   -68760
         TabIndex        =   49
         Top             =   4680
         Width           =   615
      End
      Begin VB.CommandButton Copy 
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
         Left            =   -68760
         TabIndex        =   48
         Top             =   4320
         Width           =   615
      End
      Begin VB.CommandButton Copy 
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
         Left            =   -68760
         TabIndex        =   47
         Top             =   3600
         Width           =   615
      End
      Begin VB.CommandButton Copy 
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
         Left            =   -68760
         TabIndex        =   46
         Top             =   3240
         Width           =   615
      End
      Begin VB.CommandButton Copy 
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
         Left            =   -68760
         TabIndex        =   45
         Top             =   2880
         Width           =   615
      End
      Begin VB.CommandButton Copy 
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
         Left            =   -68760
         TabIndex        =   44
         Top             =   2520
         Width           =   615
      End
      Begin VB.CommandButton Copy 
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
         Left            =   -68760
         TabIndex        =   43
         Top             =   2160
         Width           =   615
      End
      Begin VB.CommandButton Copy 
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
         Left            =   -68760
         TabIndex        =   42
         Top             =   1800
         Width           =   615
      End
      Begin VB.CommandButton Copy 
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
         Left            =   -68760
         TabIndex        =   41
         Top             =   1440
         Width           =   615
      End
      Begin VB.CommandButton Copy 
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
         Left            =   -68760
         TabIndex        =   40
         Top             =   1080
         Width           =   615
      End
      Begin VB.CommandButton Copy 
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
         Left            =   -68760
         TabIndex        =   39
         Top             =   720
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Clear Map Npc's"
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
         Left            =   -71880
         TabIndex        =   38
         Top             =   5880
         Width           =   1575
      End
      Begin VB.Frame Frame4 
         Caption         =   "Music"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3975
         Left            =   2040
         TabIndex        =   37
         Top             =   1920
         Width           =   4815
         Begin VB.CommandButton Command3 
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
            TabIndex        =   55
            Top             =   720
            Width           =   1080
         End
         Begin VB.CommandButton Command2 
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
            Left            =   2640
            TabIndex        =   54
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
            Height          =   3375
            Left            =   120
            TabIndex        =   53
            Top             =   280
            Width           =   2415
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Dungeon"
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
            Caption         =   "Boot Map"
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
            Caption         =   "Boot X"
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
            Caption         =   "Boot Y"
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
         Caption         =   "Global (within current map)"
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
         Left            =   2040
         TabIndex        =   27
         Top             =   720
         Width           =   4845
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
            Left            =   120
            List            =   "frmMapProperties.frx":0045
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   600
            Width           =   4560
         End
         Begin VB.Label Label1 
            Caption         =   "Map Morality"
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
            Left            =   90
            TabIndex        =   29
            Top             =   360
            Width           =   960
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Map Switchovers"
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
            Caption         =   "Indoors"
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
            TabIndex        =   56
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
            Caption         =   "West"
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
            Left            =   285
            TabIndex        =   26
            Top             =   1275
            Width           =   405
         End
         Begin VB.Label Label15 
            Caption         =   "South"
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
            Caption         =   "East"
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
            Caption         =   "North"
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
         Width           =   5850
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
         Left            =   -74640
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   360
         Width           =   3975
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
         Left            =   -74640
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   720
         Width           =   3975
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
         Left            =   -74640
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1080
         Width           =   3975
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
         Left            =   -74640
         List            =   "frmMapProperties.frx":006E
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1440
         Width           =   3975
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
         Left            =   -74640
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1800
         Width           =   3975
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
         Left            =   -74640
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   2160
         Width           =   3975
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
         Left            =   -74640
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   2520
         Width           =   3975
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
         Left            =   -74640
         List            =   "frmMapProperties.frx":0072
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   2880
         Width           =   3975
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
         Left            =   -74640
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   3240
         Width           =   3975
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
         Left            =   -74640
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   3600
         Width           =   3975
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
         Left            =   -74640
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   3960
         Width           =   3975
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
         Left            =   -74640
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   4320
         Width           =   3975
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
         Left            =   -74640
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   4680
         Width           =   3975
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
         Left            =   -74640
         List            =   "frmMapProperties.frx":0076
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   5040
         Width           =   3975
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
         Left            =   -74640
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   5400
         Width           =   3975
      End
      Begin VB.Label Num 
         Caption         =   "#"
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
         Index           =   14
         Left            =   -74880
         TabIndex        =   86
         Top             =   5400
         Width           =   255
      End
      Begin VB.Label Num 
         Caption         =   "#"
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
         Index           =   13
         Left            =   -74880
         TabIndex        =   85
         Top             =   5040
         Width           =   255
      End
      Begin VB.Label Num 
         Caption         =   "#"
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
         Index           =   12
         Left            =   -74880
         TabIndex        =   84
         Top             =   4680
         Width           =   255
      End
      Begin VB.Label Num 
         Caption         =   "#"
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
         Index           =   11
         Left            =   -74880
         TabIndex        =   83
         Top             =   4320
         Width           =   255
      End
      Begin VB.Label Num 
         Caption         =   "#"
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
         Index           =   10
         Left            =   -74880
         TabIndex        =   82
         Top             =   3960
         Width           =   255
      End
      Begin VB.Label Num 
         Caption         =   "#"
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
         Index           =   9
         Left            =   -74880
         TabIndex        =   81
         Top             =   3600
         Width           =   255
      End
      Begin VB.Label Num 
         Caption         =   "#"
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
         Index           =   8
         Left            =   -74880
         TabIndex        =   80
         Top             =   3240
         Width           =   255
      End
      Begin VB.Label Num 
         Caption         =   "#"
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
         Index           =   7
         Left            =   -74880
         TabIndex        =   79
         Top             =   2880
         Width           =   255
      End
      Begin VB.Label Num 
         Caption         =   "#"
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
         Index           =   6
         Left            =   -74880
         TabIndex        =   78
         Top             =   2520
         Width           =   255
      End
      Begin VB.Label Num 
         Caption         =   "#"
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
         Index           =   5
         Left            =   -74880
         TabIndex        =   77
         Top             =   2160
         Width           =   255
      End
      Begin VB.Label Num 
         Caption         =   "#"
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
         Index           =   4
         Left            =   -74880
         TabIndex        =   76
         Top             =   1800
         Width           =   255
      End
      Begin VB.Label Num 
         Caption         =   "#"
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
         Index           =   3
         Left            =   -74880
         TabIndex        =   75
         Top             =   1440
         Width           =   255
      End
      Begin VB.Label Num 
         Caption         =   "#"
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
         Index           =   2
         Left            =   -74880
         TabIndex        =   74
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label Num 
         Caption         =   "#"
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
         Index           =   1
         Left            =   -74880
         TabIndex        =   73
         Top             =   720
         Width           =   255
      End
      Begin VB.Label Num 
         Caption         =   "#"
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
         Index           =   0
         Left            =   -74880
         TabIndex        =   72
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label13 
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
         Left            =   6120
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

' Copyright (c) 2007-2008 Elysium Source. All rights reserved.
' This code is licensed under the Elysium General License.

Option Explicit

Private Sub Command1_Click()
Dim I As Long
    For I = 1 To MAX_MAP_NPCS
        cmbNpc(I - 1).ListIndex = 0
        
        TempNpcSpawn(I).Used = 0
        TempNpcSpawn(I).x = 0
        TempNpcSpawn(I).y = 0
        
        Spawn(I - 1).Caption = "Set Spawn Point"
    Next I
End Sub

Private Sub Command2_Click()
    'Call StopMidi
    Call PlayMidi(lstMusic.Text)
End Sub

Private Sub Command3_Click()
    Call StopMidi
End Sub

Private Sub Copy_Click(Index As Integer)
    cmbNpc(Index + 1).ListIndex = cmbNpc(Index).ListIndex
    
    TempNpcSpawn(Index + 2) = TempNpcSpawn(Index + 1)
    If TempNpcSpawn(Index + 2).Used = 1 Then
        Spawn(Index + 1).Caption = "(" & TempNpcSpawn(Index + 2).x & ", " & TempNpcSpawn(Index + 2).y & ")"
    Else
        Spawn(Index + 1).Caption = "Set Spawn Point"
    End If
End Sub

Private Sub Form_Load()
Dim x As Long, y As Long, I As Long

    InSpawnEditor = True

    txtName.Text = Trim$(Map(GetPlayerMap(MyIndex)).Name)
    txtUp.Text = STR(Map(GetPlayerMap(MyIndex)).Up)
    txtDown.Text = STR(Map(GetPlayerMap(MyIndex)).Down)
    txtLeft.Text = STR(Map(GetPlayerMap(MyIndex)).Left)
    txtRight.Text = STR(Map(GetPlayerMap(MyIndex)).Right)
    cmbMoral.ListIndex = Map(GetPlayerMap(MyIndex)).Moral
    txtBootMap.Text = STR(Map(GetPlayerMap(MyIndex)).BootMap)
    txtBootX.Text = STR(Map(GetPlayerMap(MyIndex)).BootX)
    txtBootY.Text = STR(Map(GetPlayerMap(MyIndex)).BootY)
    ListMusic (App.Path & "\Music\")
    lstMusic = Trim$(Map(GetPlayerMap(MyIndex)).Music)
    lstMusic.Text = Trim$(Map(GetPlayerMap(MyIndex)).Music)
    chkIndoors.Value = STR(Map(GetPlayerMap(MyIndex)).Indoors)
    
    For x = 1 To MAX_MAP_NPCS
        cmbNpc(x - 1).AddItem "No NPC"
        Num(x - 1).Caption = x & "."
    Next x
    
    For y = 1 To MAX_NPCS
        For x = 1 To MAX_MAP_NPCS
            cmbNpc(x - 1).AddItem y & ": " & Trim$(Npc(y).Name)
        Next x
    Next y
    
    For I = 1 To MAX_MAP_NPCS
        cmbNpc(I - 1).ListIndex = Map(GetPlayerMap(MyIndex)).Npc(I)
        
        TempNpcSpawn(I).Used = Map(GetPlayerMap(MyIndex)).NpcSpawn(I).Used
        TempNpcSpawn(I).x = Map(GetPlayerMap(MyIndex)).NpcSpawn(I).x
        TempNpcSpawn(I).y = Map(GetPlayerMap(MyIndex)).NpcSpawn(I).y
        
        If TempNpcSpawn(I).Used = 1 Then
            Spawn(I - 1).Caption = "(" & TempNpcSpawn(I).x & ", " & TempNpcSpawn(I).y & ")"
        Else
            Spawn(I - 1).Caption = "Set Spawn Point"
        End If
    Next I
    
    SpawnLocator = 0
    
    Call StopMidi
End Sub

Private Sub cmdOk_Click()
Dim x As Long, y As Long, I As Long

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
    
    For I = 1 To MAX_MAP_NPCS
        Map(GetPlayerMap(MyIndex)).Npc(I) = cmbNpc(I - 1).ListIndex
        Map(GetPlayerMap(MyIndex)).NpcSpawn(I) = TempNpcSpawn(I)
    Next I
    
    InSpawnEditor = False
    
    Call StopMidi
    frmMapEditor.Visible = True
    Unload Me
End Sub

Private Sub cmdCancel_Click()
Dim I As Long

    InSpawnEditor = False

    Call StopMidi
    frmMapEditor.Visible = True
    Unload Me
End Sub

Private Sub Spawn_Click(Index As Integer)
    If SpawnLocator = Index + 1 Then
        SpawnLocator = 0
        TempNpcSpawn(Index + 1).Used = 0
        Spawn(Index).Caption = "Set Spawn Point"
        Exit Sub
    End If
    
    SpawnLocator = Index + 1
    Spawn(Index).Caption = "Click On Screen"
End Sub
