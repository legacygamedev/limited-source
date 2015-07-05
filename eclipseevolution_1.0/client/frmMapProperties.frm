VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmMapProperties 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Map Properties"
   ClientHeight    =   6150
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
   ScaleHeight     =   410
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   517
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   5925
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   10451
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
      TabPicture(1)   =   "frmMapProperties.frx":0FDE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Copy(9)"
      Tab(1).Control(1)=   "Copy(13)"
      Tab(1).Control(2)=   "Copy(12)"
      Tab(1).Control(3)=   "Copy(11)"
      Tab(1).Control(4)=   "Copy(10)"
      Tab(1).Control(5)=   "Copy(8)"
      Tab(1).Control(6)=   "Copy(7)"
      Tab(1).Control(7)=   "Copy(6)"
      Tab(1).Control(8)=   "Copy(5)"
      Tab(1).Control(9)=   "Copy(4)"
      Tab(1).Control(10)=   "Copy(3)"
      Tab(1).Control(11)=   "Copy(2)"
      Tab(1).Control(12)=   "Copy(1)"
      Tab(1).Control(13)=   "Copy(0)"
      Tab(1).Control(14)=   "Command1"
      Tab(1).Control(15)=   "cmbNpc(0)"
      Tab(1).Control(16)=   "cmbNpc(1)"
      Tab(1).Control(17)=   "cmbNpc(2)"
      Tab(1).Control(18)=   "cmbNpc(3)"
      Tab(1).Control(19)=   "cmbNpc(4)"
      Tab(1).Control(20)=   "cmbNpc(5)"
      Tab(1).Control(21)=   "cmbNpc(6)"
      Tab(1).Control(22)=   "cmbNpc(7)"
      Tab(1).Control(23)=   "cmbNpc(8)"
      Tab(1).Control(24)=   "cmbNpc(9)"
      Tab(1).Control(25)=   "cmbNpc(10)"
      Tab(1).Control(26)=   "cmbNpc(11)"
      Tab(1).Control(27)=   "cmbNpc(12)"
      Tab(1).Control(28)=   "cmbNpc(13)"
      Tab(1).Control(29)=   "cmbNpc(14)"
      Tab(1).ControlCount=   30
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
         Left            =   480
         TabIndex        =   54
         Top             =   4920
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
         TabIndex        =   53
         Top             =   4560
         Width           =   1080
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
         Left            =   -70680
         TabIndex        =   51
         Top             =   3840
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
         Left            =   -70680
         TabIndex        =   50
         Top             =   5280
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
         Left            =   -70680
         TabIndex        =   49
         Top             =   4920
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
         Left            =   -70680
         TabIndex        =   48
         Top             =   4560
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
         Left            =   -70680
         TabIndex        =   47
         Top             =   4200
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
         Left            =   -70680
         TabIndex        =   46
         Top             =   3480
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
         Left            =   -70680
         TabIndex        =   45
         Top             =   3120
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
         Left            =   -70680
         TabIndex        =   44
         Top             =   2760
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
         Left            =   -70680
         TabIndex        =   43
         Top             =   2400
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
         Left            =   -70680
         TabIndex        =   42
         Top             =   2040
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
         Left            =   -70680
         TabIndex        =   41
         Top             =   1680
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
         Left            =   -70680
         TabIndex        =   40
         Top             =   1320
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
         Left            =   -70680
         TabIndex        =   39
         Top             =   960
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
         Left            =   -70680
         TabIndex        =   38
         Top             =   600
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
         Left            =   -69360
         TabIndex        =   37
         Top             =   5280
         Width           =   1575
      End
      Begin VB.Frame Frame4 
         Caption         =   "Music"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3735
         Left            =   2040
         TabIndex        =   36
         Top             =   1920
         Width           =   5175
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
            Left            =   3840
            TabIndex        =   56
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
            Left            =   3840
            TabIndex        =   55
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
            TabIndex        =   52
            Top             =   280
            Width           =   3495
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Dungeon"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1245
         Left            =   120
         TabIndex        =   29
         Top             =   3120
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
            TabIndex        =   32
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
            TabIndex        =   31
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
            TabIndex        =   30
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
            TabIndex        =   35
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
            TabIndex        =   34
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
            TabIndex        =   33
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1140
         Left            =   2040
         TabIndex        =   26
         Top             =   720
         Width           =   5205
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
            ItemData        =   "frmMapProperties.frx":0FFA
            Left            =   75
            List            =   "frmMapProperties.frx":100A
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   600
            Width           =   4800
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
            TabIndex        =   28
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2355
         Left            =   120
         TabIndex        =   18
         Top             =   720
         Width           =   1815
         Begin VB.TextBox txtup 
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
            Height          =   285
            Left            =   720
            TabIndex        =   58
            Text            =   "0"
            Top             =   480
            Width           =   975
         End
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
            Left            =   480
            TabIndex        =   57
            Top             =   1920
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
            Left            =   720
            TabIndex        =   24
            Text            =   "0"
            Top             =   1200
            Width           =   975
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
            Left            =   720
            TabIndex        =   22
            Text            =   "0"
            Top             =   840
            Width           =   975
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
            Left            =   720
            TabIndex        =   20
            Text            =   "0"
            Top             =   1560
            Width           =   975
         End
         Begin VB.Label Label16 
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
            Top             =   1200
            Width           =   405
         End
         Begin VB.Label Label15 
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
            Top             =   840
            Width           =   435
         End
         Begin VB.Label Label2 
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
            Top             =   1560
            Width           =   375
         End
         Begin VB.Label Label14 
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
            Top             =   480
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
         ItemData        =   "frmMapProperties.frx":1038
         Left            =   -74880
         List            =   "frmMapProperties.frx":103A
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
         ItemData        =   "frmMapProperties.frx":103C
         Left            =   -74880
         List            =   "frmMapProperties.frx":103E
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
         ItemData        =   "frmMapProperties.frx":1040
         Left            =   -74880
         List            =   "frmMapProperties.frx":1042
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
         Left            =   6375
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
    
    Select Case Right(lstMusic.Text, 4)
    
    Case ".mp3"
    frmMirage.Mp3musicplayer.currentPlaylist.Clear
    frmMirage.Mp3musicplayer.Controls.Stop
    frmMirage.Mp3musicplayer.URL = App.Path & "\Music\" & lstMusic.Text & ""
    frmMirage.Mp3musicplayer.Controls.Play
    frmMirage.Mp3timer.Enabled = True
    
    Case ".wma"
    frmMirage.Mp3musicplayer.currentPlaylist.Clear
    frmMirage.Mp3musicplayer.Controls.Stop
    frmMirage.Mp3musicplayer.URL = App.Path & "\Music\" & lstMusic.Text & ""
    frmMirage.Mp3musicplayer.Controls.Play
    frmMirage.Mp3timer.Enabled = True
    
    Case ".mid"
    Call PlayMidi(lstMusic.Text)
    
    End Select

End Sub

Private Sub Command3_Click()
    Call StopMidi
    frmMirage.Mp3musicplayer.currentPlaylist.Clear
    frmMirage.Mp3musicplayer.Controls.Stop
    frmMirage.Mp3timer.Enabled = False
End Sub

Private Sub Copy_Click(index As Integer)
    cmbNpc(index + 1).ListIndex = cmbNpc(index).ListIndex
End Sub

Private Sub Form_Load()
Dim X As Long, Y As Long, i As Long

    txtName.Text = Trim$(Map(GetPlayerMap(MyIndex)).Name)
    txtup.Text = STR(Map(GetPlayerMap(MyIndex)).Up)
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
        
    For X = 1 To MAX_MAP_NPCS
        cmbNpc(X - 1).AddItem "No NPC"
    Next X
    
    For Y = 1 To MAX_NPCS
        For X = 1 To MAX_MAP_NPCS
            cmbNpc(X - 1).AddItem Y & ": " & Trim$(Npc(Y).Name)
        Next X
    Next Y
    
    For i = 1 To MAX_MAP_NPCS
        cmbNpc(i - 1).ListIndex = Map(GetPlayerMap(MyIndex)).Npc(i)
    Next i
    
    Call StopMidi
End Sub

Private Sub cmdOk_Click()
Dim X As Long, Y As Long, i As Long

    Map(GetPlayerMap(MyIndex)).Name = txtName.Text
    Map(GetPlayerMap(MyIndex)).Up = Val#(txtup.Text)
    Map(GetPlayerMap(MyIndex)).Down = Val#(txtDown.Text)
    Map(GetPlayerMap(MyIndex)).Left = Val#(txtLeft.Text)
    Map(GetPlayerMap(MyIndex)).Right = Val#(txtRight.Text)
    Map(GetPlayerMap(MyIndex)).Moral = cmbMoral.ListIndex
    Map(GetPlayerMap(MyIndex)).Music = lstMusic.Text
    Map(GetPlayerMap(MyIndex)).BootMap = Val#(txtBootMap.Text)
    Map(GetPlayerMap(MyIndex)).BootX = Val#(txtBootX.Text)
    Map(GetPlayerMap(MyIndex)).BootY = Val#(txtBootY.Text)
    Map(GetPlayerMap(MyIndex)).Indoors = Val#(chkIndoors.Value)
    
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
