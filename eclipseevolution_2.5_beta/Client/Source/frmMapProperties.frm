VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmMapProperties 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Map Properties"
   ClientHeight    =   6000
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
   ScaleHeight     =   400
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   517
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   5805
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   10239
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
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
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cmdCancel"
      Tab(0).Control(1)=   "cmdOk"
      Tab(0).Control(2)=   "Frame4"
      Tab(0).Control(3)=   "Frame3"
      Tab(0).Control(4)=   "Frame2"
      Tab(0).Control(5)=   "Frame1"
      Tab(0).Control(6)=   "txtName"
      Tab(0).Control(7)=   "Label13"
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "NPC'S"
      TabPicture(1)   =   "frmMapProperties.frx":0FDE
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label6"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cmbNpc(14)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cmbNpc(13)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cmbNpc(12)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "cmbNpc(11)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "cmbNpc(10)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "cmbNpc(9)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "cmbNpc(8)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "cmbNpc(7)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "cmbNpc(6)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "cmbNpc(5)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "cmbNpc(4)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "cmbNpc(3)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "cmbNpc(2)"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "cmbNpc(1)"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "cmbNpc(0)"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "Command1"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "Copy(0)"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "Copy(1)"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "Copy(2)"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "Copy(3)"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "Copy(4)"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "Copy(5)"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "Copy(6)"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "Copy(7)"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "Copy(8)"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "Copy(10)"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "Copy(11)"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "Copy(12)"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "Copy(13)"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).Control(30)=   "Copy(9)"
      Tab(1).Control(30).Enabled=   0   'False
      Tab(1).ControlCount=   31
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
         Left            =   -74760
         TabIndex        =   53
         Top             =   4920
         Width           =   1320
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
         Left            =   -74760
         TabIndex        =   52
         Top             =   4560
         Width           =   1320
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
         Left            =   4320
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
         Left            =   4320
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
         Left            =   4320
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
         Left            =   4320
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
         Left            =   4320
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
         Left            =   4320
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
         Left            =   4320
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
         Left            =   4320
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
         Left            =   4320
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
         Left            =   4320
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
         Left            =   4320
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
         Left            =   4320
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
         Left            =   4320
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
         Left            =   4320
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
         Left            =   5640
         TabIndex        =   37
         Top             =   5280
         Width           =   1575
      End
      Begin VB.Frame Frame4 
         Caption         =   "Sound"
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
         Left            =   -72960
         TabIndex        =   36
         Top             =   2280
         Width           =   5175
         Begin VB.CheckBox Check1 
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
            Left            =   120
            TabIndex        =   64
            Top             =   2520
            Width           =   1095
         End
         Begin VB.TextBox Text2 
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
            Left            =   120
            MaxLength       =   40
            TabIndex        =   63
            Top             =   2160
            Width           =   4905
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
            Left            =   120
            TabIndex        =   60
            Top             =   2880
            Width           =   2280
         End
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
            Left            =   2520
            TabIndex        =   59
            Top             =   2880
            Width           =   2520
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
            Left            =   120
            TabIndex        =   58
            Top             =   480
            Width           =   4935
         End
         Begin VB.Label Label5 
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
            Left            =   120
            TabIndex        =   62
            Top             =   1920
            Width           =   495
         End
         Begin VB.Label Label4 
            Caption         =   "Background Music"
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
            TabIndex        =   61
            Top             =   240
            Width           =   2415
         End
      End
      Begin VB.Frame Frame3 
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
         Height          =   1245
         Left            =   -74880
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
         Height          =   1500
         Left            =   -72960
         TabIndex        =   26
         Top             =   720
         Width           =   5205
         Begin VB.ComboBox CmbWeather 
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
            Left            =   120
            List            =   "frmMapProperties.frx":100A
            Style           =   2  'Dropdown List
            TabIndex        =   57
            Top             =   1080
            Width           =   4800
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
            Left            =   75
            List            =   "frmMapProperties.frx":1039
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   480
            Width           =   4800
         End
         Begin VB.Label Label3 
            Caption         =   "Map weather"
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
            TabIndex        =   56
            Top             =   840
            Width           =   960
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
            Top             =   240
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
         Left            =   -74880
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
            TabIndex        =   55
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
            Left            =   720
            TabIndex        =   54
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
         Left            =   -74835
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
         Left            =   120
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
         Left            =   120
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
         Left            =   120
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
         ItemData        =   "frmMapProperties.frx":1067
         Left            =   120
         List            =   "frmMapProperties.frx":1069
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
         Left            =   120
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
         Left            =   120
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
         Left            =   120
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
         ItemData        =   "frmMapProperties.frx":106B
         Left            =   120
         List            =   "frmMapProperties.frx":106D
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
         Left            =   120
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
         Left            =   120
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
         Left            =   120
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
         Left            =   120
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
         Left            =   120
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
         ItemData        =   "frmMapProperties.frx":106F
         Left            =   120
         List            =   "frmMapProperties.frx":1071
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
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   5280
         Width           =   4095
      End
      Begin VB.Label Label6 
         Caption         =   $"frmMapProperties.frx":1073
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   5160
         TabIndex        =   65
         Top             =   480
         Width           =   2055
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
         Left            =   -68625
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
If Check1.Value = 0 Then
Call frmMirage.MusicPlayer.PlayMedia(App.Path & "/music/" & lstMusic.List(lstMusic.ListIndex), False)
Else
Call frmMirage.MusicPlayer.PlayMedia(Text2.Text, False)
End If
End Sub

Private Sub Command3_Click()
Call frmMirage.MusicPlayer.StopMedia
End Sub

Private Sub Command4_Click()
Call frmMirage.BGSPlayer.PlayMedia(App.Path & "/BGS/" & lstMusic.List(lstMusic.ListIndex), True)
End Sub

Private Sub Command5_Click()
Call frmMirage.BGSPlayer.StopMedia
End Sub

Private Sub Copy_Click(Index As Integer)
    cmbNpc(Index + 1).ListIndex = cmbNpc(Index).ListIndex
End Sub

Private Sub Form_Load()
Dim x As Long, y As Long, i As Long

    txtName.Text = Trim$(Map(GetPlayerMap(MyIndex)).Name)
    txtup.Text = STR(Map(GetPlayerMap(MyIndex)).Up)
    txtDown.Text = STR(Map(GetPlayerMap(MyIndex)).Down)
    txtLeft.Text = STR(Map(GetPlayerMap(MyIndex)).Left)
    txtRight.Text = STR(Map(GetPlayerMap(MyIndex)).right)
    cmbMoral.ListIndex = Map(GetPlayerMap(MyIndex)).Moral
    txtBootMap.Text = STR(Map(GetPlayerMap(MyIndex)).BootMap)
    txtBootX.Text = STR(Map(GetPlayerMap(MyIndex)).BootX)
    txtBootY.Text = STR(Map(GetPlayerMap(MyIndex)).BootY)
    ListMusic (App.Path & "\Music\")
    lstMusic = Trim$(Map(GetPlayerMap(MyIndex)).Music)
    lstMusic.Text = Trim$(Map(GetPlayerMap(MyIndex)).Music)
        ListBGS (App.Path & "\BGS\")
    chkIndoors.Value = STR(Map(GetPlayerMap(MyIndex)).Indoors)
    CmbWeather.ListIndex = Map(GetPlayerMap(MyIndex)).Weather
    
    For x = 1 To 15
        cmbNpc(x - 1).addItem "No NPC"
    Next x
    
    For y = 1 To MAX_NPCS
        For x = 1 To 15
            cmbNpc(x - 1).addItem y & ": " & Trim$(Npc(y).Name)
        Next x
    Next y
    
    For i = 1 To 15
        cmbNpc(i - 1).ListIndex = Map(GetPlayerMap(MyIndex)).Npc(i)
    Next i
    
    Call StopBGM
End Sub

Private Sub cmdOk_Click()
Dim i As Integer

    Map(GetPlayerMap(MyIndex)).Name = txtName.Text
    Map(GetPlayerMap(MyIndex)).Up = val#(txtup.Text)
    Map(GetPlayerMap(MyIndex)).Down = val#(txtDown.Text)
    Map(GetPlayerMap(MyIndex)).Left = val#(txtLeft.Text)
    Map(GetPlayerMap(MyIndex)).right = val#(txtRight.Text)
    Map(GetPlayerMap(MyIndex)).Moral = cmbMoral.ListIndex
    If Check1.Value = 0 Then
    Map(GetPlayerMap(MyIndex)).Music = lstMusic.Text
    Else
    If Not Left(Text2.Text, 7) = "http://" Then
    Text2.Text = "http://" & Text2.Text
    End If
    Map(GetPlayerMap(MyIndex)).Music = Text2.Text
    End If
    Map(GetPlayerMap(MyIndex)).BootMap = val#(txtBootMap.Text)
    Map(GetPlayerMap(MyIndex)).BootX = val#(txtBootX.Text)
    Map(GetPlayerMap(MyIndex)).BootY = val#(txtBootY.Text)
    Map(GetPlayerMap(MyIndex)).Indoors = val#(chkIndoors.Value)
    Map(GetPlayerMap(MyIndex)).Weather = CmbWeather.ListIndex
    
    For i = 1 To 15
        Map(GetPlayerMap(MyIndex)).Npc(i) = cmbNpc(i - 1).ListIndex
    Next i
    
    Call StopBGM
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Call StopBGM
    Unload Me
End Sub

Private Sub Text2_Change()
Text2.MaxLength = 1024
End Sub
