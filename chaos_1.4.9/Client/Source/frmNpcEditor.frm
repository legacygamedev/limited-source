VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmNpcEditor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Npc Editor"
   ClientHeight    =   7230
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   10080
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
   ScaleHeight     =   482
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   672
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   7215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10005
      _ExtentX        =   17648
      _ExtentY        =   12726
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "frmNpcEditor.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label14"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtAttackSay"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtName"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "fraSpell"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "fraQuest"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Drop Editor"
      TabPicture(1)   =   "frmNpcEditor.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label30"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Picture9"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame3"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin VB.Frame Frame3 
         Caption         =   "Settings"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Left            =   120
         TabIndex        =   86
         Top             =   480
         Width           =   4695
         Begin VB.TextBox txtChance 
            Alignment       =   1  'Right Justify
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
            Left            =   2640
            TabIndex        =   90
            Text            =   "0"
            Top             =   1800
            Width           =   1815
         End
         Begin VB.HScrollBar scrlValue 
            Height          =   255
            Left            =   960
            Max             =   10000
            TabIndex        =   89
            Top             =   1440
            Value           =   1
            Width           =   3255
         End
         Begin VB.HScrollBar scrlNum 
            Height          =   255
            Left            =   960
            Max             =   500
            TabIndex        =   88
            Top             =   1080
            Value           =   1
            Width           =   3255
         End
         Begin VB.HScrollBar scrlDropItem 
            Height          =   255
            Left            =   960
            Max             =   5
            Min             =   1
            TabIndex        =   87
            Top             =   360
            Value           =   1
            Width           =   3255
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Drop Item Chance 1 out of :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   720
            TabIndex        =   99
            Top             =   1800
            Width           =   1815
         End
         Begin VB.Label lblValue 
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
            Left            =   4320
            TabIndex        =   98
            Top             =   1440
            Width           =   75
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "Value :"
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
            Left            =   120
            TabIndex        =   97
            Top             =   1440
            Width           =   735
         End
         Begin VB.Label lblNum 
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
            Left            =   4320
            TabIndex        =   96
            Top             =   1080
            Width           =   75
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            Caption         =   "Number :"
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
            Left            =   120
            TabIndex        =   95
            Top             =   1080
            Width           =   735
         End
         Begin VB.Label lblItemName 
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
            Left            =   960
            TabIndex        =   94
            Top             =   720
            Width           =   3495
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            Caption         =   "Item :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   93
            Top             =   720
            Width           =   735
         End
         Begin VB.Label lblDropItem 
            AutoSize        =   -1  'True
            Caption         =   "1"
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
            Left            =   4320
            TabIndex        =   92
            Top             =   360
            Width           =   75
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            Caption         =   "Dropping :"
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
            Left            =   120
            TabIndex        =   91
            Top             =   360
            Width           =   735
         End
      End
      Begin VB.Frame fraQuest 
         Caption         =   "Quest-Scripts-Speech"
         Height          =   1815
         Left            =   -70080
         TabIndex        =   66
         Top             =   1080
         Width           =   4695
         Begin VB.HScrollBar scrlScript 
            Height          =   135
            Left            =   120
            Max             =   10000
            TabIndex        =   73
            Top             =   1560
            Value           =   1
            Width           =   3255
         End
         Begin VB.HScrollBar scrlSpeech 
            Height          =   135
            Left            =   120
            Max             =   10
            TabIndex        =   70
            Top             =   1080
            Width           =   3255
         End
         Begin VB.HScrollBar scrlquest 
            Height          =   135
            Left            =   120
            TabIndex        =   67
            Top             =   600
            Width           =   2775
         End
         Begin VB.Label lblSpeechName 
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
            Left            =   840
            TabIndex        =   81
            Top             =   720
            Width           =   3495
         End
         Begin VB.Label lblScript 
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
            Height          =   255
            Left            =   3480
            TabIndex        =   74
            Top             =   1440
            Width           =   255
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            Caption         =   "Script:"
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
            Left            =   120
            TabIndex        =   72
            Top             =   1320
            Width           =   495
         End
         Begin VB.Label lblSpeech 
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
            Height          =   255
            Left            =   3480
            TabIndex        =   71
            Top             =   960
            Width           =   135
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            Caption         =   "Speech :"
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
            Left            =   120
            TabIndex        =   69
            Top             =   840
            Width           =   735
         End
         Begin VB.Label Label21 
            Caption         =   "Label21"
            Height          =   375
            Left            =   120
            TabIndex        =   68
            Top             =   240
            Width           =   3735
         End
      End
      Begin VB.Frame fraSpell 
         Caption         =   "Spell"
         Height          =   975
         Left            =   -70080
         TabIndex        =   63
         Top             =   2880
         Width           =   4695
         Begin VB.HScrollBar scrlSpell 
            Height          =   135
            Left            =   120
            TabIndex        =   64
            Top             =   600
            Width           =   2895
         End
         Begin VB.Label lblSpell 
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
            Height          =   255
            Left            =   3120
            TabIndex        =   65
            Top             =   480
            Width           =   495
         End
      End
      Begin VB.TextBox txtName 
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
         Left            =   -74040
         TabIndex        =   60
         Top             =   360
         Width           =   3975
      End
      Begin VB.TextBox txtAttackSay 
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
         Left            =   -74040
         TabIndex        =   59
         Top             =   720
         Width           =   3975
      End
      Begin VB.Frame Frame1 
         Caption         =   "General Information"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6255
         Left            =   -74880
         TabIndex        =   14
         Top             =   960
         Width           =   4695
         Begin VB.CheckBox chkDay 
            Caption         =   "Day"
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
            TabIndex        =   85
            Top             =   1200
            Value           =   1  'Checked
            Width           =   615
         End
         Begin VB.CheckBox chkNight 
            Caption         =   "Night"
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
            TabIndex        =   84
            Top             =   960
            Value           =   1  'Checked
            Width           =   735
         End
         Begin VB.HScrollBar scrlElement 
            Height          =   135
            Left            =   960
            Max             =   1000
            TabIndex        =   80
            Top             =   6000
            Value           =   1
            Width           =   2775
         End
         Begin VB.ComboBox cmbBehavior 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            ItemData        =   "frmNpcEditor.frx":0038
            Left            =   2760
            List            =   "frmNpcEditor.frx":0057
            Style           =   2  'Dropdown List
            TabIndex        =   78
            Top             =   1560
            Width           =   1935
         End
         Begin VB.TextBox txtSpawnSecs 
            Alignment       =   1  'Right Justify
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
            Left            =   2760
            TabIndex        =   76
            Text            =   "0"
            Top             =   960
            Width           =   1215
         End
         Begin VB.HScrollBar scrlMAGI 
            Enabled         =   0   'False
            Height          =   135
            Left            =   1080
            Max             =   10000
            TabIndex        =   33
            Top             =   3000
            Width           =   2895
         End
         Begin VB.HScrollBar scrlSPEED 
            Enabled         =   0   'False
            Height          =   135
            Left            =   1080
            Max             =   10000
            TabIndex        =   32
            Top             =   2760
            Width           =   2895
         End
         Begin VB.HScrollBar scrlDEF 
            Height          =   135
            Left            =   1080
            Max             =   10000
            TabIndex        =   31
            Top             =   2520
            Width           =   2895
         End
         Begin VB.HScrollBar scrlSTR 
            Height          =   135
            Left            =   1080
            Max             =   10000
            TabIndex        =   30
            Top             =   2280
            Width           =   2895
         End
         Begin VB.HScrollBar scrlRange 
            Height          =   135
            Left            =   1080
            Max             =   30
            TabIndex        =   29
            Top             =   2040
            Value           =   1
            Width           =   2895
         End
         Begin VB.HScrollBar scrlSprite 
            Height          =   255
            Left            =   1080
            Max             =   500
            TabIndex        =   28
            Top             =   360
            Width           =   2895
         End
         Begin VB.PictureBox picSprites 
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
            Height          =   960
            Left            =   1800
            ScaleHeight     =   64
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   27
            Top             =   960
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.CheckBox BigNpc 
            Caption         =   "Big NPC"
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
            Left            =   120
            TabIndex        =   26
            Top             =   1560
            Width           =   855
         End
         Begin VB.HScrollBar StartHP 
            Height          =   135
            Left            =   1080
            TabIndex        =   25
            Top             =   3240
            Value           =   1
            Width           =   2895
         End
         Begin VB.HScrollBar ExpGive 
            Height          =   135
            Left            =   1080
            TabIndex        =   24
            Top             =   3480
            Width           =   2895
         End
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   1125
            Left            =   1440
            ScaleHeight     =   73
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   73
            TabIndex        =   22
            Top             =   840
            Width           =   1125
            Begin VB.PictureBox picSprite 
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
               Height          =   960
               Left            =   330
               ScaleHeight     =   64
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   23
               Top             =   75
               Width           =   480
            End
         End
         Begin VB.HScrollBar scrlAlign 
            Height          =   135
            Left            =   1080
            TabIndex        =   21
            Top             =   3720
            Width           =   2895
         End
         Begin VB.HScrollBar scrlDIR 
            Height          =   135
            Left            =   1080
            TabIndex        =   20
            Top             =   3960
            Width           =   2895
         End
         Begin VB.CheckBox chkPoison 
            Caption         =   "Poison Player"
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
            Left            =   720
            TabIndex        =   19
            Top             =   4320
            Width           =   1335
         End
         Begin VB.CheckBox chkDisease 
            Caption         =   "Disease Player"
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
            Left            =   2400
            TabIndex        =   18
            Top             =   4320
            Width           =   1455
         End
         Begin VB.HScrollBar scrlElementDamage 
            Height          =   135
            Left            =   120
            Max             =   1000
            TabIndex        =   17
            Top             =   4920
            Value           =   1
            Width           =   2895
         End
         Begin VB.TextBox txtMS 
            Height          =   390
            Left            =   120
            TabIndex        =   16
            Top             =   5400
            Width           =   1815
         End
         Begin VB.TextBox txtInterval 
            Height          =   390
            Left            =   2520
            TabIndex        =   15
            Top             =   5400
            Width           =   1815
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Spawn Time :"
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
            Left            =   120
            TabIndex        =   83
            Top             =   720
            Width           =   885
         End
         Begin VB.Label lblElement 
            AutoSize        =   -1  'True
            Caption         =   "None"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3840
            TabIndex        =   82
            Top             =   5880
            Width           =   690
         End
         Begin VB.Label Label2134343 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Element:"
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
            TabIndex        =   79
            Top             =   6000
            Width           =   555
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Behavior :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   135
            Left            =   2640
            TabIndex        =   77
            Top             =   1320
            Width           =   735
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            Caption         =   "Spawn Rate (in seconds) :"
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
            TabIndex        =   75
            Top             =   720
            Width           =   1815
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            Caption         =   "Magic :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   135
            Left            =   360
            TabIndex        =   58
            Top             =   3000
            Width           =   615
         End
         Begin VB.Label lblMAGI 
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
            Height          =   255
            Left            =   4080
            TabIndex        =   57
            Top             =   3000
            Width           =   495
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            Caption         =   "Speed :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   135
            Left            =   360
            TabIndex        =   56
            Top             =   2760
            Width           =   615
         End
         Begin VB.Label lblSPEED 
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
            Height          =   255
            Left            =   4080
            TabIndex        =   55
            Top             =   2760
            Width           =   495
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            Caption         =   "Defence :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1200
            TabIndex        =   54
            Top             =   2520
            Width           =   615
         End
         Begin VB.Label lblDEF 
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
            Height          =   255
            Left            =   4080
            TabIndex        =   53
            Top             =   2520
            Width           =   495
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Strength :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   135
            Left            =   360
            TabIndex        =   52
            Top             =   2280
            Width           =   615
         End
         Begin VB.Label lblSTR 
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
            Height          =   255
            Left            =   4080
            TabIndex        =   51
            Top             =   2280
            Width           =   495
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Sight (By Tile):"
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
            Left            =   120
            TabIndex        =   50
            Top             =   1920
            Width           =   975
         End
         Begin VB.Label lblRange 
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
            Height          =   255
            Left            =   4080
            TabIndex        =   49
            Top             =   2040
            Width           =   495
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "Sprite :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   360
            TabIndex        =   48
            Top             =   360
            Width           =   615
         End
         Begin VB.Label lblSprite 
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
            Height          =   375
            Left            =   4080
            TabIndex        =   47
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            Caption         =   "Starting Hp :"
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
            Left            =   120
            TabIndex        =   46
            Top             =   3240
            Width           =   855
         End
         Begin VB.Label lblStartHP 
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
            Height          =   255
            Left            =   4080
            TabIndex        =   45
            Top             =   3240
            Width           =   495
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            Caption         =   "Exp Given :"
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
            Left            =   120
            TabIndex        =   44
            Top             =   3480
            Width           =   855
         End
         Begin VB.Label lblExpGiven 
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
            Height          =   255
            Left            =   4080
            TabIndex        =   43
            Top             =   3480
            Width           =   495
         End
         Begin VB.Label Label23 
            Alignment       =   1  'Right Justify
            Caption         =   "Align Given :"
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
            Left            =   120
            TabIndex        =   42
            Top             =   3720
            Width           =   855
         End
         Begin VB.Label lblAlign 
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
            Height          =   255
            Left            =   4080
            TabIndex        =   41
            Top             =   3720
            Width           =   495
         End
         Begin VB.Label Label22 
            Caption         =   "S. Direction :"
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
            Left            =   120
            TabIndex        =   40
            Top             =   3960
            Width           =   1095
         End
         Begin VB.Label lblDIR 
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
            Height          =   255
            Left            =   4080
            TabIndex        =   39
            Top             =   3960
            Width           =   495
         End
         Begin VB.Label Label24 
            Alignment       =   1  'Right Justify
            Caption         =   "Defense :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   135
            Left            =   360
            TabIndex        =   38
            Top             =   2520
            Width           =   615
         End
         Begin VB.Label Label51 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Poison/Disease Damage:"
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
            Left            =   120
            TabIndex        =   37
            Top             =   4680
            Width           =   1545
         End
         Begin VB.Label lblElementDamage 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
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
            Height          =   375
            Left            =   2520
            TabIndex        =   36
            Top             =   4680
            Width           =   495
         End
         Begin VB.Label Label54 
            AutoSize        =   -1  'True
            Caption         =   "Time to Effect Target (Seconds):"
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
            Left            =   120
            TabIndex        =   35
            Top             =   5160
            Width           =   2025
         End
         Begin VB.Label Label53 
            AutoSize        =   -1  'True
            Caption         =   "Effect Target every (Milliseconds):"
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
            Left            =   2400
            TabIndex        =   34
            Top             =   5160
            Width           =   2130
         End
      End
      Begin VB.PictureBox Picture9 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1935
         Left            =   960
         ScaleHeight     =   129
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   169
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   3600
         Width           =   2535
         Begin VB.PictureBox picInv 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   480
            Index           =   10
            Left            =   720
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   1320
            Width           =   480
         End
         Begin VB.PictureBox picInv 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   480
            Index           =   9
            Left            =   1920
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   720
            Width           =   480
         End
         Begin VB.PictureBox picInv 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   480
            Index           =   2
            Left            =   120
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   120
            Width           =   480
         End
         Begin VB.PictureBox picInv 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   480
            Index           =   3
            Left            =   720
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   120
            Width           =   480
         End
         Begin VB.PictureBox picInv 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   480
            Index           =   4
            Left            =   1320
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   120
            Width           =   480
         End
         Begin VB.PictureBox picInv 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   480
            Index           =   5
            Left            =   1920
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   120
            Width           =   480
         End
         Begin VB.PictureBox picInv 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   480
            Index           =   6
            Left            =   120
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   720
            Width           =   480
         End
         Begin VB.PictureBox picInv 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   480
            Index           =   7
            Left            =   720
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   720
            Width           =   480
         End
         Begin VB.PictureBox picInv 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   480
            Index           =   8
            Left            =   1320
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   720
            Width           =   480
         End
         Begin VB.PictureBox picInv 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   480
            Index           =   11
            Left            =   1320
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   1320
            Width           =   480
         End
         Begin VB.PictureBox picItems 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   480
            Left            =   2400
            ScaleHeight     =   480
            ScaleWidth      =   465
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   1920
            Visible         =   0   'False
            Width           =   465
         End
         Begin VB.Shape SelectedItem 
            BorderColor     =   &H000000FF&
            BorderWidth     =   2
            Height          =   510
            Left            =   120
            Top             =   120
            Width           =   510
         End
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Name :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74760
         TabIndex        =   62
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         Caption         =   "Speak :"
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
         Left            =   -74880
         TabIndex        =   61
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label30 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Npc Inventory :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   960
         TabIndex        =   13
         Top             =   3360
         Width           =   1155
      End
   End
   Begin VB.Timer tmrSprite 
      Interval        =   50
      Left            =   7200
      Top             =   5160
   End
   Begin VB.Menu FileMenu 
      Caption         =   "File"
      Begin VB.Menu SaveMenu 
         Caption         =   "Save && Exit"
      End
      Begin VB.Menu ExitMenu 
         Caption         =   "Exit Without Saving"
      End
   End
End
Attribute VB_Name = "frmNpcEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Copyright (c) 2006 Chaos Engine Source. All rights reserved.
' This code is licensed under the Chaos Engine General License.

Option Explicit

Private Sub BigNpc_Click()
Dim sDc As Long

    frmNpcEditor.ScaleMode = 3
    If BigNpc.Value = Checked Then
        sDc = DD_BigSpriteSurf.GetDC
        With picSprite
            .Width = 64
            .Height = 64
            .Left = (73 - 64) / 2 ' "73" is the scale width/height of Picture 1
            .Top = (73 - 64) / 2
            .Cls
            Call BitBlt(.hDC, 0, 0, .Width, .Height, sDc, 0, 0, SRCCOPY)
        End With
        Call DD_BigSpriteSurf.ReleaseDC(sDc)
    Else
        sDc = DD_SpriteSurf.GetDC
        With picSprite
            .Width = 32
            .Height = 64
            .Left = (73 - SIZE_X) / 2
            .Top = (73 - SIZE_Y) / 2
            .Cls
            Call BitBlt(.hDC, 0, 0, .Width, .Height, sDc, 0, 0, SRCCOPY)
        End With
        Call DD_SpriteSurf.ReleaseDC(sDc)
    End If
End Sub

Private Sub chkDay_Click()
    If chkNight.Value = Unchecked Then
        If chkDay.Value = Unchecked Then
            chkDay.Value = Checked
        End If
    End If
End Sub

Private Sub chkNight_Click()
    If chkDay.Value = Unchecked Then
        If chkNight.Value = Unchecked Then
            chkNight.Value = Checked
        End If
    End If
End Sub

Private Sub cmbBehavior_Click()
If cmbBehavior.ListIndex = NPC_BEHAVIOR_SCRIPTED Then
Label20.Visible = True
lblScript.Visible = True
scrlScript.Visible = True
Else
Label20.Visible = False
lblScript.Visible = False
scrlScript.Visible = False
End If
End Sub

Private Sub ExitMenu_Click()
Call NpcEditorCancel
End Sub

Private Sub ExpGive_Change()
    lblExpGiven.Caption = ExpGive.Value
End Sub

Private Sub Form_Load()
    scrlElement.Max = MAX_ELEMENTS
    scrlDropItem.Max = MAX_NPC_DROPS
    frmNpcEditor.scrlquest.Max = MAX_QUESTS
    scrlSpell.Max = MAX_SPELLS
End Sub

Private Sub picInv_Click(Index As Integer)
If Index = 1 Then Exit Sub

DropIndex = Index + 1

If Index = 2 Then
scrlDropItem.Value = 1
End If

If Index = 3 Then
scrlDropItem.Value = 2
End If

If Index = 4 Then
scrlDropItem.Value = 3
End If

If Index = 5 Then
scrlDropItem.Value = 4
End If

If Index = 6 Then
scrlDropItem.Value = 5
End If

If Index = 7 Then
scrlDropItem.Value = 6
End If

If Index = 8 Then
scrlDropItem.Value = 7
End If

If Index = 9 Then
scrlDropItem.Value = 8
End If

If Index = 10 Then
scrlDropItem.Value = 9
End If

If Index = 11 Then
scrlDropItem.Value = 10
End If

    frmNpcEditor.SelectedItem.Top = frmNpcEditor.picInv(DropIndex - 1).Top - 1
    frmNpcEditor.SelectedItem.Left = frmNpcEditor.picInv(DropIndex - 1).Left - 1
'scrlDropItem.Value = DropIndex
End Sub

Private Sub SaveMenu_Click()
Call NpcEditorOk
End Sub

Private Sub scrlAlign_Change()
lblAlign.Caption = scrlAlign.Value
End Sub

Private Sub scrlDIR_Change()
lblDIR.Caption = scrlDIR.Value
End Sub

Private Sub scrlElementDamage_Change()
lblElementDamage.Caption = scrlElementDamage.Value
End Sub

Private Sub scrlquest_Change()
If scrlquest.Value = 0 Then
frmNpcEditor.Label21.Caption = "None"
Else
frmNpcEditor.Label21.Caption = "Quest Number " & frmNpcEditor.scrlquest.Value
End If
End Sub

Private Sub scrlDropItem_Change()
    txtChance.Text = Npc(EditorIndex).ItemNPC(scrlDropItem.Value).Chance
    scrlNum.Value = Npc(EditorIndex).ItemNPC(scrlDropItem.Value).itemnum
    scrlValue.Value = Npc(EditorIndex).ItemNPC(scrlDropItem.Value).itemvalue
    lblDropItem.Caption = scrlDropItem.Value
End Sub

Private Sub scrlElement_Change()
    lblElement.Caption = Element(scrlElement.Value).Name
End Sub

Private Sub scrlSpell_Change()
If scrlSpell.Value > 0 Then
lblSpell.Caption = Trim(Spell(scrlSpell.Value).Name)
Else
lblSpell.Caption = "None"
End If
End Sub

Private Sub scrlSprite_Change()
    lblSprite.Caption = STR(scrlSprite.Value)
End Sub

Private Sub scrlRange_Change()
    lblRange.Caption = STR(scrlRange.Value)
End Sub

Private Sub scrlSTR_Change()
    lblSTR.Caption = STR(scrlSTR.Value)
End Sub

Private Sub scrlDEF_Change()
    lblDEF.Caption = STR(scrlDEF.Value)
End Sub

Private Sub scrlSPEED_Change()
    lblSPEED.Caption = STR(scrlSPEED.Value)
End Sub

Private Sub scrlMAGI_Change()
    lblMAGI.Caption = STR(scrlMAGI.Value)
End Sub

Private Sub scrlNum_Change()
    lblNum.Caption = STR(scrlNum.Value)
    LblItemName.Caption = ""
    If scrlNum.Value > 0 Then
      
      LblItemName.Caption = Trim(Item(scrlNum.Value).Name)
      
     If scrlDropItem.Value = 1 Then
       picDrop(1) = scrlNum.Value
     End If
     
     If scrlDropItem.Value = 2 Then
       picDrop(2) = scrlNum.Value
     End If
     
     If scrlDropItem.Value = 3 Then
       picDrop(3) = scrlNum.Value
     End If
       
     If scrlDropItem.Value = 4 Then
       picDrop(4) = scrlNum.Value
     End If
     
     If scrlDropItem.Value = 5 Then
       picDrop(5) = scrlNum.Value
     End If
     
     If scrlDropItem.Value = 6 Then
       picDrop(6) = scrlNum.Value
     End If
     
     If scrlDropItem.Value = 7 Then
       picDrop(7) = scrlNum.Value
     End If
     
     If scrlDropItem.Value = 8 Then
       picDrop(8) = scrlNum.Value
     End If
     
     If scrlDropItem.Value = 9 Then
       picDrop(9) = scrlNum.Value
     End If
     
     If scrlDropItem.Value = 10 Then
       picDrop(10) = scrlNum.Value
     End If
       
    End If
    Npc(EditorIndex).ItemNPC(scrlDropItem.Value).itemnum = scrlNum.Value
End Sub

Private Sub scrlValue_Change()
    Npc(EditorIndex).ItemNPC(scrlDropItem.Value).itemvalue = scrlValue.Value
    lblValue.Caption = STR(scrlValue.Value)
End Sub

Private Sub scrlSpeech_Change()
    Npc(EditorIndex).Speech = scrlSpeech.Value
    lblSpeech.Caption = STR(scrlSpeech.Value)
    If scrlSpeech.Value > 0 Then
        lblSpeechName.Caption = Speech(scrlSpeech.Value).Name
    Else
        lblSpeechName.Caption = ""
    End If
End Sub

Private Sub StartHP_Change()
    lblStartHP.Caption = StartHP.Value
End Sub

Private Sub tmrSprite_Timer()
    Call NpcEditorBltSprite
    Call NpcEditorBltInv
End Sub

Private Sub txtChance_Change()
    Npc(EditorIndex).ItemNPC(scrlDropItem.Value).Chance = Val(txtChance.Text)
End Sub

Private Sub scrlScript_Change()
lblScript.Caption = STR(scrlScript.Value)
End Sub
