VERSION 5.00
Begin VB.Form frmSpellEditor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Spell Editor"
   ClientHeight    =   6255
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13950
   ControlBox      =   0   'False
   DrawMode        =   14  'Copy Pen
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSpellEditor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   13950
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmLevelReq 
      Caption         =   "Level Req"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   47
      Top             =   840
      Width           =   4815
      Begin VB.HScrollBar scrlLevel 
         Height          =   255
         Left            =   1080
         Max             =   100
         TabIndex        =   48
         Top             =   360
         Value           =   1
         Width           =   2655
      End
      Begin VB.Label lblLevel 
         Alignment       =   1  'Right Justify
         Caption         =   "Level Req:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   50
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblLevelReq 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3960
         TabIndex        =   49
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.Frame frmClassReq 
      Caption         =   "Class Reqs."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   45
      Top             =   1560
      Width           =   4815
      Begin VB.CheckBox chkClass 
         Caption         =   "Class Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   120
         TabIndex        =   46
         Top             =   360
         Width           =   4215
      End
   End
   Begin VB.Frame fraVitalReq 
      Caption         =   "Vital Reqs. (Casting)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   44
      Top             =   2280
      Width           =   4815
      Begin VB.HScrollBar scrlVitalReq 
         Height          =   255
         Index           =   0
         Left            =   1080
         Max             =   1000
         TabIndex        =   54
         Top             =   480
         Visible         =   0   'False
         Width           =   2640
      End
      Begin VB.TextBox txtVitalReq 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   3840
         TabIndex        =   53
         Text            =   "0"
         Top             =   480
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "How much of a vital to take away when the spell is cast."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   56
         Top             =   240
         Width           =   4335
      End
      Begin VB.Label lblVitalReqName 
         Caption         =   "Vital Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   55
         Top             =   495
         Visible         =   0   'False
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "Calc Level - Vital only"
      Height          =   270
      Left            =   10200
      TabIndex        =   43
      Top             =   1920
      Width           =   2535
   End
   Begin VB.Frame fra 
      Caption         =   "Targeting"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   9840
      TabIndex        =   34
      Top             =   0
      Width           =   3255
      Begin VB.CheckBox chkTargets 
         Caption         =   "Targets"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   5
         Left            =   240
         TabIndex        =   57
         Top             =   1440
         Width           =   2055
      End
      Begin VB.CheckBox chkTargets 
         Caption         =   "Targets"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   4
         Left            =   240
         TabIndex        =   42
         Top             =   1200
         Width           =   2055
      End
      Begin VB.CheckBox chkTargets 
         Caption         =   "Targets"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   3
         Left            =   240
         TabIndex        =   38
         Top             =   960
         Width           =   2055
      End
      Begin VB.CheckBox chkTargets 
         Caption         =   "Targets"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   2
         Left            =   240
         TabIndex        =   37
         Top             =   720
         Width           =   2055
      End
      Begin VB.CheckBox chkTargets 
         Caption         =   "Targets"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   240
         TabIndex        =   36
         Top             =   480
         Width           =   2055
      End
      Begin VB.CheckBox chkTargets 
         Caption         =   "Targets"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   240
         TabIndex        =   35
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Overtime and Buff Spells"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   5040
      TabIndex        =   24
      Top             =   1440
      Width           =   4695
      Begin VB.HScrollBar scrlTickCount 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   26
         Top             =   1080
         Value           =   1
         Width           =   3855
      End
      Begin VB.HScrollBar scrlTickUpdate 
         Height          =   255
         Left            =   120
         Max             =   20
         TabIndex        =   25
         Top             =   1680
         Value           =   1
         Width           =   3855
      End
      Begin VB.Label Label5 
         Caption         =   $"frmSpellEditor.frx":0E42
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   33
         Top             =   240
         Width           =   4335
      End
      Begin VB.Label lblTickCount 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3960
         TabIndex        =   32
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label lblTickCountName 
         Caption         =   "Tick Counts (How many Ticks)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   31
         Top             =   840
         Width           =   2655
      End
      Begin VB.Label lblTickUpdate 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3960
         TabIndex        =   30
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Tick Update (Time between ticks in seconds)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   1440
         Width           =   3975
      End
      Begin VB.Label lblTotalTickTimeName 
         Caption         =   "Total Time for spell (Tick Count * Tick Update)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   2040
         Width           =   3735
      End
      Begin VB.Label lblTotalTime 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3960
         TabIndex        =   27
         Top             =   2040
         Width           =   495
      End
   End
   Begin VB.Frame fraCasting 
      Caption         =   "Casting"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   5040
      TabIndex        =   15
      Top             =   0
      Width           =   4695
      Begin VB.HScrollBar scrlRange 
         Height          =   255
         Left            =   1440
         Max             =   255
         TabIndex        =   39
         Top             =   1080
         Value           =   1
         Width           =   2535
      End
      Begin VB.HScrollBar scrlCastTime 
         Height          =   255
         Left            =   1440
         Max             =   20
         TabIndex        =   17
         Top             =   360
         Value           =   1
         Width           =   2535
      End
      Begin VB.HScrollBar scrlCooldown 
         Height          =   255
         Left            =   1440
         TabIndex        =   16
         Top             =   720
         Value           =   1
         Width           =   2535
      End
      Begin VB.Label Label9 
         Caption         =   "Range"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label lblRange 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3960
         TabIndex        =   40
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label7 
         Caption         =   "CastTime (Sec)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lblCastTime 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3960
         TabIndex        =   20
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lblCooldown 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3960
         TabIndex        =   19
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label10 
         Caption         =   "Cooldown (Sec)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   720
         Width           =   1215
      End
   End
   Begin VB.Frame fraModStat 
      Caption         =   "Mod Stats"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5040
      TabIndex        =   12
      Top             =   5160
      Width           =   4695
      Begin VB.TextBox txtModStat 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   3840
         TabIndex        =   51
         Text            =   "0"
         Top             =   480
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.HScrollBar scrlModStat 
         Height          =   255
         Index           =   0
         Left            =   1080
         Max             =   255
         Min             =   -255
         TabIndex        =   13
         Top             =   480
         Visible         =   0   'False
         Width           =   2640
      End
      Begin VB.Label Label4 
         Caption         =   "Only used for Buffs: Positive for buffs - Negative for Debuffs"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   4335
      End
      Begin VB.Label lblModStatName 
         Caption         =   "Stat Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   480
         Visible         =   0   'False
         Width           =   855
      End
   End
   Begin VB.Frame fraModVitals 
      Caption         =   "Mod Vitals"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   5040
      TabIndex        =   9
      Top             =   3840
      Width           =   4695
      Begin VB.TextBox txtModVital 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   3840
         TabIndex        =   52
         Text            =   "0"
         Top             =   960
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.HScrollBar scrlModVital 
         Height          =   255
         Index           =   0
         Left            =   1080
         Max             =   1000
         Min             =   -1000
         TabIndex        =   10
         Top             =   960
         Visible         =   0   'False
         Width           =   2640
      End
      Begin VB.Label Label2 
         Caption         =   $"frmSpellEditor.frx":0F0D
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   4335
      End
      Begin VB.Label lblModVitalName 
         Caption         =   "Vital Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   975
         Visible         =   0   'False
         Width           =   855
      End
   End
   Begin VB.HScrollBar scrlAnim 
      Height          =   255
      Left            =   9840
      Max             =   255
      TabIndex        =   6
      Top             =   4200
      Value           =   1
      Width           =   3495
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11760
      TabIndex        =   5
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12840
      TabIndex        =   4
      Top             =   5880
      Width           =   975
   End
   Begin VB.Timer tmrAnimation 
      Interval        =   5
      Left            =   13560
      Top             =   0
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   9840
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   3
      Top             =   2640
      Width           =   480
   End
   Begin VB.ComboBox cmbType 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmSpellEditor.frx":0FB0
      Left            =   120
      List            =   "frmSpellEditor.frx":0FC0
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   480
      Width           =   4815
   End
   Begin VB.TextBox txtName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   3975
   End
   Begin VB.Label Label8 
      Caption         =   "Animation"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9840
      TabIndex        =   8
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label lblAnim 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   13440
      TabIndex        =   7
      Top             =   4200
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmSpellEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private CurrFrame As Byte
Private LastUpdate As Long

Private Sub cmbType_Click()
Dim i As Long

    ' We reset these everytime because it's an easy way to do it
    For i = 1 To Vitals.Vital_Count
        scrlModVital(i).Min = -1000
        scrlModVital(i).Max = 1000
        scrlModVital(i).Value = Spell(EditorIndex).ModVital(i)
    Next
            
    Select Case cmbType.ListIndex
        ' Check if the spell is revive to change the min/max on the vitals
        Case SPELL_TYPE_REVIVE
            For i = 1 To Vitals.Vital_Count
                scrlModVital(i).Min = 0
                scrlModVital(i).Max = 100
                scrlModVital(i).Value = 0
            Next
            
    End Select
End Sub

Private Sub cmdCalc_Click()
    
    Select Case cmbType.ListIndex
        Case SPELL_TYPE_VITAL
            scrlLevel.Value = Clamp(Abs(scrlModVital(Vitals.HP - 1).Value) \ 2.5, 0, 100)
            scrlVitalReq(Vitals.MP - 1).Value = (Abs(scrlModVital(Vitals.HP - 1).Value) \ (scrlCastTime.Value + 1.5)) + (scrlLevel.Value \ 3.5)
            
    End Select
End Sub

Private Sub cmdOk_Click()
    Call SpellEditorOk
End Sub

Private Sub cmdCancel_Click()
    Call SpellEditorCancel
End Sub

Private Sub scrlAnim_Change()
    lblAnim.Caption = scrlAnim.Value
End Sub

Private Sub scrlCastTime_Change()
    lblCastTime.Caption = scrlCastTime.Value
End Sub

Private Sub scrlCooldown_Change()
    lblCooldown.Caption = scrlCooldown.Value
End Sub

Private Sub scrlLevel_Change()
    lblLevelReq.Caption = scrlLevel.Value
End Sub

Private Sub scrlModStat_Change(Index As Integer)
    txtModStat(Index).Text = scrlModStat(Index).Value
End Sub

Private Sub scrlModVital_Change(Index As Integer)
    txtModVital(Index).Text = scrlModVital(Index).Value
End Sub

Private Sub scrlRange_Change()
    lblRange.Caption = scrlRange.Value
End Sub

Private Sub scrlTickCount_Change()
    lblTickCount.Caption = scrlTickCount.Value
    lblTotalTime.Caption = scrlTickCount.Value * scrlTickUpdate.Value
End Sub

Private Sub scrlTickUpdate_Change()
    lblTickUpdate.Caption = scrlTickUpdate.Value
    lblTotalTime.Caption = scrlTickCount.Value * scrlTickUpdate.Value
End Sub

Private Sub scrlVitalReq_Change(Index As Integer)
    txtVitalReq(Index).Text = scrlVitalReq(Index).Value
End Sub

Private Sub tmrAnimation_Timer()
Dim sRECT As RECT
Dim dRECT As RECT
Dim Anim As Byte, Frames As Byte, Speed As Byte, Size As Byte

    If scrlAnim.Value = 0 Then Exit Sub
    
    Anim = Animation(scrlAnim.Value).Animation
    Frames = Animation(scrlAnim.Value).AnimationFrames
    Speed = Animation(scrlAnim.Value).AnimationSpeed
    Size = Animation(scrlAnim.Value).AnimationSize
    
    If Size = 0 Then Exit Sub

    If Size = 1 Then
        picSprite.Height = 32 * 15
        picSprite.Width = 32 * 15
        
         With dRECT
            .Top = 0
            .Bottom = PIC_Y
            .Left = 0
            .Right = PIC_X
        End With
    
        With sRECT
            .Top = Anim * PIC_Y
            .Bottom = .Top + PIC_Y
            .Left = CurrFrame * PIC_X
            .Right = .Left + PIC_X
        End With
        
        DD_AnimationSurf.BltToDC picSprite.hdc, sRECT, dRECT
    ElseIf Size = 2 Then
        picSprite.Height = 64 * 15
        picSprite.Width = 64 * 15
        
        With dRECT
            .Top = 0
            .Bottom = (PIC_Y * 2)
            .Left = 0
            .Right = (PIC_X * 2)
        End With
        
        With sRECT
            .Top = Anim * (PIC_Y * 2)
            .Bottom = .Top + (PIC_Y * 2)
            .Left = CurrFrame * (PIC_X * 2)
            .Right = .Left + (PIC_X * 2)
        End With
        
        DD_AnimationSurf2.BltToDC picSprite.hdc, sRECT, dRECT
    End If
    
    picSprite.Refresh

    If GetTickCount > LastUpdate Then
    
        CurrFrame = CurrFrame + 1
        
        If CurrFrame > Frames Then
            CurrFrame = 0
        Else
            LastUpdate = GetTickCount + Speed
        End If
    End If
End Sub

Private Sub txtModStat_Change(Index As Integer)
    SetTextBox txtModStat(Index), scrlModStat(Index)
End Sub

Private Sub txtModVital_Change(Index As Integer)
    SetTextBox txtModVital(Index), scrlModVital(Index)
End Sub

Private Sub txtVitalReq_Change(Index As Integer)
    SetTextBox txtVitalReq(Index), scrlVitalReq(Index)
End Sub
