VERSION 5.00
Begin VB.Form frmEditor_Quest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Quest Editor"
   ClientHeight    =   10170
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   18975
   Icon            =   "frmEditor_Quest.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10170
   ScaleWidth      =   18975
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdChangeData 
      Caption         =   "Change Data Size"
      Height          =   375
      Left            =   120
      TabIndex        =   147
      Top             =   8280
      Width           =   2895
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   8280
      TabIndex        =   146
      Top             =   8280
      Width           =   1455
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   13200
      TabIndex        =   145
      Top             =   8280
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   3360
      TabIndex        =   144
      Top             =   8280
      Width           =   1455
   End
   Begin VB.Frame fmeVar 
      Caption         =   "Get/Set a game variable."
      Height          =   3375
      Left            =   15720
      TabIndex        =   132
      Top             =   6360
      Visible         =   0   'False
      Width           =   5175
      Begin VB.CheckBox chkPassVar 
         Caption         =   "Move on to next task automatically."
         Height          =   255
         Left            =   120
         TabIndex        =   88
         ToolTipText     =   $"frmEditor_Quest.frx":038A
         Top             =   2520
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.TextBox txtVar 
         Height          =   285
         Left            =   120
         MaxLength       =   60
         TabIndex        =   83
         ToolTipText     =   "This text will be displayed in the mission request and mission log to show the player what needs done."
         Top             =   2160
         Width           =   4935
      End
      Begin VB.CheckBox chkSetValue 
         Alignment       =   1  'Right Justify
         Caption         =   "Set this value."
         Height          =   255
         Left            =   3600
         TabIndex        =   78
         ToolTipText     =   "Check this box if you would like to SET the value of the selected variable rather than add to it."
         Top             =   1500
         Width           =   1335
      End
      Begin VB.HScrollBar scrlValue 
         Enabled         =   0   'False
         Height          =   255
         LargeChange     =   10
         Left            =   120
         TabIndex        =   80
         Top             =   1800
         Width           =   4935
      End
      Begin VB.CommandButton btnVAccept 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Accept"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   90
         Top             =   2760
         Width           =   2175
      End
      Begin VB.CommandButton btnVCancel 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   91
         Top             =   2760
         Width           =   2175
      End
      Begin VB.ComboBox cmbVars 
         Enabled         =   0   'False
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   76
         Top             =   1080
         Width           =   4935
      End
      Begin VB.OptionButton opSwitch 
         Caption         =   "Switches"
         Height          =   255
         Left            =   120
         TabIndex        =   74
         Top             =   720
         Width           =   4815
      End
      Begin VB.OptionButton OpVariable 
         Caption         =   "Varia&ble"
         Height          =   255
         Left            =   120
         TabIndex        =   70
         Top             =   360
         Width           =   4815
      End
      Begin VB.Label lblValue 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Value: 0"
         Height          =   195
         Left            =   120
         TabIndex        =   133
         Top             =   1500
         Width           =   585
      End
   End
   Begin VB.Frame fmeWarp 
      Caption         =   "Select map and location to warp the player to"
      Height          =   3015
      Left            =   15000
      TabIndex        =   126
      Top             =   120
      Visible         =   0   'False
      Width           =   3735
      Begin VB.CommandButton btnCloseWarp 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   2520
         Width           =   1335
      End
      Begin VB.CommandButton btnAddWarp 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Accept"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   2520
         Width           =   1335
      End
      Begin VB.HScrollBar scrlMapY 
         Height          =   255
         LargeChange     =   10
         Left            =   120
         Max             =   255
         TabIndex        =   16
         Top             =   2040
         Width           =   3495
      End
      Begin VB.HScrollBar scrlMapX 
         Height          =   255
         LargeChange     =   10
         Left            =   120
         Max             =   255
         TabIndex        =   11
         Top             =   1440
         Width           =   3495
      End
      Begin VB.HScrollBar scrlMap 
         Height          =   255
         LargeChange     =   10
         Left            =   120
         Max             =   255
         Min             =   1
         TabIndex        =   5
         Top             =   600
         Value           =   1
         Width           =   3495
      End
      Begin VB.Label lblMapY 
         Caption         =   "Y: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   129
         Top             =   1800
         Width           =   3495
      End
      Begin VB.Label lblMapX 
         Caption         =   "X: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   128
         Top             =   1200
         Width           =   3495
      End
      Begin VB.Label lblMap 
         Alignment       =   2  'Center
         Caption         =   "Map: 1"
         Height          =   255
         Left            =   120
         TabIndex        =   127
         Top             =   360
         Width           =   3495
      End
   End
   Begin VB.Frame fmeObtainSKill 
      Caption         =   "Select a skill and skill level for the player to obtain."
      Height          =   1935
      Left            =   15720
      TabIndex        =   122
      Top             =   480
      Visible         =   0   'False
      Width           =   3855
      Begin VB.CommandButton btnObAccept 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Accept"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   1320
         Width           =   1695
      End
      Begin VB.CommandButton btnObCancel 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1320
         Width           =   1695
      End
      Begin VB.HScrollBar scrlObSkill 
         Height          =   255
         LargeChange     =   5
         Left            =   1560
         Max             =   255
         TabIndex        =   10
         Top             =   840
         Width           =   2055
      End
      Begin VB.ComboBox cmbObSKill 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   7
         ToolTipText     =   "Select the specific skill the player will need to reach a level for."
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label lblObSkill 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Skill Level:"
         Height          =   195
         Left            =   240
         TabIndex        =   124
         Top             =   840
         Width           =   765
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Skill:"
         Height          =   195
         Left            =   360
         TabIndex        =   123
         Top             =   360
         Width           =   330
      End
   End
   Begin VB.Frame fmeModify 
      Caption         =   "Adjust Player Stats"
      Height          =   3855
      Left            =   15360
      TabIndex        =   117
      Top             =   5520
      Visible         =   0   'False
      Width           =   4575
      Begin VB.ComboBox cboItem 
         Enabled         =   0   'False
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   75
         Top             =   1920
         Width           =   4215
      End
      Begin VB.OptionButton opSkillEXP 
         Caption         =   "Skill EXP"
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   73
         ToolTipText     =   "Select to modify the player's Level"
         Top             =   1560
         Width           =   4095
      End
      Begin VB.OptionButton opSkill 
         Caption         =   "Skill Level"
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   71
         ToolTipText     =   "Select to modify the player's Level"
         Top             =   1320
         Width           =   4095
      End
      Begin VB.OptionButton opStatP 
         Caption         =   "Stat Points"
         Height          =   255
         Left            =   240
         TabIndex        =   68
         ToolTipText     =   "Select to modify the player's Level"
         Top             =   1080
         Width           =   4095
      End
      Begin VB.OptionButton opStat 
         Caption         =   "Stat"
         Height          =   255
         Left            =   240
         TabIndex        =   65
         ToolTipText     =   "Select to modify the player's Level"
         Top             =   840
         Width           =   4095
      End
      Begin VB.CommandButton btnModCancel 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   89
         Top             =   3360
         Width           =   1455
      End
      Begin VB.CommandButton btnModAccept 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Accept"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   87
         Top             =   3360
         Width           =   2775
      End
      Begin VB.CheckBox chkSet 
         Caption         =   "Set the value instead of adding/subtracting"
         Height          =   255
         Left            =   240
         TabIndex        =   77
         ToolTipText     =   "This option will decide whether we set the amount or add/subtract to the current amount."
         Top             =   2280
         Width           =   4215
      End
      Begin VB.HScrollBar scrlModify 
         Height          =   255
         LargeChange     =   25
         Left            =   120
         Min             =   -32767
         TabIndex        =   82
         Top             =   3000
         Width           =   4335
      End
      Begin VB.OptionButton opLvl 
         Caption         =   "Level"
         Height          =   255
         Left            =   240
         TabIndex        =   58
         ToolTipText     =   "Select to modify the player's Level"
         Top             =   600
         Width           =   4095
      End
      Begin VB.OptionButton opEXP 
         Caption         =   "EXP"
         Height          =   255
         Left            =   240
         TabIndex        =   57
         ToolTipText     =   "Select to modify the player's EXP"
         Top             =   360
         Width           =   4095
      End
      Begin VB.Label lblModify 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Amount to modify: 0"
         Height          =   255
         Left            =   240
         TabIndex        =   118
         Top             =   2640
         Width           =   4095
      End
   End
   Begin VB.Frame fmeSelectItem 
      Caption         =   "Select Item To Give"
      Height          =   2415
      Left            =   15000
      TabIndex        =   113
      Top             =   3240
      Visible         =   0   'False
      Width           =   3615
      Begin VB.ComboBox cmbItem 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   360
         Width           =   3375
      End
      Begin VB.CheckBox chkPassI 
         Caption         =   "Move on to next task automatically."
         Height          =   255
         Left            =   120
         TabIndex        =   38
         ToolTipText     =   $"frmEditor_Quest.frx":041A
         Top             =   1560
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.CheckBox chkTake 
         Caption         =   "Take the item the player gathers?"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   1320
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.HScrollBar scrlItemAmount 
         Height          =   255
         LargeChange     =   5
         Left            =   120
         TabIndex        =   28
         Top             =   960
         Width           =   3375
      End
      Begin VB.CommandButton btnAddItem 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Accept"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   1920
         Width           =   1335
      End
      Begin VB.CommandButton btnAItemCancel 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label lblAmount 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount: 0"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   114
         Top             =   720
         Width           =   3360
      End
   End
   Begin VB.Frame fmeMoveItem 
      Caption         =   "Move List Items"
      Height          =   2415
      Left            =   1560
      TabIndex        =   112
      Top             =   4320
      Visible         =   0   'False
      Width           =   1815
      Begin VB.CommandButton btnDeleteAction 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Delete"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   60
         ToolTipText     =   "Move the currently selected list item down."
         Top             =   1920
         Width           =   1575
      End
      Begin VB.CommandButton btnEditAction 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Edit"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   51
         ToolTipText     =   "Move the currently selected list item down."
         Top             =   1440
         Width           =   1575
      End
      Begin VB.CommandButton btnHide 
         BackColor       =   &H00FFFFFF&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   0
         Width           =   255
      End
      Begin VB.CommandButton btnDown 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Move Item Down"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Move the currently selected list item down."
         Top             =   840
         Width           =   1575
      End
      Begin VB.CommandButton btnUp 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Move Item Up"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Move the currently selected list item up."
         Top             =   360
         Width           =   1575
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   1680
         Y1              =   1320
         Y2              =   1320
      End
   End
   Begin VB.Frame fraNPC 
      Caption         =   "Quest List"
      Height          =   8175
      Left            =   0
      TabIndex        =   92
      Top             =   0
      Width           =   3135
      Begin VB.ListBox lstIndex 
         Height          =   7275
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   2895
      End
      Begin VB.CommandButton cmdCopy 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Copy"
         Height          =   315
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox txtSearch 
         CausesValidation=   0   'False
         Height          =   270
         Left            =   120
         TabIndex        =   0
         Top             =   300
         Width           =   1455
      End
      Begin VB.CommandButton cmdPaste 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Paste"
         Height          =   315
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Details"
      Height          =   2415
      Left            =   3240
      TabIndex        =   93
      Top             =   0
      Width           =   4455
      Begin VB.TextBox txtDesc 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1305
         Left            =   120
         MaxLength       =   300
         MultiLine       =   -1  'True
         TabIndex        =   8
         ToolTipText     =   "The description of the quest.  Will be seen in the player's quest viewer."
         Top             =   960
         Width           =   4215
      End
      Begin VB.TextBox txtName 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         MaxLength       =   40
         ScrollBars      =   1  'Horizontal
         TabIndex        =   3
         ToolTipText     =   "The name of the mission.  Will be seen in the player's mission log."
         Top             =   360
         Width           =   3615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description:"
         Height          =   195
         Left            =   120
         TabIndex        =   95
         Top             =   720
         UseMnemonic     =   0   'False
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         Height          =   195
         Left            =   120
         TabIndex        =   94
         Top             =   360
         UseMnemonic     =   0   'False
         Width           =   465
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Editing Features"
      Height          =   2415
      Left            =   7800
      TabIndex        =   96
      Top             =   0
      Width           =   6975
      Begin VB.CheckBox chkUnOrder 
         Caption         =   "Allow mission log to display tasks as the player completes them. (Out of order)"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         ToolTipText     =   "This sets whether or not the quest can be retaken again and again after completion."
         Top             =   2040
         Width           =   6735
      End
      Begin VB.TextBox txtRank 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1320
         MaxLength       =   5
         TabIndex        =   12
         ToolTipText     =   "The player will see this letter/number/text and associate it with the difficulty of the mission."
         Top             =   1680
         Width           =   1215
      End
      Begin VB.CheckBox chkRetake 
         Caption         =   "Can this mission be retaken?"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         ToolTipText     =   "This sets whether or not the quest can be retaken again and again after completion."
         Top             =   1320
         Width           =   4095
      End
      Begin VB.Timer tmrMsg 
         Interval        =   3000
         Left            =   1800
         Top             =   240
      End
      Begin VB.CommandButton btnReq 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Edit Requirements"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Edit the requirements for the player to be able to start the quest."
         Top             =   840
         Width           =   6735
      End
      Begin VB.Label Label5 
         Caption         =   "Mission Rank:"
         Height          =   255
         Left            =   120
         TabIndex        =   130
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label lblMsg 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Right-Click a list below for additional options."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   315
         Left            =   120
         TabIndex        =   111
         Top             =   360
         Width           =   6675
      End
   End
   Begin VB.Frame fmeTask 
      Caption         =   "Add a new action/task to complete for this Greeter"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   6360
      TabIndex        =   99
      Top             =   2640
      Visible         =   0   'False
      Width           =   8295
      Begin VB.CommandButton btnSound 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Play Sound"
         Height          =   375
         Left            =   6120
         Style           =   1  'Graphical
         TabIndex        =   142
         ToolTipText     =   "Set the value for a variable or a switch."
         Top             =   1440
         Width           =   2055
      End
      Begin VB.CommandButton cmdSetVariable 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Set Variable"
         Height          =   375
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Set the value for a variable or a switch."
         Top             =   1440
         Width           =   2055
      End
      Begin VB.CommandButton btnVariable 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Check Variable"
         Height          =   495
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Check to see if a variable or switch has reached a certain value."
         Top             =   2160
         Width           =   3015
      End
      Begin VB.CommandButton btnWarp 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Warp the player"
         Height          =   375
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   49
         ToolTipText     =   "Show a message to the player."
         Top             =   2880
         Width           =   2055
      End
      Begin VB.CommandButton btnMsgPlayer 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show message."
         Height          =   375
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Show a message to the player."
         Top             =   2400
         Width           =   2055
      End
      Begin VB.CommandButton btnAdjustStat 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Adjust Player Stat"
         Height          =   375
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Give or take stat values from the player such as Str/End/Exp/ect."
         Top             =   1920
         Width           =   2055
      End
      Begin VB.CommandButton btnProtect 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Spawn and protect ally."
         Enabled         =   0   'False
         Height          =   255
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   52
         ToolTipText     =   "Spawn an NPC to follow the player for so long, making all NPC's want to attack it, but the player must protect it for the quest."
         Top             =   3120
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.CommandButton btnSkillLvl 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Obtain a skill level."
         Enabled         =   0   'False
         Height          =   255
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   "Require that the player obtain a skill level between the progress of this quest"
         Top             =   2880
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.CommandButton btnTaskCancel 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   61
         Top             =   3600
         Width           =   6735
      End
      Begin VB.CommandButton btnTask_Kill 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Kill enemie(s)"
         Height          =   495
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Select an NPC the player needs to kill and an amount of times to kill it"
         Top             =   1560
         Width           =   3015
      End
      Begin VB.CommandButton btnTask_Gather 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Gather items"
         Height          =   495
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Select an item for the player to gather for the quest"
         Top             =   960
         Width           =   3015
      End
      Begin VB.CommandButton btnTakeItem 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Take an item."
         Height          =   375
         Left            =   6120
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Take an item from the player"
         Top             =   960
         Width           =   2055
      End
      Begin VB.CommandButton btnGiveItem 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Give an item."
         Height          =   375
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Give the player an item"
         Top             =   960
         Width           =   2055
      End
      Begin VB.Line Line9 
         X1              =   240
         X2              =   6960
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Line Line8 
         X1              =   3720
         X2              =   3720
         Y1              =   240
         Y2              =   3480
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Actions"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4080
         TabIndex        =   121
         Top             =   480
         Width           =   2775
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Tasks"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   120
         Top             =   480
         Width           =   2775
      End
      Begin VB.Line Line5 
         X1              =   120
         X2              =   8160
         Y1              =   240
         Y2              =   240
      End
   End
   Begin VB.Frame fmeReq 
      Caption         =   "Quest Requirements"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   6480
      TabIndex        =   100
      Top             =   4200
      Visible         =   0   'False
      Width           =   8295
      Begin VB.HScrollBar scrlSkill 
         Enabled         =   0   'False
         Height          =   255
         Left            =   5280
         Max             =   255
         TabIndex        =   37
         Top             =   600
         Width           =   1575
      End
      Begin VB.HScrollBar scrlStatReq 
         Height          =   255
         Index           =   5
         LargeChange     =   10
         Left            =   6360
         TabIndex        =   64
         Top             =   2160
         Width           =   1455
      End
      Begin VB.HScrollBar scrlStatReq 
         Height          =   255
         Index           =   4
         LargeChange     =   10
         Left            =   3720
         TabIndex        =   63
         Top             =   2160
         Width           =   1455
      End
      Begin VB.HScrollBar scrlStatReq 
         Height          =   255
         Index           =   3
         LargeChange     =   10
         Left            =   1200
         TabIndex        =   62
         Top             =   2160
         Width           =   1455
      End
      Begin VB.HScrollBar scrlStatReq 
         Height          =   255
         Index           =   2
         LargeChange     =   10
         Left            =   6360
         TabIndex        =   56
         Top             =   1680
         Width           =   1455
      End
      Begin VB.HScrollBar scrlStatReq 
         Height          =   255
         Index           =   1
         LargeChange     =   10
         Left            =   3720
         TabIndex        =   55
         Top             =   1680
         Width           =   1455
      End
      Begin VB.ComboBox cmbClassReq 
         Height          =   315
         ItemData        =   "frmEditor_Quest.frx":04AA
         Left            =   6360
         List            =   "frmEditor_Quest.frx":04AC
         Style           =   2  'Dropdown List
         TabIndex        =   47
         Top             =   1200
         Width           =   1695
      End
      Begin VB.HScrollBar scrlAccessReq 
         Height          =   255
         Left            =   1200
         Max             =   5
         TabIndex        =   45
         Top             =   1200
         Width           =   1455
      End
      Begin VB.ComboBox cmbGenderReq 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3720
         Style           =   2  'Dropdown List
         TabIndex        =   46
         Top             =   1200
         Width           =   1455
      End
      Begin VB.HScrollBar scrlLevelReq 
         Height          =   255
         LargeChange     =   10
         Left            =   1200
         Max             =   255
         TabIndex        =   54
         Top             =   1680
         Width           =   1455
      End
      Begin VB.ComboBox cmbSkillReq 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Top             =   600
         Width           =   2415
      End
      Begin VB.CommandButton btnReqOk 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Accept"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   72
         Top             =   2760
         Width           =   8055
      End
      Begin VB.Label lblSkill 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Skill Level: 0"
         Height          =   195
         Left            =   3960
         TabIndex        =   116
         Top             =   600
         Width           =   900
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "For: 0"
         Height          =   195
         Index           =   5
         Left            =   5400
         TabIndex        =   110
         Top             =   2160
         UseMnemonic     =   0   'False
         Width           =   405
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cha: 0"
         Height          =   195
         Index           =   4
         Left            =   2880
         TabIndex        =   109
         Top             =   2160
         UseMnemonic     =   0   'False
         Width           =   465
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Int: 0"
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   108
         Top             =   2160
         UseMnemonic     =   0   'False
         Width           =   360
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Agi: 0"
         Height          =   195
         Index           =   2
         Left            =   5400
         TabIndex        =   107
         Top             =   1680
         UseMnemonic     =   0   'False
         Width           =   405
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Str: 0"
         Height          =   195
         Index           =   1
         Left            =   2880
         TabIndex        =   106
         Top             =   1680
         UseMnemonic     =   0   'False
         Width           =   375
      End
      Begin VB.Label lblClassReq 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Class:"
         Height          =   195
         Left            =   5400
         TabIndex        =   105
         Top             =   1200
         Width           =   420
      End
      Begin VB.Label lblAccessReq 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Access: 0"
         Height          =   195
         Left            =   240
         TabIndex        =   104
         Top             =   1200
         Width           =   705
      End
      Begin VB.Label lblGenderReq 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gender:"
         Height          =   195
         Left            =   2880
         TabIndex        =   103
         Top             =   1200
         Width           =   570
      End
      Begin VB.Label lblLevelReq 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Level: 0"
         Height          =   195
         Left            =   240
         TabIndex        =   102
         Top             =   1680
         Width           =   570
      End
      Begin VB.Label lblSkillReq 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Skill:"
         Height          =   195
         Left            =   240
         TabIndex        =   101
         Top             =   600
         Width           =   330
      End
      Begin VB.Line Line6 
         X1              =   120
         X2              =   8160
         Y1              =   240
         Y2              =   240
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Mission Editing"
      Height          =   5775
      Left            =   3240
      TabIndex        =   97
      Top             =   2400
      Width           =   11535
      Begin VB.ListBox CLI 
         Height          =   5325
         Left            =   120
         TabIndex        =   19
         ToolTipText     =   "List of Greeters AKA NPC's the player will need to meet with throughout the quest.  Right click for more options."
         Top             =   360
         Width           =   2895
      End
      Begin VB.ListBox lstTasks 
         Height          =   5325
         Left            =   3240
         TabIndex        =   20
         ToolTipText     =   "List of all the actions and tasks that will be completed for the selected Greeter.   Right click for more options."
         Top             =   360
         Width           =   8175
      End
      Begin VB.Line Line2 
         X1              =   3120
         X2              =   3120
         Y1              =   600
         Y2              =   5520
      End
   End
   Begin VB.Frame fmeCLI 
      Caption         =   "Add a new NPC/Event the player will need to meet with"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   3480
      TabIndex        =   98
      Top             =   4920
      Visible         =   0   'False
      Width           =   5055
      Begin VB.ComboBox cmbNPC 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   1440
         Width           =   4815
      End
      Begin VB.HScrollBar scrlKillAmnt 
         Height          =   255
         LargeChange     =   10
         Left            =   120
         Min             =   1
         TabIndex        =   59
         Top             =   3600
         Value           =   1
         Width           =   4815
      End
      Begin VB.CommandButton btnCLICancel 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2760
         Style           =   1  'Graphical
         TabIndex        =   67
         Top             =   3960
         Width           =   2175
      End
      Begin VB.OptionButton opEvent 
         Caption         =   "Interact using Events."
         Height          =   255
         Left            =   240
         TabIndex        =   34
         ToolTipText     =   "Selecting this option removes the need for an NPC to start/continue the mission."
         Top             =   2160
         Width           =   4455
      End
      Begin VB.OptionButton opNPC 
         Caption         =   "NPC's"
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   1920
         Value           =   -1  'True
         Width           =   4455
      End
      Begin VB.CommandButton btnAddCLI 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Accept"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   66
         Top             =   3960
         Width           =   2175
      End
      Begin VB.CheckBox chkReset 
         Caption         =   "Reset pre-existing kill amounts."
         Height          =   255
         Left            =   240
         TabIndex        =   44
         ToolTipText     =   "This option resets the kill count for the selected NPC('s) when the task begins."
         Top             =   2760
         Width           =   4695
      End
      Begin VB.OptionButton opThis 
         Caption         =   "Just this NPC."
         Height          =   255
         Left            =   480
         TabIndex        =   50
         ToolTipText     =   "Reset the kill count value for this specific NPC."
         Top             =   3000
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.OptionButton opAll 
         Caption         =   "All NPC's"
         Height          =   255
         Left            =   480
         TabIndex        =   53
         ToolTipText     =   "Reset the kill count values for all NPC's.  (Only applys to this quest.)"
         Top             =   3240
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CheckBox chkPass 
         Caption         =   "Move on to next task automatically."
         Height          =   255
         Left            =   240
         TabIndex        =   40
         ToolTipText     =   $"frmEditor_Quest.frx":04AE
         Top             =   2520
         Width           =   4695
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmEditor_Quest.frx":053E
         Height          =   855
         Left            =   120
         TabIndex        =   131
         Top             =   480
         Width           =   4815
      End
      Begin VB.Label lblKillAmnt 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount: 1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2100
         TabIndex        =   119
         Top             =   3360
         Width           =   840
      End
      Begin VB.Line Line3 
         X1              =   120
         X2              =   4920
         Y1              =   240
         Y2              =   240
      End
   End
   Begin VB.Frame fmeShowMsg 
      Caption         =   "Show player a message"
      Height          =   3015
      Left            =   9360
      TabIndex        =   115
      Top             =   5280
      Visible         =   0   'False
      Width           =   4455
      Begin VB.CheckBox chkComplete 
         Caption         =   "This is the response if the quest cannot be retaken."
         Height          =   255
         Left            =   120
         TabIndex        =   143
         ToolTipText     =   "Check this box if this message will be shown to the player when he/she talks to this NPC after completing the mission."
         Top             =   1800
         Width           =   4215
      End
      Begin VB.ComboBox cmbColor 
         Height          =   315
         ItemData        =   "frmEditor_Quest.frx":0630
         Left            =   720
         List            =   "frmEditor_Quest.frx":0632
         Style           =   2  'Dropdown List
         TabIndex        =   81
         Top             =   2040
         Width           =   3615
      End
      Begin VB.CommandButton btnMsgCancel 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   85
         Top             =   2400
         Width           =   2055
      End
      Begin VB.CommandButton btnMsgAccept 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Accept"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   84
         Top             =   2400
         Width           =   2055
      End
      Begin VB.CheckBox chkRes 
         Caption         =   "This is the response if the last task is not done."
         Height          =   255
         Left            =   120
         TabIndex        =   79
         ToolTipText     =   "Check this box if this message will be shown to the player if the first task before this message isn't completed yet."
         Top             =   1560
         Width           =   4215
      End
      Begin VB.CheckBox chkStart 
         Caption         =   "This is just a placeholder"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   86
         ToolTipText     =   "Check this box if this message will be the first message the player sees before starting the quest."
         Top             =   2400
         Width           =   4215
      End
      Begin VB.TextBox txtMsg 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1305
         Left            =   120
         MaxLength       =   300
         MultiLine       =   -1  'True
         TabIndex        =   69
         ToolTipText     =   "Enter a message that will be shown to the player."
         Top             =   240
         Width           =   4215
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Color:"
         Height          =   255
         Left            =   120
         TabIndex        =   125
         Top             =   2085
         Width           =   495
      End
   End
   Begin VB.Frame fmeSound 
      Caption         =   "Play a sound."
      Height          =   2175
      Left            =   9000
      TabIndex        =   134
      Top             =   7200
      Visible         =   0   'False
      Width           =   3975
      Begin VB.CommandButton btnSoundAccept 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Accept"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   141
         Top             =   1680
         Width           =   1695
      End
      Begin VB.CommandButton btnSoundCancel 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   140
         Top             =   1680
         Width           =   1695
      End
      Begin VB.OptionButton opEveryone 
         Caption         =   "Play for everyone in the game."
         Height          =   255
         Left            =   120
         TabIndex        =   139
         ToolTipText     =   "This will play the sound for all players in the game."
         Top             =   1320
         Width           =   3735
      End
      Begin VB.OptionButton opMap 
         Caption         =   "Play for entire map that the player is on."
         Height          =   255
         Left            =   120
         TabIndex        =   138
         ToolTipText     =   "This will play the sound for all players on the same map as the player triggering this action."
         Top             =   1080
         Width           =   3735
      End
      Begin VB.OptionButton opPlayer 
         Caption         =   "Play for player."
         Height          =   255
         Left            =   120
         TabIndex        =   137
         ToolTipText     =   "This will play the sound for the player only."
         Top             =   840
         Width           =   3735
      End
      Begin VB.ComboBox cmbSound 
         Height          =   315
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   136
         Top             =   360
         Width           =   3135
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sound:"
         Height          =   195
         Left            =   120
         TabIndex        =   135
         Top             =   400
         Width           =   510
      End
   End
   Begin VB.Menu mnuCLI 
      Caption         =   "CLIMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuEdit 
         Caption         =   "Edit Item"
      End
      Begin VB.Menu mnuACLI 
         Caption         =   "Add Greeter"
      End
      Begin VB.Menu mnuRCLI 
         Caption         =   "Remove Greeter"
      End
      Begin VB.Menu mnuAAction 
         Caption         =   "Add Action/Task"
      End
      Begin VB.Menu mnuRTask 
         Caption         =   "Remove Action/Task"
      End
   End
End
Attribute VB_Name = "frmEditor_Quest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private CLIHasFocus        As Boolean

Private Editing_CLI        As Boolean

Private Editing_CLI_Index  As Long

Private Editing_Task       As Boolean

Private Editing_Task_Index As Long

Private Gather             As Boolean

Private GiveItem           As Boolean

Private TakeItem           As Boolean

Private KillNPC            As Boolean

Private ChkVar             As Boolean

Private TmpIndex           As Long

Private Sub btnAddCLI_Click()

    Dim Index As Long, I As Long, tmpID As Long, TmpStr As String, NPCIndex As Long
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    If EditorIndex < 1 Or EditorIndex > MAX_QUESTS Then Exit Sub
    
    TmpStr = Replace(cmbNPC.List(cmbNPC.ListIndex), val(cmbNPC.List(cmbNPC.ListIndex)) & ": ", vbNullString)

    If Not Len(TmpStr) > 0 Then Exit Sub
    
    NPCIndex = 0

    For I = 1 To MAX_NPCS

        If Trim$(NPC(I).Name) = TmpStr Then
            NPCIndex = I

            Exit For

        End If

    Next I

    If opNPC.Value = True Then
        If Not NPCIndex > 0 Then
            Call QMsg("Please select an NPC from the drop down menu.")

            Exit Sub

        End If
    End If
    
    With Quest(EditorIndex)

        If Not KillNPC Then
            If Not Editing_CLI Then
                tmpID = Editing_CLI_Index
                
                .Max_CLI = .Max_CLI + 1
                Index = .Max_CLI

                'Prevent subscript out of range.
                If .Max_CLI = 1 Then
                    ReDim .CLI(1 To 1)
                Else
                    ReDim Preserve .CLI(1 To Index)
                End If
                
                .CLI(Index).ItemIndex = NPCIndex
                .CLI(Index).isNPC = Abs(opNPC.Value)
                
                'add in a start message automatically for the first CLI element created.
                If Index = 1 Then
                    .CLI(Index).Max_Actions = .CLI(Index).Max_Actions + 1
                    I = .CLI(Index).Max_Actions
                    ReDim Preserve .CLI(Index).Action(1 To I)
                    .CLI(Index).Action(I).ActionID = ACTION_SHOWMSG
                    .CLI(Index).Action(I).MainData = vbChecked
                    .CLI(Index).Action(I).TextHolder = "Double click to edit the start message."
                    .CLI(Index).Action(I).TertiaryData = BrightGreen
                End If

            Else
                .CLI(Editing_CLI_Index).ItemIndex = NPCIndex
                .CLI(Editing_CLI_Index).isNPC = opNPC.Value
                Editing_CLI_Index = 0
                Editing_CLI = False
            End If

            Call QuestEditorInit
        Else
            Index = CLI.ListIndex + 1

            If Index < 1 Then Exit Sub
            If Editing_Task Then
                tmpID = Editing_Task_Index
            Else
                .CLI(Index).Max_Actions = .CLI(Index).Max_Actions + 1
                ReDim Preserve .CLI(Index).Action(1 To .CLI(Index).Max_Actions)
                tmpID = .CLI(Index).Max_Actions
            End If
            
            .CLI(Index).Action(tmpID).ActionID = TASK_KILL
            .CLI(Index).Action(tmpID).MainData = NPCIndex
            .CLI(Index).Action(tmpID).SecondaryData = opNPC.Value
            .CLI(Index).Action(tmpID).amount = scrlKillAmnt.Value
            .CLI(Index).Action(tmpID).TertiaryData = chkPass.Value
            
            If chkReset.Value = vbChecked Then
                If opThis.Value Then
                    .CLI(Index).Action(tmpID).QuadData = NPCIndex
                Else
                    .CLI(Index).Action(tmpID).QuadData = -1
                End If

            Else
                .CLI(Index).Action(tmpID).QuadData = 0
            End If
            
            Editing_Task_Index = 0
            Editing_Task = False
        End If

    End With
    
    opNPC.Value = True
    fmeCLI.Visible = False
    chkPass.Value = vbUnchecked
    chkReset.Value = vbUnchecked
    opThis.Value = False
    opAll.Value = False
    ResetEditButtons
    Editing_CLI_Index = 0
    Editing_CLI = False

    If Not Editing_CLI Then CLI.ListIndex = CLI.ListCount - 1 Else CLI.ListIndex = Editing_CLI_Index
    Call QuestEditorInitCLI

    Exit Sub
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "btnAddCLI_Click", "frmEditor_Quest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub
    
End Sub

Private Sub btnAddItem_Click()

    Dim Index  As Long, Amnt As Long, Itm As Long, ID As Long, I As Long

    Dim TmpStr As String, ItemIndex As Long
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    TmpStr = cmbItem.List(cmbItem.ListIndex)
    ItemIndex = 0

    For I = 1 To MAX_NPCS

        If Trim$(Item(I).Name) = Replace$(TmpStr, I & ": ", vbNullString) Then
            ItemIndex = I

            Exit For

        End If

    Next I

    If Not ItemIndex > 0 Then
        Call QMsg("Unable to locate item within the item database.")

        Exit Sub

    End If
    
    Index = CLI.ListIndex + 1
    Itm = ItemIndex
    Amnt = scrlItemAmount.Value

    If Index < 1 Then
        Call QMsg("Please select a greeter first.")

        Exit Sub

    End If

    If Amnt < 1 Then
        Call QMsg("Please select an amount first.")

        Exit Sub

    End If
    
    If Gather Then
        ID = TASK_GATHER
    ElseIf GiveItem Then
        ID = ACTION_GIVE_ITEM
    ElseIf TakeItem Then
        ID = ACTION_TAKE_ITEM
    End If
    
    'add the item to the list
    
    With Quest(EditorIndex).CLI(Index)

        If Editing_Task Then
            I = Editing_Task_Index
        Else
            .Max_Actions = .Max_Actions + 1
            ReDim Preserve .Action(1 To .Max_Actions)
            I = .Max_Actions
        End If
        
        .Action(I).ActionID = ID
        .Action(I).MainData = Itm
        .Action(I).amount = Amnt
        .Action(I).SecondaryData = chkTake.Value
        .Action(I).TertiaryData = chkPassI.Value
        Editing_Task_Index = 0
        Editing_Task = False
        
        Call QuestEditorInitCLI
        Gather = False
        TakeItem = False
        GiveItem = False
        chkPassI.Value = vbUnchecked
        CLI.ListIndex = Index - 1
        fmeSelectItem.Visible = False
        Call ResetEditButtons
    End With
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "btnAddItem_Click", "frmEditor_Quest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub
    
End Sub

Private Sub btnAddWarp_Click()

    Dim Index As Long, X As Long, Y As Long, MapNum As Long, I As Long
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    Index = CLI.ListIndex + 1
    MapNum = scrlMap.Value
    X = scrlMapX.Value
    Y = scrlMapY.Value

    If Index < 1 Then Exit Sub
    If MapNum < 1 Then Exit Sub
    
    'add the item to the list
    
    With Quest(EditorIndex).CLI(Index)

        If Editing_Task Then
            I = Editing_Task_Index
        Else
            .Max_Actions = .Max_Actions + 1
            ReDim Preserve .Action(1 To .Max_Actions)
            I = .Max_Actions
        End If
        
        .Action(I).ActionID = ACTION_WARP
        .Action(I).amount = MapNum
        .Action(I).MainData = X
        .Action(I).SecondaryData = Y
        Editing_Task_Index = 0
        Editing_Task = False
        
        Call QuestEditorInitCLI
        CLI.ListIndex = Index - 1
        fmeWarp.Visible = False
        Call ResetEditButtons
    End With
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "btnAddWarp_Click", "frmEditor_Quest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub
    
End Sub

Private Sub btnAdjustStat_Click()

    Dim Index As Long
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    Index = CLI.ListIndex + 1

    If Index < 1 Then Exit Sub
    
    'open the panel to give an item.
    fmeModify.Visible = True
    Call BTF(fmeModify)
    cboItem.Clear
    cboItem.AddItem "None"
    cboItem.ListIndex = 0
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "btnAdjustStat_Click", "frmEditor_Quest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub

End Sub

Private Sub btnAItemCancel_Click()
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    scrlItemAmount.Value = 0
    chkTake.Value = vbUnchecked
    chkPassI.Value = vbUnchecked
    fmeSelectItem.Visible = False
    Editing_Task = False
    Editing_Task_Index = 0
    Call ResetEditButtons
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "btnAItemCancel_Click", "frmEditor_Quest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub

End Sub

Private Sub btnCLICancel_Click()
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    opNPC.Value = True
    fmeCLI.Visible = False
    lblKillAmnt.Visible = False
    scrlKillAmnt.Visible = False
    chkPass.Value = vbUnchecked
    chkReset.Value = vbUnchecked
    opThis.Value = False
    opAll.Value = False
    Editing_CLI = False
    Editing_CLI_Index = 0
    Editing_Task = False
    Editing_Task_Index = 0
    Call ResetEditButtons
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "btnCLICancel_Click", "frmEditor_Quest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub

End Sub

Private Sub btnCloseWarp_Click()
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    scrlMap.Value = 1
    scrlMapX.Value = 1
    scrlMapY.Value = 1
    fmeWarp.Visible = False
    Editing_Task = False
    Editing_Task_Index = 0
    Call ResetEditButtons
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "btnCloseWarp_Click", "frmEditor_Quest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub

End Sub

Private Sub btnDeleteAction_Click()
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    If CLIHasFocus Then
        Call mnuRCLI_Click
    Else
        Call mnuRTask_Click
    End If

    fmeMoveItem.Visible = False
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "btnDeleteAction_Click", "frmEditor_Quest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub

End Sub

Private Sub btnDown_Click()

    Dim tempSel As Long, tempSel2 As Long
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    If CLIHasFocus Then
        'move item up within the CLI
        tempSel = CLI.ListIndex

        If tempSel < 0 Or tempSel > CLI.ListCount - 1 Then
            btnDown.Enabled = False

            Exit Sub

        End If

        If Not CLI.ListCount > 1 Then Exit Sub
        
        Call MoveListItem(LIST_CLI, EditorIndex, 0, tempSel + 1, 1)
        CLI.ListIndex = tempSel + 1
        Call QuestEditorInitCLI
    Else
        'move item up within the Task List
        tempSel = lstTasks.ListIndex

        If tempSel < 0 Or tempSel > lstTasks.ListCount - 1 Then
            btnDown.Enabled = False

            Exit Sub

        End If

        If CLI.ListCount < 0 Then Exit Sub
        tempSel2 = CLI.ListIndex

        If tempSel2 < 0 Then Exit Sub
        
        'tempsel/2 is +1 because the array for the data starts at 1 whereas the listbox starts at 0
        Call MoveListItem(LIST_TASK, EditorIndex, tempSel2 + 1, tempSel + 1, 1)
        CLI.ListIndex = tempSel2
        Call QuestEditorInitCLI
        lstTasks.ListIndex = tempSel + 1
    End If
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "btnDown_Click", "frmEditor_Quest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub

End Sub

Private Sub btnEditAction_Click()
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    Call mnuEdit_Click
    fmeMoveItem.Visible = False
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "btnEditAction_Click", "frmEditor_Quest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub

End Sub

Private Sub btnGiveItem_Click()

    Dim Index As Long
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    Index = CLI.ListIndex + 1

    If Index < 1 Then Exit Sub
    
    'open the panel to give an item.
    GiveItem = True
    TakeItem = False
    Gather = False
    chkPassI.Visible = False
    fmeSelectItem.Caption = "Select an item to give."
    chkTake.Visible = False
    fmeSelectItem.Visible = True
    Call BTF(fmeSelectItem)
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "btnGiveItem_Click", "frmEditor_Quest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub

End Sub

Private Sub btnHide_Click()
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    fmeMoveItem.Visible = False
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "btnHide_Click", "frmEditor_Quest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub

End Sub

Private Sub btnModAccept_Click()

    Dim Index As Long, ID As Long
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    If EditorIndex < 1 Or EditorIndex > MAX_QUESTS Then Exit Sub
        
    Index = CLI.ListIndex + 1

    If Index < 1 Then
        QMsg ("Please be sure a greeter is selected.")

        Exit Sub

    End If

    If chkSet.Value = vbChecked And scrlModify.Value < 1 Then Exit Sub
    If cboItem.ListIndex < 0 Then Exit Sub
        
    With Quest(EditorIndex).CLI(Index)

        If Editing_Task Then
            ID = Editing_Task_Index
        Else
            .Max_Actions = .Max_Actions + 1
            ReDim Preserve .Action(1 To .Max_Actions)
            ID = .Max_Actions
        End If
                
        If opEXP.Value = True Then
            .Action(ID).ActionID = ACTION_ADJUST_EXP
        ElseIf opLvl.Value = True Then
            .Action(ID).ActionID = ACTION_ADJUST_LVL
        ElseIf opStat.Value = True Then
            .Action(ID).ActionID = ACTION_ADJUST_STAT_LVL
        ElseIf opSkill.Value = True Then
            .Action(ID).ActionID = ACTION_ADJUST_SKILL_LVL
        ElseIf opSkillEXP.Value = True Then
            .Action(ID).ActionID = ACTION_ADJUST_SKILL_EXP
        ElseIf opStatP.Value = True Then
            .Action(ID).ActionID = ACTION_ADJUST_STAT_POINTS
        End If
            
        .Action(ID).amount = scrlModify.Value
        .Action(ID).MainData = chkSet.Value
        .Action(ID).SecondaryData = cboItem.ListIndex
        Editing_Task_Index = 0
        Editing_Task = False
    End With
        
    chkSet.Value = vbUnchecked
    scrlModify.Value = 0
    opEXP.Value = True
    opEXP.Value = False
    Call QuestEditorInitCLI
    CLI.ListIndex = Index - 1
    fmeModify.Visible = False
    Call ResetEditButtons

    Exit Sub
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "btnModAccept_Click", "frmEditor_Quest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub
    
End Sub

Private Sub btnModCancel_Click()
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    chkSet.Value = vbUnchecked
    opEXP.Value = False
    opLvl.Value = False
    scrlModify.Value = 0
    fmeModify.Visible = False
    Editing_Task = False
    Editing_Task_Index = 0
    Call ResetEditButtons
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "btnModCancel_Click", "frmEditor_Quest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub

End Sub

Private Sub btnMsgAccept_Click()

    Dim Index As Long, Msg As String

    Dim I     As Long, II As Long, III As Long, ID As Long
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    Index = CLI.ListIndex + 1
    Msg = txtMsg.text

    If Index < 1 Then Exit Sub
    If Len(Msg) < 1 Then
        Call QMsg("Please type a message to show the player.")

        Exit Sub

    End If

    If cmbColor.ListIndex < 0 Then
        Call QMsg("Please select a message color.")

        Exit Sub

    End If

    txtMsg.text = vbNullString
    
    'add the item to the list
    With Quest(EditorIndex).CLI(Index)

        If Editing_Task Then
            ID = Editing_Task_Index
        Else
            .Max_Actions = .Max_Actions + 1
            ReDim Preserve .Action(1 To .Max_Actions)
            ID = .Max_Actions
        End If
        
        .Action(ID).ActionID = ACTION_SHOWMSG
        .Action(ID).MainData = chkStart.Value
        .Action(ID).SecondaryData = chkRes.Value
        .Action(ID).TertiaryData = cmbColor.ListIndex
        .Action(ID).QuadData = chkComplete.Value
        .Action(ID).TextHolder = Msg
        Editing_Task_Index = 0
        Editing_Task = False
            
        CLI.ListIndex = Index - 1
        fmeShowMsg.Visible = False
        chkStart.Value = vbUnchecked
        chkRes.Value = vbUnchecked
        Call QuestEditorInitCLI
        Call ResetEditButtons
    End With
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "btnMsgAccept_Click", "frmEditor_Quest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub
    
End Sub

Private Sub btnMsgCancel_Click()
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    chkStart.Value = vbUnchecked
    chkRes.Value = vbUnchecked
    txtMsg.text = vbNullString
    fmeShowMsg.Visible = False
    Editing_Task = False
    Editing_Task_Index = 0
    Call ResetEditButtons
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "btnMsgCancel_Click", "frmEditor_Quest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub

End Sub

Private Sub btnMsgPlayer_Click()

    Dim Index As Long
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    Index = CLI.ListIndex + 1

    If Index < 1 Then Exit Sub
    
    'open the panel to give an item.
    Call CheckResponseMsg(EditorIndex, Index, Quest(EditorIndex).CLI(Index).Max_Actions)
    chkComplete.Enabled = CanShowCompleteCheck
    fmeShowMsg.Visible = True
    Call BTF(fmeShowMsg)
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "btnMsgPlayer_Click", "frmEditor_Quest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub

End Sub

Private Sub btnObAccept_Click()

    Dim Index As Long, Amnt As Long, ID As Long, SkillID As Long
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    Index = CLI.ListIndex + 1
    SkillID = cmbObSKill.ListIndex + 1
    Amnt = scrlObSkill.Value

    If Index < 1 Then Exit Sub
    If SkillID < 1 Then Exit Sub
    If Amnt < 1 Then Exit Sub
    
    'add the item to the list
    
    With Quest(EditorIndex).CLI(Index)

        If Editing_Task Then
            ID = Editing_Task_Index
        Else
            .Max_Actions = .Max_Actions + 1
            ReDim Preserve .Action(1 To .Max_Actions)
            ID = .Max_Actions
        End If
        
        .Action(ID).ActionID = TASK_GETSKILL
        .Action(ID).MainData = SkillID
        .Action(ID).amount = Amnt
        Editing_Task_Index = 0
        Editing_Task = False
        
        Call QuestEditorInitCLI
        CLI.ListIndex = Index - 1
        fmeObtainSKill.Visible = False
        Call ResetEditButtons
    End With
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "btnObAccept_Click", "frmEditor_Quest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub
    
End Sub

Private Sub btnObCancel_Click()
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    cmbObSKill.ListIndex = 0
    scrlObSkill.Value = 0
    fmeObtainSKill.Visible = False
    Editing_Task = False
    Editing_Task_Index = 0
    Call ResetEditButtons
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "btnObCancel_Click", "frmEditor_Quest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub

End Sub

Private Sub btnReq_Click()
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    fmeReq.Visible = True
    Call BTF(fmeReq)
    fmeMoveItem.Visible = False
    DisableEditButtons
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "btnReq_Click", "frmEditor_Quest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub

End Sub

Private Sub btnReqOk_Click()
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    fmeReq.Visible = False
    ResetEditButtons
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "btnReqOk_Click", "frmEditor_Quest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub

End Sub

Private Sub btnSkillLvl_Click()
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    fmeObtainSKill.Visible = True
    Call BTF(fmeObtainSKill)
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "btnSkillLvl_Click", "frmEditor_Quest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub

End Sub

Private Sub btnSound_Click()

    Dim Index As Long
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    Index = CLI.ListIndex + 1

    If Index < 1 Then Exit Sub
    
    'open the panel to give an item.
    fmeSound.Visible = True
    Call BTF(fmeSound)
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "btnSound_Click", "frmEditor_Quest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub

End Sub

Private Sub btnSoundAccept_Click()

    Dim Index As Long, Snd As Long, ID As Long, I As Long
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    Index = CLI.ListIndex + 1
    Snd = cmbSound.ListIndex

    If Index < 1 Then
        Call QMsg("Please select a greeter first.")

        Exit Sub

    End If

    If Snd < 0 Then
        Call QMsg("Please select a sound to play from the list.")

        Exit Sub

    End If

    If Not opPlayer.Value And Not opMap.Value And Not opEveryone.Value Then
        Call QMsg("Please choose who hears the sound.")

        Exit Sub

    End If
    
    'add the item to the list
    
    With Quest(EditorIndex).CLI(Index)

        If Editing_Task Then
            I = Editing_Task_Index
        Else
            .Max_Actions = .Max_Actions + 1
            ReDim Preserve .Action(1 To .Max_Actions)
            I = .Max_Actions
        End If
        
        .Action(I).ActionID = ACTION_PLAYSOUND
        .Action(I).MainData = Snd + 1

        If opPlayer.Value Then
            .Action(I).SecondaryData = 0
        ElseIf opMap.Value Then
            .Action(I).SecondaryData = 1
        ElseIf opEveryone.Value Then
            .Action(I).SecondaryData = 2
        End If

        Editing_Task_Index = 0
        Editing_Task = False
        
        Call QuestEditorInitCLI
        opPlayer.Value = False
        opMap.Value = False
        opEveryone.Value = False
        CLI.ListIndex = Index - 1
        fmeSound.Visible = False
        Call ResetEditButtons
    End With
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "btnSoundAccept_Click", "frmEditor_Quest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub
    
End Sub

Private Sub btnSoundCancel_Click()
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    opPlayer.Value = False
    opMap.Value = False
    opEveryone.Value = False
    fmeSound.Visible = False
    Call ResetEditButtons
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "btnSoundCancel_Click", "frmEditor_Quest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub

End Sub

Private Sub btnTakeItem_Click()

    Dim Index As Long
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    Index = CLI.ListIndex + 1

    If Index < 1 Then Exit Sub
    
    'open the panel to give an item.
    TakeItem = True
    GiveItem = False
    Gather = False
    chkPassI.Visible = False
    fmeSelectItem.Caption = "Select an item to take."
    chkTake.Visible = False
    fmeSelectItem.Visible = True
    Call BTF(fmeSelectItem)
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "btnTakeItem_Click", "frmEditor_Quest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub

End Sub

Private Sub btnTask_Gather_Click()

    Dim Index As Long
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    Index = CLI.ListIndex + 1

    If Index < 1 Then Exit Sub
    
    'open the panel to give an item.
    Gather = True
    GiveItem = False
    TakeItem = False
    chkPassI.Visible = True
    fmeSelectItem.Caption = "Select an item to gather."
    chkTake.Visible = True
    fmeSelectItem.Visible = True
    Call BTF(fmeSelectItem)
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "btnTask_Gather_Click", "frmEditor_Quest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub

End Sub

Private Sub btnTask_Kill_Click()
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    opEvent.Enabled = False
    KillNPC = True
    fmeCLI.Caption = "Select the enemy the player will need to kill"
    lblKillAmnt.Visible = True
    scrlKillAmnt.Visible = True
    cmbNPC.ListIndex = -1
    Call SetNPCBox(True, CLI.ListIndex + 1)
    fmeCLI.Visible = True
    Call BTF(fmeCLI)
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "btnTask_Kill_Click", "frmEditor_Quest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub

End Sub

Private Sub btnTaskCancel_Click()
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    fmeTask.Visible = False
    ResetEditButtons
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "btnTaskCancel_Click", "frmEditor_Quest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub

End Sub

Private Sub btnUp_Click()

    Dim tempSel As Long, tempSel2 As Long
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    If CLIHasFocus Then
        'move item up within the CLI
        tempSel = CLI.ListIndex

        If tempSel <= 0 Then
            btnUp.Enabled = False

            Exit Sub

        End If

        If Not CLI.ListCount > 1 Then Exit Sub
        
        Call MoveListItem(LIST_CLI, EditorIndex, 0, tempSel + 1, -1)
        DoEvents
        CLI.ListIndex = tempSel - 1
        Call QuestEditorInitCLI
    Else
        'move item up within the Task List
        tempSel = lstTasks.ListIndex

        If tempSel <= 0 Then
            btnUp.Enabled = False

            Exit Sub

        End If

        If Not lstTasks.ListCount > 1 Then Exit Sub
        tempSel2 = CLI.ListIndex

        If tempSel2 < 0 Then Exit Sub
        
        'tempsel/2 is +1 because the array for the data starts at 1 whereas the listbox starts at 0
        Call MoveListItem(LIST_TASK, EditorIndex, tempSel2 + 1, tempSel + 1, -1)
        DoEvents
        CLI.ListIndex = tempSel2
        Call QuestEditorInitCLI
        lstTasks.ListIndex = tempSel - 1
    End If
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "btnUp_Click", "frmEditor_Quest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub

End Sub

Private Sub btnVariable_Click()
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    ChkVar = True
    fmeVar.Visible = True
    txtVar.Visible = True
    chkSetValue.Visible = False
    chkPassVar.Visible = True
    Call BTF(fmeVar)
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "btnVariable_Click", "frmEditor_Quest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub

End Sub

Private Sub btnWarp_Click()

    Dim Index As Long
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    Index = CLI.ListIndex + 1

    If Index < 1 Then Exit Sub
    
    'open the panel to warp the player.
    fmeWarp.Visible = True
    Call BTF(fmeWarp)
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "btnWarp_Click", "frmEditor_Quest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub

End Sub

Private Sub chkComplete_Click()
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    chkRes.Enabled = Not chkComplete.Value
    chkRes.Value = vbUnchecked

    If chkComplete.Value = vbUnchecked Then chkComplete.Enabled = CanShowCompleteCheck
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "chkComplete_Click", "frmEditor_Quest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub

End Sub

Private Sub chkReset_Click()
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    opThis.Visible = chkReset.Value
    opAll.Visible = chkReset.Value
    opThis.Value = chkReset.Value
    opAll.Value = chkReset.Value
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "chkReset_Click", "frmEditor_Quest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub

End Sub

Private Sub chkRetake_Click()
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    If EditorIndex < 1 Or EditorIndex > MAX_QUESTS Then Exit Sub
    
    fmeMoveItem.Visible = False
    Quest(EditorIndex).CanBeRetaken = chkRetake.Value
    chkComplete.Enabled = CanShowCompleteCheck

    Exit Sub
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "chkRetake_Click", "frmEditor_Quest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub
    
End Sub

Private Sub chkSet_Click()
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    If chkSet.Value = vbChecked Then
        scrlModify.Value = 0
        scrlModify.min = 0
    Else
        scrlModify.Value = 0
        scrlModify.min = -32767
    End If
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "chkSet_Click", "frmEditor_Quest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub

End Sub

Private Sub chkUnOrder_Click()
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    If EditorIndex < 1 Or EditorIndex > MAX_QUESTS Then Exit Sub
    
    fmeMoveItem.Visible = False
    Quest(EditorIndex).OutOfOrder = chkUnOrder.Value

    Exit Sub
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "chkUnOrder_Click", "frmEditor_Quest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub
    
End Sub

Private Sub CLI_Click()
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    If EditorIndex < 1 Or EditorIndex > MAX_QUESTS Then Exit Sub
    
    CLIHasFocus = True
    QuestEditorInitCLI

    If CLI.ListCount > 1 Then
        fmeMoveItem.Visible = True
        btnUp.Enabled = True
        btnDown.Enabled = True
        fmeMoveItem.Left = 1560
        btnDeleteAction.Enabled = True

        If CLI.ListIndex = 0 Then
            btnUp.Enabled = False
            btnDown.Enabled = False
            
            If CLI.ListCount > 1 Then
                btnDeleteAction.Enabled = False
            Else
                btnDeleteAction.Enabled = True
            End If

        ElseIf CLI.ListIndex = CLI.ListCount - 1 Then
            btnDown.Enabled = False
        Else
            btnUp.Enabled = True
        End If
        
        'Don't allow the swap between the 1st CLI containing the start message.
        If CLI.ListIndex = 1 Then
            btnUp.Enabled = False
        End If
        
    Else
        fmeMoveItem.Visible = False
    End If

    Exit Sub
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "CLI_Click", "frmEditor_Quest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub
    
End Sub

Private Sub CLI_DblClick()

    Dim Index As Long, I As Long
    
    'we're gonna edit this list item instead of creating one.
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    CLIHasFocus = True
    Index = CLI.ListIndex + 1

    If Index < 1 Then
        Call QMsg("Please select a greeter to edit first.")

        Exit Sub

    End If

    opNPC.Value = CBool(Quest(EditorIndex).CLI(Index).isNPC)
    opEvent.Value = Not CBool(Quest(EditorIndex).CLI(Index).isNPC)
    Editing_CLI = True
    fmeCLI.Visible = True
    Editing_CLI_Index = Index
    Call SetNPCBox(False, Index)
    
    If opNPC.Value Then

        For I = 0 To cmbNPC.ListCount - 1

            If Replace$(cmbNPC.List(I), val(cmbNPC.List(I)) & ": ", vbNullString) = Trim$(NPC(Quest(EditorIndex).CLI(Index).ItemIndex).Name) Then
                cmbNPC.ListIndex = I

                Exit For

            End If

        Next I

    End If

    Exit Sub
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "CLI_DblClick", "frmEditor_Quest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub
    
End Sub

Private Sub CLI_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    If Button = vbRightButton Then
        CLIHasFocus = True
        mnuACLI.Visible = True

        If CLI.ListIndex = 0 And Not CLI.ListCount > 1 Or CLI.ListIndex > 0 Then mnuRCLI.Visible = True Else mnuRCLI.Visible = False
        mnuRTask.Visible = False

        If CLI.ListCount > 0 Then mnuEdit.Visible = True Else mnuEdit.Visible = False
        PopupMenu mnuCLI
    End If

    Exit Sub
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "CLI_MouseDown", "frmEditor_Quest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub

End Sub

Private Sub cmbClassReq_Click()
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    Quest(EditorIndex).Requirements.ClassReq = cmbClassReq.ListIndex

    Exit Sub
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "cmbClassReq_Click", "frmEditor_Quest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub

End Sub

Private Sub cmbGenderReq_Click()
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    Quest(EditorIndex).Requirements.GenderReq = cmbGenderReq.ListIndex

    Exit Sub
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "cmbGenderReq_Click", "frmEditor_Quest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub

End Sub

Private Sub cmbItem_Click()
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    scrlItemAmount.Value = 0

    If cmbItem.ListIndex > 0 And cmbItem.ListIndex < MAX_ITEMS Then
        If Len(Trim$(Item(cmbItem.ListIndex).Name)) > 0 Then
            If Gather Then
                If Not Item(cmbItem.ListIndex).stackable > 0 Then
                    scrlItemAmount.max = MAX_INV
                Else
                    scrlItemAmount.max = 32767
                End If
            End If
        End If
    End If
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "cmbItem_Click", "frmEditor_Quest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub

End Sub

Private Sub cmbSkillReq_Click()
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    Quest(EditorIndex).Requirements.SkillReq = cmbSkillReq.ListIndex

    Exit Sub
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "cmbSkillReq_Click", "frmEditor_Quest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub

End Sub

Private Sub cmbSound_Click()
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    Audio.PlaySound SoundCache(cmbSound.ListIndex + 1), -1, -1, True

    Exit Sub
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "cmbSound_Click", "frmEditor_Quest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub

End Sub

Private Sub cmdChangeData_Click()
    Dim Res As VbMsgBoxResult, val As String
    Dim dataModified As Boolean, I As Long
    
    If EditorIndex < 1 Or EditorIndex > MAX_QUESTS Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    For I = 1 To MAX_QUESTS
        If Quest_Changed(I) And I <> EditorIndex Then
        
            dataModified = True
            Exit For
        End If
    Next
    
    If dataModified Then
        Res = MsgBox("Do you want to continue and discard the changes you made to your data?", vbYesNo)
        
        If Res = vbNo Then Exit Sub
    End If
    
    val = InputBox("Enter the amount you want the new data size to be.", "Change Data Size", MAX_QUESTS)
    
    If Not IsNumeric(val) Then
        Exit Sub
    End If
    
    Res = Abs(val)
    
    If Res = MAX_QUESTS Then Exit Sub
    
    Call SendChangeDataSize(Res, EDITOR_QUEST)
    
    Unload frmEditor_Quest
    MAX_QUESTS = Res
    ReDim Quest(MAX_QUESTS)
    
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdChangeDataSize_Click", "frmEditor_Quest", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdClose_Click()
    If EditorIndex < 1 Or EditorIndex > MAX_RESOURCES Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    frmAdmin.chkEditor(EDITOR_QUEST).FontBold = False
    frmAdmin.picEye(EDITOR_QUEST).Visible = False
    BringWindowToTop (frmAdmin.hWnd)
    Unload frmEditor_Quest
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdClose_Click", "frmEditor_Quest", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdCopy_Click()
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    TmpIndex = lstIndex.ListIndex

    Exit Sub
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "cmdCopy_Click", "frmEditor_Quest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub
    
End Sub

Private Sub cmdDelete_Click()
    Dim TmpIndex As Long
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    If EditorIndex < 1 Or EditorIndex > MAX_QUESTS Then Exit Sub

    ClearQuest EditorIndex
    
    TmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Trim$(Quest(EditorIndex).Name), EditorIndex - 1
    lstIndex.ListIndex = TmpIndex

    QuestEditorInit

    Exit Sub
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "cmdDelete_Click", "frmEditor_Quest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub
End Sub

Private Sub cmdPaste_Click()
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    lstIndex.RemoveItem EditorIndex - 1
    Call CopyMemory(ByVal VarPtr(Quest(EditorIndex)), ByVal VarPtr(Quest(TmpIndex + 1)), LenB(Quest(TmpIndex + 1)))
    lstIndex.AddItem EditorIndex & ": " & Trim$(Quest(EditorIndex).Name), EditorIndex - 1
    lstIndex.ListIndex = EditorIndex - 1

    Exit Sub
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "cmdPaste_Click", "frmEditor_Quest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub
    
End Sub

Private Sub cmdSave_Click()
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    If EditorIndex < 1 Or EditorIndex > MAX_QUESTS Then Exit Sub

    EditorSave = True
    Call QuestEditorSave
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "cmdSave_Click", "frmEditor_Quest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub
End Sub

Private Sub cmdSetVariable_Click()
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    ChkVar = False
    fmeVar.Visible = True
    chkSetValue.Visible = True
    txtVar.Visible = False
    chkPassVar.Visible = False
    Call BTF(fmeVar)
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "cmdSetVariable_Click", "frmEditor_Quest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub

End Sub

Private Sub btnVAccept_Click()

    Dim Index  As Long, Amnt As Long, tVar As String, ID As Long, I As Long

    Dim TmpStr As String, Tmp1 As Long, Tmp2 As Long
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    Index = CLI.ListIndex + 1
    Amnt = scrlValue.Value
    TmpStr = txtVar.text

    If Index < 1 Then
        Call QMsg("Please select a greeter.")

        Exit Sub

    End If

    If OpVariable.Value Then
        If Amnt < 1 Then
            Call QMsg("You must move the slider to select an amount.")

            Exit Sub

        End If
    End If
    
    If ChkVar Then
        ID = TASK_VARIABLE

        If Len(TmpStr) < 1 Then Exit Sub
    Else
        ID = ACTION_SETVARIABLE
    End If
    
    Tmp1 = MAX_VARIABLES
    Tmp2 = MAX_SWITCHES
    tVar = cmbVars.List(cmbVars.ListIndex)

    If Len(tVar) < Len(Tmp1) + 2 Or Len(tVar) < Len(Tmp2) + 2 Then
        Call QMsg("Please make sure you select a named variable/switch")

        Exit Sub

    End If
    
    'add the item to the list
    
    With Quest(EditorIndex).CLI(Index)

        If Editing_Task Then
            I = Editing_Task_Index
        Else
            .Max_Actions = .Max_Actions + 1
            ReDim Preserve .Action(1 To .Max_Actions)
            I = .Max_Actions
        End If
        
        .Action(I).ActionID = ID
        .Action(I).MainData = Abs(OpVariable.Value)
        .Action(I).amount = scrlValue.Value
        .Action(I).SecondaryData = cmbVars.ListIndex + 1
        .Action(I).TextHolder = txtVar.text
        .Action(I).TertiaryData = chkSetValue.Value
        .Action(I).QuadData = chkPassVar.Value
        Editing_Task_Index = 0
        Editing_Task = False
        
        Call QuestEditorInitCLI
        
        scrlValue.Value = 0
        chkPassVar.Value = vbUnchecked
        chkSetValue.Value = vbUnchecked
        txtVar.text = vbNullString
        CLI.ListIndex = Index - 1
        fmeVar.Visible = False
        Call ResetEditButtons
    End With
    
    Exit Sub
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "btnVAccept_Click", "frmEditor_Quest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub

End Sub

Private Sub btnVCancel_Click()
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    scrlValue.Value = 0
    chkSetValue.Value = vbUnchecked
    fmeVar.Visible = False
    Call ResetEditButtons
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "btnVCancel_Click", "frmEditor_Quest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub

End Sub

Private Sub Form_Load()

    Dim I As Long
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    frmEditor_Quest.Width = 14895
    frmEditor_Quest.Height = 9200
    frmEditor_Quest.Caption = "Mission Editor"
    frmEditor_Quest.scrlLevelReq.max = MAX_LEVEL
    txtName.MaxLength = QUESTNAME_LENGTH
    txtSearch.MaxLength = NAME_LENGTH
    txtDesc.MaxLength = QUESTDESC_LENGTH
    
    cmbSkillReq.Clear
    cmbClassReq.Clear
    cmbGenderReq.Clear
    cmbSkillReq.AddItem "None"
    cmbClassReq.AddItem "None"
    cmbGenderReq.AddItem "None"
    
    cmbSkillReq.Enabled = False
    cmbObSKill.Enabled = False
    
    'For I = 1 To Skill_Count - 1
    '    cmbSkillReq.AddItem GetSkillName(I)
    '    cmbObSKill.AddItem GetSkillName(I)
    'Next I
    
    For I = 1 To UBound(SoundCache)
        cmbSound.AddItem SoundCache(I)
    Next
    
    For I = 0 To 18
        cmbColor.AddItem GetColorName(I)
    Next I
    
    For I = 1 To MAX_CLASSES

        If Len(Trim$(Replace(Class(I).Name, Chr(0), ""))) > 0 Then
            cmbClassReq.AddItem Trim$(Class(I).Name)
        End If

    Next I
    
    cmbGenderReq.AddItem "Male"
    cmbGenderReq.AddItem "Female"
    
    Call PositionFrames
    Call SetItemBox

    Exit Sub
    
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "Form_Load", "frmEditor_Quest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    If EditorSave = False Then
        'QuestEditorCancel
    Else
        EditorSave = False
    End If
    
    frmAdmin.chkEditor(EDITOR_QUEST).Value = False
    BringWindowToTop (frmAdmin.hWnd)
    Exit Sub
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "Form_Unload", "frmEditor_Quest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub
    
End Sub

Private Sub lblCredit_Click()
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    fmeMoveItem.Visible = False
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "lblCredit_Click", "frmEditor_Quest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub

End Sub

Private Sub lstIndex_Click()
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    If EditorIndex < 1 Or EditorIndex > MAX_QUESTS Then Exit Sub
    
    fmeMoveItem.Visible = False
    QuestEditorInit

    Exit Sub
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "lstIndex_Click", "frmEditor_Quest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub
    
End Sub

Private Sub lstTasks_Click()

    Dim CLIID As Long, TaskID As Long
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    If EditorIndex < 1 Or EditorIndex > MAX_QUESTS Then Exit Sub
    
    CLIHasFocus = False
    CLIID = CLI.ListIndex + 1
    TaskID = lstTasks.ListIndex + 1
    
    If lstTasks.ListCount > 1 Then
        If Quest(EditorIndex).CLI(CLIID).Action(TaskID).ActionID = ACTION_SHOWMSG Then
            If Quest(EditorIndex).CLI(CLIID).Action(TaskID).MainData = vbChecked Then
                fmeMoveItem.Visible = False

                Exit Sub

            End If
        End If
        
        fmeMoveItem.Visible = True
        fmeMoveItem.Left = 4680

        If lstTasks.ListIndex = 0 Then btnUp.Enabled = False Else btnUp.Enabled = True
        If lstTasks.ListIndex = lstTasks.ListCount - 1 Then btnDown.Enabled = False Else btnDown.Enabled = True
        
        If lstTasks.ListIndex = 1 Then
            If Quest(EditorIndex).CLI(CLIID).Action(TaskID - 1).ActionID = ACTION_SHOWMSG Then
                If Quest(EditorIndex).CLI(CLIID).Action(TaskID - 1).MainData = vbChecked Then
                    btnUp.Enabled = False
                End If
            End If
        End If

    Else
        fmeMoveItem.Visible = False
    End If

    Exit Sub
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "lstTasks_Click", "frmEditor_Quest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub
    
End Sub

Private Sub lstTasks_DblClick()

    Dim Index As Long, I As Long, II As Long
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    If EditorIndex < 1 Or EditorIndex > MAX_QUESTS Then Exit Sub
    
    CLIHasFocus = False
    'edit the selected list item instead of creating a new one
    Index = CLI.ListIndex + 1
    I = lstTasks.ListIndex + 1

    If Index < 1 Then
        Call QMsg("Please select a greeter first.")

        Exit Sub

    End If

    If I < 1 Then
        Call QMsg("Please select a task to edit first.")

        Exit Sub

    End If
    
    Editing_Task = True
    Editing_Task_Index = I
    
    Gather = False
    GiveItem = False
    TakeItem = False
    KillNPC = False
    ChkVar = False
    
    Call DisableEditButtons
    
    With Quest(EditorIndex).CLI(Index).Action(I)

        Select Case .ActionID

            Case TASK_GATHER
                Gather = True
                chkPassI.Visible = True
                chkPassI.Value = .TertiaryData
                chkTake.Value = .SecondaryData
                fmeSelectItem.Visible = True
                Call BTF(fmeSelectItem)
                chkTake.Visible = True

                For I = 0 To cmbItem.ListCount - 1

                    If cmbItem.List(I) = Trim$(Item(.MainData).Name) Then
                        cmbItem.ListIndex = I

                        Exit For

                    End If

                Next I

                scrlItemAmount.Value = .amount

            Case TASK_KILL
                KillNPC = True
                fmeCLI.Visible = True
                Call BTF(fmeCLI)
                lblKillAmnt.Visible = True
                scrlKillAmnt.Visible = True
                Call SetNPCBox(True, CLI.ListIndex + 1)

                For I = 0 To cmbNPC.ListCount - 1

                    If Replace$(cmbNPC.List(I), val(cmbNPC.List(I)) & ": ", vbNullString) = Trim$(NPC(.MainData).Name) Then
                        cmbNPC.ListIndex = I

                        Exit For

                    End If

                Next I

                If .QuadData <> 0 Then
                    chkReset.Value = vbChecked

                    If .QuadData = -1 Then opAll.Value = True Else opThis.Value = True
                End If

                scrlKillAmnt.Value = .amount
                chkPass.Value = .TertiaryData

            Case TASK_VARIABLE
                ChkVar = True
                fmeVar.Visible = True
                Call BTF(fmeVar)
                OpVariable.Value = CBool(.MainData)
                opSwitch.Value = Not CBool(.MainData)
                chkSetValue.Value = 0
                chkSetValue.Visible = False
                scrlValue.Value = .amount
                cmbVars.ListIndex = .SecondaryData - 1
                txtVar.Visible = True
                txtVar.text = Trim$(.TextHolder)
                chkPassVar.Visible = True
                chkPassVar.Value = .QuadData

            Case TASK_GETSKILL
                fmeObtainSKill.Visible = True
                Call BTF(fmeObtainSKill)
                cmbObSKill.ListIndex = .MainData - 1
                scrlObSkill.Value = .amount

            Case ACTION_SETVARIABLE
                fmeVar.Visible = True
                Call BTF(fmeVar)
                OpVariable.Value = CBool(.MainData)
                opSwitch.Value = Not CBool(.MainData)
                chkSetValue.Visible = True
                chkSetValue.Value = .TertiaryData
                scrlValue.Value = .amount
                cmbVars.ListIndex = .SecondaryData - 1
                txtVar.Visible = False
                txtVar.text = vbNullString
                chkPassVar.Visible = False
                chkPassVar.Value = vbUnchecked

            Case ACTION_SHOWMSG
                Call CheckResponseMsg(EditorIndex, Index, I - 1)
                fmeShowMsg.Visible = True
                Call BTF(fmeShowMsg)
                txtMsg.text = Trim$(.TextHolder)
                chkStart.Value = .MainData
                chkRes.Value = .SecondaryData
                cmbColor.ListIndex = .TertiaryData

                If .QuadData > 0 Then chkComplete.Enabled = .QuadData Else chkComplete.Enabled = CanShowCompleteCheck
                chkComplete.Value = .QuadData

                If .MainData = vbChecked Then chkStart.Enabled = True

            Case ACTION_ADJUST_EXP
                fmeModify.Visible = True
                Call BTF(fmeModify)
                opEXP.Value = True
                chkSet.Value = .MainData
                scrlModify.Value = .amount

            Case ACTION_ADJUST_LVL
                fmeModify.Visible = True
                Call BTF(fmeModify)
                opLvl.Value = True
                chkSet.Value = .MainData
                scrlModify.Value = .amount

            Case ACTION_ADJUST_STAT_LVL
                fmeModify.Visible = True
                Call BTF(fmeModify)
                opStat.Value = True
                chkSet.Value = .MainData
                scrlModify.Value = .amount
                cboItem.ListIndex = .SecondaryData

            Case ACTION_ADJUST_EXP
                fmeModify.Visible = True
                Call BTF(fmeModify)
                opEXP.Value = True
                chkSet.Value = .MainData
                scrlModify.Value = .amount
                cboItem.ListIndex = .SecondaryData

            Case ACTION_ADJUST_STAT_POINTS
                fmeModify.Visible = True
                Call BTF(fmeModify)
                opStatP.Value = True
                chkSet.Value = .MainData
                scrlModify.Value = .amount
                cboItem.ListIndex = .SecondaryData

            Case ACTION_ADJUST_SKILL_LVL
                fmeModify.Visible = True
                Call BTF(fmeModify)
                opSkill.Value = True
                chkSet.Value = .MainData
                scrlModify.Value = .amount
                cboItem.ListIndex = .SecondaryData

            Case ACTION_ADJUST_SKILL_EXP
                fmeModify.Visible = True
                Call BTF(fmeModify)
                opSkillEXP.Value = True
                chkSet.Value = .MainData
                scrlModify.Value = .amount
                cboItem.ListIndex = .SecondaryData

            Case ACTION_GIVE_ITEM
                GiveItem = True
                fmeSelectItem.Visible = True
                Call BTF(fmeSelectItem)

                For I = 0 To cmbItem.ListCount - 1

                    If cmbItem.List(I) = Trim$(Item(.MainData).Name) Then
                        cmbItem.ListIndex = I

                        Exit For

                    End If

                Next I

                scrlItemAmount.Value = .amount

            Case ACTION_TAKE_ITEM
                TakeItem = True
                fmeSelectItem.Visible = True
                Call BTF(fmeSelectItem)

                For I = 0 To cmbItem.ListCount - 1

                    If cmbItem.List(I) = Trim$(Item(.MainData).Name) Then
                        cmbItem.ListIndex = I

                        Exit For

                    End If

                Next I

                scrlItemAmount.Value = .amount

            Case ACTION_WARP
                fmeWarp.Visible = True
                Call BTF(fmeWarp)
                scrlMap.Value = .amount
                scrlMapX.Value = .MainData
                scrlMapY.Value = .SecondaryData
                
            Case ACTION_PLAYSOUND
                fmeSound.Visible = True
                Call BTF(fmeSound)
                cmbSound.ListIndex = .MainData - 1

                If .SecondaryData = 0 Then
                    opPlayer.Value = True
                ElseIf .SecondaryData = 1 Then
                    opMap.Value = True
                ElseIf .SecondaryData = 2 Then
                    opEveryone.Value = True
                End If

            Case Else

                Exit Sub

        End Select

    End With

    Exit Sub
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "lstTasks_DblClick", "frmEditor_Quest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub
    
End Sub

Private Sub lstTasks_MouseDown(Button As Integer, _
                               Shift As Integer, _
                               X As Single, _
                               Y As Single)
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    If Button = vbRightButton Then
        CLIHasFocus = False

        If lstTasks.ListCount > 0 Then mnuEdit.Visible = True Else mnuEdit.Visible = False
        mnuACLI.Visible = False
        mnuRCLI.Visible = False

        If lstTasks.ListIndex = 0 Then mnuRTask.Visible = False Else mnuRTask.Visible = True
        PopupMenu mnuCLI
    End If
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "lstTasks_MouseDown", "frmEditor_Quest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub

End Sub

Private Sub mnuAAction_Click()
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    If CLI.ListIndex < 0 Then
        Call QMsg("Please select one of the NPC's to meet with from the list below.")

        Exit Sub

    End If

    fmeMoveItem.Visible = False
    fmeTask.Visible = True
    Call BTF(fmeTask)
    DisableEditButtons
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "mnuAAction_Click", "frmEditor_Quest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub

End Sub

Private Sub mnuACLI_Click()
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    KillNPC = False
    fmeCLI.Caption = "Add a new NPC/Event the player will need to meet with"
    fmeCLI.Visible = True
    lblKillAmnt.Visible = False
    scrlKillAmnt.Visible = False
    Call SetNPCBox(False, CLI.ListIndex + 1, True)
    Call BTF(fmeCLI)
    DisableEditButtons
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "mnuACLI_Click", "frmEditor_Quest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub

End Sub

Private Sub mnuEdit_Click()
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    If CLIHasFocus Then Call CLI_DblClick Else Call lstTasks_DblClick
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "mnuEdit_Click", "frmEditor_Quest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub

End Sub

Private Sub mnuRCLI_Click()

    Dim Index As Long

    Dim Res   As VbMsgBoxResult
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    If EditorIndex < 1 Or EditorIndex > MAX_QUESTS Then Exit Sub
    
    Index = CLI.ListIndex + 1

    If Index < 1 Then
        Call QMsg("Please select the Greeter you would like to delete.")

        Exit Sub

    End If

    fmeMoveItem.Visible = False
    Exit Sub

' Error Handler
ErrorHandler:
    HandleError "mnuRCLI_Click", "frmEditor_Quest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub
    
End Sub

Private Sub mnuRTask_Click()

    Dim Index As Long, TaskID As Long

    Dim Res   As VbMsgBoxResult
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    If EditorIndex < 1 Or EditorIndex > MAX_QUESTS Then Exit Sub
    
    Index = CLI.ListIndex + 1
    TaskID = lstTasks.ListIndex + 1

    If Index < 1 Then
        Call QMsg("Please select a Greeter first, and then an Action/Task to be removed..")

        Exit Sub

    End If

    If TaskID < 1 Then
        Call QMsg("Please select an Action/Task to remove from the list.")

        Exit Sub

    End If

    fmeMoveItem.Visible = False
    
    'lets delete the selected action/task
    Res = MsgBox("Are you sure you want to delete this Action/Task?", vbYesNo, "Confirm Deletion?")

    If Res = vbYes Then Call DeleteAction(EditorIndex, Index, TaskID)
    CLI.ListIndex = Index - 1
    Call QuestEditorInitCLI

    Exit Sub
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "mnuRTask_Click", "frmEditor_Quest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub
    
End Sub

Private Sub opEvent_Click()
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    cmbNPC.Enabled = False
    chkPass.Value = vbUnchecked
    chkPass.Enabled = False
    chkReset.Value = vbUnchecked
    chkReset.Enabled = False
    
    Exit Sub
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "opEvent_Click", "frmEditor_Quest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub

End Sub

Private Sub opEXP_Click()
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    cboItem.Enabled = False
    cboItem.Clear
    cboItem.AddItem "None"
    cboItem.ListIndex = 0
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "opEXP_Click", "frmEditor_Quest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub

End Sub

Private Sub opLvl_Click()
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    cboItem.Enabled = False
    cboItem.Clear
    cboItem.AddItem "None"
    cboItem.ListIndex = 0
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "opLvl_Click", "frmEditor_Quest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub

End Sub

Private Sub opNPC_Click()

    Dim I As Long
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    Call SetNPCBox(False, CLI.ListIndex + 1, True)
    cmbNPC.Enabled = True
    chkPass.Enabled = True
    chkReset.Enabled = True
        
    Exit Sub
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "opNPC_Click", "frmEditor_Quest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub

End Sub

Private Sub opSkill_Click()

    Dim I As Long
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    cboItem.Enabled = True
    cboItem.Clear
    cboItem.AddItem "None"
    'For I = 1 To Skills.Skill_Count - 1
    '    cboItem.AddItem GetSkillName(I)
    'Next I
    cboItem.ListIndex = 0
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "opSkill_Click", "frmEditor_Quest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub

End Sub

Private Sub opSkillEXP_Click()

    Dim I As Long
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    cboItem.Enabled = True
    cboItem.Clear
    cboItem.AddItem "None"
    'For I = 1 To Skills.Skill_Count - 1
    '    cboItem.AddItem GetSkillName(I)
    'Next I
    cboItem.ListIndex = 0
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "opSkillEXP_Click", "frmEditor_Quest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub

End Sub

Private Sub opStat_Click()

    Dim I As Long
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    cboItem.Enabled = True
    cboItem.Clear
    cboItem.AddItem "None"

    For I = 1 To Stats.Stat_Count - 1
        cboItem.AddItem GetStatName(I)
    Next I

    cboItem.ListIndex = 0
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "opStat_Click", "frmEditor_Quest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub

End Sub

Private Sub opStatP_Click()
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    cboItem.Enabled = False
    cboItem.Clear
    cboItem.AddItem "None"
    cboItem.ListIndex = 0
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "opStatP_Click", "frmEditor_Quest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub

End Sub

Private Sub opSwitch_Click()

    Dim I As Long
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    scrlValue.Enabled = True
    scrlValue.min = 0
    scrlValue.max = 1
    scrlValue.Value = 0
    Call scrlValue_Change
    cmbVars.Enabled = True
    cmbVars.Clear

    For I = 1 To MAX_SWITCHES
        cmbVars.AddItem I & ": " & Switches(I)
    Next I
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "opSwitch_Click", "frmEditor_Quest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub

End Sub

Private Sub OpVariable_Click()

    Dim I As Long
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    scrlValue.Enabled = True
    scrlValue.min = CLng("-" & MAX_INTEGER)
    scrlValue.max = MAX_INTEGER
    scrlValue.Value = 0
    Call scrlValue_Change
    cmbVars.Enabled = True
    cmbVars.Clear

    For I = 1 To MAX_VARIABLES
        cmbVars.AddItem I & ": " & Variables(I)
    Next I
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "OpVariable_Click", "frmEditor_Quest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub

End Sub

Private Sub scrlAccessReq_Change()
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    If EditorIndex < 1 Or EditorIndex > MAX_QUESTS Then Exit Sub

    lblAccessReq.Caption = "Access: " & scrlAccessReq.Value
    Quest(EditorIndex).Requirements.AccessReq = scrlAccessReq.Value

    Exit Sub
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "scrlAccessReq_Change", "frmEditor_Quest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub
    
End Sub

Private Sub scrlItemAmount_Change()
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    lblAmount.Caption = "Amount: " & scrlItemAmount.Value
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "scrlItemAmount_Change", "frmEditor_Quest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub

End Sub

Private Sub scrlKillAmnt_Change()
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    lblKillAmnt.Caption = "Amount: " & scrlKillAmnt.Value
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "scrlKillAmnt_Change", "frmEditor_Quest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub

End Sub

Private Sub scrlLevelReq_Change()
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    If EditorIndex < 1 Or EditorIndex > MAX_QUESTS Then Exit Sub

    lblLevelReq.Caption = "Level: " & scrlLevelReq.Value
    Quest(EditorIndex).Requirements.LevelReq = scrlLevelReq.Value

    Exit Sub
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "scrlLevelReq_Change", "frmEditor_Quest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub
    
End Sub

Private Sub scrlMap_Change()
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    lblMap.Caption = "Map: " & scrlMap.Value
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "scrlMap_Change", "frmEditor_Quest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub

End Sub

Private Sub scrlMapX_Change()
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    lblMapX.Caption = "X: " & scrlMapX.Value
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "scrlMapX_Change", "frmEditor_Quest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub

End Sub

Private Sub scrlMapY_Change()
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    lblMapY.Caption = "Y: " & scrlMapY.Value
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "scrlMapY_Change", "frmEditor_Quest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub

End Sub

Private Sub scrlModify_Change()
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    If EditorIndex < 1 Or EditorIndex > MAX_QUESTS Then Exit Sub

    lblModify.Caption = "Amount to modify: " & scrlModify.Value

    Exit Sub
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "scrlModify_Change", "frmEditor_Quest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub
    
End Sub

Private Sub ResetEditButtons()
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    mnuACLI.Enabled = True
    mnuRCLI.Enabled = True
    mnuAAction.Enabled = True
    mnuRTask.Enabled = True
    mnuEdit.Enabled = True
    btnReq.Enabled = True
    CLI.Enabled = True
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "ResetEditButtons", "frmEditor_Quest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub

End Sub

Private Sub DisableEditButtons()
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    mnuACLI.Enabled = False
    mnuRCLI.Enabled = False
    mnuAAction.Enabled = False
    mnuRTask.Enabled = False
    mnuEdit.Enabled = False
    btnReq.Enabled = False
    CLI.Enabled = False
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "DisableEditButtons", "frmEditor_Quest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub

End Sub

Private Sub scrlObSkill_Change()
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    lblObSkill.Caption = "Skill Level: " & scrlObSkill.Value
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "scrlObSkill_Change", "frmEditor_Quest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub

End Sub

Private Sub scrlSkill_Change()
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    If EditorIndex < 1 Or EditorIndex > MAX_QUESTS Then Exit Sub

    lblSkill.Caption = "Skill Level: " & scrlSkill.Value
    Quest(EditorIndex).Requirements.SkillLevelReq = scrlSkill.Value

    Exit Sub
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "scrlSkill_Change", "frmEditor_Quest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub
    
End Sub

Private Sub scrlStatReq_Change(Index As Integer)

    Dim TmpStr As String
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    If EditorIndex < 1 Or EditorIndex > MAX_QUESTS Then Exit Sub

    Select Case Index

        Case 1
            TmpStr = "Str: "

        Case 2
            TmpStr = "Agi: "

        Case 3
            TmpStr = "Int: "

        Case 4
            TmpStr = "Cha: "

        Case 5
            TmpStr = "For: "

        Case Else

            Exit Sub

    End Select

    lblStatReq(Index).Caption = TmpStr & scrlStatReq(Index).Value
    Quest(EditorIndex).Requirements.Stat_Req(Index) = scrlStatReq(Index).Value

    Exit Sub
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "scrlStatReq_Change", "frmEditor_Quest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub
    
End Sub

Private Sub scrlValue_Change()
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    If OpVariable.Value Then
        lblValue.Caption = "Value: " & scrlValue.Value
    Else

        Select Case scrlValue.Value
        
            Case 0 'false
                lblValue.Caption = "Value: False"

            Case 1 'true
                lblValue.Caption = "Value: True"
        End Select

    End If
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "scrlValue_Change", "frmEditor_Quest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub

End Sub

Private Sub tmrMsg_Timer()
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    lblMsg.Caption = vbNullString
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "tmrMsg_Timer", "frmEditor_Quest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub

End Sub

Private Sub txtDesc_Change()
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    If EditorIndex < 1 Or EditorIndex > MAX_QUESTS Then Exit Sub
    
    Quest(EditorIndex).Description = Trim$(txtDesc.text)

    Exit Sub
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "txtDesc_Change", "frmEditor_Quest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub
    
End Sub

Private Sub txtDesc_Click()
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    fmeMoveItem.Visible = False
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "txtDesc_Click", "frmEditor_Quest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub

End Sub

Private Sub txtName_GotFocus()
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    fmeMoveItem.Visible = False
    txtName.SelStart = Len(txtName)

    Exit Sub
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "txtName_GotFocus", "frmEditor_Quest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub
    
End Sub

Private Sub txtName_Validate(Cancel As Boolean)

    Dim I        As Long

    Dim TmpIndex As Long
    
    'be sure we don't use the same name twice.
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    For I = 1 To MAX_QUESTS

        If I <> EditorIndex Then
            If LCase$(Trim$(txtName.text)) = LCase$(Trim$(Quest(I).Name)) And Len(Trim$(txtName.text)) > 0 Then
                txtName.text = vbNullString
                Call MsgBox("Duplicate quest name found.  Quest names must be unique.  Please change it.", vbOKOnly, "Duplicate Quest Name")

                Exit Sub

            End If
        End If

    Next I
    
    If EditorIndex < 1 Or EditorIndex > MAX_QUESTS Then Exit Sub
    
    TmpIndex = lstIndex.ListIndex
    Quest(EditorIndex).Name = Trim$(txtName.text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Trim$(Quest(EditorIndex).Name), EditorIndex - 1
    lstIndex.ListIndex = TmpIndex

    Exit Sub
    
    ' Error handlerin
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "txtName_Validate", "frmEditor_Quest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub

End Sub

Private Sub QMsg(ByVal Msg As String)
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    tmrMsg.Enabled = False
    lblMsg.Caption = Msg
    lblMsg.Visible = True
    tmrMsg.Enabled = True
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "QMsg", "frmEditor_Quest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub

End Sub

Private Sub txtRank_Change()
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    If EditorIndex < 1 Or EditorIndex > MAX_QUESTS Then Exit Sub
    
    Quest(EditorIndex).Rank = Trim$(txtRank.text)

    Exit Sub
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "txtRank_Change", "frmEditor_Quest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub
    
End Sub

Private Sub txtSearch_Click()
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    fmeMoveItem.Visible = False
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "txtSearch_Click", "frmEditor_Quest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub

End Sub

Private Sub PositionFrames()

    Dim tTop As Integer, tLeft As Integer

    Dim FME  As Control
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    tTop = 2640
    tLeft = 6480
    
    For Each FME In frmEditor_Quest.Controls

        If (TypeOf FME Is Frame) Then
            If FME.Name <> "Frame1" And FME.Name <> "Frame2" And FME.Name <> "Frame3" And FME.Name <> "fraNPC" And FME.Name <> "fmeMoveItem" Then
                FME.Top = tTop
                FME.Left = tLeft
            End If
        End If

    Next
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "PositionFrames", "frmEditor_Quest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub

End Sub

Private Sub BTF(ByVal FrameID As Frame)
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    Call FrameID.ZOrder(vbBringToFront)
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "BTF", "frmEditor_Quest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub

End Sub

Private Sub SetNPCBox(ByVal All As Boolean, _
                      Optional ByVal CurCLI As Long = 0, _
                      Optional ByVal Adding As Boolean = False)

    Dim I As Long, ShowItem As Boolean
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    cmbNPC.Clear
    cmbNPC.AddItem "Select An NPC"
    
    If All Then
        chkPass.Visible = True
        opEvent.Top = 2160
        chkReset.Visible = True
        opThis.Visible = True
        opAll.Visible = True
        lblKillAmnt.Top = 3360
        scrlKillAmnt.Top = 3600
        btnAddCLI.Top = 3960
        btnCLICancel.Top = 3960
    Else
        chkPass.Visible = False
        opEvent.Top = 2280
        chkReset.Visible = False
        opThis.Visible = False
        opAll.Visible = False
        lblKillAmnt.Top = 2520
        scrlKillAmnt.Top = 2760
        btnAddCLI.Top = 3120
        btnCLICancel.Top = 3120
    End If
    
    For I = 1 To MAX_NPCS

        If Len(Trim$(NPC(I).Name)) > 0 Then
            If All Then
                If Not NPC(I).Behavior = NPC_BEHAVIOR_QUEST Then
                    cmbNPC.AddItem I & ": " & Trim$(NPC(I).Name)
                End If

            Else

                If NPC(I).Behavior = NPC_BEHAVIOR_QUEST Then
                    ShowItem = True
                    
                    If Quest(EditorIndex).Max_CLI = 0 Then

                        'Don't allow the same NPC to be used for the beginning of more than one quest
                        If IsNPCInAnotherQuest(I, EditorIndex) Then ShowItem = False
                    Else

                        'Don't show the NPC if it used in the previous slot
                        If Adding Then
                            If Quest(EditorIndex).CLI(Quest(EditorIndex).Max_CLI).ItemIndex = I Then ShowItem = False
                        Else

                            If CurCLI > 1 Then
                                If Quest(EditorIndex).CLI(CurCLI - 1).ItemIndex = I Then ShowItem = False
                            ElseIf CurCLI = 1 Then

                                'Don't allow the same NPC to be used for the beginning of more than one quest
                                If IsNPCInAnotherQuest(I, EditorIndex) Then ShowItem = False
                            End If
                        End If
                    End If
                    
                    If ShowItem Then cmbNPC.AddItem I & ": " & Trim$(NPC(I).Name)
                End If
            End If
        End If

    Next I
    
    cmbNPC.ListIndex = 0
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "SetNPCBox", "frmEditor_Quest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub

End Sub

Private Sub SetItemBox()

    Dim I     As Long

    Dim SCHAR As String * 5
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    cmbItem.Clear
    cmbItem.AddItem "Select An Item"
    
    For I = 1 To MAX_ITEMS

        If Not InStr(Trim$(Item(I).Name), SCHAR) > 0 Then
            cmbItem.AddItem I & ": " & Trim$(Item(I).Name)
        End If

    Next I
    
    cmbItem.ListIndex = 0
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "SetItemBox", "frmEditor_Quest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub

End Sub

Private Function CanShowCompleteCheck() As Byte

    Dim I As Long, II As Long
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    CanShowCompleteCheck = vbChecked

    If Not EditorIndex > 0 Then Exit Function
    
    If chkRetake.Value = vbChecked Then
        CanShowCompleteCheck = vbUnchecked

        Exit Function

    End If
    
    For I = 1 To Quest(EditorIndex).Max_CLI
        For II = 1 To Quest(EditorIndex).CLI(I).Max_Actions

            If Quest(EditorIndex).CLI(I).Action(II).ActionID = ACTION_SHOWMSG Then
                If Quest(EditorIndex).CLI(I).Action(II).QuadData = vbChecked Then
                    'Found one all ready used.  Deny it
                    CanShowCompleteCheck = vbUnchecked

                    Exit Function

                End If
            End If

        Next II
    Next I
    
    ' Error Handler
    Exit Function

ErrorHandler:
    HandleError "CanShowCompleteCheck", "frmEditor_Quest", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Function

End Function
