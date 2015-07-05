VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEditor_Events 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Event Editor"
   ClientHeight    =   8970
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12885
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   598
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   859
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraGraphic 
      Caption         =   "Graphic Selection"
      Height          =   375
      Left            =   120
      TabIndex        =   74
      Top             =   120
      Visible         =   0   'False
      Width           =   495
      Begin VB.HScrollBar hScrlGraphicSel 
         Height          =   255
         LargeChange     =   64
         Left            =   240
         SmallChange     =   32
         TabIndex        =   107
         Top             =   7920
         Visible         =   0   'False
         Width           =   11895
      End
      Begin VB.VScrollBar vScrlGraphicSel 
         Height          =   7095
         LargeChange     =   64
         Left            =   12240
         SmallChange     =   32
         TabIndex        =   106
         Top             =   720
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox picGraphicSel 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   7080
         Left            =   240
         ScaleHeight     =   472
         ScaleMode       =   0  'User
         ScaleWidth      =   792.01
         TabIndex        =   81
         Top             =   720
         Width           =   11895
         Begin VB.Shape shpLoc 
            BorderColor     =   &H00FF0000&
            BorderWidth     =   2
            Height          =   480
            Left            =   0
            Top             =   0
            Width           =   480
         End
      End
      Begin VB.CommandButton cmdGraphicCancel 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   11040
         TabIndex        =   80
         Top             =   8280
         Width           =   1455
      End
      Begin VB.CommandButton cmdGraphicOK 
         Caption         =   "OK"
         Height          =   375
         Left            =   9480
         TabIndex        =   79
         Top             =   8280
         Width           =   1455
      End
      Begin VB.ComboBox cmbGraphic 
         Height          =   315
         ItemData        =   "frmEditor_Events.frx":0000
         Left            =   720
         List            =   "frmEditor_Events.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   76
         Top             =   240
         Width           =   2175
      End
      Begin VB.HScrollBar scrlGraphic 
         Height          =   255
         Left            =   4440
         TabIndex        =   75
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label lblRandomLabel 
         Caption         =   "Type:"
         Height          =   255
         Index           =   33
         Left            =   120
         TabIndex        =   78
         Top             =   270
         Width           =   855
      End
      Begin VB.Label lblGraphic 
         Caption         =   "Number: 1"
         Height          =   255
         Left            =   3000
         TabIndex        =   77
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame fraLabeling 
      Caption         =   "Labeling Variables and Switches"
      Height          =   495
      Left            =   120
      TabIndex        =   325
      Top             =   120
      Visible         =   0   'False
      Width           =   615
      Begin VB.Frame fraRenaming 
         Caption         =   "Renaming Variable/Switch"
         Height          =   8535
         Left            =   120
         TabIndex        =   334
         Top             =   120
         Visible         =   0   'False
         Width           =   12615
         Begin VB.Frame fraRandom 
            Caption         =   "Editing Variable/Switch"
            Height          =   2295
            Index           =   10
            Left            =   3600
            TabIndex        =   335
            Top             =   2520
            Width           =   5055
            Begin VB.TextBox txtRename 
               Height          =   375
               Left            =   120
               TabIndex        =   338
               Top             =   720
               Width           =   4815
            End
            Begin VB.CommandButton cmdRename_Cancel 
               Caption         =   "Cancel"
               Height          =   375
               Left            =   3720
               TabIndex        =   337
               Top             =   1800
               Width           =   1215
            End
            Begin VB.CommandButton cmdRename_Ok 
               Caption         =   "Ok"
               Height          =   375
               Left            =   2280
               TabIndex        =   336
               Top             =   1800
               Width           =   1215
            End
            Begin VB.Label lblEditing 
               Caption         =   "Naming Variable #1"
               Height          =   375
               Left            =   120
               TabIndex        =   339
               Top             =   360
               Width           =   4815
            End
         End
      End
      Begin VB.CommandButton cmdRenameSwitch 
         Caption         =   "Rename Switch"
         Height          =   375
         Left            =   8280
         TabIndex        =   333
         Top             =   7320
         Width           =   1935
      End
      Begin VB.CommandButton cmdRenameVariable 
         Caption         =   "Rename Variable"
         Height          =   375
         Left            =   360
         TabIndex        =   332
         Top             =   7320
         Width           =   1935
      End
      Begin VB.ListBox lstSwitches 
         Height          =   6495
         Left            =   8280
         TabIndex        =   330
         Top             =   720
         Width           =   3855
      End
      Begin VB.ListBox lstVariables 
         Height          =   6495
         Left            =   360
         TabIndex        =   328
         Top             =   720
         Width           =   3855
      End
      Begin VB.CommandButton cmbLabel_Ok 
         Caption         =   "OK"
         Height          =   375
         Left            =   9480
         TabIndex        =   327
         Top             =   8400
         Width           =   1455
      End
      Begin VB.CommandButton cmdLabel_Cancel 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   11040
         TabIndex        =   326
         Top             =   8400
         Width           =   1455
      End
      Begin VB.Label lblRandomLabel 
         Caption         =   "Player Switches"
         Height          =   255
         Index           =   36
         Left            =   8280
         TabIndex        =   331
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label lblRandomLabel 
         Caption         =   "Player Variables"
         Height          =   255
         Index           =   25
         Left            =   360
         TabIndex        =   329
         Top             =   480
         Width           =   2175
      End
   End
   Begin VB.Frame fraMoveRoute 
      Caption         =   "Move Route"
      Height          =   375
      Left            =   120
      TabIndex        =   108
      Top             =   120
      Visible         =   0   'False
      Width           =   855
      Begin VB.Frame Frame18 
         Caption         =   "Commands"
         Height          =   6615
         Left            =   3120
         TabIndex        =   115
         Top             =   480
         Width           =   9255
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Set Graphic..."
            Height          =   375
            Index           =   42
            Left            =   6720
            TabIndex        =   158
            Top             =   3240
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Set Position Above Players"
            Height          =   375
            Index           =   41
            Left            =   6720
            TabIndex        =   157
            Top             =   2760
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Set Position with Players"
            Height          =   375
            Index           =   40
            Left            =   6720
            TabIndex        =   156
            Top             =   2280
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Set Position Below Players"
            Height          =   375
            Index           =   39
            Left            =   6720
            TabIndex        =   155
            Top             =   1800
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Turn Off Walk-Through"
            Height          =   375
            Index           =   38
            Left            =   6720
            TabIndex        =   154
            Top             =   1320
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Turn On Walk-Through"
            Height          =   375
            Index           =   37
            Left            =   6720
            TabIndex        =   153
            Top             =   840
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Turn Fixed Dir Off"
            Height          =   375
            Index           =   36
            Left            =   6720
            TabIndex        =   152
            Top             =   360
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Turn Fixed Dir On"
            Height          =   375
            Index           =   35
            Left            =   4560
            TabIndex        =   151
            Top             =   5640
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Turn Walking Anim Off"
            Height          =   375
            Index           =   34
            Left            =   4560
            TabIndex        =   150
            Top             =   5160
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Turn Walking Anim On"
            Height          =   375
            Index           =   33
            Left            =   4560
            TabIndex        =   149
            Top             =   4680
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Set Freq. To Highest"
            Height          =   375
            Index           =   32
            Left            =   4560
            TabIndex        =   148
            Top             =   4200
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Set Freq. To Higher"
            Height          =   375
            Index           =   31
            Left            =   4560
            TabIndex        =   147
            Top             =   3720
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Set Freq. To Normal"
            Height          =   375
            Index           =   30
            Left            =   4560
            TabIndex        =   146
            Top             =   3240
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Set Freq. To Lower"
            Height          =   375
            Index           =   29
            Left            =   4560
            TabIndex        =   145
            Top             =   2760
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Set Freq. To Lowest"
            Height          =   375
            Index           =   28
            Left            =   4560
            TabIndex        =   144
            Top             =   2280
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Set Speed 4x Faster"
            Height          =   375
            Index           =   27
            Left            =   4560
            TabIndex        =   143
            Top             =   1800
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Set Speed 2x Faster"
            Height          =   375
            Index           =   26
            Left            =   4560
            TabIndex        =   142
            Top             =   1320
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Set Speed to Normal"
            Height          =   375
            Index           =   25
            Left            =   4560
            TabIndex        =   141
            Top             =   840
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Set Speed 2x Slower"
            Height          =   375
            Index           =   24
            Left            =   4560
            TabIndex        =   140
            Top             =   360
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Set Speed 4x Slower"
            Height          =   375
            Index           =   23
            Left            =   2400
            TabIndex        =   139
            Top             =   5640
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Set Speed 8x Slower"
            Height          =   375
            Index           =   22
            Left            =   2400
            TabIndex        =   138
            Top             =   5160
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Turn Away From Player***"
            Height          =   375
            Index           =   21
            Left            =   2400
            TabIndex        =   137
            Top             =   4680
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Turn Toward Player***"
            Height          =   375
            Index           =   20
            Left            =   2400
            TabIndex        =   136
            Top             =   4200
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Turn Randomly"
            Height          =   375
            Index           =   19
            Left            =   2400
            TabIndex        =   135
            Top             =   3720
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Turn 180 Degrees"
            Height          =   375
            Index           =   18
            Left            =   2400
            TabIndex        =   134
            Top             =   3240
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Turn 90 Degrees to the Left"
            Height          =   375
            Index           =   17
            Left            =   2400
            TabIndex        =   133
            Top             =   2760
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Turn 90 Degrees to the Right"
            Height          =   375
            Index           =   16
            Left            =   2400
            TabIndex        =   132
            Top             =   2280
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Turn Right"
            Height          =   375
            Index           =   15
            Left            =   2400
            TabIndex        =   131
            Top             =   1800
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Turn Left"
            Height          =   375
            Index           =   14
            Left            =   2400
            TabIndex        =   130
            Top             =   1320
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Turn Down"
            Height          =   375
            Index           =   13
            Left            =   2400
            TabIndex        =   129
            Top             =   840
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Turn Up"
            Height          =   375
            Index           =   12
            Left            =   2400
            TabIndex        =   128
            Top             =   360
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Wait 1000Ms"
            Height          =   375
            Index           =   11
            Left            =   240
            TabIndex        =   127
            Top             =   5640
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Wait 500Ms"
            Height          =   375
            Index           =   10
            Left            =   240
            TabIndex        =   126
            Top             =   5160
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Wait 100Ms"
            Height          =   375
            Index           =   9
            Left            =   240
            TabIndex        =   125
            Top             =   4680
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Step Back"
            Height          =   375
            Index           =   8
            Left            =   240
            TabIndex        =   124
            Top             =   4200
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Step Forward"
            Height          =   375
            Index           =   7
            Left            =   240
            TabIndex        =   123
            Top             =   3720
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Move Away From Player***"
            Height          =   375
            Index           =   6
            Left            =   240
            TabIndex        =   122
            Top             =   3240
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Move Towards Player***"
            Height          =   375
            Index           =   5
            Left            =   240
            TabIndex        =   121
            Top             =   2760
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Move Randomly"
            Height          =   375
            Index           =   4
            Left            =   240
            TabIndex        =   120
            Top             =   2280
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Move Right"
            Height          =   375
            Index           =   3
            Left            =   240
            TabIndex        =   119
            Top             =   1800
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Move Left"
            Height          =   375
            Index           =   2
            Left            =   240
            TabIndex        =   118
            Top             =   1320
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Move Down"
            Height          =   375
            Index           =   1
            Left            =   240
            TabIndex        =   117
            Top             =   840
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Move Up"
            Height          =   375
            Index           =   0
            Left            =   240
            TabIndex        =   116
            Top             =   360
            Width           =   1935
         End
         Begin VB.Label lblRandomLabel 
            Caption         =   "*** These commands will not process on global events."
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   15
            Left            =   240
            TabIndex        =   159
            Top             =   6240
            Width           =   8535
         End
      End
      Begin VB.ComboBox cmbEvent 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmEditor_Events.frx":002B
         Left            =   120
         List            =   "frmEditor_Events.frx":002D
         Style           =   2  'Dropdown List
         TabIndex        =   114
         Top             =   480
         Width           =   2655
      End
      Begin VB.CheckBox chkRepeatRoute 
         Caption         =   "Repeat Route"
         Height          =   255
         Left            =   120
         TabIndex        =   113
         Top             =   7560
         Width           =   2655
      End
      Begin VB.CheckBox chkIgnoreMove 
         Caption         =   "Ignore if event can't move."
         Height          =   255
         Left            =   120
         TabIndex        =   112
         Top             =   7200
         Width           =   2655
      End
      Begin VB.ListBox lstMoveRoute 
         Height          =   6105
         Left            =   120
         TabIndex        =   111
         Top             =   960
         Width           =   2655
      End
      Begin VB.CommandButton cmdMoveRouteOk 
         Caption         =   "OK"
         Height          =   375
         Left            =   9480
         TabIndex        =   110
         Top             =   8160
         Width           =   1455
      End
      Begin VB.CommandButton cmdMoveRouteCancel 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   11040
         TabIndex        =   109
         Top             =   8160
         Width           =   1455
      End
   End
   Begin VB.Frame fraDialogue 
      Height          =   6975
      Left            =   6240
      TabIndex        =   73
      Top             =   1320
      Visible         =   0   'False
      Width           =   6375
      Begin VB.Frame fraGiveExp 
         Caption         =   "Give Experience"
         Height          =   1695
         Left            =   1200
         TabIndex        =   341
         Top             =   1680
         Visible         =   0   'False
         Width           =   4095
         Begin VB.CommandButton cmdGiveExp_Ok 
            Caption         =   "Ok"
            Height          =   375
            Left            =   1320
            TabIndex        =   344
            Top             =   1200
            Width           =   1215
         End
         Begin VB.CommandButton cmdGiveExp_Cancel 
            Caption         =   "Cancel"
            Height          =   375
            Left            =   2640
            TabIndex        =   343
            Top             =   1200
            Width           =   1215
         End
         Begin VB.HScrollBar scrlGiveExp 
            Height          =   255
            Left            =   120
            Max             =   32000
            TabIndex        =   342
            Top             =   480
            Width           =   3735
         End
         Begin VB.Label lblGiveExp 
            Caption         =   "Give Exp: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   345
            Top             =   240
            Width           =   3735
         End
      End
      Begin VB.Frame fraChangeSprite 
         Caption         =   "Change Player Sprite"
         Height          =   1695
         Left            =   1200
         TabIndex        =   257
         Top             =   1680
         Visible         =   0   'False
         Width           =   4095
         Begin VB.HScrollBar scrlChangeSprite 
            Height          =   255
            Left            =   1200
            Max             =   100
            Min             =   1
            TabIndex        =   261
            Top             =   360
            Value           =   1
            Width           =   2535
         End
         Begin VB.CommandButton cmdChangeSprite_Cancel 
            Caption         =   "Cancel"
            Height          =   375
            Left            =   2520
            TabIndex        =   259
            Top             =   960
            Width           =   1215
         End
         Begin VB.CommandButton cmdChangeSprite_Ok 
            Caption         =   "Ok"
            Height          =   375
            Left            =   1200
            TabIndex        =   258
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label lblChangeSprite 
            Caption         =   "Sprite: 1"
            Height          =   255
            Left            =   120
            TabIndex        =   260
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.Frame fraChangeClass 
         Caption         =   "Change Player Class"
         Height          =   1695
         Left            =   1200
         TabIndex        =   252
         Top             =   1680
         Visible         =   0   'False
         Width           =   4095
         Begin VB.CommandButton cmdChangeClass_Ok 
            Caption         =   "Ok"
            Height          =   375
            Left            =   1200
            TabIndex        =   255
            Top             =   960
            Width           =   1215
         End
         Begin VB.CommandButton cmdChangeClass_Cancel 
            Caption         =   "Cancel"
            Height          =   375
            Left            =   2520
            TabIndex        =   254
            Top             =   960
            Width           =   1215
         End
         Begin VB.ComboBox cmbChangeClass 
            Height          =   315
            ItemData        =   "frmEditor_Events.frx":002F
            Left            =   120
            List            =   "frmEditor_Events.frx":0031
            Style           =   2  'Dropdown List
            TabIndex        =   253
            Top             =   480
            Width           =   3735
         End
         Begin VB.Label lblRandomLabel 
            Caption         =   "Class:"
            Height          =   255
            Index           =   29
            Left            =   120
            TabIndex        =   256
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Frame fraChangeSkills 
         Caption         =   "Change Player Skills"
         Height          =   2175
         Left            =   1200
         TabIndex        =   245
         Top             =   1440
         Visible         =   0   'False
         Width           =   4095
         Begin VB.OptionButton optChangeSkillsRemove 
            Caption         =   "Remove"
            Height          =   255
            Left            =   1800
            TabIndex        =   251
            Top             =   960
            Width           =   1455
         End
         Begin VB.OptionButton optChangeSkillsAdd 
            Caption         =   "Teach"
            Height          =   255
            Left            =   120
            TabIndex        =   250
            Top             =   960
            Width           =   1455
         End
         Begin VB.ComboBox cmbChangeSkills 
            Height          =   315
            ItemData        =   "frmEditor_Events.frx":0033
            Left            =   120
            List            =   "frmEditor_Events.frx":0035
            Style           =   2  'Dropdown List
            TabIndex        =   249
            Top             =   480
            Width           =   3735
         End
         Begin VB.CommandButton cmdChangeSkills_Cancel 
            Caption         =   "Cancel"
            Height          =   375
            Left            =   2520
            TabIndex        =   247
            Top             =   1680
            Width           =   1215
         End
         Begin VB.CommandButton cmdChangeSkills_Ok 
            Caption         =   "Ok"
            Height          =   375
            Left            =   1200
            TabIndex        =   246
            Top             =   1680
            Width           =   1215
         End
         Begin VB.Label lblRandomLabel 
            Caption         =   "Skill"
            Height          =   255
            Index           =   28
            Left            =   120
            TabIndex        =   248
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.Frame fraChangeLevel 
         Caption         =   "Change Level"
         Height          =   1815
         Left            =   1200
         TabIndex        =   240
         Top             =   1560
         Visible         =   0   'False
         Width           =   4095
         Begin VB.HScrollBar scrlChangeLevel 
            Height          =   255
            Left            =   120
            TabIndex        =   244
            Top             =   600
            Width           =   3615
         End
         Begin VB.CommandButton cmdChangeLevel_Ok 
            Caption         =   "Ok"
            Height          =   375
            Left            =   1200
            TabIndex        =   242
            Top             =   1080
            Width           =   1215
         End
         Begin VB.CommandButton cmdChangeLevel_Cancel 
            Caption         =   "Cancel"
            Height          =   375
            Left            =   2520
            TabIndex        =   241
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label lblChangeLevel 
            Caption         =   "Level: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   243
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.Frame fraChangeItems 
         Caption         =   "Change Items"
         Height          =   2415
         Left            =   1200
         TabIndex        =   231
         Top             =   1320
         Visible         =   0   'False
         Width           =   4095
         Begin VB.ComboBox cmbChangeItemIndex 
            Height          =   315
            ItemData        =   "frmEditor_Events.frx":0037
            Left            =   120
            List            =   "frmEditor_Events.frx":0039
            Style           =   2  'Dropdown List
            TabIndex        =   239
            Top             =   480
            Width           =   1695
         End
         Begin VB.TextBox txtChangeItemsAmount 
            Height          =   375
            Left            =   120
            TabIndex        =   238
            Text            =   "0"
            Top             =   1320
            Width           =   3735
         End
         Begin VB.CommandButton cmdChangeItems_Cancel 
            Caption         =   "Cancel"
            Height          =   375
            Left            =   2640
            TabIndex        =   236
            Top             =   1800
            Width           =   1215
         End
         Begin VB.CommandButton cmdChangeItems_Ok 
            Caption         =   "Ok"
            Height          =   375
            Left            =   1200
            TabIndex        =   235
            Top             =   1800
            Width           =   1215
         End
         Begin VB.OptionButton optChangeItemRemove 
            Caption         =   "Take Away"
            Height          =   255
            Left            =   2640
            TabIndex        =   234
            Top             =   960
            Width           =   1335
         End
         Begin VB.OptionButton optChangeItemAdd 
            Caption         =   "Give"
            Height          =   255
            Left            =   1680
            TabIndex        =   233
            Top             =   960
            Width           =   735
         End
         Begin VB.OptionButton optChangeItemSet 
            Caption         =   "Set Amount"
            Height          =   255
            Left            =   120
            TabIndex        =   232
            Top             =   960
            Value           =   -1  'True
            Width           =   1455
         End
         Begin VB.Label lblRandomLabel 
            Caption         =   "Item Index:"
            Height          =   255
            Index           =   27
            Left            =   120
            TabIndex        =   237
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.Frame fraShowChoices 
         Caption         =   "Show Choices"
         Height          =   4335
         Left            =   1320
         TabIndex        =   192
         Top             =   1080
         Visible         =   0   'False
         Width           =   4095
         Begin VB.TextBox txtChoices 
            Height          =   375
            Index           =   4
            Left            =   2160
            TabIndex        =   203
            Text            =   "0"
            Top             =   3240
            Width           =   1455
         End
         Begin VB.TextBox txtChoices 
            Height          =   375
            Index           =   3
            Left            =   120
            TabIndex        =   201
            Text            =   "0"
            Top             =   3240
            Width           =   1455
         End
         Begin VB.TextBox txtChoices 
            Height          =   375
            Index           =   2
            Left            =   2160
            TabIndex        =   199
            Text            =   "0"
            Top             =   2520
            Width           =   1455
         End
         Begin VB.TextBox txtChoices 
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   197
            Text            =   "0"
            Top             =   2520
            Width           =   1455
         End
         Begin VB.CommandButton cmdChoices_Cancel 
            Caption         =   "Cancel"
            Height          =   375
            Left            =   2760
            TabIndex        =   195
            Top             =   3840
            Width           =   1215
         End
         Begin VB.CommandButton cmdChoices_Ok 
            Caption         =   "Ok"
            Height          =   375
            Left            =   1440
            TabIndex        =   194
            Top             =   3840
            Width           =   1215
         End
         Begin VB.TextBox txtChoicePrompt 
            Height          =   1695
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   193
            Top             =   480
            Width           =   3855
         End
         Begin VB.Label lblRandomLabel 
            Caption         =   "Choice 4"
            Height          =   255
            Index           =   21
            Left            =   2160
            TabIndex        =   204
            Top             =   3000
            Width           =   1455
         End
         Begin VB.Label lblRandomLabel 
            Caption         =   "Choice 3"
            Height          =   255
            Index           =   20
            Left            =   120
            TabIndex        =   202
            Top             =   3000
            Width           =   1455
         End
         Begin VB.Label lblRandomLabel 
            Caption         =   "Choice 2"
            Height          =   255
            Index           =   19
            Left            =   2160
            TabIndex        =   200
            Top             =   2280
            Width           =   1455
         End
         Begin VB.Label lblRandomLabel 
            Caption         =   "Choice 1"
            Height          =   255
            Index           =   17
            Left            =   120
            TabIndex        =   198
            Top             =   2280
            Width           =   1455
         End
         Begin VB.Label lblRandomLabel 
            Caption         =   "Prompt:"
            Height          =   255
            Index           =   16
            Left            =   120
            TabIndex        =   196
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.Frame fraShowText 
         Caption         =   "Show Text"
         Height          =   4095
         Left            =   1320
         TabIndex        =   187
         Top             =   1200
         Visible         =   0   'False
         Width           =   4095
         Begin VB.TextBox txtShowText 
            Height          =   2775
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   190
            Top             =   480
            Width           =   3855
         End
         Begin VB.CommandButton cmdShowText_Ok 
            Caption         =   "Ok"
            Height          =   375
            Left            =   1440
            TabIndex        =   189
            Top             =   3600
            Width           =   1215
         End
         Begin VB.CommandButton cmdShowText_Cancel 
            Caption         =   "Cancel"
            Height          =   375
            Left            =   2760
            TabIndex        =   188
            Top             =   3600
            Width           =   1215
         End
         Begin VB.Label lblRandomLabel 
            Caption         =   "Text:"
            Height          =   255
            Index           =   18
            Left            =   120
            TabIndex        =   191
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.Frame fraAddText 
         Caption         =   "Add Text"
         Height          =   4095
         Left            =   1200
         TabIndex        =   220
         Top             =   600
         Visible         =   0   'False
         Width           =   4095
         Begin VB.TextBox txtAddText_Text 
            Height          =   1815
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   227
            Top             =   480
            Width           =   3855
         End
         Begin VB.HScrollBar scrlAddText_Colour 
            Height          =   255
            Left            =   120
            Max             =   18
            TabIndex        =   226
            Top             =   2640
            Width           =   3855
         End
         Begin VB.OptionButton optAddText_Player 
            Caption         =   "Player"
            Height          =   255
            Left            =   120
            TabIndex        =   225
            Top             =   3240
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton optAddText_Map 
            Caption         =   "Map"
            Height          =   255
            Left            =   1080
            TabIndex        =   224
            Top             =   3240
            Width           =   735
         End
         Begin VB.OptionButton optAddText_Global 
            Caption         =   "Global"
            Height          =   255
            Left            =   1920
            TabIndex        =   223
            Top             =   3240
            Width           =   855
         End
         Begin VB.CommandButton cmdAddText_Ok 
            Caption         =   "Ok"
            Height          =   375
            Left            =   1440
            TabIndex        =   222
            Top             =   3600
            Width           =   1215
         End
         Begin VB.CommandButton cmdAddText_Cancel 
            Caption         =   "Cancel"
            Height          =   375
            Left            =   2760
            TabIndex        =   221
            Top             =   3600
            Width           =   1215
         End
         Begin VB.Label lblRandomLabel 
            Caption         =   "Text:"
            Height          =   255
            Index           =   34
            Left            =   120
            TabIndex        =   230
            Top             =   240
            Width           =   1935
         End
         Begin VB.Label lblAddText_Colour 
            Caption         =   "Colour: Black"
            Height          =   255
            Left            =   120
            TabIndex        =   229
            Top             =   2400
            Width           =   3255
         End
         Begin VB.Label lblRandomLabel 
            Caption         =   "Channel:"
            Height          =   255
            Index           =   10
            Left            =   120
            TabIndex        =   228
            Top             =   3000
            Width           =   1575
         End
      End
      Begin VB.Frame fraPlayAnimation 
         Caption         =   "Play Animation"
         Height          =   2775
         Left            =   720
         TabIndex        =   274
         Top             =   1320
         Visible         =   0   'False
         Width           =   5055
         Begin VB.ComboBox cmbPlayAnim 
            Height          =   315
            ItemData        =   "frmEditor_Events.frx":003B
            Left            =   1680
            List            =   "frmEditor_Events.frx":003D
            Style           =   2  'Dropdown List
            TabIndex        =   287
            Top             =   300
            Width           =   3135
         End
         Begin VB.HScrollBar scrlPlayAnimTileY 
            Height          =   255
            Left            =   1920
            TabIndex        =   285
            Top             =   1800
            Width           =   2895
         End
         Begin VB.HScrollBar scrlPlayAnimTileX 
            Height          =   255
            Left            =   1920
            TabIndex        =   284
            Top             =   1455
            Width           =   2895
         End
         Begin VB.CommandButton cmdPlayAnim_Cancel 
            Caption         =   "Cancel"
            Height          =   375
            Left            =   3600
            TabIndex        =   280
            Top             =   2280
            Width           =   1215
         End
         Begin VB.CommandButton cmdPlayAnim_Ok 
            Caption         =   "Ok"
            Height          =   375
            Left            =   2160
            TabIndex        =   279
            Top             =   2280
            Width           =   1215
         End
         Begin VB.OptionButton optPlayAnimPlayer 
            Caption         =   "Player"
            Height          =   255
            Left            =   120
            TabIndex        =   278
            Top             =   1080
            Width           =   1695
         End
         Begin VB.OptionButton optPlayAnimEvent 
            Caption         =   "Event"
            Height          =   255
            Left            =   1920
            TabIndex        =   277
            Top             =   1080
            Width           =   1335
         End
         Begin VB.OptionButton optPlayAnimTile 
            Caption         =   "Tile"
            Height          =   255
            Left            =   3720
            TabIndex        =   276
            Top             =   1080
            Width           =   975
         End
         Begin VB.ComboBox cmbPlayAnimEvent 
            Height          =   315
            ItemData        =   "frmEditor_Events.frx":003F
            Left            =   1920
            List            =   "frmEditor_Events.frx":0041
            Style           =   2  'Dropdown List
            TabIndex        =   275
            Top             =   1440
            Visible         =   0   'False
            Width           =   2895
         End
         Begin VB.Label lblRandomLabel 
            Caption         =   "Animation"
            Height          =   255
            Index           =   30
            Left            =   120
            TabIndex        =   286
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label lblPlayAnimY 
            Caption         =   "Map Tile Y:"
            Height          =   255
            Left            =   240
            TabIndex        =   283
            Top             =   1800
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label lblPlayAnimX 
            Caption         =   "Map Tile X:"
            Height          =   255
            Left            =   240
            TabIndex        =   282
            Top             =   1440
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label lblRandomLabel 
            Caption         =   "Target Type:"
            Height          =   255
            Index           =   31
            Left            =   120
            TabIndex        =   281
            Top             =   840
            Width           =   1335
         End
      End
      Begin VB.Frame fraChangePK 
         Caption         =   "Set Player PK"
         Height          =   1455
         Left            =   1200
         TabIndex        =   267
         Top             =   1800
         Visible         =   0   'False
         Width           =   4095
         Begin VB.CommandButton cmdChangePK_Cancel 
            Caption         =   "Cancel"
            Height          =   375
            Left            =   2520
            TabIndex        =   271
            Top             =   840
            Width           =   1215
         End
         Begin VB.CommandButton cmdChangePK_Ok 
            Caption         =   "Ok"
            Height          =   375
            Left            =   1200
            TabIndex        =   270
            Top             =   840
            Width           =   1215
         End
         Begin VB.OptionButton optChangePKYes 
            Caption         =   "Yes"
            Height          =   255
            Left            =   240
            TabIndex        =   269
            Top             =   360
            Width           =   1455
         End
         Begin VB.OptionButton optChangePKNo 
            Caption         =   "No"
            Height          =   255
            Left            =   1920
            TabIndex        =   268
            Top             =   360
            Width           =   1455
         End
      End
      Begin VB.Frame fraWarpPlayer 
         Caption         =   "Warp Player"
         Height          =   3015
         Left            =   1320
         TabIndex        =   89
         Top             =   1320
         Visible         =   0   'False
         Width           =   4095
         Begin VB.ComboBox cmbWarpPlayerDir 
            Height          =   315
            ItemData        =   "frmEditor_Events.frx":0043
            Left            =   120
            List            =   "frmEditor_Events.frx":0056
            Style           =   2  'Dropdown List
            TabIndex        =   305
            Top             =   2040
            Width           =   3855
         End
         Begin VB.CommandButton cmdWPCancel 
            Caption         =   "Cancel"
            Height          =   375
            Left            =   2760
            TabIndex        =   97
            Top             =   2520
            Width           =   1215
         End
         Begin VB.CommandButton cmdWPOkay 
            Caption         =   "Ok"
            Height          =   375
            Left            =   1440
            TabIndex        =   96
            Top             =   2520
            Width           =   1215
         End
         Begin VB.HScrollBar scrlWPY 
            Height          =   255
            Left            =   120
            Max             =   255
            TabIndex        =   95
            Top             =   1680
            Width           =   3855
         End
         Begin VB.HScrollBar scrlWPX 
            Height          =   255
            Left            =   120
            Max             =   255
            TabIndex        =   93
            Top             =   1080
            Width           =   3855
         End
         Begin VB.HScrollBar scrlWPMap 
            Height          =   255
            Left            =   120
            TabIndex        =   91
            Top             =   480
            Width           =   3855
         End
         Begin VB.Label lblWPY 
            Caption         =   "Y: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   94
            Top             =   1440
            Width           =   3135
         End
         Begin VB.Label lblWPX 
            Caption         =   "X: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   92
            Top             =   840
            Width           =   3135
         End
         Begin VB.Label lblWPMap 
            Caption         =   "Map: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   90
            Top             =   240
            Width           =   3135
         End
      End
      Begin VB.Frame fraConditionalBranch 
         Caption         =   "Conditional Branch"
         Height          =   4815
         Left            =   120
         TabIndex        =   160
         Top             =   480
         Visible         =   0   'False
         Width           =   6135
         Begin VB.OptionButton optCondition_Index 
            Caption         =   "Self Switch"
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   303
            Top             =   3720
            Width           =   1695
         End
         Begin VB.ComboBox cmbCondition_SelfSwitch 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmEditor_Events.frx":0097
            Left            =   1920
            List            =   "frmEditor_Events.frx":00A7
            Style           =   2  'Dropdown List
            TabIndex        =   302
            Top             =   3720
            Width           =   1695
         End
         Begin VB.ComboBox cmbCondition_SelfSwitchCondition 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmEditor_Events.frx":00B7
            Left            =   3960
            List            =   "frmEditor_Events.frx":00C1
            Style           =   2  'Dropdown List
            TabIndex        =   301
            Top             =   3720
            Width           =   1095
         End
         Begin VB.ComboBox cmbCondition_LearntSkill 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmEditor_Events.frx":00D2
            Left            =   1920
            List            =   "frmEditor_Events.frx":00D4
            Style           =   2  'Dropdown List
            TabIndex        =   180
            Top             =   2760
            Width           =   1695
         End
         Begin VB.ComboBox cmbCondition_ClassIs 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmEditor_Events.frx":00D6
            Left            =   1920
            List            =   "frmEditor_Events.frx":00D8
            Style           =   2  'Dropdown List
            TabIndex        =   179
            Top             =   2280
            Width           =   1695
         End
         Begin VB.ComboBox cmbCondition_HasItem 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmEditor_Events.frx":00DA
            Left            =   1920
            List            =   "frmEditor_Events.frx":00DC
            Style           =   2  'Dropdown List
            TabIndex        =   178
            Top             =   1800
            Width           =   1695
         End
         Begin VB.ComboBox cmbCondtion_PlayerSwitchCondition 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmEditor_Events.frx":00DE
            Left            =   3960
            List            =   "frmEditor_Events.frx":00E8
            Style           =   2  'Dropdown List
            TabIndex        =   177
            Top             =   1320
            Width           =   1095
         End
         Begin VB.ComboBox cmbCondition_PlayerSwitch 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmEditor_Events.frx":00F9
            Left            =   1920
            List            =   "frmEditor_Events.frx":00FB
            Style           =   2  'Dropdown List
            TabIndex        =   176
            Top             =   1320
            Width           =   1695
         End
         Begin VB.TextBox txtCondition_LevelAmount 
            Enabled         =   0   'False
            Height          =   285
            Left            =   3480
            TabIndex        =   175
            Text            =   "0"
            Top             =   3240
            Width           =   855
         End
         Begin VB.ComboBox cmbCondition_LevelCompare 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmEditor_Events.frx":00FD
            Left            =   1440
            List            =   "frmEditor_Events.frx":0113
            Style           =   2  'Dropdown List
            TabIndex        =   174
            Top             =   3240
            Width           =   1695
         End
         Begin VB.OptionButton optCondition_Index 
            Caption         =   "Level"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   172
            Top             =   3240
            Width           =   975
         End
         Begin VB.ComboBox cmbCondition_PlayerVarCompare 
            Height          =   315
            ItemData        =   "frmEditor_Events.frx":0179
            Left            =   1920
            List            =   "frmEditor_Events.frx":018F
            Style           =   2  'Dropdown List
            TabIndex        =   170
            Top             =   840
            Width           =   1695
         End
         Begin VB.TextBox txtCondition_PlayerVarCondition 
            Height          =   285
            Left            =   3840
            TabIndex        =   169
            Text            =   "0"
            Top             =   840
            Width           =   855
         End
         Begin VB.ComboBox cmbCondition_PlayerVarIndex 
            Height          =   315
            ItemData        =   "frmEditor_Events.frx":01F5
            Left            =   1920
            List            =   "frmEditor_Events.frx":01F7
            Style           =   2  'Dropdown List
            TabIndex        =   168
            Top             =   480
            Width           =   1695
         End
         Begin VB.OptionButton optCondition_Index 
            Caption         =   "Knows Skill"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   167
            Top             =   2760
            Width           =   1695
         End
         Begin VB.OptionButton optCondition_Index 
            Caption         =   "Class Is"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   166
            Top             =   2280
            Width           =   1695
         End
         Begin VB.OptionButton optCondition_Index 
            Caption         =   "Has Item"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   165
            Top             =   1800
            Width           =   1695
         End
         Begin VB.OptionButton optCondition_Index 
            Caption         =   "Player Switch"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   164
            Top             =   1320
            Width           =   1695
         End
         Begin VB.OptionButton optCondition_Index 
            Caption         =   "Player Variable"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   163
            Top             =   480
            Width           =   1695
         End
         Begin VB.CommandButton cmdCondition_Ok 
            Caption         =   "Ok"
            Height          =   375
            Left            =   3360
            TabIndex        =   162
            Top             =   4320
            Width           =   1215
         End
         Begin VB.CommandButton cmdCondition_Cancel 
            Caption         =   "Cancel"
            Height          =   375
            Left            =   4680
            TabIndex        =   161
            Top             =   4320
            Width           =   1215
         End
         Begin VB.Label lblRandomLabel 
            Caption         =   "is"
            Height          =   255
            Index           =   35
            Left            =   3720
            TabIndex        =   304
            Top             =   3720
            Width           =   255
         End
         Begin VB.Label lblRandomLabel 
            Caption         =   "is"
            Height          =   255
            Index           =   1
            Left            =   3720
            TabIndex        =   186
            Top             =   1320
            Width           =   255
         End
         Begin VB.Label lblRandomLabel 
            Caption         =   "is"
            Height          =   255
            Index           =   2
            Left            =   1080
            TabIndex        =   173
            Top             =   3240
            Width           =   255
         End
         Begin VB.Label lblRandomLabel 
            Caption         =   "is"
            Height          =   255
            Index           =   0
            Left            =   3840
            TabIndex        =   171
            Top             =   480
            Width           =   615
         End
      End
      Begin VB.Frame fraOpenShop 
         Caption         =   "Open Shop"
         Height          =   1575
         Left            =   1200
         TabIndex        =   320
         Top             =   1800
         Visible         =   0   'False
         Width           =   4215
         Begin VB.CommandButton cmdOpenShop_Ok 
            Caption         =   "Ok"
            Height          =   375
            Left            =   1440
            TabIndex        =   323
            Top             =   1080
            Width           =   1215
         End
         Begin VB.CommandButton cmdOpenShop_Cancel 
            Caption         =   "Cancel"
            Height          =   375
            Left            =   2880
            TabIndex        =   322
            Top             =   1080
            Width           =   1215
         End
         Begin VB.ComboBox cmbOpenShop 
            Height          =   315
            ItemData        =   "frmEditor_Events.frx":01F9
            Left            =   960
            List            =   "frmEditor_Events.frx":020C
            Style           =   2  'Dropdown List
            TabIndex        =   321
            Top             =   360
            Width           =   3135
         End
      End
      Begin VB.Frame fraSetAccess 
         Caption         =   "Set Access"
         Height          =   1575
         Left            =   1200
         TabIndex        =   316
         Top             =   1800
         Visible         =   0   'False
         Width           =   4215
         Begin VB.ComboBox cmbSetAccess 
            Height          =   315
            ItemData        =   "frmEditor_Events.frx":024F
            Left            =   960
            List            =   "frmEditor_Events.frx":0262
            Style           =   2  'Dropdown List
            TabIndex        =   319
            Top             =   360
            Width           =   3135
         End
         Begin VB.CommandButton cmdSetAccess_Cancel 
            Caption         =   "Cancel"
            Height          =   375
            Left            =   2880
            TabIndex        =   318
            Top             =   1080
            Width           =   1215
         End
         Begin VB.CommandButton cmdSetAccess_Ok 
            Caption         =   "Ok"
            Height          =   375
            Left            =   1440
            TabIndex        =   317
            Top             =   1080
            Width           =   1215
         End
      End
      Begin VB.Frame fraPlaySound 
         Caption         =   "Play Sound"
         Height          =   1575
         Left            =   1200
         TabIndex        =   297
         Top             =   1800
         Visible         =   0   'False
         Width           =   4215
         Begin VB.CommandButton cmdPlaySound_Ok 
            Caption         =   "Ok"
            Height          =   375
            Left            =   1440
            TabIndex        =   300
            Top             =   1080
            Width           =   1215
         End
         Begin VB.CommandButton cmdPlaySound_Cancel 
            Caption         =   "Cancel"
            Height          =   375
            Left            =   2880
            TabIndex        =   299
            Top             =   1080
            Width           =   1215
         End
         Begin VB.ComboBox cmbPlaySound 
            Height          =   315
            ItemData        =   "frmEditor_Events.frx":02A5
            Left            =   960
            List            =   "frmEditor_Events.frx":02A7
            Style           =   2  'Dropdown List
            TabIndex        =   298
            Top             =   360
            Width           =   3135
         End
      End
      Begin VB.Frame fraPlayBGM 
         Caption         =   "Play BGM"
         Height          =   1575
         Left            =   1080
         TabIndex        =   293
         Top             =   1800
         Visible         =   0   'False
         Width           =   4335
         Begin VB.ComboBox cmbPlayBGM 
            Height          =   315
            ItemData        =   "frmEditor_Events.frx":02A9
            Left            =   1080
            List            =   "frmEditor_Events.frx":02AB
            Style           =   2  'Dropdown List
            TabIndex        =   296
            Top             =   360
            Width           =   3135
         End
         Begin VB.CommandButton cmdPlayBGM_Cancel 
            Caption         =   "Cancel"
            Height          =   375
            Left            =   3000
            TabIndex        =   295
            Top             =   1080
            Width           =   1215
         End
         Begin VB.CommandButton cmdPlayBGM_Ok 
            Caption         =   "Ok"
            Height          =   375
            Left            =   1560
            TabIndex        =   294
            Top             =   1080
            Width           =   1215
         End
      End
      Begin VB.Frame fraCustomScript 
         Caption         =   "Execute Custom Script"
         Height          =   1575
         Left            =   1080
         TabIndex        =   288
         Top             =   1800
         Visible         =   0   'False
         Width           =   4335
         Begin VB.HScrollBar scrlCustomScript 
            Height          =   255
            Left            =   1560
            Max             =   255
            Min             =   1
            TabIndex        =   292
            Top             =   360
            Value           =   1
            Width           =   2655
         End
         Begin VB.CommandButton cmdCustomScript_Ok 
            Caption         =   "Ok"
            Height          =   375
            Left            =   1560
            TabIndex        =   290
            Top             =   840
            Width           =   1215
         End
         Begin VB.CommandButton cmdCustomScript_Cancel 
            Caption         =   "Cancel"
            Height          =   375
            Left            =   3000
            TabIndex        =   289
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label lblCustomScript 
            Caption         =   "Case: 1"
            Height          =   255
            Left            =   120
            TabIndex        =   291
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Frame fraSelfSwitch 
         Caption         =   "Self Switch"
         Height          =   1695
         Left            =   1320
         TabIndex        =   212
         Top             =   1920
         Visible         =   0   'False
         Width           =   4095
         Begin VB.CommandButton cmdSelfSwitch_Cancel 
            Caption         =   "Cancel"
            Height          =   375
            Left            =   2760
            TabIndex        =   216
            Top             =   1200
            Width           =   1215
         End
         Begin VB.CommandButton cmdSelfSwitch_Ok 
            Caption         =   "Ok"
            Height          =   375
            Left            =   1440
            TabIndex        =   215
            Top             =   1200
            Width           =   1215
         End
         Begin VB.ComboBox cmbSetSelfSwitch 
            Height          =   315
            ItemData        =   "frmEditor_Events.frx":02AD
            Left            =   1440
            List            =   "frmEditor_Events.frx":02BD
            Style           =   2  'Dropdown List
            TabIndex        =   214
            Top             =   360
            Width           =   2535
         End
         Begin VB.ComboBox cmbSetSelfSwitchTo 
            Height          =   315
            ItemData        =   "frmEditor_Events.frx":02CD
            Left            =   960
            List            =   "frmEditor_Events.frx":02D7
            Style           =   2  'Dropdown List
            TabIndex        =   213
            Top             =   800
            Width           =   3015
         End
         Begin VB.Label lblRandomLabel 
            Caption         =   "Set to:"
            Height          =   255
            Index           =   26
            Left            =   120
            TabIndex        =   218
            Top             =   840
            Width           =   1815
         End
         Begin VB.Label lblRandomLabel 
            Caption         =   "Self Switch:"
            Height          =   255
            Index           =   24
            Left            =   120
            TabIndex        =   217
            Top             =   360
            Width           =   3855
         End
      End
      Begin VB.Frame fraPlayerSwitch 
         Caption         =   "Player Switch"
         Height          =   1695
         Left            =   1320
         TabIndex        =   205
         Top             =   1920
         Visible         =   0   'False
         Width           =   4095
         Begin VB.ComboBox cmbPlayerSwitchSet 
            Height          =   315
            ItemData        =   "frmEditor_Events.frx":02E4
            Left            =   960
            List            =   "frmEditor_Events.frx":02EE
            Style           =   2  'Dropdown List
            TabIndex        =   211
            Top             =   800
            Width           =   3015
         End
         Begin VB.ComboBox cmbSwitch 
            Height          =   315
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   208
            Top             =   360
            Width           =   3015
         End
         Begin VB.CommandButton cmbPlayerSwitch_Ok 
            Caption         =   "Ok"
            Height          =   375
            Left            =   1440
            TabIndex        =   207
            Top             =   1200
            Width           =   1215
         End
         Begin VB.CommandButton cmdPlayerSwitch_Cancel 
            Caption         =   "Cancel"
            Height          =   375
            Left            =   2760
            TabIndex        =   206
            Top             =   1200
            Width           =   1215
         End
         Begin VB.Label lblRandomLabel 
            Caption         =   "Switch:"
            Height          =   255
            Index           =   23
            Left            =   120
            TabIndex        =   210
            Top             =   360
            Width           =   3855
         End
         Begin VB.Label lblRandomLabel 
            Caption         =   "Set to:"
            Height          =   255
            Index           =   22
            Left            =   120
            TabIndex        =   209
            Top             =   840
            Width           =   1815
         End
      End
      Begin VB.Frame fraPlayerVar 
         Caption         =   "Player Variable"
         Height          =   1695
         Left            =   1320
         TabIndex        =   82
         Top             =   1920
         Visible         =   0   'False
         Width           =   4095
         Begin VB.CommandButton cmdVariableCancel 
            Caption         =   "Cancel"
            Height          =   375
            Left            =   2760
            TabIndex        =   88
            Top             =   1200
            Width           =   1215
         End
         Begin VB.CommandButton cmdVariableOK 
            Caption         =   "Ok"
            Height          =   375
            Left            =   1440
            TabIndex        =   87
            Top             =   1200
            Width           =   1215
         End
         Begin VB.TextBox txtVariable 
            Height          =   285
            Left            =   960
            TabIndex        =   86
            Top             =   840
            Width           =   3015
         End
         Begin VB.ComboBox cmbVariable 
            Height          =   315
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   84
            Top             =   360
            Width           =   3015
         End
         Begin VB.Label lblRandomLabel 
            Caption         =   "Set to:"
            Height          =   255
            Index           =   13
            Left            =   120
            TabIndex        =   85
            Top             =   840
            Width           =   1815
         End
         Begin VB.Label lblRandomLabel 
            Caption         =   "Variable:"
            Height          =   255
            Index           =   12
            Left            =   120
            TabIndex        =   83
            Top             =   360
            Width           =   3855
         End
      End
      Begin VB.Frame fraChangeSex 
         Caption         =   "Change Player Sex"
         Height          =   1455
         Left            =   1200
         TabIndex        =   262
         Top             =   1800
         Visible         =   0   'False
         Width           =   4095
         Begin VB.OptionButton optChangeSexFemale 
            Caption         =   "Female"
            Height          =   255
            Left            =   1920
            TabIndex        =   266
            Top             =   360
            Width           =   1455
         End
         Begin VB.OptionButton optChangeSexMale 
            Caption         =   "Male"
            Height          =   255
            Left            =   240
            TabIndex        =   265
            Top             =   360
            Width           =   1455
         End
         Begin VB.CommandButton cmdChangeSex_Ok 
            Caption         =   "Ok"
            Height          =   375
            Left            =   1200
            TabIndex        =   264
            Top             =   840
            Width           =   1215
         End
         Begin VB.CommandButton cmdChangeSex_Cancel 
            Caption         =   "Cancel"
            Height          =   375
            Left            =   2520
            TabIndex        =   263
            Top             =   840
            Width           =   1215
         End
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Positioning"
      Height          =   855
      Left            =   2760
      TabIndex        =   104
      Top             =   5880
      Width           =   3375
      Begin VB.ComboBox cmbPositioning 
         Height          =   315
         ItemData        =   "frmEditor_Events.frx":02FF
         Left            =   120
         List            =   "frmEditor_Events.frx":030C
         Style           =   2  'Dropdown List
         TabIndex        =   105
         Top             =   360
         Width           =   3135
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Global?"
      Height          =   615
      Left            =   2760
      TabIndex        =   101
      Top             =   7680
      Width           =   3375
      Begin VB.CheckBox chkGlobal 
         Caption         =   "Global**"
         Height          =   255
         Left            =   120
         TabIndex        =   102
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   9720
      TabIndex        =   36
      Top             =   8520
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   11280
      TabIndex        =   35
      Top             =   8520
      Width           =   1455
   End
   Begin VB.Frame Frame8 
      Caption         =   "General"
      Height          =   735
      Left            =   120
      TabIndex        =   26
      Top             =   120
      Width           =   12615
      Begin VB.CommandButton cmdClearPage 
         Caption         =   "Clear Page"
         Height          =   375
         Left            =   10920
         TabIndex        =   33
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdDeletePage 
         Caption         =   "Delete Page"
         Enabled         =   0   'False
         Height          =   375
         Left            =   9360
         TabIndex        =   32
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdPastePage 
         Caption         =   "Paste Page"
         Enabled         =   0   'False
         Height          =   375
         Left            =   7800
         TabIndex        =   31
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdCopyPage 
         Caption         =   "Copy Page"
         Height          =   375
         Left            =   6240
         TabIndex        =   30
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdNewPage 
         Caption         =   "New Page"
         Height          =   375
         Left            =   4680
         TabIndex        =   29
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   840
         TabIndex        =   28
         Top             =   300
         Width           =   3135
      End
      Begin VB.Label lblRandomLabel 
         Caption         =   "Name:"
         Height          =   255
         Index           =   32
         Left            =   120
         TabIndex        =   27
         Top             =   330
         Width           =   735
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Trigger"
      Height          =   735
      Left            =   2760
      TabIndex        =   24
      Top             =   6840
      Width           =   3375
      Begin VB.ComboBox cmbTrigger 
         Height          =   315
         ItemData        =   "frmEditor_Events.frx":0348
         Left            =   120
         List            =   "frmEditor_Events.frx":0355
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Options"
      Height          =   1455
      Left            =   360
      TabIndex        =   20
      Top             =   6840
      Width           =   2295
      Begin VB.CheckBox chkShowName 
         Caption         =   "Show Name"
         Height          =   255
         Left            =   120
         TabIndex        =   346
         Top             =   960
         Width           =   1695
      End
      Begin VB.CheckBox chkWalkThrough 
         Caption         =   "Walk Through"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   720
         Width           =   1695
      End
      Begin VB.CheckBox chkDirFix 
         Caption         =   "Direction Fix"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   480
         Width           =   1575
      End
      Begin VB.CheckBox chkWalkAnim 
         Caption         =   "No Walking Anim."
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Movement"
      Height          =   2175
      Left            =   2760
      TabIndex        =   13
      Top             =   3480
      Width           =   3375
      Begin VB.CommandButton cmdMoveRoute 
         Caption         =   "Move Route"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1680
         TabIndex        =   100
         Top             =   720
         Width           =   1575
      End
      Begin VB.ComboBox cmbMoveFreq 
         Height          =   315
         ItemData        =   "frmEditor_Events.frx":039A
         Left            =   840
         List            =   "frmEditor_Events.frx":03AD
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   1680
         Width           =   2415
      End
      Begin VB.ComboBox cmbMoveSpeed 
         Height          =   315
         ItemData        =   "frmEditor_Events.frx":03D9
         Left            =   840
         List            =   "frmEditor_Events.frx":03EF
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   1200
         Width           =   2415
      End
      Begin VB.ComboBox cmbMoveType 
         Height          =   315
         ItemData        =   "frmEditor_Events.frx":0432
         Left            =   840
         List            =   "frmEditor_Events.frx":043F
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label lblRandomLabel 
         Caption         =   "Freq:"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   18
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label lblRandomLabel 
         Caption         =   "Speed:"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   16
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label lblRandomLabel 
         Caption         =   "Type:"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   14
         Top             =   390
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Graphic"
      Height          =   3255
      Left            =   360
      TabIndex        =   11
      Top             =   3480
      Width           =   2295
      Begin VB.PictureBox picGraphic 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2895
         Left            =   240
         ScaleHeight     =   193
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   121
         TabIndex        =   12
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame fraRandom 
      Caption         =   "Conditions"
      Height          =   2055
      Index           =   0
      Left            =   360
      TabIndex        =   2
      Top             =   1320
      Width           =   5775
      Begin VB.ComboBox cmbPlayerVarCompare 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmEditor_Events.frx":0467
         Left            =   3720
         List            =   "frmEditor_Events.frx":047D
         Style           =   2  'Dropdown List
         TabIndex        =   311
         Top             =   240
         Width           =   1335
      End
      Begin VB.ComboBox cmbSelfSwitchCompare 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmEditor_Events.frx":04E3
         Left            =   3720
         List            =   "frmEditor_Events.frx":04ED
         Style           =   2  'Dropdown List
         TabIndex        =   310
         Top             =   1680
         Width           =   1095
      End
      Begin VB.ComboBox cmbPlayerSwitchCompare 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmEditor_Events.frx":04FE
         Left            =   3720
         List            =   "frmEditor_Events.frx":0508
         Style           =   2  'Dropdown List
         TabIndex        =   307
         Top             =   720
         Width           =   1095
      End
      Begin VB.ComboBox cmbSelfSwitch 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmEditor_Events.frx":0519
         Left            =   1920
         List            =   "frmEditor_Events.frx":052C
         Style           =   2  'Dropdown List
         TabIndex        =   99
         Top             =   1680
         Width           =   1455
      End
      Begin VB.CheckBox chkSelfSwitch 
         Caption         =   "Self Switch*"
         Height          =   255
         Left            =   120
         TabIndex        =   98
         Top             =   1680
         Width           =   1695
      End
      Begin VB.CheckBox chkHasItem 
         Caption         =   "Has Item"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1200
         Width           =   1695
      End
      Begin VB.ComboBox cmbHasItem 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmEditor_Events.frx":0552
         Left            =   1920
         List            =   "frmEditor_Events.frx":0554
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1200
         Width           =   1455
      End
      Begin VB.CheckBox chkPlayerSwitch 
         Caption         =   "Player Switch"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   1695
      End
      Begin VB.ComboBox cmbPlayerSwitch 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmEditor_Events.frx":0556
         Left            =   1920
         List            =   "frmEditor_Events.frx":0558
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   720
         Width           =   1455
      End
      Begin VB.CheckBox chkPlayerVar 
         Caption         =   "Player Variable"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1695
      End
      Begin VB.ComboBox cmbPlayerVar 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmEditor_Events.frx":055A
         Left            =   1920
         List            =   "frmEditor_Events.frx":055C
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txtPlayerVariable 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5160
         TabIndex        =   3
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lblRandomLabel 
         Caption         =   "is"
         Height          =   255
         Index           =   4
         Left            =   3480
         TabIndex        =   309
         Top             =   1760
         Width           =   255
      End
      Begin VB.Label lblRandomLabel 
         Caption         =   "is"
         Height          =   255
         Index           =   3
         Left            =   3480
         TabIndex        =   308
         Top             =   800
         Width           =   255
      End
      Begin VB.Label lblRandomLabel 
         Caption         =   "is"
         Height          =   255
         Index           =   5
         Left            =   3480
         TabIndex        =   6
         Top             =   340
         Width           =   255
      End
   End
   Begin VB.Frame fraCommands 
      Caption         =   "Commands"
      Height          =   6975
      Left            =   6240
      TabIndex        =   37
      Top             =   1320
      Visible         =   0   'False
      Width           =   6375
      Begin VB.CommandButton cmdCancelCommand 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   4560
         TabIndex        =   72
         Top             =   6360
         Width           =   1455
      End
      Begin VB.PictureBox picCommands 
         BorderStyle     =   0  'None
         Height          =   6015
         Index           =   1
         Left            =   240
         ScaleHeight     =   6015
         ScaleWidth      =   5775
         TabIndex        =   39
         Top             =   600
         Width           =   5775
         Begin VB.Frame fraRandom 
            Caption         =   "Player Control"
            Height          =   5535
            Index           =   3
            Left            =   3000
            TabIndex        =   52
            Top             =   0
            Width           =   2775
            Begin VB.CommandButton cmdGiveExp 
               Caption         =   "Give EXP"
               Height          =   375
               Left            =   120
               TabIndex        =   340
               Top             =   5040
               Width           =   2535
            End
            Begin VB.CommandButton cmdChangePK 
               Caption         =   "Change PK"
               Height          =   375
               Left            =   120
               TabIndex        =   219
               Top             =   4560
               Width           =   2535
            End
            Begin VB.CommandButton cmdChangeSex 
               Caption         =   "Change Sex"
               Height          =   375
               Left            =   120
               TabIndex        =   61
               Top             =   4080
               Width           =   2535
            End
            Begin VB.CommandButton cmdChangeSprite 
               Caption         =   "Change Sprite"
               Height          =   375
               Left            =   120
               TabIndex        =   60
               Top             =   3600
               Width           =   2535
            End
            Begin VB.CommandButton cmdChangeClass 
               Caption         =   "Change Class"
               Height          =   375
               Left            =   120
               TabIndex        =   59
               Top             =   3120
               Width           =   2535
            End
            Begin VB.CommandButton cmbChangeLevel 
               Caption         =   "Change Level"
               Height          =   375
               Left            =   120
               TabIndex        =   57
               Top             =   2160
               Width           =   2535
            End
            Begin VB.CommandButton cmdLevelUp 
               Caption         =   "Level Up"
               Height          =   375
               Left            =   120
               TabIndex        =   56
               Top             =   1680
               Width           =   2535
            End
            Begin VB.CommandButton cmdRestoreMp 
               Caption         =   "Restore Mp"
               Height          =   375
               Left            =   120
               TabIndex        =   55
               Top             =   1200
               Width           =   2535
            End
            Begin VB.CommandButton cmbRestoreHp 
               Caption         =   "Restore Hp"
               Height          =   375
               Left            =   120
               TabIndex        =   54
               Top             =   720
               Width           =   2535
            End
            Begin VB.CommandButton cmdChangeItems 
               Caption         =   "Change Items"
               Height          =   375
               Left            =   120
               TabIndex        =   53
               Top             =   240
               Width           =   2535
            End
            Begin VB.CommandButton cmdChangeSkills 
               Caption         =   "Change Skills"
               Height          =   375
               Left            =   120
               TabIndex        =   58
               Top             =   2640
               Width           =   2535
            End
         End
         Begin VB.Frame fraRandom 
            Caption         =   "Flow Control"
            Height          =   1335
            Index           =   2
            Left            =   0
            TabIndex        =   49
            Top             =   3720
            Width           =   2775
            Begin VB.CommandButton cmdAddConditionalBranch 
               Caption         =   "Conditional Branch"
               Height          =   375
               Left            =   120
               TabIndex        =   51
               Top             =   240
               Width           =   2535
            End
            Begin VB.CommandButton cmdExitEventProcess 
               Caption         =   "Exit Event Process"
               Height          =   375
               Left            =   120
               TabIndex        =   50
               Top             =   720
               Width           =   2535
            End
         End
         Begin VB.Frame fraRandom 
            Caption         =   "Event Progression"
            Height          =   1695
            Index           =   1
            Left            =   0
            TabIndex        =   45
            Top             =   1920
            Width           =   2775
            Begin VB.CommandButton cmdAddSelfSwitch 
               Caption         =   "Self Switch"
               Height          =   375
               Left            =   120
               TabIndex        =   48
               Top             =   1200
               Width           =   2535
            End
            Begin VB.CommandButton cmdPlayerSwitch 
               Caption         =   "Player Switch"
               Height          =   375
               Left            =   120
               TabIndex        =   47
               Top             =   720
               Width           =   2535
            End
            Begin VB.CommandButton cmdPlayerVar 
               Caption         =   "Player Variable"
               Height          =   375
               Left            =   120
               TabIndex        =   46
               Top             =   240
               Width           =   2535
            End
         End
         Begin VB.Frame Frame9 
            Caption         =   "Message"
            Height          =   1815
            Left            =   0
            TabIndex        =   41
            Top             =   0
            Width           =   2775
            Begin VB.CommandButton cmdShowChoices 
               Caption         =   "Show Choices"
               Height          =   375
               Left            =   120
               TabIndex        =   44
               Top             =   1200
               Width           =   2535
            End
            Begin VB.CommandButton cmdShowText 
               Caption         =   "Show Text"
               Height          =   375
               Left            =   120
               TabIndex        =   43
               Top             =   720
               Width           =   2535
            End
            Begin VB.CommandButton cmdAddText 
               Caption         =   "Add Chatbox Text"
               Height          =   375
               Left            =   120
               TabIndex        =   42
               Top             =   240
               Width           =   2535
            End
         End
      End
      Begin VB.PictureBox picCommands 
         BorderStyle     =   0  'None
         Height          =   5775
         Index           =   2
         Left            =   240
         ScaleHeight     =   5775
         ScaleWidth      =   5775
         TabIndex        =   40
         Top             =   600
         Visible         =   0   'False
         Width           =   5775
         Begin VB.Frame fraRandom 
            Caption         =   "Shop and Bank..."
            Height          =   1215
            Index           =   6
            Left            =   0
            TabIndex        =   312
            Top             =   2400
            Width           =   2775
            Begin VB.CommandButton cmdOpenShop 
               Caption         =   "Open Shop"
               Height          =   375
               Left            =   120
               TabIndex        =   314
               Top             =   720
               Width           =   2535
            End
            Begin VB.CommandButton cmdOpenBank 
               Caption         =   "Open Bank"
               Height          =   375
               Left            =   120
               TabIndex        =   313
               Top             =   240
               Width           =   2535
            End
         End
         Begin VB.Frame fraRandom 
            Caption         =   "Etc..."
            Height          =   1215
            Index           =   8
            Left            =   3000
            TabIndex        =   272
            Top             =   2400
            Width           =   2775
            Begin VB.CommandButton cmdSetAccess 
               Caption         =   "Set Access"
               Height          =   375
               Left            =   120
               TabIndex        =   315
               Top             =   240
               Width           =   2535
            End
            Begin VB.CommandButton cmdCustomScript 
               Caption         =   "Custom Script"
               Height          =   375
               Left            =   120
               TabIndex        =   273
               Top             =   720
               Width           =   2535
            End
         End
         Begin VB.Frame fraRandom 
            Caption         =   "Music and Sound"
            Height          =   2295
            Index           =   7
            Left            =   3000
            TabIndex        =   67
            Top             =   0
            Width           =   2775
            Begin VB.CommandButton cmdStopSound 
               Caption         =   "Stop Sound"
               Height          =   375
               Left            =   120
               TabIndex        =   71
               Top             =   1680
               Width           =   2535
            End
            Begin VB.CommandButton cmdPlaySound 
               Caption         =   "Play Sound"
               Height          =   375
               Left            =   120
               TabIndex        =   70
               Top             =   1200
               Width           =   2535
            End
            Begin VB.CommandButton cmdFadeuutBGM 
               Caption         =   "Fadeout BGM"
               Height          =   375
               Left            =   120
               TabIndex        =   69
               Top             =   720
               Width           =   2535
            End
            Begin VB.CommandButton cmdPlayBGM 
               Caption         =   "Play BGM"
               Height          =   375
               Left            =   120
               TabIndex        =   68
               Top             =   240
               Width           =   2535
            End
         End
         Begin VB.Frame fraRandom 
            Caption         =   "Animation"
            Height          =   855
            Index           =   5
            Left            =   0
            TabIndex        =   65
            Top             =   1440
            Width           =   2775
            Begin VB.CommandButton cmdPlayAnimation 
               Caption         =   "Play Animation"
               Height          =   375
               Left            =   120
               TabIndex        =   66
               Top             =   240
               Width           =   2535
            End
         End
         Begin VB.Frame fraRandom 
            Caption         =   "Movement"
            Height          =   1335
            Index           =   4
            Left            =   0
            TabIndex        =   62
            Top             =   0
            Width           =   2775
            Begin VB.CommandButton cmdMoveRouteCommand 
               Caption         =   "Set Move Route"
               Height          =   375
               Left            =   120
               TabIndex        =   64
               Top             =   720
               Width           =   2535
            End
            Begin VB.CommandButton cmdWarpPlayer 
               Caption         =   "Warp Player"
               Height          =   375
               Left            =   120
               TabIndex        =   63
               Top             =   240
               Width           =   2535
            End
         End
      End
      Begin MSComctlLib.TabStrip tabCommands 
         Height          =   6615
         Left            =   120
         TabIndex        =   38
         Top             =   240
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   11668
         MultiRow        =   -1  'True
         TabMinWidth     =   1764
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   2
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "1"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "2"
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame fraRandom 
      Caption         =   "Commands"
      Height          =   735
      Index           =   9
      Left            =   6240
      TabIndex        =   181
      Top             =   7560
      Width           =   6255
      Begin VB.CommandButton cmdClearCommand 
         Caption         =   "Clear"
         Height          =   375
         Left            =   4680
         TabIndex        =   185
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdDeleteCommand 
         Caption         =   "Delete"
         Height          =   375
         Left            =   3120
         TabIndex        =   184
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdEditCommand 
         Caption         =   "Edit"
         Height          =   375
         Left            =   1560
         TabIndex        =   183
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdAddCommand 
         Caption         =   "Add"
         Height          =   375
         Left            =   120
         TabIndex        =   182
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdLabel 
      Caption         =   "Label Variables/Switches"
      Height          =   375
      Left            =   120
      TabIndex        =   324
      Top             =   8520
      Width           =   2415
   End
   Begin VB.ListBox lstCommands 
      Height          =   6105
      Left            =   6240
      TabIndex        =   1
      Top             =   1440
      Width           =   6255
   End
   Begin MSComctlLib.TabStrip tabPages 
      Height          =   7455
      Left            =   120
      TabIndex        =   34
      Top             =   960
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   13150
      MultiRow        =   -1  'True
      ShowTips        =   0   'False
      TabMinWidth     =   529
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "1"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label lblRandomLabel 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "* Self Switches are always global and will reset on server restart."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   11
      Left            =   4680
      TabIndex        =   306
      Top             =   8520
      Width           =   4935
   End
   Begin VB.Label lblRandomLabel 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "** If global, only the first page will be processed. For shop keepers and such.(Experimental)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   14
      Left            =   2640
      TabIndex        =   103
      Top             =   8700
      Width           =   6975
   End
   Begin VB.Label lblRandomLabel 
      Caption         =   "List of commands:"
      Height          =   255
      Index           =   9
      Left            =   6240
      TabIndex        =   0
      Top             =   1560
      Width           =   6255
   End
End
Attribute VB_Name = "frmEditor_Events"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private copyPage As EventPageRec

Private Sub chkDirFix_Click()
    tmpEvent.Pages(curPageNum).DirFix = chkDirFix.Value
End Sub

Private Sub chkGlobal_Click()
    tmpEvent.Global = chkGlobal.Value
End Sub

Private Sub chkHasItem_Click()
    tmpEvent.Pages(curPageNum).chkHasItem = chkHasItem.Value
    If chkHasItem.Value = 0 Then cmbHasItem.Enabled = False Else cmbHasItem.Enabled = True
End Sub

Private Sub chkIgnoreMove_Click()
    tmpEvent.Pages(curPageNum).IgnoreMoveRoute = chkIgnoreMove.Value
End Sub

Private Sub chkPlayerSwitch_Click()
    tmpEvent.Pages(curPageNum).chkSwitch = chkPlayerSwitch.Value
    If chkPlayerSwitch.Value = 0 Then
        cmbPlayerSwitch.Enabled = False
        cmbPlayerSwitchCompare.Enabled = False
    Else
        cmbPlayerSwitch.Enabled = True
        cmbPlayerSwitchCompare.Enabled = True
    End If
End Sub

Private Sub chkPlayerVar_Click()
    tmpEvent.Pages(curPageNum).chkVariable = chkPlayerVar.Value
    If chkPlayerVar.Value = 0 Then
        cmbPlayerVar.Enabled = False
        txtPlayerVariable.Enabled = False
        cmbPlayerVarCompare.Enabled = False
    Else
        cmbPlayerVar.Enabled = True
        txtPlayerVariable.Enabled = True
        cmbPlayerVarCompare.Enabled = True
    End If
End Sub

Private Sub chkRepeatRoute_Click()
    tmpEvent.Pages(curPageNum).RepeatMoveRoute = chkRepeatRoute.Value
End Sub

Private Sub chkSelfSwitch_Click()
    tmpEvent.Pages(curPageNum).chkSelfSwitch = chkSelfSwitch.Value
    If chkSelfSwitch.Value = 0 Then
        cmbSelfSwitch.Enabled = False
        cmbSelfSwitchCompare.Enabled = False
    Else
        cmbSelfSwitch.Enabled = True
        cmbSelfSwitchCompare.Enabled = True
    End If
End Sub

Private Sub chkShowName_Click()
    tmpEvent.Pages(curPageNum).ShowName = chkShowName.Value
End Sub

Private Sub chkWalkAnim_Click()
    tmpEvent.Pages(curPageNum).WalkAnim = chkWalkAnim.Value
End Sub

Private Sub chkWalkThrough_Click()
    tmpEvent.Pages(curPageNum).WalkThrough = chkWalkThrough.Value
End Sub

Private Sub cmbChangeLevel_Click()
    scrlChangeLevel.Value = 1
    lblChangeLevel.Caption = "Level: 1"
    ' show
    fraDialogue.Visible = True
    fraChangeLevel.Visible = True
    ' hide
    fraCommands.Visible = False
End Sub

Private Sub cmbGraphic_Click()
    If cmbGraphic.ListIndex = -1 Then Exit Sub
    tmpEvent.Pages(curPageNum).GraphicType = cmbGraphic.ListIndex
    ' set the max on the scrollbar
    Select Case cmbGraphic.ListIndex
        Case 0 ' None
            scrlGraphic.Value = 1
            scrlGraphic.Enabled = False
        Case 1 ' character
            scrlGraphic.Max = NumCharacters
            scrlGraphic.Enabled = True
        Case 2 ' Tileset
            scrlGraphic.Max = NumTileSets
            scrlGraphic.Enabled = True
    End Select
    
    If scrlGraphic.Value = 0 Then
        lblGraphic.Caption = "Number: None"
    Else
        lblGraphic.Caption = "Number: " & scrlGraphic.Value
    End If
    
    If tmpEvent.Pages(curPageNum).GraphicType = 1 Then
        If frmEditor_Events.scrlGraphic.Value <= 0 Or frmEditor_Events.scrlGraphic.Value > NumCharacters Then Exit Sub
        CharacterTimer(frmEditor_Events.scrlGraphic.Value) = GetTickCount + SurfaceTimerMax
        If DDS_Character(frmEditor_Events.scrlGraphic.Value) Is Nothing Then
            Call InitDDSurf("Characters\" & frmEditor_Events.scrlGraphic.Value, DDSD_Character(frmEditor_Events.scrlGraphic.Value), DDS_Character(frmEditor_Events.scrlGraphic.Value))
        End If
                    
        If DDSD_Character(frmEditor_Events.scrlGraphic.Value).lWidth > 800 Then
            frmEditor_Events.hScrlGraphicSel.Visible = True
            frmEditor_Events.hScrlGraphicSel.Value = 0
            frmEditor_Events.hScrlGraphicSel.Max = DDSD_Character(frmEditor_Events.scrlGraphic.Value).lWidth - 800
        Else
            frmEditor_Events.hScrlGraphicSel.Visible = False
        End If
                    
        If DDSD_Character(frmEditor_Events.scrlGraphic.Value).lHeight > 512 Then
            frmEditor_Events.vScrlGraphicSel.Visible = True
            frmEditor_Events.vScrlGraphicSel.Value = 0
            frmEditor_Events.vScrlGraphicSel.Max = DDSD_Character(frmEditor_Events.scrlGraphic.Value).lHeight - 512
        Else
            frmEditor_Events.vScrlGraphicSel.Visible = False
        End If
    ElseIf tmpEvent.Pages(curPageNum).GraphicType = 2 Then
        If frmEditor_Events.scrlGraphic.Value <= 0 Or frmEditor_Events.scrlGraphic.Value > NumTileSets Then Exit Sub
        If DDS_Tileset(frmEditor_Events.scrlGraphic.Value) Is Nothing Then
            Call InitDDSurf("tilesets\" & frmEditor_Events.scrlGraphic.Value, DDSD_Tileset(frmEditor_Events.scrlGraphic.Value), DDS_Tileset(frmEditor_Events.scrlGraphic.Value))
        End If
                    
        If DDSD_Tileset(frmEditor_Events.scrlGraphic.Value).lWidth > 800 Then
            frmEditor_Events.hScrlGraphicSel.Visible = True
            frmEditor_Events.hScrlGraphicSel.Value = 0
            frmEditor_Events.hScrlGraphicSel.Max = DDSD_Tileset(frmEditor_Events.scrlGraphic.Value).lWidth - 800
        Else
            frmEditor_Events.hScrlGraphicSel.Visible = False
        End If
                    
        If DDSD_Tileset(frmEditor_Events.scrlGraphic.Value).lHeight > 512 Then
            frmEditor_Events.vScrlGraphicSel.Visible = True
            frmEditor_Events.vScrlGraphicSel.Value = 0
            frmEditor_Events.vScrlGraphicSel.Max = DDSD_Tileset(frmEditor_Events.scrlGraphic.Value).lHeight - 512
        Else
            frmEditor_Events.vScrlGraphicSel.Visible = False
        End If
    End If
End Sub

Private Sub cmbHasItem_Click()
    If cmbHasItem.ListIndex = -1 Then Exit Sub
    tmpEvent.Pages(curPageNum).HasItemIndex = cmbHasItem.ListIndex
End Sub

Private Sub cmbLabel_Ok_Click()
    fraLabeling.Visible = False
    SendSwitchesAndVariables
End Sub

Private Sub cmbMoveFreq_Click()
    If cmbMoveFreq.ListIndex = -1 Then Exit Sub
    tmpEvent.Pages(curPageNum).MoveFreq = cmbMoveFreq.ListIndex
End Sub

Private Sub cmbMoveSpeed_Click()
    If cmbMoveSpeed.ListIndex = -1 Then Exit Sub
    tmpEvent.Pages(curPageNum).MoveSpeed = cmbMoveSpeed.ListIndex
End Sub

Private Sub cmbMoveType_Click()
    If cmbMoveType.ListIndex = -1 Then Exit Sub
    tmpEvent.Pages(curPageNum).MoveType = cmbMoveType.ListIndex
    If cmbMoveType.ListIndex = 2 Then
        cmdMoveRoute.Enabled = True
    Else
        cmdMoveRoute.Enabled = False
    End If
End Sub

Private Sub cmbPlayerSwitch_Click()
    If cmbPlayerSwitch.ListIndex = -1 Then Exit Sub
    tmpEvent.Pages(curPageNum).SwitchIndex = cmbPlayerSwitch.ListIndex
End Sub

Private Sub cmbPlayerSwitch_Ok_Click()
    If Not isEdit Then
        AddCommand EventType.evPlayerSwitch
    Else
        EditCommand
    End If
    ' hide
    fraDialogue.Visible = False
    fraPlayerSwitch.Visible = False
    fraCommands.Visible = False
End Sub

Private Sub cmbPlayerSwitchCompare_Click()
    If cmbPlayerSwitchCompare.ListIndex = -1 Then Exit Sub
    tmpEvent.Pages(curPageNum).SwitchCompare = cmbPlayerSwitchCompare.ListIndex
End Sub

Private Sub cmbPlayerVar_Click()
    If cmbPlayerVar.ListIndex = -1 Then Exit Sub
    tmpEvent.Pages(curPageNum).VariableIndex = cmbPlayerVar.ListIndex
End Sub

Private Sub cmbPlayerVarCompare_Click()
    If cmbPlayerVarCompare.ListIndex = -1 Then Exit Sub
    tmpEvent.Pages(curPageNum).VariableCompare = cmbPlayerVarCompare.ListIndex
End Sub

Private Sub cmbPositioning_Click()
    If cmbPositioning.ListIndex = -1 Then Exit Sub
    tmpEvent.Pages(curPageNum).Position = cmbPositioning.ListIndex
End Sub

Private Sub cmbRestoreHp_Click()
    AddCommand EventType.evRestoreHP
    fraCommands.Visible = False
    fraDialogue.Visible = False
End Sub

Private Sub cmbSelfSwitch_Click()
    If cmbSelfSwitch.ListIndex = -1 Then Exit Sub
    tmpEvent.Pages(curPageNum).SelfSwitchIndex = cmbSelfSwitch.ListIndex
End Sub

Private Sub cmbSelfSwitchCompare_Click()
    If cmbSelfSwitchCompare.ListIndex = -1 Then Exit Sub
    tmpEvent.Pages(curPageNum).SelfSwitchCompare = cmbSelfSwitchCompare.ListIndex
End Sub

Private Sub cmbTrigger_Click()
    If cmbTrigger.ListIndex = -1 Then Exit Sub
    tmpEvent.Pages(curPageNum).Trigger = cmbTrigger.ListIndex
End Sub

Private Sub cmdAddCommand_Click()
    If lstCommands.ListIndex > -1 Then
        isEdit = False
        tabCommands.SelectedItem = tabCommands.Tabs(1)
        fraCommands.Visible = True
        picCommands(1).Visible = True
        picCommands(2).Visible = False
    End If
End Sub

Private Sub cmdAddConditionalBranch_Click()
    ' show
    fraDialogue.Visible = True
    fraConditionalBranch.Visible = True
    optCondition_Index(0).Value = True
    ClearConditionFrame
    cmbCondition_PlayerVarIndex.Enabled = True
    cmbCondition_PlayerVarCompare.Enabled = True
    txtCondition_PlayerVarCondition.Enabled = True
    ' hide
    fraCommands.Visible = False
End Sub

Private Sub cmdAddMoveRoute_Click(Index As Integer)
    If Index = 42 Then
        fraGraphic.width = 841
        fraGraphic.height = 585
        fraGraphic.Visible = True
        GraphicSelType = 1
    Else
        AddMoveRouteCommand Index
    End If
End Sub

Private Sub cmdAddSelfSwitch_Click()
    cmbSetSelfSwitch.ListIndex = 0
    cmbSetSelfSwitchTo.ListIndex = 0
    ' show
    fraDialogue.Visible = True
    fraSelfSwitch.Visible = True
    ' hide
    fraCommands.Visible = False
End Sub

Private Sub cmdAddText_Cancel_Click()
    If Not isEdit Then fraCommands.Visible = True Else fraCommands.Visible = False
    fraDialogue.Visible = False
    fraAddText.Visible = False
End Sub

Private Sub cmdAddText_Click()
    ' reset form
    txtAddText_Text.text = vbNullString
    scrlAddText_Colour.Value = 0
    optAddText_Player.Value = True
    ' show
    fraDialogue.Visible = True
    fraAddText.Visible = True
    ' hide
    fraCommands.Visible = False
End Sub

Private Sub cmdAddText_Ok_Click()
    If Not isEdit Then
        AddCommand EventType.evAddText
    Else
        EditCommand
    End If
    ' hide
    fraDialogue.Visible = False
    fraAddText.Visible = False
    fraCommands.Visible = False
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdCancelCommand_Click()
    fraCommands.Visible = False
End Sub

Private Sub cmdChangeClass_Cancel_Click()
    If Not isEdit Then fraCommands.Visible = True Else fraCommands.Visible = False
    fraDialogue.Visible = False
    fraChangeClass.Visible = False
End Sub

Private Sub cmdChangeClass_Click()
    cmbChangeClass.ListIndex = 0
    ' show
    fraDialogue.Visible = True
    fraChangeClass.Visible = True
    ' hide
    fraCommands.Visible = False
End Sub

Private Sub cmdChangeClass_Ok_Click()
    If isEdit = False Then
        AddCommand EventType.evChangeClass
    Else
        EditCommand
    End If
    ' hide
    fraDialogue.Visible = False
    fraChangeClass.Visible = False
    fraCommands.Visible = False
End Sub

Private Sub cmdChangeItems_Cancel_Click()
    If Not isEdit Then fraCommands.Visible = True Else fraCommands.Visible = False
    fraDialogue.Visible = False
    fraChangeItems.Visible = False
End Sub

Private Sub cmdChangeItems_Click()
    cmbChangeItemIndex.ListIndex = 0
    optChangeItemSet.Value = True
    txtChangeItemsAmount.text = "0"
    ' show
    fraDialogue.Visible = True
    fraChangeItems.Visible = True
    ' hide
    fraCommands.Visible = False
End Sub

Private Sub cmdChangeItems_Ok_Click()
    If isEdit = False Then
        AddCommand EventType.evChangeItems
    Else
        EditCommand
    End If
    ' hide
    fraDialogue.Visible = False
    fraCommands.Visible = False
    fraChangeItems.Visible = False
End Sub

Private Sub cmdChangeLevel_Cancel_Click()
    If Not isEdit Then fraCommands.Visible = True Else fraCommands.Visible = False
    fraDialogue.Visible = False
    fraChangeLevel.Visible = False
End Sub

Private Sub cmdChangeLevel_Ok_Click()
    If isEdit = False Then
        AddCommand EventType.evChangeLevel
    Else
        EditCommand
    End If
    ' hide
    fraDialogue.Visible = False
    fraChangeLevel.Visible = False
    fraCommands.Visible = False
End Sub

Private Sub cmdChangePK_Cancel_Click()
    If Not isEdit Then fraCommands.Visible = True Else fraCommands.Visible = False
    fraDialogue.Visible = False
    fraChangePK.Visible = False
End Sub

Private Sub cmdChangePK_Click()
    optChangePKYes.Value = True
    ' show
    fraDialogue.Visible = True
    fraChangePK.Visible = True
    ' hide
    fraCommands.Visible = False
End Sub

Private Sub cmdChangePK_Ok_Click()
    If isEdit = False Then
        AddCommand EventType.evChangePK
    Else
        EditCommand
    End If
    ' hide
    fraDialogue.Visible = False
    fraChangePK.Visible = False
    fraCommands.Visible = False
End Sub

Private Sub cmdChangeSex_Cancel_Click()
    If Not isEdit Then fraCommands.Visible = True Else fraCommands.Visible = False
    fraDialogue.Visible = False
    fraChangeSex.Visible = False
End Sub

Private Sub cmdChangeSex_Click()
    optChangeSexMale.Value = True
    ' show
    fraDialogue.Visible = True
    fraChangeSex.Visible = True
    ' hide
    fraCommands.Visible = False
End Sub

Private Sub cmdChangeSex_Ok_Click()
    If isEdit = False Then
        AddCommand EventType.evChangeSex
    Else
        EditCommand
    End If
    ' hide
    fraDialogue.Visible = False
    fraChangeSex.Visible = False
    fraCommands.Visible = False
End Sub

Private Sub cmdChangeSkills_Cancel_Click()
    If Not isEdit Then fraCommands.Visible = True Else fraCommands.Visible = False
    fraDialogue.Visible = False
    fraChangeSkills.Visible = False
End Sub

Private Sub cmdChangeSkills_Click()
    cmbChangeSkills.ListIndex = 0
    ' show
    fraDialogue.Visible = True
    fraChangeSkills.Visible = True
    ' hide
    fraCommands.Visible = False
End Sub

Private Sub cmdChangeSkills_Ok_Click()
    If isEdit = False Then
        AddCommand EventType.evChangeSkills
    Else
        EditCommand
    End If
    ' hide
    fraDialogue.Visible = False
    fraChangeSkills.Visible = False
    fraCommands.Visible = False
End Sub

Private Sub cmdChangeSprite_Cancel_Click()
    If Not isEdit Then fraCommands.Visible = True Else fraCommands.Visible = False
    fraDialogue.Visible = False
    fraChangeSprite.Visible = False
End Sub

Private Sub cmdChangeSprite_Click()
    scrlChangeSprite.Value = 1
    lblChangeSprite.Caption = "Sprite: 1"
    ' show
    fraDialogue.Visible = True
    fraChangeSprite.Visible = True
    ' hide
    fraCommands.Visible = False
End Sub

Private Sub cmdChangeSprite_Ok_Click()
    If isEdit = False Then
        AddCommand EventType.evChangeSprite
    Else
        EditCommand
    End If
    ' hide
    fraDialogue.Visible = False
    fraChangeSprite.Visible = False
    fraCommands.Visible = False
End Sub

Private Sub cmdChoices_Cancel_Click()
    If Not isEdit Then fraCommands.Visible = True Else fraCommands.Visible = False
    fraDialogue.Visible = False
    fraShowChoices.Visible = False
End Sub

Private Sub cmdChoices_Ok_Click()
    If Not isEdit Then
        AddCommand EventType.evShowChoices
    Else
        EditCommand
    End If
    ' hide
    fraDialogue.Visible = False
    fraShowChoices.Visible = False
    fraCommands.Visible = False
End Sub

Private Sub cmdClearCommand_Click()
    If MsgBox("Are you sure you want to clear all event commands?", vbYesNo, "Clear Event Commands?") = vbYes Then
        ClearEventCommands
    End If
End Sub

Private Sub cmdClearPage_Click()
    ZeroMemory ByVal VarPtr(tmpEvent.Pages(curPageNum)), LenB(tmpEvent.Pages(curPageNum))
    'EventEditorLoadPage curPageNum
End Sub

Private Sub cmdCondition_Cancel_Click()
    If Not isEdit Then fraCommands.Visible = True Else fraCommands.Visible = False
    fraDialogue.Visible = False
    fraConditionalBranch.Visible = False
End Sub

Private Sub cmdCondition_Ok_Click()
    If isEdit = False Then
        AddCommand EventType.evCondition
    Else
        EditCommand
    End If
    ' hide
    fraDialogue.Visible = False
    fraAddText.Visible = False
    fraConditionalBranch.Visible = False
End Sub

Private Sub cmdCopyPage_Click()
    CopyMemory ByVal VarPtr(copyPage), ByVal VarPtr(tmpEvent.Pages(curPageNum)), LenB(tmpEvent.Pages(curPageNum))
    cmdPastePage.Enabled = True
End Sub

Private Sub cmdCustomScript_Cancel_Click()
    If Not isEdit Then fraCommands.Visible = True Else fraCommands.Visible = False
    fraDialogue.Visible = False
    fraCustomScript.Visible = False
End Sub

Private Sub cmdCustomScript_Click()
    scrlCustomScript.Value = 1
    ' show
    fraDialogue.Visible = True
    fraCustomScript.Visible = True
    ' hide
    fraCommands.Visible = False
End Sub

Private Sub cmdCustomScript_Ok_Click()
    If Not isEdit Then
        AddCommand EventType.evCustomScript
    Else
        EditCommand
    End If
    ' hide
    fraDialogue.Visible = False
    fraCustomScript.Visible = False
    fraCommands.Visible = False
End Sub

Private Sub cmdDeleteCommand_Click()
    If lstCommands.ListIndex > -1 Then
        DeleteEventCommand
    End If
End Sub

Private Sub cmdDeletePage_Click()
Dim i As Long
    ZeroMemory ByVal VarPtr(tmpEvent.Pages(curPageNum)), LenB(tmpEvent.Pages(curPageNum))
    ' move everything else down a notch
    If curPageNum < tmpEvent.pageCount Then
        For i = curPageNum To tmpEvent.pageCount - 1
            CopyMemory ByVal VarPtr(tmpEvent.Pages(i)), ByVal VarPtr(tmpEvent.Pages(i + 1)), LenB(tmpEvent.Pages(i + 1))
        Next
    End If
    tmpEvent.pageCount = tmpEvent.pageCount - 1
    ' set the tabs
    tabPages.Tabs.Clear
    For i = 1 To tmpEvent.pageCount
        tabPages.Tabs.Add , , str(i)
    Next
    ' set the tab back
    If curPageNum <= tmpEvent.pageCount Then
        tabPages.SelectedItem = tabPages.Tabs(curPageNum)
    Else
        tabPages.SelectedItem = tabPages.Tabs(tmpEvent.pageCount)
    End If
    ' make sure we disable
    If tmpEvent.pageCount <= 1 Then
        cmdDeletePage.Enabled = False
    End If
End Sub

Private Sub cmdEditCommand_Click()
Dim i As Long
If lstCommands.ListIndex > -1 Then
    EditEventCommand
End If
End Sub

Private Sub cmdExitEventProcess_Click()
    AddCommand EventType.evExitProcess
    fraCommands.Visible = False
    fraDialogue.Visible = False
End Sub

Private Sub cmdFadeuutBGM_Click()
    AddCommand EventType.evFadeoutBGM
    fraCommands.Visible = False
    fraDialogue.Visible = False
End Sub

Private Sub cmdGiveExp_Cancel_Click()
    If Not isEdit Then fraCommands.Visible = True Else fraCommands.Visible = False
    fraDialogue.Visible = False
    fraGiveExp.Visible = False
End Sub

Private Sub cmdGiveExp_Click()
    scrlGiveExp.Value = 0
    lblGiveExp.Caption = "Give Exp: 0"
    ' show
    fraDialogue.Visible = True
    fraGiveExp.Visible = True
    ' hide
    fraCommands.Visible = False
End Sub

Private Sub cmdGiveExp_Ok_Click()
    If isEdit = False Then
        AddCommand EventType.evGiveExp
    Else
        EditCommand
    End If
    ' hide
    fraDialogue.Visible = False
    fraGiveExp.Visible = False
    fraCommands.Visible = False
End Sub

Private Sub cmdGraphicCancel_Click()
    fraGraphic.Visible = False
End Sub

Private Sub cmdGraphicOK_Click()
    If GraphicSelType = 0 Then
        tmpEvent.Pages(curPageNum).GraphicType = cmbGraphic.ListIndex
        tmpEvent.Pages(curPageNum).Graphic = scrlGraphic.Value
        tmpEvent.Pages(curPageNum).GraphicX = GraphicSelX
        tmpEvent.Pages(curPageNum).GraphicY = GraphicSelY
        tmpEvent.Pages(curPageNum).GraphicX2 = GraphicSelX2
        tmpEvent.Pages(curPageNum).GraphicY2 = GraphicSelY2
    Else
        AddMoveRouteCommand 42
        GraphicSelType = 0
    End If
    fraGraphic.Visible = False
End Sub

Private Sub cmdLabel_Cancel_Click()
    fraLabeling.Visible = False
    RequestSwitchesAndVariables
End Sub

Private Sub cmdLabel_Click()
Dim i As Long
    fraLabeling.Visible = True
    fraLabeling.width = 849
    fraLabeling.height = 593
    lstSwitches.Clear
    For i = 1 To MAX_SWITCHES
        lstSwitches.AddItem CStr(i) & ". " & Trim$(Switches(i))
    Next
    lstSwitches.ListIndex = 0
    
    lstVariables.Clear
    For i = 1 To MAX_VARIABLES
        lstVariables.AddItem CStr(i) & ". " & Trim$(Variables(i))
    Next
    lstVariables.ListIndex = 0
End Sub

Private Sub cmdLevelUp_Click()
    AddCommand EventType.evLevelUp
    fraCommands.Visible = False
    fraDialogue.Visible = False
End Sub

Private Sub cmdMoveRoute_Click()
Dim i As Long
    fraMoveRoute.Visible = True
    lstMoveRoute.Clear
    cmbEvent.Clear
    cmbEvent.AddItem "This Event"
    cmbEvent.ListIndex = 0
    cmbEvent.Enabled = False
    
    IsMoveRouteCommand = False
    
    chkIgnoreMove.Value = tmpEvent.Pages(curPageNum).IgnoreMoveRoute
    chkRepeatRoute.Value = tmpEvent.Pages(curPageNum).RepeatMoveRoute
    
    TempMoveRouteCount = tmpEvent.Pages(curPageNum).MoveRouteCount
    'Will it let me do this?
    TempMoveRoute = tmpEvent.Pages(curPageNum).MoveRoute
    
    For i = 1 To TempMoveRouteCount
        Select Case TempMoveRoute(i).Index
            Case 1
                lstMoveRoute.AddItem "Move Up"
            Case 2
                lstMoveRoute.AddItem "Move Down"
            Case 3
                lstMoveRoute.AddItem "Move Left"
            Case 4
                lstMoveRoute.AddItem "Move Right"
            Case 5
                lstMoveRoute.AddItem "Move Randomly"
            Case 6
                lstMoveRoute.AddItem "Move Towards Player"
            Case 7
                lstMoveRoute.AddItem "Move Away From Player"
            Case 8
                lstMoveRoute.AddItem "Step Forward"
            Case 9
                lstMoveRoute.AddItem "Step Back"
            Case 10
                lstMoveRoute.AddItem "Wait 100ms"
            Case 11
                lstMoveRoute.AddItem "Wait 500ms"
            Case 12
                lstMoveRoute.AddItem "Wait 1000ms"
            Case 13
                lstMoveRoute.AddItem "Turn Up"
            Case 14
                lstMoveRoute.AddItem "Turn Down"
            Case 15
                lstMoveRoute.AddItem "Turn Left"
            Case 16
                lstMoveRoute.AddItem "Turn Right"
            Case 17
                lstMoveRoute.AddItem "Turn 90 Degrees To the Right"
            Case 18
                lstMoveRoute.AddItem "Turn 90 Degrees To the Left"
            Case 19
                lstMoveRoute.AddItem "Turn Around 180 Degrees"
            Case 20
                lstMoveRoute.AddItem "Turn Randomly"
            Case 21
                lstMoveRoute.AddItem "Turn Towards Player"
            Case 22
                lstMoveRoute.AddItem "Turn Away from Player"
            Case 23
                lstMoveRoute.AddItem "Set Speed 8x Slower"
            Case 24
                lstMoveRoute.AddItem "Set Speed 4x Slower"
            Case 25
                lstMoveRoute.AddItem "Set Speed 2x Slower"
            Case 26
                lstMoveRoute.AddItem "Set Speed to Normal"
            Case 27
                lstMoveRoute.AddItem "Set Speed 2x Faster"
            Case 28
                lstMoveRoute.AddItem "Set Speed 4x Faster"
            Case 29
                lstMoveRoute.AddItem "Set Frequency Lowest"
            Case 30
                lstMoveRoute.AddItem "Set Frequency Lower"
            Case 31
                lstMoveRoute.AddItem "Set Frequency Normal"
            Case 32
                lstMoveRoute.AddItem "Set Frequency Higher"
            Case 33
                lstMoveRoute.AddItem "Set Frequency Highest"
            Case 34
                lstMoveRoute.AddItem "Turn On Walking Animation"
            Case 35
                lstMoveRoute.AddItem "Turn Off Walking Animation"
            Case 36
                lstMoveRoute.AddItem "Turn On Fixed Direction"
            Case 37
                lstMoveRoute.AddItem "Turn Off Fixed Direction"
            Case 38
                lstMoveRoute.AddItem "Turn On Walk Through"
            Case 39
                lstMoveRoute.AddItem "Turn Off Walk Through"
            Case 40
                lstMoveRoute.AddItem "Set Position Below Player"
            Case 41
                lstMoveRoute.AddItem "Set Position at Player Level"
            Case 42
                lstMoveRoute.AddItem "Set Position Above Player"
            Case 43
                lstMoveRoute.AddItem "Set Graphic"
        End Select
    Next
    
    fraMoveRoute.width = 841
    fraMoveRoute.height = 585
    fraMoveRoute.Visible = True
    
End Sub

Sub PopulateMoveRouteList()
Dim i As Long
    lstMoveRoute.Clear
    For i = 1 To TempMoveRouteCount
        Select Case TempMoveRoute(i).Index
            Case 1
                lstMoveRoute.AddItem "Move Up"
            Case 2
                lstMoveRoute.AddItem "Move Down"
            Case 3
                lstMoveRoute.AddItem "Move Left"
            Case 4
                lstMoveRoute.AddItem "Move Right"
            Case 5
                lstMoveRoute.AddItem "Move Randomly"
            Case 6
                lstMoveRoute.AddItem "Move Towards Player"
            Case 7
                lstMoveRoute.AddItem "Move Away From Player"
            Case 8
                lstMoveRoute.AddItem "Step Forward"
            Case 9
                lstMoveRoute.AddItem "Step Back"
            Case 10
                lstMoveRoute.AddItem "Wait 100ms"
            Case 11
                lstMoveRoute.AddItem "Wait 500ms"
            Case 12
                lstMoveRoute.AddItem "Wait 1000ms"
            Case 13
                lstMoveRoute.AddItem "Turn Up"
            Case 14
                lstMoveRoute.AddItem "Turn Down"
            Case 15
                lstMoveRoute.AddItem "Turn Left"
            Case 16
                lstMoveRoute.AddItem "Turn Right"
            Case 17
                lstMoveRoute.AddItem "Turn 90 Degrees To the Right"
            Case 18
                lstMoveRoute.AddItem "Turn 90 Degrees To the Left"
            Case 19
                lstMoveRoute.AddItem "Turn Around 180 Degrees"
            Case 20
                lstMoveRoute.AddItem "Turn Randomly"
            Case 21
                lstMoveRoute.AddItem "Turn Towards Player"
            Case 22
                lstMoveRoute.AddItem "Turn Away from Player"
            Case 23
                lstMoveRoute.AddItem "Set Speed 8x Slower"
            Case 24
                lstMoveRoute.AddItem "Set Speed 4x Slower"
            Case 25
                lstMoveRoute.AddItem "Set Speed 2x Slower"
            Case 26
                lstMoveRoute.AddItem "Set Speed to Normal"
            Case 27
                lstMoveRoute.AddItem "Set Speed 2x Faster"
            Case 28
                lstMoveRoute.AddItem "Set Speed 4x Faster"
            Case 29
                lstMoveRoute.AddItem "Set Frequency Lowest"
            Case 30
                lstMoveRoute.AddItem "Set Frequency Lower"
            Case 31
                lstMoveRoute.AddItem "Set Frequency Normal"
            Case 32
                lstMoveRoute.AddItem "Set Frequency Higher"
            Case 33
                lstMoveRoute.AddItem "Set Frequency Highest"
            Case 34
                lstMoveRoute.AddItem "Turn On Walking Animation"
            Case 35
                lstMoveRoute.AddItem "Turn Off Walking Animation"
            Case 36
                lstMoveRoute.AddItem "Turn On Fixed Direction"
            Case 37
                lstMoveRoute.AddItem "Turn Off Fixed Direction"
            Case 38
                lstMoveRoute.AddItem "Turn On Walk Through"
            Case 39
                lstMoveRoute.AddItem "Turn Off Walk Through"
            Case 40
                lstMoveRoute.AddItem "Set Position Below Player"
            Case 41
                lstMoveRoute.AddItem "Set Position at Player Level"
            Case 42
                lstMoveRoute.AddItem "Set Position Above Player"
            Case 43
                lstMoveRoute.AddItem "Set Graphic"
        End Select
    Next
End Sub

Private Sub cmdMoveRouteCancel_Click()
    TempMoveRouteCount = 0
    ReDim TempMoveRoute(0)
    fraMoveRoute.Visible = False
End Sub

Private Sub cmdMoveRouteCommand_Click()
Dim i As Long, x As Long
    fraMoveRoute.Visible = True
    lstMoveRoute.Clear
    cmbEvent.Clear
    ReDim ListOfEvents(0 To Map.EventCount)
    ListOfEvents(0) = EditorEvent
    cmbEvent.AddItem "This Event"
    cmbEvent.ListIndex = 0
    cmbEvent.Enabled = True

    For i = 1 To Map.EventCount
        If i <> EditorEvent Then
            cmbEvent.AddItem Trim$(Map.Events(i).Name)
            x = x + 1
            ListOfEvents(x) = i
        End If
    Next

    
    IsMoveRouteCommand = True
    
    chkIgnoreMove.Value = 0
    chkRepeatRoute.Value = 0
    
    TempMoveRouteCount = 0

    ReDim TempMoveRoute(0)
    
    fraMoveRoute.width = 841
    fraMoveRoute.height = 585
    fraMoveRoute.Visible = True
    
    ' hide
    fraCommands.Visible = False
End Sub

Private Sub cmdMoveRouteOk_Click()
    If IsMoveRouteCommand = True Then
        If Not isEdit Then
            AddCommand EventType.evSetMoveRoute
        Else
            EditCommand
        End If
        TempMoveRouteCount = 0
        ReDim TempMoveRoute(0)
        fraMoveRoute.Visible = False
    Else
        tmpEvent.Pages(curPageNum).MoveRouteCount = TempMoveRouteCount
        tmpEvent.Pages(curPageNum).MoveRoute = TempMoveRoute
        TempMoveRouteCount = 0
        ReDim TempMoveRoute(0)
        fraMoveRoute.Visible = False
    End If
End Sub

Private Sub cmdNewPage_Click()
Dim pageCount As Long, i As Long
    pageCount = tmpEvent.pageCount + 1
    ' redim the array
    ReDim Preserve tmpEvent.Pages(pageCount)
    tmpEvent.pageCount = pageCount
    ' set the tabs
    tabPages.Tabs.Clear
    For i = 1 To tmpEvent.pageCount
        tabPages.Tabs.Add , , str(i)
    Next
    cmdDeletePage.Enabled = True
End Sub

Private Sub cmdOk_Click()
    EventEditorOK
End Sub

Private Sub cmdOpenBank_Click()
    AddCommand EventType.evOpenBank
    fraCommands.Visible = False
    fraDialogue.Visible = False
End Sub

Private Sub cmdOpenShop_Cancel_Click()
    If Not isEdit Then fraCommands.Visible = True Else fraCommands.Visible = False
    fraDialogue.Visible = False
    fraOpenShop.Visible = False
End Sub

Private Sub cmdOpenShop_Click()
    ' show
    fraDialogue.Visible = True
    fraOpenShop.Visible = True
    cmbOpenShop.ListIndex = 0
    ' hide
    fraCommands.Visible = False
End Sub

Private Sub cmdOpenShop_Ok_Click()
    If Not isEdit Then
        AddCommand EventType.evOpenShop
    Else
        EditCommand
    End If
    ' hide
    fraDialogue.Visible = False
    fraOpenShop.Visible = False
    fraCommands.Visible = False
End Sub

Private Sub cmdPastePage_Click()
    CopyMemory ByVal VarPtr(tmpEvent.Pages(curPageNum)), ByVal VarPtr(copyPage), LenB(tmpEvent.Pages(curPageNum))
    EventEditorLoadPage curPageNum
End Sub

Private Sub cmdPlayAnim_Cancel_Click()
    If Not isEdit Then fraCommands.Visible = True Else fraCommands.Visible = False
    fraDialogue.Visible = False
    fraPlayAnimation.Visible = False
End Sub

Private Sub cmdPlayAnim_Ok_Click()
    If Not isEdit Then
        AddCommand EventType.evPlayAnimation
    Else
        EditCommand
    End If
    ' hide
    fraDialogue.Visible = False
    fraPlayAnimation.Visible = False
    fraCommands.Visible = False
End Sub

Private Sub cmdPlayAnimation_Click()
Dim i As Long
    cmbPlayAnimEvent.Clear
    For i = 1 To Map.EventCount
        cmbPlayAnimEvent.AddItem i & ". " & Trim$(Map.Events(i).Name)
    Next
    cmbPlayAnimEvent.ListIndex = 0
    
    optPlayAnimPlayer.Value = True
    
    cmbPlayAnim.ListIndex = 0
    
    lblPlayAnimX.Caption = "Map Tile X: 0"
    lblPlayAnimY.Caption = "Map Tile Y: 0"
    scrlPlayAnimTileX.Value = 0
    scrlPlayAnimTileY.Value = 0
    scrlPlayAnimTileX.Max = Map.MaxX
    scrlPlayAnimTileY.Max = Map.MaxY
    ' show
    fraDialogue.Visible = True
    fraPlayAnimation.Visible = True
    ' hide
    fraCommands.Visible = False
    lblPlayAnimX.Visible = False
    lblPlayAnimY.Visible = False
    scrlPlayAnimTileX.Visible = False
    scrlPlayAnimTileY.Visible = False
    cmbPlayAnimEvent.Visible = False
End Sub

Private Sub cmdPlayBGM_Cancel_Click()
    If Not isEdit Then fraCommands.Visible = True Else fraCommands.Visible = False
    fraDialogue.Visible = False
    fraPlayBGM.Visible = False
End Sub

Private Sub cmdPlayBGM_Click()
    cmbPlayBGM.ListIndex = 0
    ' show
    fraDialogue.Visible = True
    fraPlayBGM.Visible = True
    ' hide
    fraCommands.Visible = False
End Sub

Private Sub cmdPlayBGM_Ok_Click()
    If Not isEdit Then
        AddCommand EventType.evPlayBGM
    Else
        EditCommand
    End If
    ' hide
    fraDialogue.Visible = False
    fraPlayBGM.Visible = False
    fraCommands.Visible = False
End Sub

Private Sub cmdPlayerSwitch_Cancel_Click()
    If Not isEdit Then fraCommands.Visible = True Else fraCommands.Visible = False
    fraDialogue.Visible = False
    fraPlayerSwitch.Visible = False
End Sub

Private Sub cmdPlayerSwitch_Click()
    cmbPlayerSwitchSet.ListIndex = 0
    cmbSwitch.ListIndex = 0
    ' show
    fraDialogue.Visible = True
    fraPlayerSwitch.Visible = True
    ' hide
    fraCommands.Visible = False
End Sub

Private Sub cmdPlayerVar_Click()
Dim i As Long
    txtVariable.text = vbNullString
    cmbVariable.ListIndex = 0
    ' show
    fraDialogue.Visible = True
    fraPlayerVar.Visible = True
    ' hide
    fraCommands.Visible = False
End Sub

Private Sub cmdPlaySound_Cancel_Click()
    If Not isEdit Then fraCommands.Visible = True Else fraCommands.Visible = False
    fraDialogue.Visible = False
    fraPlaySound.Visible = False
End Sub

Private Sub cmdPlaySound_Click()
    cmbPlaySound.ListIndex = 0
    ' show
    fraDialogue.Visible = True
    fraPlaySound.Visible = True
    ' hide
    fraCommands.Visible = False
End Sub

Private Sub cmdPlaySound_Ok_Click()
    If Not isEdit Then
        AddCommand EventType.evPlaySound
    Else
        EditCommand
    End If
    ' hide
    fraDialogue.Visible = False
    fraPlaySound.Visible = False
    fraCommands.Visible = False
End Sub

Private Sub cmdRename_Cancel_Click()
Dim i As Long
    fraRenaming.Visible = False
    RenameType = 0
    RenameIndex = 0
    lstSwitches.Clear
    For i = 1 To MAX_SWITCHES
        lstSwitches.AddItem CStr(i) & ". " & Trim$(Switches(i))
    Next
    lstSwitches.ListIndex = 0
    
    lstVariables.Clear
    For i = 1 To MAX_VARIABLES
        lstVariables.AddItem CStr(i) & ". " & Trim$(Variables(i))
    Next
    lstVariables.ListIndex = 0
End Sub

Private Sub cmdRename_Ok_Click()
Dim i As Long
    Select Case RenameType
        Case 1
            'Variable
            If RenameIndex > 0 And RenameIndex <= MAX_VARIABLES + 1 Then
                Variables(RenameIndex) = txtRename.text
                fraRenaming.Visible = False
                RenameType = 0
                RenameIndex = 0
            End If
        Case 2
            'Switch
            If RenameIndex > 0 And RenameIndex <= MAX_SWITCHES + 1 Then
                Switches(RenameIndex) = txtRename.text
                fraRenaming.Visible = False
                RenameType = 0
                RenameIndex = 0
            End If
    End Select
    
    lstSwitches.Clear
    For i = 1 To MAX_SWITCHES
        lstSwitches.AddItem CStr(i) & ". " & Trim$(Switches(i))
    Next
    lstSwitches.ListIndex = 0
    
    lstVariables.Clear
    For i = 1 To MAX_VARIABLES
        lstVariables.AddItem CStr(i) & ". " & Trim$(Variables(i))
    Next
    lstVariables.ListIndex = 0
End Sub

Private Sub cmdRenameSwitch_Click()
    If lstSwitches.ListIndex > -1 And lstSwitches.ListIndex < MAX_SWITCHES Then
        fraRenaming.Visible = True
        lblEditing.Caption = "Editing Switch #" & CStr(lstSwitches.ListIndex + 1)
        txtRename.text = Switches(lstSwitches.ListIndex + 1)
        RenameType = 2
        RenameIndex = lstSwitches.ListIndex + 1
    End If
End Sub

Private Sub cmdRenameVariable_Click()
    If lstVariables.ListIndex > -1 And lstVariables.ListIndex < MAX_VARIABLES Then
        fraRenaming.Visible = True
        lblEditing.Caption = "Editing Variable #" & CStr(lstVariables.ListIndex + 1)
        txtRename.text = Variables(lstVariables.ListIndex + 1)
        RenameType = 1
        RenameIndex = lstVariables.ListIndex + 1
    End If
End Sub

Private Sub cmdRestoreMp_Click()
    AddCommand EventType.evRestoreMP
    fraCommands.Visible = False
    fraDialogue.Visible = False
End Sub

Private Sub cmdSelfSwitch_Cancel_Click()
    If Not isEdit Then fraCommands.Visible = True Else fraCommands.Visible = False
    fraDialogue.Visible = False
    fraSelfSwitch.Visible = False
End Sub

Private Sub cmdSelfSwitch_Ok_Click()
    If Not isEdit Then
        AddCommand EventType.evSelfSwitch
    Else
        EditCommand
    End If
    ' hide
    fraDialogue.Visible = False
    fraSelfSwitch.Visible = False
    fraCommands.Visible = False
End Sub

Private Sub cmdSetAccess_Cancel_Click()
    If Not isEdit Then fraCommands.Visible = True Else fraCommands.Visible = False
    fraDialogue.Visible = False
    fraSetAccess.Visible = False
End Sub

Private Sub cmdSetAccess_Click()
    cmbSetAccess.ListIndex = 0
    ' show
    fraDialogue.Visible = True
    fraSetAccess.Visible = True
    ' hide
    fraCommands.Visible = False
End Sub

Private Sub cmdSetAccess_Ok_Click()
    If Not isEdit Then
        AddCommand EventType.evSetAccess
    Else
        EditCommand
    End If
    ' hide
    fraDialogue.Visible = False
    fraSetAccess.Visible = False
    fraCommands.Visible = False
End Sub

Private Sub cmdShowChoices_Click()
    ' reset form
    txtChoicePrompt.text = vbNullString
    txtChoices(1).text = vbNullString
    txtChoices(2).text = vbNullString
    txtChoices(3).text = vbNullString
    txtChoices(4).text = vbNullString
    ' show
    fraDialogue.Visible = True
    fraShowChoices.Visible = True
    ' hide
    fraCommands.Visible = False
End Sub

Private Sub cmdShowText_Cancel_Click()
    If Not isEdit Then fraCommands.Visible = True Else fraCommands.Visible = False
    fraDialogue.Visible = False
    fraShowText.Visible = False
End Sub

Private Sub cmdShowText_Click()
    ' reset form
    txtShowText.text = vbNullString
    ' show
    fraDialogue.Visible = True
    fraShowText.Visible = True
    ' hide
    fraCommands.Visible = False
End Sub

Private Sub cmdShowText_Ok_Click()
    If Not isEdit Then
        AddCommand EventType.evShowText
    Else
        EditCommand
    End If
    ' hide
    fraDialogue.Visible = False
    fraShowText.Visible = False
    fraCommands.Visible = False
End Sub

Private Sub cmdStopSound_Click()
    AddCommand EventType.evStopSound
    fraCommands.Visible = False
    fraDialogue.Visible = False
End Sub

Private Sub cmdVariableCancel_Click()
    If Not isEdit Then fraCommands.Visible = True Else fraCommands.Visible = False
    fraDialogue.Visible = False
    fraPlayerVar.Visible = False
End Sub

Private Sub cmdVariableOK_Click()
    If Not isEdit Then
        AddCommand EventType.evPlayerVar
    Else
        EditCommand
    End If
    ' hide
    fraDialogue.Visible = False
    fraPlayerVar.Visible = False
    fraCommands.Visible = False
End Sub

Private Sub cmdWarpPlayer_Click()
    ' reset form
    scrlWPMap.Value = 0
    scrlWPX.Value = 0
    scrlWPY.Value = 0
    cmbWarpPlayerDir.ListIndex = 0
    ' show
    fraDialogue.Visible = True
    fraWarpPlayer.Visible = True
    ' hide
    fraCommands.Visible = False
End Sub

Private Sub cmdWPCancel_Click()
    If Not isEdit Then fraCommands.Visible = True Else fraCommands.Visible = False
    fraDialogue.Visible = False
    fraWarpPlayer.Visible = False
End Sub

Private Sub cmdWPOkay_Click()
    If Not isEdit Then
        AddCommand EventType.evWarpPlayer
    Else
        EditCommand
    End If
    ' hide
    fraDialogue.Visible = False
    fraWarpPlayer.Visible = False
    fraCommands.Visible = False
End Sub

Private Sub Command4_Click()

End Sub

Private Sub Command15_Click()

End Sub

Private Sub Command16_Click()

End Sub

Private Sub Command17_Click()

End Sub

Private Sub Command21_Click()

End Sub

Private Sub Command25_Click()

End Sub

Private Sub Command30_Click()

End Sub

Private Sub Command33_Click()

End Sub

Private Sub Form_Load()
Dim i As Long
    cmbSwitch.Clear
    For i = 1 To MAX_SWITCHES
        cmbSwitch.AddItem i & ". " & Switches(i)
    Next
    
    cmbSwitch.ListIndex = 0
    
    cmbVariable.Clear
    For i = 1 To MAX_VARIABLES
        cmbVariable.AddItem i & ". " & Variables(i)
    Next
    
    cmbVariable.ListIndex = 0
    
    cmbChangeItemIndex.Clear
    For i = 1 To MAX_ITEMS
        cmbChangeItemIndex.AddItem Trim$(Item(i).Name)
    Next
    
    cmbChangeItemIndex.ListIndex = 0
    
    scrlChangeLevel.Min = 1
    scrlChangeLevel.Max = MAX_LEVELS
    scrlChangeLevel.Value = 1
    lblChangeLevel.Caption = "Level: 1"
    
    cmbChangeSkills.Clear
    For i = 1 To MAX_SPELLS
        cmbChangeSkills.AddItem Trim$(Spell(i).Name)
    Next
    
    cmbChangeSkills.ListIndex = 0
    
    cmbChangeClass.Clear
    For i = 1 To Max_Classes
        cmbChangeClass.AddItem Trim$(Class(i).Name)
    Next
    cmbChangeClass.ListIndex = 0
    
    scrlChangeSprite.Max = NumCharacters
    
    cmbPlayAnim.Clear
    For i = 1 To MAX_ANIMATIONS
        cmbPlayAnim.AddItem i & ". " & Trim$(Animation(i).Name)
    Next
    cmbPlayAnim.ListIndex = 0
    PopulateLists
    cmbPlayBGM.Clear
    For i = 1 To UBound(musicCache)
        cmbPlayBGM.AddItem (musicCache(i))
    Next
    cmbPlayBGM.ListIndex = 0
    
    cmbPlaySound.Clear
    For i = 1 To UBound(soundCache)
        cmbPlaySound.AddItem (soundCache(i))
    Next
    cmbPlaySound.ListIndex = 0
    
    cmbOpenShop.Clear
    For i = 1 To MAX_SHOPS
        cmbOpenShop.AddItem i & ". " & Trim$(Shop(i).Name)
    Next
    
    cmbOpenShop.ListIndex = 0
End Sub

Private Sub lstCommands_Click()
    curCommand = lstCommands.ListIndex + 1
End Sub

Sub AddMoveRouteCommand(Index As Integer)
Dim i As Long, x As Long, z As Long
    Index = Index + 1
    If lstMoveRoute.ListIndex > -1 Then
        i = lstMoveRoute.ListIndex + 1
        TempMoveRouteCount = TempMoveRouteCount + 1
        ReDim Preserve TempMoveRoute(TempMoveRouteCount)
        For x = TempMoveRouteCount - 1 To i Step -1
            TempMoveRoute(x + 1) = TempMoveRoute(x)
        Next
        TempMoveRoute(i).Index = Index
        'if set graphic then...
        If Index = 43 Then
            TempMoveRoute(i).Data1 = cmbGraphic.ListIndex
            TempMoveRoute(i).Data2 = scrlGraphic.Value
            TempMoveRoute(i).Data3 = GraphicSelX
            TempMoveRoute(i).Data4 = GraphicSelX2
            TempMoveRoute(i).Data5 = GraphicSelY
            TempMoveRoute(i).Data6 = GraphicSelY2
        End If
        PopulateMoveRouteList
    Else
        TempMoveRouteCount = TempMoveRouteCount + 1
        ReDim Preserve TempMoveRoute(TempMoveRouteCount)
        TempMoveRoute(TempMoveRouteCount).Index = Index
        PopulateMoveRouteList
        'if set graphic then....
        If Index = 43 Then
            TempMoveRoute(TempMoveRouteCount).Data1 = cmbGraphic.ListIndex
            TempMoveRoute(TempMoveRouteCount).Data2 = scrlGraphic.Value
            TempMoveRoute(TempMoveRouteCount).Data3 = GraphicSelX
            TempMoveRoute(TempMoveRouteCount).Data4 = GraphicSelX2
            TempMoveRoute(TempMoveRouteCount).Data5 = GraphicSelY
            TempMoveRoute(TempMoveRouteCount).Data6 = GraphicSelY2
        End If
    End If
End Sub

Sub RemoveMoveRouteCommand(Index As Long)
Dim i As Long
    Index = Index + 1
    If Index > 0 And Index <= TempMoveRouteCount Then
        For i = Index + 1 To TempMoveRouteCount
            TempMoveRoute(i - 1) = TempMoveRoute(i)
        Next
        TempMoveRouteCount = TempMoveRouteCount - 1
        If TempMoveRouteCount = 0 Then
            ReDim TempMoveRoute(0)
        Else
            ReDim Preserve TempMoveRoute(TempMoveRouteCount)
        End If
        PopulateMoveRouteList
    End If
End Sub

Private Sub lstCommands_DblClick()
    cmdAddCommand_Click
End Sub

Private Sub lstCommands_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then
        'remove move route command lol
        cmdDeleteCommand_Click
    End If
End Sub

Private Sub lstMoveRoute_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then
        'remove move route command lol
        If lstMoveRoute.ListIndex > -1 Then
            Call RemoveMoveRouteCommand(lstMoveRoute.ListIndex)
        End If
    End If
End Sub

Private Sub optAddText_Game_Click()

End Sub

Private Sub lstSwitches_DblClick()
    If lstSwitches.ListIndex > -1 And lstSwitches.ListIndex < MAX_SWITCHES Then
        fraRenaming.Visible = True
        lblEditing.Caption = "Editing Switch #" & CStr(lstSwitches.ListIndex + 1)
        txtRename.text = Switches(lstSwitches.ListIndex + 1)
        RenameType = 2
        RenameIndex = lstSwitches.ListIndex + 1
    End If
End Sub

Private Sub lstVariables_DblClick()
    If lstVariables.ListIndex > -1 And lstVariables.ListIndex < MAX_VARIABLES Then
        fraRenaming.Visible = True
        lblEditing.Caption = "Editing Variable #" & CStr(lstVariables.ListIndex + 1)
        txtRename.text = Variables(lstVariables.ListIndex + 1)
        RenameType = 1
        RenameIndex = lstVariables.ListIndex + 1
    End If
End Sub

Private Sub optCondition_Index_Click(Index As Integer)
Dim i As Long, x As Long
    For i = 0 To 6
        If optCondition_Index(i).Value = True Then x = i
    Next
    ClearConditionFrame
    Select Case x
        Case 0
            cmbCondition_PlayerVarIndex.Enabled = True
            cmbCondition_PlayerVarCompare.Enabled = True
            txtCondition_PlayerVarCondition.Enabled = True
        Case 1
            cmbCondition_PlayerSwitch.Enabled = True
            cmbCondtion_PlayerSwitchCondition.Enabled = True
        Case 2
            cmbCondition_HasItem.Enabled = True
        Case 3
            cmbCondition_ClassIs.Enabled = True
        Case 4
            cmbCondition_LearntSkill.Enabled = True
        Case 5
            cmbCondition_LevelCompare.Enabled = True
            txtCondition_LevelAmount.Enabled = True
        Case 6
            cmbCondition_SelfSwitch.Enabled = True
            cmbCondition_SelfSwitchCondition.Enabled = True
    End Select
End Sub
Sub ClearConditionFrame()
Dim i As Long
    cmbCondition_PlayerVarIndex.Enabled = False
    cmbCondition_PlayerVarIndex.Clear
    For i = 1 To MAX_VARIABLES
        cmbCondition_PlayerVarIndex.AddItem i & ". " & Variables(i)
    Next
    cmbCondition_PlayerVarIndex.ListIndex = 0
    
    cmbCondition_PlayerVarCompare.ListIndex = 0
    cmbCondition_PlayerVarCompare.Enabled = False
    
    txtCondition_PlayerVarCondition.Enabled = False
    txtCondition_PlayerVarCondition.text = "0"
    
    cmbCondition_PlayerSwitch.Enabled = False
    cmbCondition_PlayerSwitch.Clear
    For i = 1 To MAX_SWITCHES
        cmbCondition_PlayerSwitch.AddItem i & ". " & Switches(i)
    Next
    cmbCondition_PlayerSwitch.ListIndex = 0
    
    cmbCondtion_PlayerSwitchCondition.Enabled = False
    cmbCondtion_PlayerSwitchCondition.ListIndex = 0
    
    cmbCondition_HasItem.Enabled = False
    cmbCondition_HasItem.Clear
    For i = 1 To MAX_ITEMS
        cmbCondition_HasItem.AddItem i & ". " & Trim$(Item(i).Name)
    Next
    cmbCondition_HasItem.ListIndex = 0
    
    cmbCondition_ClassIs.Enabled = False
    cmbCondition_ClassIs.Clear
    For i = 1 To Max_Classes
        cmbCondition_ClassIs.AddItem i & ". " & CStr(Class(i).Name)
    Next
    cmbCondition_ClassIs.ListIndex = 0
    
    cmbCondition_LearntSkill.Enabled = False
    cmbCondition_LearntSkill.Clear
    For i = 1 To MAX_SPELLS
        cmbCondition_LearntSkill.AddItem i & ". " & Trim$(Spell(i).Name)
    Next
    cmbCondition_LearntSkill.ListIndex = 0
    cmbCondition_LevelCompare.Enabled = False
    cmbCondition_LevelCompare.ListIndex = 0
    txtCondition_LevelAmount.Enabled = False
    txtCondition_LevelAmount.text = "0"
    
    cmbCondition_SelfSwitch.ListIndex = 0
    cmbCondition_SelfSwitch.Enabled = False
    cmbCondition_SelfSwitchCondition.ListIndex = 0
    cmbCondition_SelfSwitchCondition.Enabled = False
End Sub

Private Sub optPlayAnimEvent_Click()
    lblPlayAnimX.Visible = False
    lblPlayAnimY.Visible = False
    scrlPlayAnimTileX.Visible = False
    scrlPlayAnimTileY.Visible = False
    cmbPlayAnimEvent.Visible = True
End Sub

Private Sub optPlayAnimPlayer_Click()
    lblPlayAnimX.Visible = False
    lblPlayAnimY.Visible = False
    scrlPlayAnimTileX.Visible = False
    scrlPlayAnimTileY.Visible = False
    cmbPlayAnimEvent.Visible = False
End Sub

Private Sub optPlayAnimTile_Click()
    lblPlayAnimX.Visible = True
    lblPlayAnimY.Visible = True
    scrlPlayAnimTileX.Visible = True
    scrlPlayAnimTileY.Visible = True
    cmbPlayAnimEvent.Visible = False
End Sub

Private Sub picGraphic_Click()
    fraGraphic.width = 841
    fraGraphic.height = 585
    fraGraphic.Visible = True
    GraphicSelType = 0
End Sub

Private Sub picGraphicSel_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim i As Long
    If frmEditor_Events.cmbGraphic.ListIndex = 2 Then
        'Tileset... hard one....
        If ShiftDown Then
            If GraphicSelX > -1 And GraphicSelY > -1 Then
                If CLng(x + frmEditor_Events.hScrlGraphicSel.Value) / 32 > GraphicSelX And CLng(y + frmEditor_Events.vScrlGraphicSel.Value) / 32 > GraphicSelY Then
                    GraphicSelX2 = CLng(x + frmEditor_Events.hScrlGraphicSel.Value) / 32
                    GraphicSelY2 = CLng(y + frmEditor_Events.vScrlGraphicSel.Value) / 32
                End If
            End If
        Else
            GraphicSelX = CLng(x + frmEditor_Events.hScrlGraphicSel.Value) \ 32
            GraphicSelY = CLng(y + frmEditor_Events.vScrlGraphicSel.Value) \ 32
            GraphicSelX2 = 0
            GraphicSelY2 = 0
        End If
    ElseIf frmEditor_Events.cmbGraphic.ListIndex = 1 Then
        GraphicSelX = CLng(x + frmEditor_Events.hScrlGraphicSel.Value)
        GraphicSelY = CLng(y + frmEditor_Events.vScrlGraphicSel.Value)
        GraphicSelX2 = 0
        GraphicSelY2 = 0
        
        If frmEditor_Events.scrlGraphic.Value <= 0 Or frmEditor_Events.scrlGraphic.Value > NumCharacters Then Exit Sub
        
        If DDS_Character(frmEditor_Events.scrlGraphic.Value) Is Nothing Then
            Call InitDDSurf("Characters\" & frmEditor_Events.scrlGraphic.Value, DDSD_Character(frmEditor_Events.scrlGraphic.Value), DDS_Character(frmEditor_Events.scrlGraphic.Value))
        End If
        
        For i = 0 To 3
            If GraphicSelX >= ((DDSD_Character(frmEditor_Events.scrlGraphic.Value).lWidth / 4) * i) And GraphicSelX < ((DDSD_Character(frmEditor_Events.scrlGraphic.Value).lWidth / 4) * (i + 1)) Then
                GraphicSelX = i
            End If
        Next
        
        For i = 0 To 3
            If GraphicSelY >= ((DDSD_Character(frmEditor_Events.scrlGraphic.Value).lHeight / 4) * i) And GraphicSelY < ((DDSD_Character(frmEditor_Events.scrlGraphic.Value).lHeight / 4) * (i + 1)) Then
                GraphicSelY = i
            End If
        Next
        
        
    End If
End Sub

Private Sub scrlGraphic_Click()
    lblGraphic.Caption = "Graphic: " & scrlGraphic.Value
    tmpEvent.Pages(curPageNum).Graphic = scrlGraphic.Value
    
    If tmpEvent.Pages(curPageNum).GraphicType = 1 Then
        If frmEditor_Events.scrlGraphic.Value <= 0 Or frmEditor_Events.scrlGraphic.Value > NumCharacters Then Exit Sub
        CharacterTimer(frmEditor_Events.scrlGraphic.Value) = GetTickCount + SurfaceTimerMax
        If DDS_Character(frmEditor_Events.scrlGraphic.Value) Is Nothing Then
            Call InitDDSurf("Characters\" & frmEditor_Events.scrlGraphic.Value, DDSD_Character(frmEditor_Events.scrlGraphic.Value), DDS_Character(frmEditor_Events.scrlGraphic.Value))
        End If
                    
        If DDSD_Character(frmEditor_Events.scrlGraphic.Value).lWidth > 800 Then
            frmEditor_Events.hScrlGraphicSel.Visible = True
            frmEditor_Events.hScrlGraphicSel.Value = 0
            frmEditor_Events.hScrlGraphicSel.Max = DDSD_Character(frmEditor_Events.scrlGraphic.Value).lWidth - 800
        Else
            frmEditor_Events.hScrlGraphicSel.Visible = False
        End If
                    
        If DDSD_Character(frmEditor_Events.scrlGraphic.Value).lHeight > 512 Then
            frmEditor_Events.vScrlGraphicSel.Visible = True
            frmEditor_Events.vScrlGraphicSel.Value = 0
            frmEditor_Events.vScrlGraphicSel.Max = DDSD_Character(frmEditor_Events.scrlGraphic.Value).lHeight - 512
        Else
            frmEditor_Events.vScrlGraphicSel.Visible = False
        End If
    ElseIf tmpEvent.Pages(curPageNum).GraphicType = 2 Then
        If frmEditor_Events.scrlGraphic.Value <= 0 Or frmEditor_Events.scrlGraphic.Value > NumTileSets Then Exit Sub
        If DDS_Tileset(frmEditor_Events.scrlGraphic.Value) Is Nothing Then
            Call InitDDSurf("tilesets\" & frmEditor_Events.scrlGraphic.Value, DDSD_Tileset(frmEditor_Events.scrlGraphic.Value), DDS_Tileset(frmEditor_Events.scrlGraphic.Value))
        End If
                    
        If DDSD_Tileset(frmEditor_Events.scrlGraphic.Value).lWidth > 800 Then
            frmEditor_Events.hScrlGraphicSel.Visible = True
            frmEditor_Events.hScrlGraphicSel.Value = 0
            frmEditor_Events.hScrlGraphicSel.Max = DDSD_Tileset(frmEditor_Events.scrlGraphic.Value).lWidth - 800
        Else
            frmEditor_Events.hScrlGraphicSel.Visible = False
        End If
                    
        If DDSD_Tileset(frmEditor_Events.scrlGraphic.Value).lHeight > 512 Then
            frmEditor_Events.vScrlGraphicSel.Visible = True
            frmEditor_Events.vScrlGraphicSel.Value = 0
            frmEditor_Events.vScrlGraphicSel.Max = DDSD_Tileset(frmEditor_Events.scrlGraphic.Value).lHeight - 512
        Else
            frmEditor_Events.vScrlGraphicSel.Visible = False
        End If
    End If
End Sub

Private Sub scrlAddText_Colour_Click()
    frmEditor_Events.lblAddText_Colour.Caption = "Color: " & GetColorString(frmEditor_Events.scrlAddText_Colour.Value)
End Sub

Private Sub scrlAddText_Colour_Change()
    frmEditor_Events.lblAddText_Colour.Caption = "Color: " & GetColorString(frmEditor_Events.scrlAddText_Colour.Value)
End Sub

Private Sub scrlChangeLevel_Change()
    lblChangeLevel.Caption = "Level: " & scrlChangeLevel.Value
End Sub

Private Sub scrlChangeSprite_Change()
    lblChangeSprite.Caption = "Sprite: " & scrlChangeSprite.Value
End Sub

Private Sub scrlCustomScript_Change()
    lblCustomScript.Caption = "Case: " & scrlCustomScript.Value
End Sub

Private Sub scrlGiveExp_Click()
    lblGiveExp.Caption = "Give Exp: " & scrlGiveExp.Value
End Sub

Private Sub scrlGiveExp_Change()
    lblGiveExp.Caption = "Give Exp: " & scrlGiveExp.Value
End Sub

Private Sub scrlGraphic_Change()
    If scrlGraphic.Value = 0 Then
        lblGraphic.Caption = "Number: None"
    Else
        lblGraphic.Caption = "Number: " & scrlGraphic.Value
    End If
    Call cmbGraphic_Click
End Sub

Private Sub scrlPlayAnimTileX_Change()
    lblPlayAnimX.Caption = "Map Tile X: " & scrlPlayAnimTileX.Value
End Sub

Private Sub scrlPlayAnimTileY_Change()
    lblPlayAnimY.Caption = "Map Tile Y: " & scrlPlayAnimTileY.Value
End Sub

Private Sub scrlWPMap_Change()
    lblWPMap.Caption = "Map: " & scrlWPMap.Value
End Sub

Private Sub scrlWPX_Change()
    lblWPX.Caption = "X: " & scrlWPX.Value
End Sub

Private Sub scrlWPY_Change()
    lblWPY.Caption = "Y: " & scrlWPY.Value
End Sub

Private Sub tabCommands_Click()
Dim i As Long
    For i = 1 To 2
        picCommands(i).Visible = False
    Next
    picCommands(tabCommands.SelectedItem.Index).Visible = True
End Sub

Private Sub tabPages_Click()
    If tabPages.SelectedItem.Index <> curPageNum Then
        curPageNum = tabPages.SelectedItem.Index
        EventEditorLoadPage curPageNum
    End If
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
    tmpEvent.Name = Trim$(txtName.text)
End Sub

Private Sub txtPlayerVariable_Validate(Cancel As Boolean)
    tmpEvent.Pages(curPageNum).VariableCondition = Val(Trim$(txtPlayerVariable.text))
End Sub


