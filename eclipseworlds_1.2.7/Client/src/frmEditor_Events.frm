VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmEditor_Events 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Event Editor"
   ClientHeight    =   8970
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12885
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEditor_Events.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   598
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   859
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2880
      TabIndex        =   435
      Top             =   8400
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Frame fraGraphic 
      Caption         =   "Graphic Selection"
      Height          =   615
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
         TabIndex        =   102
         Top             =   7920
         Visible         =   0   'False
         Width           =   11895
      End
      Begin VB.VScrollBar vScrlGraphicSel 
         Height          =   7095
         LargeChange     =   64
         Left            =   12240
         SmallChange     =   32
         TabIndex        =   101
         Top             =   720
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox picGraphicSel 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   7080
         Left            =   240
         ScaleHeight     =   472
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   793
         TabIndex        =   81
         Top             =   720
         Width           =   11895
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
         ItemData        =   "frmEditor_Events.frx":038A
         Left            =   720
         List            =   "frmEditor_Events.frx":0397
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
      TabIndex        =   319
      Top             =   120
      Visible         =   0   'False
      Width           =   495
      Begin VB.Frame fraRenaming 
         Caption         =   "Renaming Variable/Switch"
         Height          =   8535
         Left            =   120
         TabIndex        =   328
         Top             =   120
         Visible         =   0   'False
         Width           =   12615
         Begin VB.Frame fraRandom 
            Caption         =   "Editing Variable/Switch"
            Height          =   2295
            Index           =   10
            Left            =   3600
            TabIndex        =   329
            Top             =   2520
            Width           =   5055
            Begin VB.TextBox txtRename 
               Height          =   375
               Left            =   120
               TabIndex        =   332
               Top             =   720
               Width           =   4815
            End
            Begin VB.CommandButton cmdRename_Cancel 
               Caption         =   "Cancel"
               Height          =   375
               Left            =   3720
               TabIndex        =   331
               Top             =   1800
               Width           =   1215
            End
            Begin VB.CommandButton cmdRename_Ok 
               Caption         =   "Ok"
               Height          =   375
               Left            =   2280
               TabIndex        =   330
               Top             =   1800
               Width           =   1215
            End
            Begin VB.Label lblEditing 
               Caption         =   "Naming Variable #1"
               Height          =   375
               Left            =   120
               TabIndex        =   333
               Top             =   360
               Width           =   4815
            End
         End
      End
      Begin VB.CommandButton cmdRenameSwitch 
         Caption         =   "Rename Switch"
         Height          =   375
         Left            =   8280
         TabIndex        =   327
         Top             =   7320
         Width           =   1935
      End
      Begin VB.CommandButton cmdRenameVariable 
         Caption         =   "Rename Variable"
         Height          =   375
         Left            =   360
         TabIndex        =   326
         Top             =   7320
         Width           =   1935
      End
      Begin VB.ListBox lstSwitches 
         Height          =   6495
         Left            =   8280
         TabIndex        =   324
         Top             =   720
         Width           =   3855
      End
      Begin VB.ListBox lstVariables 
         Height          =   6495
         Left            =   360
         TabIndex        =   322
         Top             =   720
         Width           =   3855
      End
      Begin VB.CommandButton cmbLabel_Ok 
         Caption         =   "OK"
         Height          =   375
         Left            =   9480
         TabIndex        =   321
         Top             =   8400
         Width           =   1455
      End
      Begin VB.CommandButton cmdLabel_Cancel 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   11040
         TabIndex        =   320
         Top             =   8400
         Width           =   1455
      End
      Begin VB.Label lblRandomLabel 
         Caption         =   "Player Switches"
         Height          =   255
         Index           =   36
         Left            =   8280
         TabIndex        =   325
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label lblRandomLabel 
         Caption         =   "Player Variables"
         Height          =   255
         Index           =   25
         Left            =   360
         TabIndex        =   323
         Top             =   480
         Width           =   2175
      End
   End
   Begin VB.Frame fraMoveRoute 
      Caption         =   "Move Route"
      Height          =   375
      Left            =   120
      TabIndex        =   103
      Top             =   120
      Visible         =   0   'False
      Width           =   855
      Begin VB.Frame fraRandom 
         Caption         =   "Commands"
         Height          =   6615
         Index           =   14
         Left            =   3120
         TabIndex        =   110
         Top             =   480
         Width           =   9255
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Set Graphic..."
            Height          =   375
            Index           =   42
            Left            =   6720
            TabIndex        =   153
            Top             =   3240
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Set Position Above Players"
            Height          =   375
            Index           =   41
            Left            =   6720
            TabIndex        =   152
            Top             =   2760
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Set Position with Players"
            Height          =   375
            Index           =   40
            Left            =   6720
            TabIndex        =   151
            Top             =   2280
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Set Position Below Players"
            Height          =   375
            Index           =   39
            Left            =   6720
            TabIndex        =   150
            Top             =   1800
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Turn Off Walk-Through"
            Height          =   375
            Index           =   38
            Left            =   6720
            TabIndex        =   149
            Top             =   1320
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Turn On Walk-Through"
            Height          =   375
            Index           =   37
            Left            =   6720
            TabIndex        =   148
            Top             =   840
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Turn Fixed Dir Off"
            Height          =   375
            Index           =   36
            Left            =   6720
            TabIndex        =   147
            Top             =   360
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Turn Fixed Dir On"
            Height          =   375
            Index           =   35
            Left            =   4560
            TabIndex        =   146
            Top             =   5640
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Turn Walking Anim Off"
            Height          =   375
            Index           =   34
            Left            =   4560
            TabIndex        =   145
            Top             =   5160
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Turn Walking Anim On"
            Height          =   375
            Index           =   33
            Left            =   4560
            TabIndex        =   144
            Top             =   4680
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Set Freq. To Highest"
            Height          =   375
            Index           =   32
            Left            =   4560
            TabIndex        =   143
            Top             =   4200
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Set Freq. To Higher"
            Height          =   375
            Index           =   31
            Left            =   4560
            TabIndex        =   142
            Top             =   3720
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Set Freq. To Normal"
            Height          =   375
            Index           =   30
            Left            =   4560
            TabIndex        =   141
            Top             =   3240
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Set Freq. To Lower"
            Height          =   375
            Index           =   29
            Left            =   4560
            TabIndex        =   140
            Top             =   2760
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Set Freq. To Lowest"
            Height          =   375
            Index           =   28
            Left            =   4560
            TabIndex        =   139
            Top             =   2280
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Set Speed 4x Faster"
            Height          =   375
            Index           =   27
            Left            =   4560
            TabIndex        =   138
            Top             =   1800
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Set Speed 2x Faster"
            Height          =   375
            Index           =   26
            Left            =   4560
            TabIndex        =   137
            Top             =   1320
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Set Speed to Normal"
            Height          =   375
            Index           =   25
            Left            =   4560
            TabIndex        =   136
            Top             =   840
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Set Speed 2x Slower"
            Height          =   375
            Index           =   24
            Left            =   4560
            TabIndex        =   135
            Top             =   360
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Set Speed 4x Slower"
            Height          =   375
            Index           =   23
            Left            =   2400
            TabIndex        =   134
            Top             =   5640
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Set Speed 8x Slower"
            Height          =   375
            Index           =   22
            Left            =   2400
            TabIndex        =   133
            Top             =   5160
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Turn Away From Player***"
            Height          =   375
            Index           =   21
            Left            =   2400
            TabIndex        =   132
            Top             =   4680
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Turn Toward Player***"
            Height          =   375
            Index           =   20
            Left            =   2400
            TabIndex        =   131
            Top             =   4200
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Turn Randomly"
            Height          =   375
            Index           =   19
            Left            =   2400
            TabIndex        =   130
            Top             =   3720
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Turn 180 Degrees"
            Height          =   375
            Index           =   18
            Left            =   2400
            TabIndex        =   129
            Top             =   3240
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Turn 90 Degrees to the Left"
            Height          =   375
            Index           =   17
            Left            =   2400
            TabIndex        =   128
            Top             =   2760
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Turn 90 Degrees to the Right"
            Height          =   375
            Index           =   16
            Left            =   2400
            TabIndex        =   127
            Top             =   2280
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Turn Right"
            Height          =   375
            Index           =   15
            Left            =   2400
            TabIndex        =   126
            Top             =   1800
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Turn Left"
            Height          =   375
            Index           =   14
            Left            =   2400
            TabIndex        =   125
            Top             =   1320
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Turn Down"
            Height          =   375
            Index           =   13
            Left            =   2400
            TabIndex        =   124
            Top             =   840
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Turn Up"
            Height          =   375
            Index           =   12
            Left            =   2400
            TabIndex        =   123
            Top             =   360
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Wait 1000Ms"
            Height          =   375
            Index           =   11
            Left            =   240
            TabIndex        =   122
            Top             =   5640
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Wait 500Ms"
            Height          =   375
            Index           =   10
            Left            =   240
            TabIndex        =   121
            Top             =   5160
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Wait 100Ms"
            Height          =   375
            Index           =   9
            Left            =   240
            TabIndex        =   120
            Top             =   4680
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Step Back"
            Height          =   375
            Index           =   8
            Left            =   240
            TabIndex        =   119
            Top             =   4200
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Step Forward"
            Height          =   375
            Index           =   7
            Left            =   240
            TabIndex        =   118
            Top             =   3720
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Move Away From Player***"
            Height          =   375
            Index           =   6
            Left            =   240
            TabIndex        =   117
            Top             =   3240
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Move Towards Player***"
            Height          =   375
            Index           =   5
            Left            =   240
            TabIndex        =   116
            Top             =   2760
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Move Randomly"
            Height          =   375
            Index           =   4
            Left            =   240
            TabIndex        =   115
            Top             =   2280
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Move Right"
            Height          =   375
            Index           =   3
            Left            =   240
            TabIndex        =   114
            Top             =   1800
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Move Left"
            Height          =   375
            Index           =   2
            Left            =   240
            TabIndex        =   113
            Top             =   1320
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Move Down"
            Height          =   375
            Index           =   1
            Left            =   240
            TabIndex        =   112
            Top             =   840
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Move Up"
            Height          =   375
            Index           =   0
            Left            =   240
            TabIndex        =   111
            Top             =   360
            Width           =   1935
         End
         Begin VB.Label lblRandomLabel 
            Caption         =   "*** These commands will not process on global events."
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   15
            Left            =   240
            TabIndex        =   154
            Top             =   6240
            Width           =   8535
         End
      End
      Begin VB.ComboBox cmbEvent 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmEditor_Events.frx":03B5
         Left            =   120
         List            =   "frmEditor_Events.frx":03B7
         Style           =   2  'Dropdown List
         TabIndex        =   109
         Top             =   480
         Width           =   2655
      End
      Begin VB.CheckBox chkRepeatRoute 
         Caption         =   "Repeat Route"
         Height          =   255
         Left            =   120
         TabIndex        =   108
         Top             =   7560
         Width           =   2655
      End
      Begin VB.CheckBox chkIgnoreMove 
         Caption         =   "Ignore if event can't move."
         Height          =   255
         Left            =   120
         TabIndex        =   107
         Top             =   7200
         Width           =   2655
      End
      Begin VB.ListBox lstMoveRoute 
         Height          =   6105
         Left            =   120
         TabIndex        =   106
         Top             =   960
         Width           =   2655
      End
      Begin VB.CommandButton cmdMoveRouteOk 
         Caption         =   "OK"
         Height          =   375
         Left            =   9480
         TabIndex        =   105
         Top             =   8160
         Width           =   1455
      End
      Begin VB.CommandButton cmdMoveRouteCancel 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   11040
         TabIndex        =   104
         Top             =   8160
         Width           =   1455
      End
   End
   Begin VB.Frame fraRandom 
      Caption         =   "Positioning"
      Height          =   855
      Index           =   19
      Left            =   2760
      TabIndex        =   99
      Top             =   5640
      Width           =   3375
      Begin VB.ComboBox cmbPositioning 
         Height          =   315
         ItemData        =   "frmEditor_Events.frx":03B9
         Left            =   120
         List            =   "frmEditor_Events.frx":03C6
         Style           =   2  'Dropdown List
         TabIndex        =   100
         Top             =   360
         Width           =   3135
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   9600
      TabIndex        =   36
      Top             =   8400
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   11160
      TabIndex        =   35
      Top             =   8400
      Width           =   1455
   End
   Begin VB.Frame fraRandom 
      Caption         =   "General"
      Height          =   735
      Index           =   20
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
   Begin VB.Frame fraRandom 
      Caption         =   "Trigger"
      Height          =   855
      Index           =   18
      Left            =   2760
      TabIndex        =   24
      Top             =   7440
      Width           =   3375
      Begin VB.ComboBox cmbTrigger 
         Height          =   315
         ItemData        =   "frmEditor_Events.frx":0402
         Left            =   120
         List            =   "frmEditor_Events.frx":040F
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   360
         Width           =   3135
      End
   End
   Begin VB.Frame fraRandom 
      Caption         =   "Options"
      Height          =   1575
      Index           =   16
      Left            =   360
      TabIndex        =   20
      Top             =   6720
      Width           =   2295
      Begin VB.CheckBox chkGlobal 
         Caption         =   "Global"
         Height          =   255
         Left            =   120
         TabIndex        =   429
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CheckBox chkShowName 
         Caption         =   "Show Name"
         Height          =   255
         Left            =   120
         TabIndex        =   335
         Top             =   960
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.CheckBox chkWalkThrough 
         Caption         =   "Through"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   720
         Width           =   1095
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
   Begin VB.Frame fraRandom 
      Caption         =   "Movement"
      Height          =   2175
      Index           =   15
      Left            =   2760
      TabIndex        =   13
      Top             =   3480
      Width           =   3375
      Begin VB.CommandButton cmdMoveRoute 
         Caption         =   "Move Route"
         Enabled         =   0   'False
         Height          =   375
         Left            =   840
         TabIndex        =   98
         Top             =   720
         Width           =   2415
      End
      Begin VB.ComboBox cmbMoveFreq 
         Height          =   315
         ItemData        =   "frmEditor_Events.frx":0454
         Left            =   840
         List            =   "frmEditor_Events.frx":0467
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   1680
         Width           =   2415
      End
      Begin VB.ComboBox cmbMoveSpeed 
         Height          =   315
         ItemData        =   "frmEditor_Events.frx":0493
         Left            =   840
         List            =   "frmEditor_Events.frx":04A9
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   1200
         Width           =   2415
      End
      Begin VB.ComboBox cmbMoveType 
         Height          =   315
         ItemData        =   "frmEditor_Events.frx":04EC
         Left            =   840
         List            =   "frmEditor_Events.frx":04F9
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
   Begin VB.Frame fraRandom 
      Caption         =   "Graphic"
      Height          =   3255
      Index           =   13
      Left            =   360
      TabIndex        =   11
      Top             =   3480
      Width           =   2295
      Begin VB.PictureBox picGraphic 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000001&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2775
         Left            =   240
         ScaleHeight     =   185
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   121
         TabIndex        =   12
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame fraRandom 
      Caption         =   "Conditions"
      Height          =   2175
      Index           =   0
      Left            =   360
      TabIndex        =   2
      Top             =   1320
      Width           =   5775
      Begin VB.ComboBox cmbPlayerVarCompare 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmEditor_Events.frx":0521
         Left            =   3720
         List            =   "frmEditor_Events.frx":0537
         Style           =   2  'Dropdown List
         TabIndex        =   305
         Top             =   240
         Width           =   1335
      End
      Begin VB.ComboBox cmbSelfSwitchCompare 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmEditor_Events.frx":059D
         Left            =   3720
         List            =   "frmEditor_Events.frx":05A7
         Style           =   2  'Dropdown List
         TabIndex        =   304
         Top             =   1680
         Width           =   1095
      End
      Begin VB.ComboBox cmbPlayerSwitchCompare 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmEditor_Events.frx":05B8
         Left            =   3720
         List            =   "frmEditor_Events.frx":05C2
         Style           =   2  'Dropdown List
         TabIndex        =   301
         Top             =   720
         Width           =   1095
      End
      Begin VB.ComboBox cmbSelfSwitch 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmEditor_Events.frx":05D3
         Left            =   1920
         List            =   "frmEditor_Events.frx":05E6
         Style           =   2  'Dropdown List
         TabIndex        =   97
         Top             =   1680
         Width           =   1455
      End
      Begin VB.CheckBox chkSelfSwitch 
         Caption         =   "Self Switch"
         Height          =   255
         Left            =   120
         TabIndex        =   96
         Top             =   1680
         Width           =   1695
      End
      Begin VB.CheckBox chkHasItem 
         Caption         =   "Has Item           (In Inventory)"
         Height          =   495
         Left            =   120
         TabIndex        =   10
         Top             =   1080
         Width           =   1695
      End
      Begin VB.ComboBox cmbHasItem 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmEditor_Events.frx":060C
         Left            =   1920
         List            =   "frmEditor_Events.frx":060E
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
         ItemData        =   "frmEditor_Events.frx":0610
         Left            =   1920
         List            =   "frmEditor_Events.frx":0612
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
         ItemData        =   "frmEditor_Events.frx":0614
         Left            =   1920
         List            =   "frmEditor_Events.frx":0616
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
         TabIndex        =   303
         Top             =   1760
         Width           =   255
      End
      Begin VB.Label lblRandomLabel 
         Caption         =   "is"
         Height          =   255
         Index           =   3
         Left            =   3480
         TabIndex        =   302
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
   Begin VB.CommandButton cmdLabel 
      Caption         =   "Label Variables/Switches"
      Height          =   375
      Left            =   360
      TabIndex        =   318
      Top             =   8400
      Width           =   2415
   End
   Begin VB.Frame fraEvents 
      Caption         =   "Events List"
      Height          =   4815
      Left            =   360
      TabIndex        =   430
      Top             =   3480
      Visible         =   0   'False
      Width           =   5775
      Begin VB.TextBox txtSearch 
         CausesValidation=   0   'False
         Height          =   270
         Left            =   120
         TabIndex        =   434
         Top             =   240
         Width           =   4095
      End
      Begin VB.CommandButton cmdCopy 
         Caption         =   "Copy"
         Height          =   315
         Left            =   4320
         TabIndex        =   433
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdPaste 
         Caption         =   "Paste"
         Height          =   315
         Left            =   5040
         TabIndex        =   432
         Top             =   240
         Width           =   615
      End
      Begin VB.ListBox lstIndex 
         Height          =   3960
         Left            =   120
         TabIndex        =   431
         Top             =   600
         Width           =   5535
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
         Left            =   4440
         TabIndex        =   72
         Top             =   6360
         Width           =   1455
      End
      Begin VB.PictureBox picCommands 
         BorderStyle     =   0  'None
         Height          =   6135
         Index           =   1
         Left            =   240
         ScaleHeight     =   6135
         ScaleWidth      =   5775
         TabIndex        =   39
         Top             =   600
         Width           =   5775
         Begin VB.Frame fraRandom 
            Caption         =   "Message"
            Height          =   2175
            Index           =   21
            Left            =   0
            TabIndex        =   41
            Top             =   0
            Width           =   2775
            Begin VB.CommandButton cmdCommands 
               Caption         =   "Show Chat Bubble"
               Height          =   375
               Index           =   3
               Left            =   120
               TabIndex        =   345
               Top             =   1680
               Width           =   2535
            End
            Begin VB.CommandButton cmdCommands 
               Caption         =   "Add Chatbox Text"
               Height          =   375
               Index           =   2
               Left            =   120
               TabIndex        =   42
               Top             =   1200
               Width           =   2535
            End
            Begin VB.CommandButton cmdCommands 
               Caption         =   "Show Text"
               Height          =   375
               Index           =   0
               Left            =   120
               TabIndex        =   43
               Top             =   240
               Width           =   2535
            End
            Begin VB.CommandButton cmdCommands 
               Caption         =   "Show Choices"
               Height          =   375
               Index           =   1
               Left            =   120
               TabIndex        =   44
               Top             =   720
               Width           =   2535
            End
         End
         Begin VB.Frame fraRandom 
            Caption         =   "Flow Control"
            Height          =   2175
            Index           =   2
            Left            =   0
            TabIndex        =   49
            Top             =   3840
            Width           =   2775
            Begin VB.CommandButton cmdCommands 
               Caption         =   "Go To Label"
               Height          =   375
               Index           =   10
               Left            =   120
               TabIndex        =   350
               Top             =   1680
               Width           =   2535
            End
            Begin VB.CommandButton cmdCommands 
               Caption         =   "Label"
               Height          =   375
               Index           =   9
               Left            =   120
               TabIndex        =   349
               Top             =   1200
               Width           =   2535
            End
            Begin VB.CommandButton cmdCommands 
               Caption         =   "Conditional Branch"
               Height          =   375
               Index           =   7
               Left            =   120
               TabIndex        =   51
               Top             =   240
               Width           =   2535
            End
            Begin VB.CommandButton cmdCommands 
               Caption         =   "Exit Event Process"
               Height          =   375
               Index           =   8
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
            Top             =   2160
            Width           =   2775
            Begin VB.CommandButton cmdCommands 
               Caption         =   "Player Switch"
               Height          =   375
               Index           =   5
               Left            =   120
               TabIndex        =   47
               Top             =   720
               Width           =   2535
            End
            Begin VB.CommandButton cmdCommands 
               Caption         =   "Player Variable"
               Height          =   375
               Index           =   4
               Left            =   120
               TabIndex        =   46
               Top             =   240
               Width           =   2535
            End
            Begin VB.CommandButton cmdCommands 
               Caption         =   "Self Switch"
               Height          =   375
               Index           =   6
               Left            =   120
               TabIndex        =   48
               Top             =   1200
               Width           =   2535
            End
         End
         Begin VB.Frame fraRandom 
            Caption         =   "Player Control"
            Height          =   5535
            Index           =   3
            Left            =   3000
            TabIndex        =   52
            Top             =   0
            Width           =   2775
            Begin VB.CommandButton cmdCommands 
               Caption         =   "Change Class"
               Height          =   375
               Index           =   17
               Left            =   120
               TabIndex        =   59
               Top             =   3120
               Width           =   2535
            End
            Begin VB.CommandButton cmdCommands 
               Caption         =   "Give EXP"
               Height          =   375
               Index           =   21
               Left            =   120
               TabIndex        =   334
               Top             =   5040
               Width           =   2535
            End
            Begin VB.CommandButton cmdCommands 
               Caption         =   "Change PK"
               Height          =   375
               Index           =   20
               Left            =   120
               TabIndex        =   214
               Top             =   4560
               Width           =   2535
            End
            Begin VB.CommandButton cmdCommands 
               Caption         =   "Change Sex"
               Height          =   375
               Index           =   19
               Left            =   120
               TabIndex        =   61
               Top             =   4080
               Width           =   2535
            End
            Begin VB.CommandButton cmdCommands 
               Caption         =   "Change Sprite"
               Height          =   375
               Index           =   18
               Left            =   120
               TabIndex        =   60
               Top             =   3600
               Width           =   2535
            End
            Begin VB.CommandButton cmdCommands 
               Caption         =   "Change Level"
               Height          =   375
               Index           =   15
               Left            =   120
               TabIndex        =   57
               Top             =   2160
               Width           =   2535
            End
            Begin VB.CommandButton cmdCommands 
               Caption         =   "Level Up"
               Height          =   375
               Index           =   14
               Left            =   120
               TabIndex        =   56
               Top             =   1680
               Width           =   2535
            End
            Begin VB.CommandButton cmdCommands 
               Caption         =   "Restore MP"
               Height          =   375
               Index           =   13
               Left            =   120
               TabIndex        =   55
               Top             =   1200
               Width           =   2535
            End
            Begin VB.CommandButton cmdCommands 
               Caption         =   "Restore HP"
               Height          =   375
               Index           =   12
               Left            =   120
               TabIndex        =   54
               Top             =   720
               Width           =   2535
            End
            Begin VB.CommandButton cmdCommands 
               Caption         =   "Change Items"
               Height          =   375
               Index           =   11
               Left            =   120
               TabIndex        =   53
               Top             =   240
               Width           =   2535
            End
            Begin VB.CommandButton cmdCommands 
               Caption         =   "Change Spells"
               Height          =   375
               Index           =   16
               Left            =   120
               TabIndex        =   58
               Top             =   2640
               Width           =   2535
            End
         End
      End
      Begin VB.PictureBox picCommands 
         BorderStyle     =   0  'None
         Height          =   6015
         Index           =   2
         Left            =   240
         ScaleHeight     =   6015
         ScaleWidth      =   5775
         TabIndex        =   40
         Top             =   600
         Visible         =   0   'False
         Width           =   5775
         Begin VB.Frame fraRandom 
            Caption         =   "Map Functions"
            Height          =   1695
            Index           =   12
            Left            =   3000
            TabIndex        =   354
            Top             =   0
            Width           =   2775
            Begin VB.CommandButton cmdCommands 
               Caption         =   "Set Fog"
               Height          =   375
               Index           =   31
               Left            =   120
               TabIndex        =   357
               Top             =   240
               Width           =   2535
            End
            Begin VB.CommandButton cmdCommands 
               Caption         =   "Set Weather"
               Height          =   375
               Index           =   32
               Left            =   120
               TabIndex        =   356
               Top             =   720
               Width           =   2535
            End
            Begin VB.CommandButton cmdCommands 
               Caption         =   "Set Map Tinting"
               Height          =   375
               Index           =   33
               Left            =   120
               TabIndex        =   355
               Top             =   1200
               Width           =   2535
            End
         End
         Begin VB.Frame fraRandom 
            Caption         =   "Cut-Scene Options"
            Height          =   1695
            Index           =   11
            Left            =   0
            TabIndex        =   347
            Top             =   3840
            Width           =   2775
            Begin VB.CommandButton cmdCommands 
               Caption         =   "Flash White"
               Height          =   375
               Index           =   30
               Left            =   120
               TabIndex        =   353
               Top             =   1200
               Width           =   2535
            End
            Begin VB.CommandButton cmdCommands 
               Caption         =   "Fade Out"
               Height          =   375
               Index           =   29
               Left            =   120
               TabIndex        =   352
               Top             =   720
               Width           =   2535
            End
            Begin VB.CommandButton cmdCommands 
               Caption         =   "Fade In"
               Height          =   375
               Index           =   28
               Left            =   120
               TabIndex        =   348
               Top             =   240
               Width           =   2535
            End
         End
         Begin VB.Frame fraRandom 
            Caption         =   "Scenes"
            Height          =   1215
            Index           =   6
            Left            =   0
            TabIndex        =   306
            Top             =   2520
            Width           =   2775
            Begin VB.CommandButton cmdCommands 
               Caption         =   "Open Shop"
               Height          =   375
               Index           =   27
               Left            =   120
               TabIndex        =   308
               Top             =   720
               Width           =   2535
            End
            Begin VB.CommandButton cmdCommands 
               Caption         =   "Open Bank"
               Height          =   375
               Index           =   26
               Left            =   120
               TabIndex        =   307
               Top             =   240
               Width           =   2535
            End
         End
         Begin VB.Frame fraRandom 
            Caption         =   "Etc..."
            Height          =   1695
            Index           =   8
            Left            =   3000
            TabIndex        =   267
            Top             =   3840
            Width           =   2775
            Begin VB.CommandButton cmdCommands 
               Caption         =   "Wait..."
               Height          =   375
               Index           =   38
               Left            =   120
               TabIndex        =   346
               Top             =   240
               Width           =   2535
            End
            Begin VB.CommandButton cmdCommands 
               Caption         =   "Set Access"
               Height          =   375
               Index           =   39
               Left            =   120
               TabIndex        =   309
               Top             =   720
               Width           =   2535
            End
            Begin VB.CommandButton cmdCommands 
               Caption         =   "Custom Script"
               Height          =   375
               Index           =   40
               Left            =   120
               TabIndex        =   268
               Top             =   1200
               Width           =   2535
            End
         End
         Begin VB.Frame fraRandom 
            Caption         =   "Music and Sound"
            Height          =   2175
            Index           =   7
            Left            =   3000
            TabIndex        =   67
            Top             =   1680
            Width           =   2775
            Begin VB.CommandButton cmdCommands 
               Caption         =   "Stop Sounds"
               Height          =   375
               Index           =   37
               Left            =   120
               TabIndex        =   71
               Top             =   1680
               Width           =   2535
            End
            Begin VB.CommandButton cmdCommands 
               Caption         =   "Play Sound"
               Height          =   375
               Index           =   36
               Left            =   120
               TabIndex        =   70
               Top             =   1200
               Width           =   2535
            End
            Begin VB.CommandButton cmdCommands 
               Caption         =   "Fadeout BGM"
               Height          =   375
               Index           =   35
               Left            =   120
               TabIndex        =   69
               Top             =   720
               Width           =   2535
            End
            Begin VB.CommandButton cmdCommands 
               Caption         =   "Play BGM"
               Height          =   375
               Index           =   34
               Left            =   120
               TabIndex        =   68
               Top             =   240
               Width           =   2535
            End
         End
         Begin VB.Frame fraRandom 
            Caption         =   "Animation"
            Height          =   735
            Index           =   5
            Left            =   0
            TabIndex        =   65
            Top             =   1680
            Width           =   2775
            Begin VB.CommandButton cmdCommands 
               Caption         =   "Play Animation"
               Height          =   375
               Index           =   25
               Left            =   120
               TabIndex        =   66
               Top             =   240
               Width           =   2535
            End
         End
         Begin VB.Frame fraRandom 
            Caption         =   "Movement"
            Height          =   1695
            Index           =   4
            Left            =   0
            TabIndex        =   62
            Top             =   0
            Width           =   2775
            Begin VB.CommandButton cmdCommands 
               Caption         =   "Force Spawn NPC"
               Height          =   375
               Index           =   24
               Left            =   120
               TabIndex        =   351
               Top             =   1200
               Width           =   2535
            End
            Begin VB.CommandButton cmdCommands 
               Caption         =   "Set Move Route"
               Height          =   375
               Index           =   23
               Left            =   120
               TabIndex        =   64
               Top             =   720
               Width           =   2535
            End
            Begin VB.CommandButton cmdCommands 
               Caption         =   "Warp Player"
               Height          =   375
               Index           =   22
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
   Begin VB.Frame fraDialogue 
      Height          =   6975
      Left            =   6240
      TabIndex        =   73
      Top             =   1320
      Visible         =   0   'False
      Width           =   6375
      Begin VB.Frame fraCommand 
         Caption         =   "Play BGM"
         Height          =   1575
         Index           =   25
         Left            =   1080
         TabIndex        =   288
         Top             =   1800
         Visible         =   0   'False
         Width           =   4335
         Begin VB.ComboBox cmbPlayBGM 
            Height          =   315
            ItemData        =   "frmEditor_Events.frx":0618
            Left            =   120
            List            =   "frmEditor_Events.frx":061A
            Style           =   2  'Dropdown List
            TabIndex        =   291
            Top             =   360
            Width           =   3135
         End
         Begin VB.CommandButton cmdPlayBGM_Cancel 
            Caption         =   "Cancel"
            Height          =   375
            Left            =   3000
            TabIndex        =   290
            Top             =   1080
            Width           =   1215
         End
         Begin VB.CommandButton cmdPlayBGM_Ok 
            Caption         =   "Ok"
            Height          =   375
            Left            =   1560
            TabIndex        =   289
            Top             =   1080
            Width           =   1215
         End
      End
      Begin VB.Frame fraCommand 
         Caption         =   "Play Sound"
         Height          =   1575
         Index           =   26
         Left            =   1080
         TabIndex        =   292
         Top             =   1800
         Visible         =   0   'False
         Width           =   4215
         Begin VB.CommandButton cmdPlaySound_Ok 
            Caption         =   "Ok"
            Height          =   375
            Left            =   1440
            TabIndex        =   295
            Top             =   1080
            Width           =   1215
         End
         Begin VB.CommandButton cmdPlaySound_Cancel 
            Caption         =   "Cancel"
            Height          =   375
            Left            =   2880
            TabIndex        =   294
            Top             =   1080
            Width           =   1215
         End
         Begin VB.ComboBox cmbPlaySound 
            Height          =   315
            ItemData        =   "frmEditor_Events.frx":061C
            Left            =   120
            List            =   "frmEditor_Events.frx":061E
            Style           =   2  'Dropdown List
            TabIndex        =   293
            Top             =   360
            Width           =   3135
         End
      End
      Begin VB.Frame fraCommand 
         Caption         =   "Wait"
         Height          =   1455
         Index           =   27
         Left            =   1080
         TabIndex        =   417
         Top             =   1800
         Visible         =   0   'False
         Width           =   4335
         Begin VB.CommandButton cmdWait_Cancel 
            Caption         =   "Cancel"
            Height          =   375
            Left            =   3000
            TabIndex        =   420
            Top             =   840
            Width           =   1215
         End
         Begin VB.CommandButton cmdWait_Ok 
            Caption         =   "Ok"
            Height          =   375
            Left            =   1560
            TabIndex        =   419
            Top             =   840
            Width           =   1215
         End
         Begin VB.HScrollBar scrlWaitAmount 
            Height          =   255
            Left            =   120
            Min             =   1
            TabIndex        =   418
            Top             =   480
            Value           =   1
            Width           =   4095
         End
         Begin VB.Label lblRandomLabel 
            Caption         =   "Hint: 1000 ms = 1 Second"
            Height          =   255
            Index           =   44
            Left            =   1920
            TabIndex        =   422
            Top             =   240
            Width           =   2295
         End
         Begin VB.Label lblWaitAmount 
            Caption         =   "Wait: 0 ms"
            Height          =   255
            Left            =   120
            TabIndex        =   421
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.Frame fraCommand 
         Caption         =   "Set Access"
         Height          =   1575
         Index           =   28
         Left            =   1080
         TabIndex        =   310
         Top             =   1800
         Visible         =   0   'False
         Width           =   4215
         Begin VB.ComboBox cmbSetAccess 
            Height          =   315
            ItemData        =   "frmEditor_Events.frx":0620
            Left            =   120
            List            =   "frmEditor_Events.frx":0636
            Style           =   2  'Dropdown List
            TabIndex        =   313
            Top             =   360
            Width           =   3135
         End
         Begin VB.CommandButton cmdSetAccess_Cancel 
            Caption         =   "Cancel"
            Height          =   375
            Left            =   2880
            TabIndex        =   312
            Top             =   1080
            Width           =   1215
         End
         Begin VB.CommandButton cmdSetAccess_Ok 
            Caption         =   "Ok"
            Height          =   375
            Left            =   1440
            TabIndex        =   311
            Top             =   1080
            Width           =   1215
         End
      End
      Begin VB.Frame fraCommand 
         Caption         =   "Execute Custom Script"
         Height          =   1575
         Index           =   29
         Left            =   1080
         TabIndex        =   283
         Top             =   1800
         Visible         =   0   'False
         Width           =   4335
         Begin VB.HScrollBar scrlCustomScript 
            Height          =   255
            Left            =   1560
            Max             =   255
            Min             =   1
            TabIndex        =   287
            Top             =   360
            Value           =   1
            Width           =   2655
         End
         Begin VB.CommandButton cmdCustomScript_Ok 
            Caption         =   "Ok"
            Height          =   375
            Left            =   1560
            TabIndex        =   285
            Top             =   840
            Width           =   1215
         End
         Begin VB.CommandButton cmdCustomScript_Cancel 
            Caption         =   "Cancel"
            Height          =   375
            Left            =   3000
            TabIndex        =   284
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label lblCustomScript 
            Caption         =   "Case: 1"
            Height          =   255
            Left            =   120
            TabIndex        =   286
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Frame fraCommand 
         Caption         =   "Show Chat Bubble"
         Height          =   2775
         Index           =   3
         Left            =   720
         TabIndex        =   360
         Top             =   1320
         Visible         =   0   'False
         Width           =   5055
         Begin VB.TextBox txtChatbubbleText 
            Height          =   285
            Left            =   1680
            TabIndex        =   369
            Top             =   360
            Width           =   3135
         End
         Begin VB.ComboBox cmbChatBubbleTarget 
            Height          =   315
            ItemData        =   "frmEditor_Events.frx":0683
            Left            =   1920
            List            =   "frmEditor_Events.frx":0685
            Style           =   2  'Dropdown List
            TabIndex        =   366
            Top             =   1560
            Visible         =   0   'False
            Width           =   2895
         End
         Begin VB.OptionButton optChatBubbleTarget 
            Caption         =   "Event"
            Height          =   255
            Index           =   2
            Left            =   3720
            TabIndex        =   365
            Top             =   1080
            Width           =   975
         End
         Begin VB.OptionButton optChatBubbleTarget 
            Caption         =   "NPC"
            Height          =   255
            Index           =   1
            Left            =   1920
            TabIndex        =   364
            Top             =   1080
            Width           =   1335
         End
         Begin VB.OptionButton optChatBubbleTarget 
            Caption         =   "Player"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   363
            Top             =   1080
            Width           =   1695
         End
         Begin VB.CommandButton cmdShowChatBubble_Ok 
            Caption         =   "Ok"
            Height          =   375
            Left            =   2160
            TabIndex        =   362
            Top             =   2280
            Width           =   1215
         End
         Begin VB.CommandButton cmdShowChatBubble_Cancel 
            Caption         =   "Cancel"
            Height          =   375
            Left            =   3600
            TabIndex        =   361
            Top             =   2280
            Width           =   1215
         End
         Begin VB.Label lblRandomLabel 
            Caption         =   "Target Type:"
            Height          =   255
            Index           =   39
            Left            =   120
            TabIndex        =   368
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label lblRandomLabel 
            Caption         =   "Chatbubble Text:"
            Height          =   255
            Index           =   38
            Left            =   120
            TabIndex        =   367
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.Frame fraCommand 
         Caption         =   "Player Variable"
         Height          =   2535
         Index           =   4
         Left            =   1320
         TabIndex        =   82
         Top             =   1560
         Visible         =   0   'False
         Width           =   4095
         Begin VB.TextBox txtVariableData 
            Enabled         =   0   'False
            Height          =   285
            Index           =   4
            Left            =   3240
            TabIndex        =   344
            Text            =   "0"
            Top             =   1560
            Width           =   735
         End
         Begin VB.TextBox txtVariableData 
            Enabled         =   0   'False
            Height          =   285
            Index           =   3
            Left            =   1920
            TabIndex        =   343
            Text            =   "0"
            Top             =   1590
            Width           =   855
         End
         Begin VB.TextBox txtVariableData 
            Enabled         =   0   'False
            Height          =   285
            Index           =   2
            Left            =   1680
            TabIndex        =   342
            Text            =   "0"
            Top             =   1320
            Width           =   2295
         End
         Begin VB.TextBox txtVariableData 
            Enabled         =   0   'False
            Height          =   285
            Index           =   1
            Left            =   1680
            TabIndex        =   341
            Text            =   "0"
            Top             =   1080
            Width           =   2295
         End
         Begin VB.TextBox txtVariableData 
            Height          =   285
            Index           =   0
            Left            =   1680
            TabIndex        =   340
            Text            =   "0"
            Top             =   840
            Width           =   2295
         End
         Begin VB.OptionButton optVariableAction 
            Caption         =   "Random"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   339
            Top             =   1560
            Width           =   1095
         End
         Begin VB.OptionButton optVariableAction 
            Caption         =   "Subtract"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   338
            Top             =   1320
            Width           =   1095
         End
         Begin VB.OptionButton optVariableAction 
            Caption         =   "Add"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   337
            Top             =   1080
            Width           =   1095
         End
         Begin VB.OptionButton optVariableAction 
            Caption         =   "Set"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   336
            Top             =   840
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.CommandButton cmdVariableCancel 
            Caption         =   "Cancel"
            Height          =   375
            Left            =   2760
            TabIndex        =   86
            Top             =   1920
            Width           =   1215
         End
         Begin VB.CommandButton cmdVariableOK 
            Caption         =   "Ok"
            Height          =   375
            Left            =   1440
            TabIndex        =   85
            Top             =   1920
            Width           =   1215
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
            Caption         =   "High:"
            Height          =   255
            Index           =   37
            Left            =   2760
            TabIndex        =   359
            Top             =   1590
            Width           =   495
         End
         Begin VB.Label lblRandomLabel 
            Caption         =   "Low:"
            Height          =   255
            Index           =   13
            Left            =   1440
            TabIndex        =   358
            Top             =   1590
            Width           =   495
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
      Begin VB.Frame fraCommand 
         Caption         =   "Create Label"
         Height          =   1695
         Index           =   8
         Left            =   1200
         TabIndex        =   375
         Top             =   1680
         Visible         =   0   'False
         Width           =   4095
         Begin VB.TextBox txtLabelName 
            Height          =   375
            Left            =   120
            TabIndex        =   379
            Top             =   480
            Width           =   3855
         End
         Begin VB.CommandButton cmdCreateLabel_Ok 
            Caption         =   "Ok"
            Height          =   375
            Left            =   1320
            TabIndex        =   377
            Top             =   1200
            Width           =   1215
         End
         Begin VB.CommandButton cmdCreateLabel_Cancel 
            Caption         =   "Cancel"
            Height          =   375
            Left            =   2640
            TabIndex        =   376
            Top             =   1200
            Width           =   1215
         End
         Begin VB.Label lblRandomLabel 
            Caption         =   "Name:"
            Height          =   255
            Index           =   40
            Left            =   120
            TabIndex        =   378
            Top             =   240
            Width           =   3735
         End
      End
      Begin VB.Frame fraCommand 
         Caption         =   "Player Switch"
         Height          =   1695
         Index           =   5
         Left            =   1320
         TabIndex        =   200
         Top             =   1920
         Visible         =   0   'False
         Width           =   4095
         Begin VB.ComboBox cmbPlayerSwitchSet 
            Height          =   315
            ItemData        =   "frmEditor_Events.frx":0687
            Left            =   960
            List            =   "frmEditor_Events.frx":0691
            Style           =   2  'Dropdown List
            TabIndex        =   206
            Top             =   800
            Width           =   3015
         End
         Begin VB.ComboBox cmbSwitch 
            Height          =   315
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   203
            Top             =   360
            Width           =   3015
         End
         Begin VB.CommandButton cmbPlayerSwitch_Ok 
            Caption         =   "Ok"
            Height          =   375
            Left            =   1440
            TabIndex        =   202
            Top             =   1200
            Width           =   1215
         End
         Begin VB.CommandButton cmdPlayerSwitch_Cancel 
            Caption         =   "Cancel"
            Height          =   375
            Left            =   2760
            TabIndex        =   201
            Top             =   1200
            Width           =   1215
         End
         Begin VB.Label lblRandomLabel 
            Caption         =   "Switch:"
            Height          =   255
            Index           =   23
            Left            =   120
            TabIndex        =   205
            Top             =   360
            Width           =   3855
         End
         Begin VB.Label lblRandomLabel 
            Caption         =   "Set to:"
            Height          =   255
            Index           =   22
            Left            =   120
            TabIndex        =   204
            Top             =   840
            Width           =   1815
         End
      End
      Begin VB.Frame fraCommand 
         Caption         =   "Self Switch"
         Height          =   1695
         Index           =   6
         Left            =   1320
         TabIndex        =   207
         Top             =   1920
         Visible         =   0   'False
         Width           =   4095
         Begin VB.CommandButton cmdSelfSwitch_Cancel 
            Caption         =   "Cancel"
            Height          =   375
            Left            =   2760
            TabIndex        =   211
            Top             =   1200
            Width           =   1215
         End
         Begin VB.CommandButton cmdSelfSwitch_Ok 
            Caption         =   "Ok"
            Height          =   375
            Left            =   1440
            TabIndex        =   210
            Top             =   1200
            Width           =   1215
         End
         Begin VB.ComboBox cmbSetSelfSwitch 
            Height          =   315
            ItemData        =   "frmEditor_Events.frx":06A2
            Left            =   1440
            List            =   "frmEditor_Events.frx":06B2
            Style           =   2  'Dropdown List
            TabIndex        =   209
            Top             =   360
            Width           =   2535
         End
         Begin VB.ComboBox cmbSetSelfSwitchTo 
            Height          =   315
            ItemData        =   "frmEditor_Events.frx":06C2
            Left            =   960
            List            =   "frmEditor_Events.frx":06CC
            Style           =   2  'Dropdown List
            TabIndex        =   208
            Top             =   800
            Width           =   3015
         End
         Begin VB.Label lblRandomLabel 
            Caption         =   "Set to:"
            Height          =   255
            Index           =   26
            Left            =   120
            TabIndex        =   213
            Top             =   840
            Width           =   1815
         End
         Begin VB.Label lblRandomLabel 
            Caption         =   "Self Switch:"
            Height          =   255
            Index           =   24
            Left            =   120
            TabIndex        =   212
            Top             =   360
            Width           =   3855
         End
      End
      Begin VB.Frame fraCommand 
         Caption         =   "Conditional Branch"
         Height          =   4815
         Index           =   7
         Left            =   120
         TabIndex        =   155
         Top             =   480
         Visible         =   0   'False
         Width           =   6135
         Begin VB.OptionButton optCondition_Index 
            Caption         =   "Self Switch"
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   298
            Top             =   3720
            Width           =   1695
         End
         Begin VB.ComboBox cmbCondition_SelfSwitch 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmEditor_Events.frx":06D9
            Left            =   1920
            List            =   "frmEditor_Events.frx":06E9
            Style           =   2  'Dropdown List
            TabIndex        =   297
            Top             =   3720
            Width           =   1695
         End
         Begin VB.ComboBox cmbCondition_SelfSwitchCondition 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmEditor_Events.frx":06F9
            Left            =   3960
            List            =   "frmEditor_Events.frx":0703
            Style           =   2  'Dropdown List
            TabIndex        =   296
            Top             =   3720
            Width           =   1095
         End
         Begin VB.ComboBox cmbCondition_LearntSkill 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmEditor_Events.frx":0714
            Left            =   1920
            List            =   "frmEditor_Events.frx":0716
            Style           =   2  'Dropdown List
            TabIndex        =   175
            Top             =   2760
            Width           =   1695
         End
         Begin VB.ComboBox cmbCondition_ClassIs 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmEditor_Events.frx":0718
            Left            =   1920
            List            =   "frmEditor_Events.frx":071A
            Style           =   2  'Dropdown List
            TabIndex        =   174
            Top             =   2280
            Width           =   1695
         End
         Begin VB.ComboBox cmbCondition_HasItem 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmEditor_Events.frx":071C
            Left            =   1920
            List            =   "frmEditor_Events.frx":071E
            Style           =   2  'Dropdown List
            TabIndex        =   173
            Top             =   1800
            Width           =   1695
         End
         Begin VB.ComboBox cmbCondtion_PlayerSwitchCondition 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmEditor_Events.frx":0720
            Left            =   3960
            List            =   "frmEditor_Events.frx":072A
            Style           =   2  'Dropdown List
            TabIndex        =   172
            Top             =   1320
            Width           =   1095
         End
         Begin VB.ComboBox cmbCondition_PlayerSwitch 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmEditor_Events.frx":073B
            Left            =   1920
            List            =   "frmEditor_Events.frx":073D
            Style           =   2  'Dropdown List
            TabIndex        =   171
            Top             =   1320
            Width           =   1695
         End
         Begin VB.TextBox txtCondition_LevelAmount 
            Enabled         =   0   'False
            Height          =   285
            Left            =   3480
            TabIndex        =   170
            Text            =   "0"
            Top             =   3240
            Width           =   855
         End
         Begin VB.ComboBox cmbCondition_LevelCompare 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmEditor_Events.frx":073F
            Left            =   1440
            List            =   "frmEditor_Events.frx":0755
            Style           =   2  'Dropdown List
            TabIndex        =   169
            Top             =   3240
            Width           =   1695
         End
         Begin VB.OptionButton optCondition_Index 
            Caption         =   "Level"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   167
            Top             =   3240
            Width           =   975
         End
         Begin VB.ComboBox cmbCondition_PlayerVarCompare 
            Height          =   315
            ItemData        =   "frmEditor_Events.frx":07BB
            Left            =   1920
            List            =   "frmEditor_Events.frx":07D1
            Style           =   2  'Dropdown List
            TabIndex        =   165
            Top             =   840
            Width           =   1695
         End
         Begin VB.TextBox txtCondition_PlayerVarCondition 
            Height          =   285
            Left            =   3840
            TabIndex        =   164
            Text            =   "0"
            Top             =   840
            Width           =   855
         End
         Begin VB.OptionButton optCondition_Index 
            Caption         =   "Knows Skill"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   162
            Top             =   2760
            Width           =   1695
         End
         Begin VB.OptionButton optCondition_Index 
            Caption         =   "Class Is"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   161
            Top             =   2280
            Width           =   1695
         End
         Begin VB.OptionButton optCondition_Index 
            Caption         =   "Has Item"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   160
            Top             =   1800
            Width           =   1695
         End
         Begin VB.OptionButton optCondition_Index 
            Caption         =   "Player Switch"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   159
            Top             =   1320
            Width           =   1695
         End
         Begin VB.OptionButton optCondition_Index 
            Caption         =   "Player Variable"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   158
            Top             =   480
            Width           =   1695
         End
         Begin VB.CommandButton cmdCondition_Ok 
            Caption         =   "Ok"
            Height          =   375
            Left            =   3360
            TabIndex        =   157
            Top             =   4320
            Width           =   1215
         End
         Begin VB.CommandButton cmdCondition_Cancel 
            Caption         =   "Cancel"
            Height          =   375
            Left            =   4680
            TabIndex        =   156
            Top             =   4320
            Width           =   1215
         End
         Begin VB.ComboBox cmbCondition_PlayerVarIndex 
            Height          =   315
            ItemData        =   "frmEditor_Events.frx":0837
            Left            =   1920
            List            =   "frmEditor_Events.frx":0839
            Style           =   2  'Dropdown List
            TabIndex        =   163
            Top             =   480
            Width           =   1695
         End
         Begin VB.Label lblRandomLabel 
            Caption         =   "is"
            Height          =   255
            Index           =   35
            Left            =   3720
            TabIndex        =   299
            Top             =   3720
            Width           =   255
         End
         Begin VB.Label lblRandomLabel 
            Caption         =   "is"
            Height          =   255
            Index           =   1
            Left            =   3720
            TabIndex        =   181
            Top             =   1320
            Width           =   255
         End
         Begin VB.Label lblRandomLabel 
            Caption         =   "is"
            Height          =   255
            Index           =   2
            Left            =   1080
            TabIndex        =   168
            Top             =   3240
            Width           =   255
         End
         Begin VB.Label lblRandomLabel 
            Caption         =   "is"
            Height          =   255
            Index           =   0
            Left            =   3840
            TabIndex        =   166
            Top             =   480
            Width           =   615
         End
      End
      Begin VB.Frame fraCommand 
         Caption         =   "Go to Label"
         Height          =   1695
         Index           =   9
         Left            =   1200
         TabIndex        =   380
         Top             =   1680
         Visible         =   0   'False
         Width           =   4095
         Begin VB.CommandButton cmdGotoLabel_Cancel 
            Caption         =   "Cancel"
            Height          =   375
            Left            =   2640
            TabIndex        =   383
            Top             =   1200
            Width           =   1215
         End
         Begin VB.CommandButton cmdGotoLabel_Ok 
            Caption         =   "Ok"
            Height          =   375
            Left            =   1320
            TabIndex        =   382
            Top             =   1200
            Width           =   1215
         End
         Begin VB.TextBox txtGotoLabel 
            Height          =   375
            Left            =   120
            TabIndex        =   381
            Top             =   480
            Width           =   3855
         End
         Begin VB.Label lblRandomLabel 
            Caption         =   "Label Name:"
            Height          =   255
            Index           =   41
            Left            =   120
            TabIndex        =   384
            Top             =   240
            Width           =   3735
         End
      End
      Begin VB.Frame fraCommand 
         Caption         =   "Add Text"
         Height          =   4095
         Index           =   2
         Left            =   1200
         TabIndex        =   215
         Top             =   600
         Visible         =   0   'False
         Width           =   4095
         Begin VB.TextBox txtAddText_Text 
            Height          =   1815
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   222
            Top             =   480
            Width           =   3855
         End
         Begin VB.HScrollBar scrlAddText_Colour 
            Height          =   255
            Left            =   120
            Max             =   17
            TabIndex        =   221
            Top             =   2640
            Width           =   3855
         End
         Begin VB.OptionButton optAddText_Player 
            Caption         =   "Player"
            Height          =   255
            Left            =   120
            TabIndex        =   220
            Top             =   3240
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton optAddText_Map 
            Caption         =   "Map"
            Height          =   255
            Left            =   1080
            TabIndex        =   219
            Top             =   3240
            Width           =   735
         End
         Begin VB.OptionButton optAddText_Global 
            Caption         =   "Global"
            Height          =   255
            Left            =   1920
            TabIndex        =   218
            Top             =   3240
            Width           =   855
         End
         Begin VB.CommandButton cmdAddText_Ok 
            Caption         =   "Ok"
            Height          =   375
            Left            =   1440
            TabIndex        =   217
            Top             =   3600
            Width           =   1215
         End
         Begin VB.CommandButton cmdAddText_Cancel 
            Caption         =   "Cancel"
            Height          =   375
            Left            =   2760
            TabIndex        =   216
            Top             =   3600
            Width           =   1215
         End
         Begin VB.Label lblRandomLabel 
            Caption         =   "Text:"
            Height          =   255
            Index           =   34
            Left            =   120
            TabIndex        =   225
            Top             =   240
            Width           =   1935
         End
         Begin VB.Label lblAddText_Colour 
            Caption         =   "Color: Black"
            Height          =   255
            Left            =   120
            TabIndex        =   224
            Top             =   2400
            Width           =   3255
         End
         Begin VB.Label lblRandomLabel 
            Caption         =   "Channel:"
            Height          =   255
            Index           =   10
            Left            =   120
            TabIndex        =   223
            Top             =   3000
            Width           =   1575
         End
      End
      Begin VB.Frame fraCommand 
         Caption         =   "Change Items"
         Height          =   2415
         Index           =   10
         Left            =   1200
         TabIndex        =   226
         Top             =   1320
         Visible         =   0   'False
         Width           =   4095
         Begin VB.ComboBox cmbChangeItemIndex 
            Height          =   315
            ItemData        =   "frmEditor_Events.frx":083B
            Left            =   120
            List            =   "frmEditor_Events.frx":083D
            Style           =   2  'Dropdown List
            TabIndex        =   234
            Top             =   480
            Width           =   1695
         End
         Begin VB.TextBox txtChangeItemsAmount 
            Height          =   375
            Left            =   120
            TabIndex        =   233
            Text            =   "0"
            Top             =   1320
            Width           =   3735
         End
         Begin VB.CommandButton cmdChangeItems_Cancel 
            Caption         =   "Cancel"
            Height          =   375
            Left            =   2640
            TabIndex        =   231
            Top             =   1800
            Width           =   1215
         End
         Begin VB.CommandButton cmdChangeItems_Ok 
            Caption         =   "Ok"
            Height          =   375
            Left            =   1200
            TabIndex        =   230
            Top             =   1800
            Width           =   1215
         End
         Begin VB.OptionButton optChangeItemRemove 
            Caption         =   "Take Away"
            Height          =   255
            Left            =   2640
            TabIndex        =   229
            Top             =   960
            Width           =   1335
         End
         Begin VB.OptionButton optChangeItemAdd 
            Caption         =   "Give"
            Height          =   255
            Left            =   1680
            TabIndex        =   228
            Top             =   960
            Width           =   735
         End
         Begin VB.OptionButton optChangeItemSet 
            Caption         =   "Set Amount"
            Height          =   255
            Left            =   120
            TabIndex        =   227
            Top             =   960
            Value           =   -1  'True
            Width           =   1455
         End
         Begin VB.Label lblRandomLabel 
            Caption         =   "Item Index:"
            Height          =   255
            Index           =   27
            Left            =   120
            TabIndex        =   232
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.Frame fraCommand 
         Caption         =   "Change Level"
         Height          =   1815
         Index           =   11
         Left            =   1200
         TabIndex        =   235
         Top             =   1560
         Visible         =   0   'False
         Width           =   4095
         Begin VB.HScrollBar scrlChangeLevel 
            Height          =   255
            Left            =   120
            TabIndex        =   239
            Top             =   600
            Width           =   3615
         End
         Begin VB.CommandButton cmdChangeLevel_Ok 
            Caption         =   "Ok"
            Height          =   375
            Left            =   1200
            TabIndex        =   237
            Top             =   1080
            Width           =   1215
         End
         Begin VB.CommandButton cmdChangeLevel_Cancel 
            Caption         =   "Cancel"
            Height          =   375
            Left            =   2520
            TabIndex        =   236
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label lblChangeLevel 
            Caption         =   "Level: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   238
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.Frame fraCommand 
         Caption         =   "Change Player Skills"
         Height          =   2175
         Index           =   12
         Left            =   1200
         TabIndex        =   240
         Top             =   1440
         Visible         =   0   'False
         Width           =   4095
         Begin VB.OptionButton optChangeSkillsRemove 
            Caption         =   "Remove"
            Height          =   255
            Left            =   2640
            TabIndex        =   246
            Top             =   960
            Width           =   1095
         End
         Begin VB.OptionButton optChangeSkillsAdd 
            Caption         =   "Teach"
            Height          =   255
            Left            =   120
            TabIndex        =   245
            Top             =   960
            Width           =   1455
         End
         Begin VB.ComboBox cmbChangeSkills 
            Height          =   315
            ItemData        =   "frmEditor_Events.frx":083F
            Left            =   120
            List            =   "frmEditor_Events.frx":0841
            Style           =   2  'Dropdown List
            TabIndex        =   244
            Top             =   480
            Width           =   3735
         End
         Begin VB.CommandButton cmdChangeSkills_Cancel 
            Caption         =   "Cancel"
            Height          =   375
            Left            =   2520
            TabIndex        =   242
            Top             =   1680
            Width           =   1215
         End
         Begin VB.CommandButton cmdChangeSkills_Ok 
            Caption         =   "Ok"
            Height          =   375
            Left            =   1200
            TabIndex        =   241
            Top             =   1680
            Width           =   1215
         End
         Begin VB.Label lblRandomLabel 
            Caption         =   "Skill"
            Height          =   255
            Index           =   28
            Left            =   120
            TabIndex        =   243
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.Frame fraCommand 
         Caption         =   "Change Player Class"
         Height          =   1695
         Index           =   13
         Left            =   1200
         TabIndex        =   247
         Top             =   1680
         Visible         =   0   'False
         Width           =   4095
         Begin VB.CommandButton cmdChangeClass_Ok 
            Caption         =   "Ok"
            Height          =   375
            Left            =   1200
            TabIndex        =   250
            Top             =   960
            Width           =   1215
         End
         Begin VB.CommandButton cmdChangeClass_Cancel 
            Caption         =   "Cancel"
            Height          =   375
            Left            =   2520
            TabIndex        =   249
            Top             =   960
            Width           =   1215
         End
         Begin VB.ComboBox cmbChangeClass 
            Height          =   315
            ItemData        =   "frmEditor_Events.frx":0843
            Left            =   120
            List            =   "frmEditor_Events.frx":0845
            Style           =   2  'Dropdown List
            TabIndex        =   248
            Top             =   480
            Width           =   3735
         End
         Begin VB.Label lblRandomLabel 
            Caption         =   "Class:"
            Height          =   255
            Index           =   29
            Left            =   120
            TabIndex        =   251
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Frame fraCommand 
         Caption         =   "Change Player Sprite"
         Height          =   1695
         Index           =   14
         Left            =   1200
         TabIndex        =   252
         Top             =   1680
         Visible         =   0   'False
         Width           =   4095
         Begin VB.HScrollBar scrlChangeSprite 
            Height          =   255
            Left            =   1200
            Max             =   100
            Min             =   1
            TabIndex        =   256
            Top             =   360
            Value           =   1
            Width           =   2535
         End
         Begin VB.CommandButton cmdChangeSprite_Cancel 
            Caption         =   "Cancel"
            Height          =   375
            Left            =   2520
            TabIndex        =   254
            Top             =   960
            Width           =   1215
         End
         Begin VB.CommandButton cmdChangeSprite_Ok 
            Caption         =   "Ok"
            Height          =   375
            Left            =   1200
            TabIndex        =   253
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label lblChangeSprite 
            Caption         =   "Sprite: 1"
            Height          =   255
            Left            =   120
            TabIndex        =   255
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.Frame fraCommand 
         Caption         =   "Change Player Sex"
         Height          =   1455
         Index           =   15
         Left            =   1200
         TabIndex        =   257
         Top             =   1800
         Visible         =   0   'False
         Width           =   4095
         Begin VB.OptionButton optChangeSexFemale 
            Caption         =   "Female"
            Height          =   255
            Left            =   1200
            TabIndex        =   261
            Top             =   360
            Width           =   975
         End
         Begin VB.OptionButton optChangeSexMale 
            Caption         =   "Male"
            Height          =   255
            Left            =   240
            TabIndex        =   260
            Top             =   360
            Width           =   1455
         End
         Begin VB.CommandButton cmdChangeSex_Ok 
            Caption         =   "Ok"
            Height          =   375
            Left            =   1200
            TabIndex        =   259
            Top             =   840
            Width           =   1215
         End
         Begin VB.CommandButton cmdChangeSex_Cancel 
            Caption         =   "Cancel"
            Height          =   375
            Left            =   2520
            TabIndex        =   258
            Top             =   840
            Width           =   1215
         End
      End
      Begin VB.Frame fraCommand 
         Caption         =   "Set Player PK"
         Height          =   1455
         Index           =   16
         Left            =   1200
         TabIndex        =   262
         Top             =   1800
         Visible         =   0   'False
         Width           =   4095
         Begin VB.CommandButton cmdChangePK_Cancel 
            Caption         =   "Cancel"
            Height          =   375
            Left            =   2520
            TabIndex        =   266
            Top             =   840
            Width           =   1215
         End
         Begin VB.CommandButton cmdChangePK_Ok 
            Caption         =   "Ok"
            Height          =   375
            Left            =   1200
            TabIndex        =   265
            Top             =   840
            Width           =   1215
         End
         Begin VB.OptionButton optChangePKYes 
            Caption         =   "Yes"
            Height          =   255
            Left            =   240
            TabIndex        =   264
            Top             =   360
            Width           =   735
         End
         Begin VB.OptionButton optChangePKNo 
            Caption         =   "No"
            Height          =   255
            Left            =   1080
            TabIndex        =   263
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.Frame fraCommand 
         Caption         =   "Give Experience"
         Height          =   1695
         Index           =   17
         Left            =   1200
         TabIndex        =   370
         Top             =   1680
         Visible         =   0   'False
         Width           =   4095
         Begin VB.HScrollBar scrlGiveExp 
            Height          =   255
            Left            =   120
            Max             =   32000
            TabIndex        =   373
            Top             =   480
            Width           =   3735
         End
         Begin VB.CommandButton cmdGiveExp_Cancel 
            Caption         =   "Cancel"
            Height          =   375
            Left            =   2640
            TabIndex        =   372
            Top             =   1200
            Width           =   1215
         End
         Begin VB.CommandButton cmdGiveExp_Ok 
            Caption         =   "Ok"
            Height          =   375
            Left            =   1320
            TabIndex        =   371
            Top             =   1200
            Width           =   1215
         End
         Begin VB.Label lblGiveExp 
            Caption         =   "Experience: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   374
            Top             =   240
            Width           =   3735
         End
      End
      Begin VB.Frame fraCommand 
         Caption         =   "Warp Player"
         Height          =   3015
         Index           =   18
         Left            =   1320
         TabIndex        =   87
         Top             =   1320
         Visible         =   0   'False
         Width           =   4095
         Begin VB.ComboBox cmbWarpPlayerDir 
            Height          =   315
            ItemData        =   "frmEditor_Events.frx":0847
            Left            =   120
            List            =   "frmEditor_Events.frx":085A
            Style           =   2  'Dropdown List
            TabIndex        =   300
            Top             =   2040
            Width           =   3855
         End
         Begin VB.CommandButton cmdWPCancel 
            Caption         =   "Cancel"
            Height          =   375
            Left            =   2760
            TabIndex        =   95
            Top             =   2520
            Width           =   1215
         End
         Begin VB.CommandButton cmdWPOkay 
            Caption         =   "Ok"
            Height          =   375
            Left            =   1440
            TabIndex        =   94
            Top             =   2520
            Width           =   1215
         End
         Begin VB.HScrollBar scrlWPY 
            Height          =   255
            Left            =   120
            Max             =   255
            TabIndex        =   93
            Top             =   1680
            Width           =   3855
         End
         Begin VB.HScrollBar scrlWPX 
            Height          =   255
            Left            =   120
            Max             =   255
            TabIndex        =   91
            Top             =   1080
            Width           =   3855
         End
         Begin VB.HScrollBar scrlWPMap 
            Height          =   255
            Left            =   120
            TabIndex        =   89
            Top             =   480
            Width           =   3855
         End
         Begin VB.Label lblWPY 
            Caption         =   "Y: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   92
            Top             =   1440
            Width           =   3135
         End
         Begin VB.Label lblWPX 
            Caption         =   "X: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   90
            Top             =   840
            Width           =   3135
         End
         Begin VB.Label lblWPMap 
            Caption         =   "Map: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   88
            Top             =   240
            Width           =   3135
         End
      End
      Begin VB.Frame fraCommand 
         Caption         =   "Spawn NPC"
         Height          =   1695
         Index           =   19
         Left            =   1200
         TabIndex        =   385
         Top             =   1680
         Visible         =   0   'False
         Width           =   4095
         Begin VB.ComboBox cmbSpawnNPC 
            Height          =   315
            ItemData        =   "frmEditor_Events.frx":089B
            Left            =   120
            List            =   "frmEditor_Events.frx":089D
            Style           =   2  'Dropdown List
            TabIndex        =   389
            Top             =   480
            Width           =   3735
         End
         Begin VB.CommandButton cmdSpawnNpc_Ok 
            Caption         =   "Ok"
            Height          =   375
            Left            =   1320
            TabIndex        =   387
            Top             =   1200
            Width           =   1215
         End
         Begin VB.CommandButton cmdSpawnNpc_Cancel 
            Caption         =   "Cancel"
            Height          =   375
            Left            =   2640
            TabIndex        =   386
            Top             =   1200
            Width           =   1215
         End
         Begin VB.Label lblRandomLabel 
            Caption         =   "NPC:"
            Height          =   255
            Index           =   42
            Left            =   120
            TabIndex        =   388
            Top             =   240
            Width           =   3735
         End
      End
      Begin VB.Frame fraCommand 
         Caption         =   "Play Animation"
         Height          =   2775
         Index           =   20
         Left            =   720
         TabIndex        =   269
         Top             =   1320
         Visible         =   0   'False
         Width           =   5055
         Begin VB.ComboBox cmbPlayAnim 
            Height          =   315
            ItemData        =   "frmEditor_Events.frx":089F
            Left            =   1680
            List            =   "frmEditor_Events.frx":08A1
            Style           =   2  'Dropdown List
            TabIndex        =   282
            Top             =   300
            Width           =   3135
         End
         Begin VB.HScrollBar scrlPlayAnimTileY 
            Height          =   255
            Left            =   1920
            TabIndex        =   280
            Top             =   1800
            Width           =   2895
         End
         Begin VB.HScrollBar scrlPlayAnimTileX 
            Height          =   255
            Left            =   1920
            TabIndex        =   279
            Top             =   1455
            Width           =   2895
         End
         Begin VB.CommandButton cmdPlayAnim_Cancel 
            Caption         =   "Cancel"
            Height          =   375
            Left            =   3600
            TabIndex        =   275
            Top             =   2280
            Width           =   1215
         End
         Begin VB.CommandButton cmdPlayAnim_Ok 
            Caption         =   "Ok"
            Height          =   375
            Left            =   2160
            TabIndex        =   274
            Top             =   2280
            Width           =   1215
         End
         Begin VB.OptionButton optPlayAnimPlayer 
            Caption         =   "Player"
            Height          =   255
            Left            =   120
            TabIndex        =   273
            Top             =   1080
            Width           =   1695
         End
         Begin VB.OptionButton optPlayAnimEvent 
            Caption         =   "Event"
            Height          =   255
            Left            =   1920
            TabIndex        =   272
            Top             =   1080
            Width           =   1335
         End
         Begin VB.OptionButton optPlayAnimTile 
            Caption         =   "Tile"
            Height          =   255
            Left            =   3720
            TabIndex        =   271
            Top             =   1080
            Width           =   975
         End
         Begin VB.ComboBox cmbPlayAnimEvent 
            Height          =   315
            ItemData        =   "frmEditor_Events.frx":08A3
            Left            =   1920
            List            =   "frmEditor_Events.frx":08A5
            Style           =   2  'Dropdown List
            TabIndex        =   270
            Top             =   1440
            Visible         =   0   'False
            Width           =   2895
         End
         Begin VB.Label lblRandomLabel 
            Caption         =   "Animation"
            Height          =   255
            Index           =   30
            Left            =   120
            TabIndex        =   281
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label lblPlayAnimY 
            Caption         =   "Map Tile Y:"
            Height          =   255
            Left            =   240
            TabIndex        =   278
            Top             =   1800
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label lblPlayAnimX 
            Caption         =   "Map Tile X:"
            Height          =   255
            Left            =   240
            TabIndex        =   277
            Top             =   1440
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label lblRandomLabel 
            Caption         =   "Target Type:"
            Height          =   255
            Index           =   31
            Left            =   120
            TabIndex        =   276
            Top             =   840
            Width           =   1335
         End
      End
      Begin VB.Frame fraCommand 
         Caption         =   "Open Shop"
         Height          =   1575
         Index           =   21
         Left            =   1200
         TabIndex        =   314
         Top             =   1800
         Visible         =   0   'False
         Width           =   4215
         Begin VB.CommandButton cmdOpenShop_Ok 
            Caption         =   "Ok"
            Height          =   375
            Left            =   1440
            TabIndex        =   317
            Top             =   1080
            Width           =   1215
         End
         Begin VB.CommandButton cmdOpenShop_Cancel 
            Caption         =   "Cancel"
            Height          =   375
            Left            =   2880
            TabIndex        =   316
            Top             =   1080
            Width           =   1215
         End
         Begin VB.ComboBox cmbOpenShop 
            Height          =   315
            ItemData        =   "frmEditor_Events.frx":08A7
            Left            =   960
            List            =   "frmEditor_Events.frx":08BA
            Style           =   2  'Dropdown List
            TabIndex        =   315
            Top             =   360
            Width           =   3135
         End
      End
      Begin VB.Frame fraCommand 
         Caption         =   "Set Fog"
         Height          =   2415
         Index           =   22
         Left            =   1200
         TabIndex        =   390
         Top             =   1440
         Visible         =   0   'False
         Width           =   4095
         Begin VB.CommandButton cmdSetFog_Ok 
            Caption         =   "Ok"
            Height          =   375
            Left            =   1440
            TabIndex        =   398
            Top             =   1920
            Width           =   1215
         End
         Begin VB.CommandButton cmdSetFog_Cancel 
            Caption         =   "Cancel"
            Height          =   375
            Left            =   2760
            TabIndex        =   397
            Top             =   1920
            Width           =   1215
         End
         Begin VB.HScrollBar ScrlFogData 
            Height          =   255
            Index           =   1
            Left            =   120
            Max             =   255
            TabIndex        =   393
            Top             =   1050
            Width           =   1575
         End
         Begin VB.HScrollBar ScrlFogData 
            Height          =   255
            Index           =   0
            Left            =   120
            Max             =   255
            TabIndex        =   392
            Top             =   480
            Width           =   1575
         End
         Begin VB.HScrollBar ScrlFogData 
            Height          =   255
            Index           =   2
            Left            =   120
            Max             =   255
            TabIndex        =   391
            Top             =   1620
            Width           =   1575
         End
         Begin VB.Label lblFogData 
            Caption         =   "Fog Speed: 0"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   396
            Top             =   810
            Width           =   1815
         End
         Begin VB.Label lblFogData 
            Caption         =   "Fog: None"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   395
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label lblFogData 
            Caption         =   "Fog Opacity: 0"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   394
            Top             =   1380
            Width           =   1815
         End
      End
      Begin VB.Frame fraCommand 
         Caption         =   "Set Weather"
         Height          =   1935
         Index           =   23
         Left            =   1200
         TabIndex        =   399
         Top             =   1680
         Visible         =   0   'False
         Width           =   4095
         Begin VB.CommandButton cmdSetWeather_Cancel 
            Caption         =   "Cancel"
            Height          =   375
            Left            =   2760
            TabIndex        =   405
            Top             =   1440
            Width           =   1215
         End
         Begin VB.CommandButton cmdSetWeather_Ok 
            Caption         =   "Ok"
            Height          =   375
            Left            =   1440
            TabIndex        =   404
            Top             =   1440
            Width           =   1215
         End
         Begin VB.HScrollBar scrlWeatherIntensity 
            Height          =   255
            Left            =   120
            Max             =   100
            TabIndex        =   401
            Top             =   1080
            Width           =   1815
         End
         Begin VB.ComboBox CmbWeather 
            Height          =   315
            ItemData        =   "frmEditor_Events.frx":08FD
            Left            =   120
            List            =   "frmEditor_Events.frx":0913
            Style           =   2  'Dropdown List
            TabIndex        =   400
            Top             =   480
            Width           =   1815
         End
         Begin VB.Label lblWeatherIntensity 
            Caption         =   "Intensity: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   403
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label lblRandomLabel 
            AutoSize        =   -1  'True
            Caption         =   "Weather Type:"
            Height          =   195
            Index           =   43
            Left            =   120
            TabIndex        =   402
            Top             =   240
            Width           =   1275
         End
      End
      Begin VB.Frame fraCommand 
         Caption         =   "Show Text"
         Height          =   4095
         Index           =   0
         Left            =   360
         TabIndex        =   182
         Top             =   1200
         Visible         =   0   'False
         Width           =   5655
         Begin VB.HScrollBar scrlFace 
            Height          =   255
            Left            =   240
            TabIndex        =   424
            Top             =   2280
            Width           =   1575
         End
         Begin VB.PictureBox picFace 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1500
            Left            =   240
            ScaleHeight     =   100
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   100
            TabIndex        =   423
            TabStop         =   0   'False
            Top             =   480
            Width           =   1500
         End
         Begin VB.TextBox txtShowText 
            Height          =   2775
            Left            =   1920
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   185
            Top             =   480
            Width           =   3255
         End
         Begin VB.CommandButton cmdShowText_Ok 
            Caption         =   "Ok"
            Height          =   375
            Left            =   1440
            TabIndex        =   184
            Top             =   3600
            Width           =   1215
         End
         Begin VB.CommandButton cmdShowText_Cancel 
            Caption         =   "Cancel"
            Height          =   375
            Left            =   2760
            TabIndex        =   183
            Top             =   3600
            Width           =   1215
         End
         Begin VB.Label lblFace 
            Caption         =   "Face: 0"
            Height          =   255
            Left            =   240
            TabIndex        =   425
            Top             =   2040
            Width           =   975
         End
         Begin VB.Label lblRandomLabel 
            Caption         =   "Text:"
            Height          =   255
            Index           =   18
            Left            =   1920
            TabIndex        =   186
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.Frame fraCommand 
         Caption         =   "Show Choices"
         Height          =   4335
         Index           =   1
         Left            =   240
         TabIndex        =   187
         Top             =   1080
         Visible         =   0   'False
         Width           =   5655
         Begin VB.HScrollBar scrlFace2 
            Height          =   255
            Left            =   120
            TabIndex        =   427
            Top             =   2400
            Width           =   1455
         End
         Begin VB.PictureBox picFace2 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1500
            Left            =   120
            ScaleHeight     =   100
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   100
            TabIndex        =   426
            TabStop         =   0   'False
            Top             =   600
            Width           =   1500
         End
         Begin VB.TextBox txtChoices 
            Height          =   375
            Index           =   4
            Left            =   3720
            TabIndex        =   198
            Text            =   "0"
            Top             =   3360
            Width           =   1455
         End
         Begin VB.TextBox txtChoices 
            Height          =   375
            Index           =   3
            Left            =   1680
            TabIndex        =   196
            Text            =   "0"
            Top             =   3360
            Width           =   1455
         End
         Begin VB.TextBox txtChoices 
            Height          =   375
            Index           =   2
            Left            =   3720
            TabIndex        =   194
            Text            =   "0"
            Top             =   2640
            Width           =   1455
         End
         Begin VB.TextBox txtChoices 
            Height          =   375
            Index           =   1
            Left            =   1680
            TabIndex        =   192
            Text            =   "0"
            Top             =   2640
            Width           =   1455
         End
         Begin VB.CommandButton cmdChoices_Cancel 
            Caption         =   "Cancel"
            Height          =   375
            Left            =   2760
            TabIndex        =   190
            Top             =   3840
            Width           =   1215
         End
         Begin VB.CommandButton cmdChoices_Ok 
            Caption         =   "Ok"
            Height          =   375
            Left            =   1440
            TabIndex        =   189
            Top             =   3840
            Width           =   1215
         End
         Begin VB.TextBox txtChoicePrompt 
            Height          =   1695
            Left            =   1680
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   188
            Top             =   600
            Width           =   3855
         End
         Begin VB.Label lblFace2 
            Caption         =   "Face: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   428
            Top             =   2160
            Width           =   975
         End
         Begin VB.Label lblRandomLabel 
            Caption         =   "Choice 4"
            Height          =   255
            Index           =   21
            Left            =   3720
            TabIndex        =   199
            Top             =   3120
            Width           =   1455
         End
         Begin VB.Label lblRandomLabel 
            Caption         =   "Choice 3"
            Height          =   255
            Index           =   20
            Left            =   1680
            TabIndex        =   197
            Top             =   3120
            Width           =   1455
         End
         Begin VB.Label lblRandomLabel 
            Caption         =   "Choice 2"
            Height          =   255
            Index           =   19
            Left            =   3720
            TabIndex        =   195
            Top             =   2400
            Width           =   1455
         End
         Begin VB.Label lblRandomLabel 
            Caption         =   "Choice 1"
            Height          =   255
            Index           =   17
            Left            =   1680
            TabIndex        =   193
            Top             =   2400
            Width           =   1455
         End
         Begin VB.Label lblRandomLabel 
            Caption         =   "Text:"
            Height          =   255
            Index           =   16
            Left            =   1680
            TabIndex        =   191
            Top             =   360
            Width           =   1935
         End
      End
      Begin VB.Frame fraCommand 
         Caption         =   "Map Overlay"
         Height          =   2055
         Index           =   24
         Left            =   1320
         TabIndex        =   406
         Top             =   1320
         Visible         =   0   'False
         Width           =   4095
         Begin VB.CommandButton cmdMapTint_Cancel 
            Caption         =   "Cancel"
            Height          =   375
            Left            =   2760
            TabIndex        =   416
            Top             =   1560
            Width           =   1215
         End
         Begin VB.CommandButton cmdMapTint_Ok 
            Caption         =   "Ok"
            Height          =   375
            Left            =   1440
            TabIndex        =   415
            Top             =   1560
            Width           =   1215
         End
         Begin VB.HScrollBar scrlMapTintData 
            Height          =   255
            Index           =   3
            Left            =   2280
            Max             =   255
            TabIndex        =   410
            Top             =   1200
            Width           =   855
         End
         Begin VB.HScrollBar scrlMapTintData 
            Height          =   255
            Index           =   0
            Left            =   120
            Max             =   255
            TabIndex        =   409
            Top             =   480
            Width           =   855
         End
         Begin VB.HScrollBar scrlMapTintData 
            Height          =   255
            Index           =   1
            Left            =   2280
            Max             =   255
            TabIndex        =   408
            Top             =   480
            Width           =   855
         End
         Begin VB.HScrollBar scrlMapTintData 
            Height          =   255
            Index           =   2
            Left            =   120
            Max             =   255
            TabIndex        =   407
            Top             =   1200
            Width           =   855
         End
         Begin VB.Label lblMapTintData 
            Caption         =   "Alpha: 0"
            ForeColor       =   &H80000006&
            Height          =   255
            Index           =   3
            Left            =   2280
            TabIndex        =   414
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label lblMapTintData 
            Caption         =   "Red: 0"
            ForeColor       =   &H000000C0&
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   413
            Top             =   240
            Width           =   975
         End
         Begin VB.Label lblMapTintData 
            Caption         =   "Green: 0"
            ForeColor       =   &H0000C000&
            Height          =   255
            Index           =   1
            Left            =   2280
            TabIndex        =   412
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label lblMapTintData 
            Caption         =   "Blue: 0"
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   411
            Top             =   960
            Width           =   1455
         End
      End
   End
   Begin VB.ListBox lstCommands 
      Height          =   6105
      Left            =   6240
      TabIndex        =   1
      Top             =   1440
      Width           =   6255
   End
   Begin VB.Frame fraRandom 
      Caption         =   "Commands"
      Height          =   735
      Index           =   9
      Left            =   6240
      TabIndex        =   176
      Top             =   7560
      Width           =   6255
      Begin VB.CommandButton cmdClearCommand 
         Caption         =   "Clear"
         Height          =   375
         Left            =   4680
         TabIndex        =   180
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdDeleteCommand 
         Caption         =   "Delete"
         Height          =   375
         Left            =   3120
         TabIndex        =   179
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdEditCommand 
         Caption         =   "Edit"
         Height          =   375
         Left            =   1560
         TabIndex        =   178
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdAddCommand 
         Caption         =   "Add"
         Height          =   375
         Left            =   120
         TabIndex        =   177
         Top             =   240
         Width           =   1335
      End
   End
   Begin MSComctlLib.TabStrip tabPages 
      Height          =   8055
      Left            =   120
      TabIndex        =   34
      Top             =   960
      Width           =   12735
      _ExtentX        =   22463
      _ExtentY        =   14208
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
Private TmpIndex As Long

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

Private Sub cmbGraphic_Click()
    If cmbGraphic.ListIndex = -1 Then Exit Sub
    
    tmpEvent.Pages(curPageNum).GraphicType = cmbGraphic.ListIndex
    
    ' Set the max on the scrollbar
    Select Case cmbGraphic.ListIndex
        Case 0 ' None
            scrlGraphic.Value = 1
            scrlGraphic.Enabled = False
        Case 1 ' Character
            scrlGraphic.max = NumCharacters
            scrlGraphic.Enabled = True
        Case 2 ' Tileset
            scrlGraphic.max = NumTileSets
            scrlGraphic.Enabled = True
    End Select
    
    If scrlGraphic.Value = 0 Then
        lblGraphic.Caption = "Number: None"
    Else
        lblGraphic.Caption = "Number: " & scrlGraphic.Value
    End If
    
    If tmpEvent.Pages(curPageNum).GraphicType = 1 Then
        If frmEditor_Events.scrlGraphic.Value <= 0 Or frmEditor_Events.scrlGraphic.Value > NumCharacters Then Exit Sub
                    
        If Tex_Character(frmEditor_Events.scrlGraphic.Value).Width > 793 Then
            frmEditor_Events.hScrlGraphicSel.Visible = True
            frmEditor_Events.hScrlGraphicSel.Value = 0
            frmEditor_Events.hScrlGraphicSel.max = Tex_Character(frmEditor_Events.scrlGraphic.Value).Width - 800
        Else
            frmEditor_Events.hScrlGraphicSel.Visible = False
        End If
                    
        If Tex_Character(frmEditor_Events.scrlGraphic.Value).Height > 472 Then
            frmEditor_Events.vScrlGraphicSel.Visible = True
            frmEditor_Events.vScrlGraphicSel.Value = 0
            frmEditor_Events.vScrlGraphicSel.max = Tex_Character(frmEditor_Events.scrlGraphic.Value).Height - 512
        Else
            frmEditor_Events.vScrlGraphicSel.Visible = False
        End If
    ElseIf tmpEvent.Pages(curPageNum).GraphicType = 2 Then
        If frmEditor_Events.scrlGraphic.Value <= 0 Or frmEditor_Events.scrlGraphic.Value > NumTileSets Then Exit Sub
                    
        If Tex_Tileset(frmEditor_Events.scrlGraphic.Value).Width > 793 Then
            frmEditor_Events.hScrlGraphicSel.Visible = True
            frmEditor_Events.hScrlGraphicSel.Value = 0
            frmEditor_Events.hScrlGraphicSel.max = Tex_Tileset(frmEditor_Events.scrlGraphic.Value).Width - 800
        Else
            frmEditor_Events.hScrlGraphicSel.Visible = False
        End If
                    
        If Tex_Tileset(frmEditor_Events.scrlGraphic.Value).Height > 472 Then
            frmEditor_Events.vScrlGraphicSel.Visible = True
            frmEditor_Events.vScrlGraphicSel.Value = 0
            frmEditor_Events.vScrlGraphicSel.max = Tex_Tileset(frmEditor_Events.scrlGraphic.Value).Height - 512
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
    
    ' Hide
    fraDialogue.Visible = False
    fraCommand(5).Visible = False
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

Private Sub cmdAddMoveRoute_Click(Index As Integer)
    If Index = 42 Then
        fraGraphic.Width = 841
        fraGraphic.Height = 585
        fraGraphic.Visible = True
        GraphicSelType = 1
    Else
        AddMoveRouteCommand Index
    End If
End Sub

Private Sub cmdAddText_Cancel_Click()
    If Not isEdit Then fraCommands.Visible = True Else fraCommands.Visible = False
    
    frmEditor_Events.fraRandom(9).Visible = True
    frmEditor_Events.lstCommands.Visible = True
    fraDialogue.Visible = False
    fraCommand(2).Visible = False
End Sub

Private Sub cmdAddText_Ok_Click()
    If Not isEdit Then
        AddCommand EventType.evAddText
    Else
        EditCommand
    End If
    
    ' Hide
    fraDialogue.Visible = False
    fraCommand(2).Visible = False
    fraCommands.Visible = False
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdCancelCommand_Click()
    fraCommands.Visible = False
    frmEditor_Events.fraRandom(9).Visible = True
    frmEditor_Events.lstCommands.Visible = True
End Sub

Private Sub cmdChangeClass_Cancel_Click()
    If Not isEdit Then fraCommands.Visible = True Else fraCommands.Visible = False
    
    frmEditor_Events.fraRandom(9).Visible = True
    frmEditor_Events.lstCommands.Visible = True
    fraDialogue.Visible = False
    fraCommand(13).Visible = False
End Sub

Private Sub cmdChangeClass_Ok_Click()
    If isEdit = False Then
        AddCommand EventType.evChangeClass
    Else
        EditCommand
    End If
    
    ' Hide
    fraDialogue.Visible = False
    fraCommand(13).Visible = False
    fraCommands.Visible = False
End Sub

Private Sub cmdChangeItems_Cancel_Click()
    If Not isEdit Then fraCommands.Visible = True Else fraCommands.Visible = False
    
    frmEditor_Events.fraRandom(9).Visible = True
    frmEditor_Events.lstCommands.Visible = True
    fraDialogue.Visible = False
    fraCommand(10).Visible = False
End Sub

Private Sub cmdChangeItems_Ok_Click()
    If isEdit = False Then
        AddCommand EventType.evChangeItems
    Else
        EditCommand
    End If
    
    ' Hide
    fraDialogue.Visible = False
    fraCommands.Visible = False
    fraCommand(10).Visible = False
End Sub

Private Sub cmdChangeLevel_Cancel_Click()
    If Not isEdit Then fraCommands.Visible = True Else fraCommands.Visible = False
    
    frmEditor_Events.fraRandom(9).Visible = True
    frmEditor_Events.lstCommands.Visible = True
    fraDialogue.Visible = False
    fraCommand(11).Visible = False
End Sub

Private Sub cmdChangeLevel_Ok_Click()
    If isEdit = False Then
        AddCommand EventType.evChangeLevel
    Else
        EditCommand
    End If
    
    ' Hide
    fraDialogue.Visible = False
    fraCommand(11).Visible = False
    fraCommands.Visible = False
End Sub

Private Sub cmdChangePK_Cancel_Click()
    If Not isEdit Then fraCommands.Visible = True Else fraCommands.Visible = False
    
    frmEditor_Events.fraRandom(9).Visible = True
    frmEditor_Events.lstCommands.Visible = True
    fraDialogue.Visible = False
    fraCommand(16).Visible = False
End Sub

Private Sub cmdChangePK_Ok_Click()
    If isEdit = False Then
        AddCommand EventType.evChangePK
    Else
        EditCommand
    End If
    
    ' Hide
    fraDialogue.Visible = False
    fraCommand(16).Visible = False
    fraCommands.Visible = False
End Sub

Private Sub cmdChangeSex_Cancel_Click()
    If Not isEdit Then fraCommands.Visible = True Else fraCommands.Visible = False
    
    frmEditor_Events.fraRandom(9).Visible = True
    frmEditor_Events.lstCommands.Visible = True
    fraDialogue.Visible = False
    fraCommand(15).Visible = False
End Sub

Private Sub cmdChangeSex_Ok_Click()
    If isEdit = False Then
        AddCommand EventType.evChangeGender
    Else
        EditCommand
    End If
    
    ' Hide
    fraDialogue.Visible = False
    fraCommand(15).Visible = False
    fraCommands.Visible = False
End Sub

Private Sub cmdChangeSkills_Cancel_Click()
    If Not isEdit Then fraCommands.Visible = True Else fraCommands.Visible = False
    
    frmEditor_Events.fraRandom(9).Visible = True
    frmEditor_Events.lstCommands.Visible = True
    fraDialogue.Visible = False
    fraCommand(12).Visible = False
End Sub

Private Sub cmdChangeSkills_Ok_Click()
    If isEdit = False Then
        AddCommand EventType.evChangeSkills
    Else
        EditCommand
    End If
    
    ' Hide
    fraDialogue.Visible = False
    fraCommand(12).Visible = False
    fraCommands.Visible = False
End Sub

Private Sub cmdChangeSprite_Cancel_Click()
    If Not isEdit Then fraCommands.Visible = True Else fraCommands.Visible = False
    
    frmEditor_Events.fraRandom(9).Visible = True
    frmEditor_Events.lstCommands.Visible = True
    fraDialogue.Visible = False
    fraCommand(14).Visible = False
End Sub

Private Sub cmdChangeSprite_Ok_Click()
    If isEdit = False Then
        AddCommand EventType.evChangeSprite
    Else
        EditCommand
    End If
    
    ' Hide
    fraDialogue.Visible = False
    fraCommand(14).Visible = False
    fraCommands.Visible = False
End Sub

Private Sub cmdChoices_Cancel_Click()
    If Not isEdit Then fraCommands.Visible = True Else fraCommands.Visible = False
    
    frmEditor_Events.fraRandom(9).Visible = True
    frmEditor_Events.lstCommands.Visible = True
    fraDialogue.Visible = False
    fraCommand(1).Visible = False
End Sub

Private Sub cmdChoices_Ok_Click()
    If Not isEdit Then
        AddCommand EventType.evShowChoices
    Else
        EditCommand
    End If
    
    ' Hide
    fraDialogue.Visible = False
    fraCommand(1).Visible = False
    fraCommands.Visible = False
End Sub

Private Sub cmdClearCommand_Click()
    If MsgBox("Are you sure you want to clear all event commands?", vbYesNo, "Clear Event Commands?") = vbYes Then
        ClearEventCommands
    End If
End Sub

Private Sub cmdClearPage_Click()
    ZeroMemory ByVal VarPtr(tmpEvent.Pages(curPageNum)), LenB(tmpEvent.Pages(curPageNum))
End Sub

Private Sub cmdCommands_Click(Index As Integer)
    Dim I As Long, X As Long
    
    Select Case Index
        Case 0
            txtShowText.text = vbNullString
            fraDialogue.Visible = True
            fraCommand(0).Visible = True
            fraCommands.Visible = False
            scrlFace.Value = 0
        Case 1
            txtChoicePrompt.text = vbNullString
            txtChoices(1).text = vbNullString
            txtChoices(2).text = vbNullString
            txtChoices(3).text = vbNullString
            txtChoices(4).text = vbNullString
            fraDialogue.Visible = True
            fraCommand(1).Visible = True
            fraCommands.Visible = False
            scrlFace2.Value = 0
        Case 2
            txtAddText_Text.text = vbNullString
            scrlAddText_Colour.Value = 0
            optAddText_Player.Value = True
            fraDialogue.Visible = True
            fraCommand(2).Visible = True
            fraCommands.Visible = False
        Case 3
            txtChatbubbleText.text = ""
            optChatBubbleTarget(0).Value = True
            cmbChatBubbleTarget.Visible = False
            fraDialogue.Visible = True
            fraCommand(3).Visible = True
            fraCommands.Visible = False
      Case 4
            For I = 0 To 4
                txtVariableData(I).text = 0
            Next
            cmbVariable.ListIndex = 0
            optVariableAction(0).Value = True
            fraDialogue.Visible = True
            fraCommand(4).Visible = True
            fraCommands.Visible = False
        Case 5
            cmbPlayerSwitchSet.ListIndex = 0
            cmbSwitch.ListIndex = 0
            fraDialogue.Visible = True
            fraCommand(5).Visible = True
            fraCommands.Visible = False
        Case 6
            cmbSetSelfSwitch.ListIndex = 0
            cmbSetSelfSwitchTo.ListIndex = 0
            fraDialogue.Visible = True
            fraCommand(6).Visible = True
            fraCommands.Visible = False
        Case 7
            fraDialogue.Visible = True
            fraCommand(7).Visible = True
            optCondition_Index(0).Value = True
            ClearConditionFrame
            cmbCondition_PlayerVarIndex.Enabled = True
            cmbCondition_PlayerVarCompare.Enabled = True
            txtCondition_PlayerVarCondition.Enabled = True
            fraCommands.Visible = False
        Case 8
            AddCommand EventType.evExitProcess
            fraCommands.Visible = False
            fraDialogue.Visible = False
        Case 9
            txtLabelName.text = ""
            fraCommand(8).Visible = True
            fraCommands.Visible = False
            fraDialogue.Visible = True
        Case 10
            txtGotoLabel.text = ""
            fraCommand(9).Visible = True
            fraCommands.Visible = False
            fraDialogue.Visible = True
        Case 11
            cmbChangeItemIndex.ListIndex = 0
            optChangeItemSet.Value = True
            txtChangeItemsAmount.text = "0"
            fraDialogue.Visible = True
            fraCommand(10).Visible = True
            fraCommands.Visible = False
        Case 12
            AddCommand EventType.evRestoreHP
            fraCommands.Visible = False
            fraDialogue.Visible = False
        Case 13
            AddCommand EventType.evRestoreMP
            fraCommands.Visible = False
            fraDialogue.Visible = False
        Case 14
            AddCommand EventType.evLevelUp
            fraCommands.Visible = False
            fraDialogue.Visible = False
        Case 15
            scrlChangeLevel.Value = 1
            lblChangeLevel.Caption = "Level: 1"
            fraDialogue.Visible = True
            fraCommand(11).Visible = True
            fraCommands.Visible = False
        Case 16
            cmbChangeSkills.ListIndex = 0
            fraDialogue.Visible = True
            fraCommand(12).Visible = True
            fraCommands.Visible = False
        Case 17
            If MAX_CLASSES > 0 Then
                If cmbChangeClass.ListCount = 0 Then
                cmbChangeClass.Clear
                For I = 1 To MAX_CLASSES
                    cmbChangeClass.AddItem Trim$(Class(I).Name)
                Next
                cmbChangeClass.ListIndex = 0
                End If
            End If
            fraDialogue.Visible = True
            fraCommand(13).Visible = True
            fraCommands.Visible = False
        Case 18
            scrlChangeSprite.Value = 1
            lblChangeSprite.Caption = "Sprite: 1"
            fraDialogue.Visible = True
            fraCommand(14).Visible = True
            fraCommands.Visible = False
        Case 19
            optChangeSexMale.Value = True
            fraDialogue.Visible = True
            fraCommand(15).Visible = True
            fraCommands.Visible = False
        Case 20
            optChangePKYes.Value = True
            fraDialogue.Visible = True
            fraCommand(16).Visible = True
            fraCommands.Visible = False
        Case 21
            scrlGiveExp.Value = 0
            lblGiveExp.Caption = "Give Exp: 0"
            fraDialogue.Visible = True
            fraCommand(17).Visible = True
            fraCommands.Visible = False
        Case 22
            scrlWPMap.Value = 0
            scrlWPX.Value = 0
            scrlWPY.Value = 0
            cmbWarpPlayerDir.ListIndex = 0
            fraDialogue.Visible = True
            fraCommand(18).Visible = True
            fraCommands.Visible = False
        Case 23
            fraMoveRoute.Visible = True
            lstMoveRoute.Clear
            cmbEvent.Clear
            ReDim ListOfEvents(0 To Map.EventCount)
            ListOfEvents(0) = EditorEvent
            cmbEvent.AddItem "This Event"
            cmbEvent.ListIndex = 0
            cmbEvent.Enabled = True
            For I = 1 To Map.EventCount
                If I <> EditorEvent Then
                    cmbEvent.AddItem Trim$(Map.events(I).Name)
                    X = X + 1
                    ListOfEvents(X) = I
                End If
            Next
            IsMoveRouteCommand = True
            chkIgnoreMove.Value = 0
            chkRepeatRoute.Value = 0
            TempMoveRouteCount = 0
            ReDim TempMoveRoute(0)
            fraMoveRoute.Width = 841
            fraMoveRoute.Height = 585
            fraMoveRoute.Visible = True
            fraCommands.Visible = False
        Case 24
            cmbSpawnNPC.ListIndex = 0
            fraDialogue.Visible = True
            fraCommand(19).Visible = True
            fraCommands.Visible = False
        Case 25
            cmbPlayAnimEvent.Clear
            For I = 1 To Map.EventCount
                cmbPlayAnimEvent.AddItem I & ". " & Trim$(Map.events(I).Name)
            Next
            cmbPlayAnimEvent.ListIndex = 0
            optPlayAnimPlayer.Value = True
            cmbPlayAnim.ListIndex = 0
            lblPlayAnimX.Caption = "Map Tile X: 0"
            lblPlayAnimY.Caption = "Map Tile Y: 0"
            scrlPlayAnimTileX.Value = 0
            scrlPlayAnimTileY.Value = 0
            scrlPlayAnimTileX.max = Map.MaxX
            scrlPlayAnimTileY.max = Map.MaxY
            fraDialogue.Visible = True
            fraCommand(20).Visible = True
            fraCommands.Visible = False
            lblPlayAnimX.Visible = False
            lblPlayAnimY.Visible = False
            scrlPlayAnimTileX.Visible = False
            scrlPlayAnimTileY.Visible = False
            cmbPlayAnimEvent.Visible = False
        Case 26
            AddCommand EventType.evOpenBank
            fraCommands.Visible = False
            fraDialogue.Visible = False
        Case 27
            fraDialogue.Visible = True
            fraCommand(21).Visible = True
            cmbOpenShop.ListIndex = 0
            fraCommands.Visible = False
        Case 28
            AddCommand EventType.evFadeIn
            fraCommands.Visible = False
            fraDialogue.Visible = False
        Case 29
            AddCommand EventType.evFadeOut
            fraCommands.Visible = False
            fraDialogue.Visible = False
        Case 30
            AddCommand EventType.evFlashWhite
            fraCommands.Visible = False
            fraDialogue.Visible = False
        Case 31
            ScrlFogData(0).Value = 0
            ScrlFogData(0).Value = 0
            ScrlFogData(0).Value = 0
            fraDialogue.Visible = True
            fraCommand(22).Visible = True
            fraCommands.Visible = False
        Case 32
            CmbWeather.ListIndex = 0
            scrlWeatherIntensity.Value = 0
            fraDialogue.Visible = True
            fraCommand(23).Visible = True
            fraCommands.Visible = False
        Case 33
            scrlMapTintData(0).Value = 0
            scrlMapTintData(1).Value = 0
            scrlMapTintData(2).Value = 0
            scrlMapTintData(3).Value = 0
            fraDialogue.Visible = True
            fraCommand(24).Visible = True
            fraCommands.Visible = False
        Case 34
            cmbPlayBGM.ListIndex = 0
            fraDialogue.Visible = True
            fraCommand(25).Visible = True
            fraCommands.Visible = False
        Case 35
            AddCommand EventType.evFadeoutBGM
            fraCommands.Visible = False
            fraDialogue.Visible = False
        Case 36
            cmbPlaySound.ListIndex = 0
            fraDialogue.Visible = True
            fraCommand(26).Visible = True
            fraCommands.Visible = False
        Case 37
            AddCommand EventType.evStopSound
            fraCommands.Visible = False
            fraDialogue.Visible = False
        Case 38
            scrlWaitAmount.Value = 1
            fraDialogue.Visible = True
            fraCommand(27).Visible = True
            fraCommands.Visible = False
        Case 39
            cmbSetAccess.ListIndex = 0
            fraDialogue.Visible = True
            fraCommand(28).Visible = True
            fraCommands.Visible = False
        Case 40
            scrlCustomScript.Value = 1
            fraDialogue.Visible = True
            fraCommand(29).Visible = True
            fraCommands.Visible = False
    End Select
End Sub

Private Sub cmdCondition_Cancel_Click()
    If Not isEdit Then fraCommands.Visible = True Else fraCommands.Visible = False
    
    frmEditor_Events.fraRandom(9).Visible = True
    frmEditor_Events.lstCommands.Visible = True
    fraDialogue.Visible = False
    fraCommand(7).Visible = False
End Sub

Private Sub cmdCondition_Ok_Click()
    If isEdit = False Then
        AddCommand EventType.evCondition
    Else
        EditCommand
    End If
    
    ' Hide
    fraDialogue.Visible = False
    fraCommands.Visible = False
    fraCommand(7).Visible = False
End Sub

Private Sub cmdCopyPage_Click()
    CopyMemory ByVal VarPtr(copyPage), ByVal VarPtr(tmpEvent.Pages(curPageNum)), LenB(tmpEvent.Pages(curPageNum))
    cmdPastePage.Enabled = True
End Sub

Private Sub cmdCreateLabel_Cancel_Click()
    If Not isEdit Then fraCommands.Visible = True Else fraCommands.Visible = False
    
    frmEditor_Events.fraRandom(9).Visible = True
    frmEditor_Events.lstCommands.Visible = True
    fraDialogue.Visible = False
    fraCommand(8).Visible = False
End Sub

Private Sub cmdCreateLabel_Ok_Click()
    If isEdit = False Then
        AddCommand EventType.evLabel
    Else
        EditCommand
    End If
    
    ' Hide
    fraDialogue.Visible = False
    fraCommand(8).Visible = False
    fraCommands.Visible = False
End Sub

Private Sub cmdCustomScript_Cancel_Click()
    If Not isEdit Then fraCommands.Visible = True Else fraCommands.Visible = False
    
    frmEditor_Events.fraRandom(9).Visible = True
    frmEditor_Events.lstCommands.Visible = True
    fraDialogue.Visible = False
    fraCommand(29).Visible = False
End Sub

Private Sub cmdCustomScript_Ok_Click()
    If Not isEdit Then
        AddCommand EventType.evCustomScript
    Else
        EditCommand
    End If
    
    ' Hide
    fraDialogue.Visible = False
    fraCommand(29).Visible = False
    fraCommands.Visible = False
End Sub

Private Sub cmdDeleteCommand_Click()
    If lstCommands.ListIndex > -1 Then
        DeleteEventCommand
    End If
End Sub

Private Sub cmdDeletePage_Click()
    Dim I As Long
    
    ZeroMemory ByVal VarPtr(tmpEvent.Pages(curPageNum)), LenB(tmpEvent.Pages(curPageNum))
    
    ' Move everything else down a notch
    If curPageNum < tmpEvent.PageCount Then
        For I = curPageNum To tmpEvent.PageCount - 1
            CopyMemory ByVal VarPtr(tmpEvent.Pages(I)), ByVal VarPtr(tmpEvent.Pages(I + 1)), LenB(tmpEvent.Pages(I + 1))
        Next
    End If
    
    tmpEvent.PageCount = tmpEvent.PageCount - 1
    
    ' Set the tabs
    tabPages.Tabs.Clear
    For I = 1 To tmpEvent.PageCount
        tabPages.Tabs.Add , , str$(I)
    Next
    
    ' Set the tab back
    If curPageNum <= tmpEvent.PageCount Then
        tabPages.SelectedItem = tabPages.Tabs(curPageNum)
    Else
        tabPages.SelectedItem = tabPages.Tabs(tmpEvent.PageCount)
    End If
    
    ' Make sure we disable
    If tmpEvent.PageCount <= 1 Then
        cmdDeletePage.Enabled = False
    End If
End Sub

Private Sub cmdEditCommand_Click()
    Dim I As Long
    
    If lstCommands.ListIndex > -1 Then
        EditEventCommand
    End If
End Sub

Private Sub cmdGiveExp_Cancel_Click()
    If Not isEdit Then fraCommands.Visible = True Else fraCommands.Visible = False
    
    frmEditor_Events.fraRandom(9).Visible = True
    frmEditor_Events.lstCommands.Visible = True
    fraDialogue.Visible = False
    fraCommand(17).Visible = False
End Sub

Private Sub cmdGiveExp_Ok_Click()
    If isEdit = False Then
        AddCommand EventType.evGiveExp
    Else
        EditCommand
    End If
    
    ' Hide
    fraDialogue.Visible = False
    fraCommand(17).Visible = False
    fraCommands.Visible = False
End Sub

Private Sub cmdGotoLabel_Cancel_Click()
    If Not isEdit Then fraCommands.Visible = True Else fraCommands.Visible = False
    
    frmEditor_Events.fraRandom(9).Visible = True
    frmEditor_Events.lstCommands.Visible = True
    fraDialogue.Visible = False
    fraCommand(9).Visible = False
End Sub

Private Sub cmdGotoLabel_Ok_Click()
    If isEdit = False Then
        AddCommand EventType.evGotoLabel
    Else
        EditCommand
    End If
    
    ' Hide
    fraDialogue.Visible = False
    fraCommand(9).Visible = False
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
    Dim I As Long

    fraLabeling.Visible = True
    fraLabeling.Width = 849
    fraLabeling.Height = 593
    lstSwitches.Clear
    
    For I = 1 To MAX_SWITCHES
        lstSwitches.AddItem CStr(I) & ". " & Trim$(Switches(I))
    Next
    
    lstSwitches.ListIndex = 0
    lstVariables.Clear
    
    For I = 1 To MAX_VARIABLES
        lstVariables.AddItem CStr(I) & ". " & Trim$(Variables(I))
    Next
    
    lstVariables.ListIndex = 0
End Sub

Private Sub cmdMapTint_Cancel_Click()
    If Not isEdit Then fraCommands.Visible = True Else fraCommands.Visible = False
    
    frmEditor_Events.fraRandom(9).Visible = True
    frmEditor_Events.lstCommands.Visible = True
    fraDialogue.Visible = False
    fraCommand(24).Visible = False
End Sub

Private Sub cmdMapTint_Ok_Click()
    If Not isEdit Then
        AddCommand EventType.evSetTint
    Else
        EditCommand
    End If
    
    ' Hide
    fraDialogue.Visible = False
    fraCommand(24).Visible = False
    fraCommands.Visible = False
End Sub

Private Sub cmdMoveRoute_Click()
    Dim I As Long

    fraMoveRoute.Visible = True
    lstMoveRoute.Clear
    IsMoveRouteCommand = False
    
    cmbEvent.Clear
    cmbEvent.AddItem "This Event"
    cmbEvent.ListIndex = 0
    cmbEvent.Enabled = False
    
    chkIgnoreMove.Value = tmpEvent.Pages(curPageNum).IgnoreMoveRoute
    chkRepeatRoute.Value = tmpEvent.Pages(curPageNum).RepeatMoveRoute
    TempMoveRouteCount = tmpEvent.Pages(curPageNum).MoveRouteCount
    TempMoveRoute = tmpEvent.Pages(curPageNum).MoveRoute
    
    For I = 1 To TempMoveRouteCount
        Select Case TempMoveRoute(I).Index
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
    
    fraMoveRoute.Width = 841
    fraMoveRoute.Height = 585
    fraMoveRoute.Visible = True
End Sub

Sub PopulateMoveRouteList()
    Dim I As Long

    lstMoveRoute.Clear
    
    For I = 1 To TempMoveRouteCount
        Select Case TempMoveRoute(I).Index
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
    Dim PageCount As Long, I As Long
    
    PageCount = tmpEvent.PageCount + 1
    
    ' Redim the array
    ReDim Preserve tmpEvent.Pages(PageCount)
    tmpEvent.PageCount = PageCount
    
    ' Set the tabs
    tabPages.Tabs.Clear
    
    For I = 1 To tmpEvent.PageCount
        tabPages.Tabs.Add , , str$(I)
    Next
    
    cmdDeletePage.Enabled = True
End Sub

Public Sub cmdOK_Click()
    EventEditorSave
End Sub

Private Sub cmdOpenShop_Cancel_Click()
    If Not isEdit Then fraCommands.Visible = True Else fraCommands.Visible = False
    
    frmEditor_Events.fraRandom(9).Visible = True
    frmEditor_Events.lstCommands.Visible = True
    fraDialogue.Visible = False
    fraCommand(21).Visible = False
End Sub

Private Sub cmdOpenShop_Ok_Click()
    If Not isEdit Then
        AddCommand EventType.evOpenShop
    Else
        EditCommand
    End If
    
    ' Hide
    fraDialogue.Visible = False
    fraCommand(21).Visible = False
    fraCommands.Visible = False
End Sub

Private Sub cmdPastePage_Click()
    CopyMemory ByVal VarPtr(tmpEvent.Pages(curPageNum)), ByVal VarPtr(copyPage), LenB(tmpEvent.Pages(curPageNum))
    EventEditorLoadPage curPageNum
End Sub

Private Sub cmdPlayAnim_Cancel_Click()
    If Not isEdit Then fraCommands.Visible = True Else fraCommands.Visible = False
    
    frmEditor_Events.fraRandom(9).Visible = True
    frmEditor_Events.lstCommands.Visible = True
    fraDialogue.Visible = False
    fraCommand(20).Visible = False
End Sub

Private Sub cmdPlayAnim_Ok_Click()
    If Not isEdit Then
        AddCommand EventType.evPlayAnimation
    Else
        EditCommand
    End If
    
    ' Hide
    fraDialogue.Visible = False
    fraCommand(20).Visible = False
    fraCommands.Visible = False
End Sub

Private Sub cmdPlayBGM_Cancel_Click()
    If Not isEdit Then fraCommands.Visible = True Else fraCommands.Visible = False
    
    frmEditor_Events.fraRandom(9).Visible = True
    frmEditor_Events.lstCommands.Visible = True
    fraDialogue.Visible = False
    fraCommand(25).Visible = False
End Sub

Private Sub cmdPlayBGM_Ok_Click()
    If Not isEdit Then
        AddCommand EventType.evPlayBGM
    Else
        EditCommand
    End If
    
    ' Hide
    fraDialogue.Visible = False
    fraCommand(25).Visible = False
    fraCommands.Visible = False
End Sub

Private Sub cmdPlayerSwitch_Cancel_Click()
    If Not isEdit Then fraCommands.Visible = True Else fraCommands.Visible = False
    
    frmEditor_Events.fraRandom(9).Visible = True
    frmEditor_Events.lstCommands.Visible = True
    fraDialogue.Visible = False
    fraCommand(5).Visible = False
End Sub

Private Sub cmdPlaySound_Cancel_Click()
    If Not isEdit Then fraCommands.Visible = True Else fraCommands.Visible = False
    
    frmEditor_Events.fraRandom(9).Visible = True
    frmEditor_Events.lstCommands.Visible = True
    fraDialogue.Visible = False
    fraCommand(26).Visible = False
End Sub

Private Sub cmdPlaySound_Ok_Click()
    If Not isEdit Then
        AddCommand EventType.evPlaySound
    Else
        EditCommand
    End If
    
    ' Hide
    fraDialogue.Visible = False
    fraCommand(26).Visible = False
    fraCommands.Visible = False
End Sub

Private Sub cmdRename_Cancel_Click()
    Dim I As Long

    fraRenaming.Visible = False
    RenameType = 0
    RenameIndex = 0
    lstSwitches.Clear
    
    For I = 1 To MAX_SWITCHES
        lstSwitches.AddItem CStr(I) & ". " & Trim$(Switches(I))
    Next
    
    lstSwitches.ListIndex = 0
    lstVariables.Clear
    
    For I = 1 To MAX_VARIABLES
        lstVariables.AddItem CStr(I) & ". " & Trim$(Variables(I))
    Next
    
    lstVariables.ListIndex = 0
End Sub

Private Sub cmdRename_Ok_Click()
    Dim I As Long
    
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
    
    For I = 1 To MAX_SWITCHES
        lstSwitches.AddItem CStr(I) & ". " & Trim$(Switches(I))
    Next
    
    lstSwitches.ListIndex = 0
    lstVariables.Clear
    
    For I = 1 To MAX_VARIABLES
        lstVariables.AddItem CStr(I) & ". " & Trim$(Variables(I))
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

Private Sub cmdSelfSwitch_Cancel_Click()
    If Not isEdit Then fraCommands.Visible = True Else fraCommands.Visible = False
    
    frmEditor_Events.fraRandom(9).Visible = True
    frmEditor_Events.lstCommands.Visible = True
    fraDialogue.Visible = False
    fraCommand(6).Visible = False
End Sub

Private Sub cmdSelfSwitch_Ok_Click()
    If Not isEdit Then
        AddCommand EventType.evSelfSwitch
    Else
        EditCommand
    End If
    
    ' Hide
    fraDialogue.Visible = False
    fraCommand(6).Visible = False
    fraCommands.Visible = False
End Sub

Private Sub cmdSetAccess_Cancel_Click()
    If Not isEdit Then fraCommands.Visible = True Else fraCommands.Visible = False
    
    frmEditor_Events.fraRandom(9).Visible = True
    frmEditor_Events.lstCommands.Visible = True
    fraDialogue.Visible = False
    fraCommand(28).Visible = False
End Sub

Private Sub cmdSetAccess_Ok_Click()
    If Not isEdit Then
        AddCommand EventType.evSetAccess
    Else
        EditCommand
    End If
    
    ' Hide
    fraDialogue.Visible = False
    fraCommand(28).Visible = False
    fraCommands.Visible = False
End Sub

Private Sub cmdSetFog_Cancel_Click()
    If Not isEdit Then fraCommands.Visible = True Else fraCommands.Visible = False
    
    frmEditor_Events.fraRandom(9).Visible = True
    frmEditor_Events.lstCommands.Visible = True
    fraDialogue.Visible = False
    fraCommand(22).Visible = False
End Sub

Private Sub cmdSetFog_Ok_Click()
    If Not isEdit Then
        AddCommand EventType.evSetFog
    Else
        EditCommand
    End If
    
    ' Hide
    fraDialogue.Visible = False
    fraCommand(22).Visible = False
    fraCommands.Visible = False
End Sub

Private Sub cmdSetWeather_Cancel_Click()
    If Not isEdit Then fraCommands.Visible = True Else fraCommands.Visible = False
    
    frmEditor_Events.fraRandom(9).Visible = True
    frmEditor_Events.lstCommands.Visible = True
    fraDialogue.Visible = False
    fraCommand(23).Visible = False
End Sub

Private Sub cmdSetWeather_Ok_Click()
    If Not isEdit Then
        AddCommand EventType.evSetWeather
    Else
        EditCommand
    End If
    
    ' Hide
    fraDialogue.Visible = False
    fraCommand(23).Visible = False
    fraCommands.Visible = False
End Sub

Private Sub cmdShowChatBubble_Cancel_Click()
    If Not isEdit Then fraCommands.Visible = True Else fraCommands.Visible = False
    
    frmEditor_Events.fraRandom(9).Visible = True
    frmEditor_Events.lstCommands.Visible = True
    fraDialogue.Visible = False
    fraCommand(3).Visible = False
End Sub

Private Sub cmdShowChatBubble_Ok_Click()
    If Not isEdit Then
        AddCommand EventType.evShowChatBubble
    Else
        EditCommand
    End If
    
    ' Hide
    fraDialogue.Visible = False
    fraCommand(3).Visible = False
    fraCommands.Visible = False
End Sub

Private Sub cmdShowText_Cancel_Click()
    If Not isEdit Then fraCommands.Visible = True Else fraCommands.Visible = False
    
    frmEditor_Events.fraRandom(9).Visible = True
    frmEditor_Events.lstCommands.Visible = True
    fraDialogue.Visible = False
    fraCommand(0).Visible = False
End Sub

Private Sub cmdShowText_Ok_Click()
    If Not isEdit Then
        AddCommand EventType.evShowText
    Else
        EditCommand
    End If
    
    ' Hide
    fraDialogue.Visible = False
    fraCommand(0).Visible = False
    fraCommands.Visible = False
End Sub

Private Sub cmdSpawnNPC_Cancel_Click()
    If Not isEdit Then fraCommands.Visible = True Else fraCommands.Visible = False
    
    frmEditor_Events.fraRandom(9).Visible = True
    frmEditor_Events.lstCommands.Visible = True
    fraDialogue.Visible = False
    fraCommand(19).Visible = False
End Sub

Private Sub cmdSpawnNPC_Ok_Click()
    If isEdit = False Then
        AddCommand EventType.evSpawnNPC
    Else
        EditCommand
    End If
    
    ' Hide
    fraDialogue.Visible = False
    fraCommand(19).Visible = False
    fraCommands.Visible = False
End Sub

Private Sub cmdVariableCancel_Click()
    If Not isEdit Then fraCommands.Visible = True Else fraCommands.Visible = False
    
    frmEditor_Events.fraRandom(9).Visible = True
    frmEditor_Events.lstCommands.Visible = True
    fraDialogue.Visible = False
    fraCommand(4).Visible = False
End Sub

Private Sub cmdVariableOK_Click()
    If Not isEdit Then
        AddCommand EventType.evPlayerVar
    Else
        EditCommand
    End If
    
    ' Hide
    fraDialogue.Visible = False
    fraCommand(4).Visible = False
    fraCommands.Visible = False
End Sub

Private Sub cmdWait_Cancel_Click()
    If Not isEdit Then fraCommands.Visible = True Else fraCommands.Visible = False
    
    frmEditor_Events.fraRandom(9).Visible = True
    frmEditor_Events.lstCommands.Visible = True
    fraDialogue.Visible = False
    fraCommand(27).Visible = False
End Sub

Private Sub cmdWait_Ok_Click()
    If Not isEdit Then
        AddCommand EventType.evWait
    Else
        EditCommand
    End If
    
    ' Hide
    fraDialogue.Visible = False
    fraCommand(27).Visible = False
    fraCommands.Visible = False
End Sub

Private Sub cmdWPCancel_Click()
    If Not isEdit Then fraCommands.Visible = True Else fraCommands.Visible = False
    
    frmEditor_Events.fraRandom(9).Visible = True
    frmEditor_Events.lstCommands.Visible = True
    fraDialogue.Visible = False
    fraCommand(18).Visible = False
End Sub

Private Sub cmdWPOkay_Click()
    If Not isEdit Then
        AddCommand EventType.evWarpPlayer
    Else
        EditCommand
    End If
    
    ' Hide
    fraDialogue.Visible = False
    fraCommand(18).Visible = False
    fraCommands.Visible = False
End Sub

Public Sub InitEventEditorForm()
    Dim I As Long

    If frmAdmin.Visible Then
        frmEditor_Events.Move frmAdmin.Left - frmEditor_Events.Width, frmAdmin.Top
     Else
        frmEditor_Events.Move frmMain.Left + frmMain.Width - frmEditor_Events.Width, frmMain.Top
    End If
    
    ' Level
    scrlChangeLevel.min = 1
    scrlChangeLevel.Value = 1
    lblChangeLevel.Caption = "Level: 1"
    
    ' Switches
    cmbSwitch.Clear
    
    For I = 1 To MAX_SWITCHES
        cmbSwitch.AddItem I & ". " & Switches(I)
    Next
    cmbSwitch.ListIndex = 0
    
    ' Variables
    cmbVariable.Clear
    
    For I = 1 To MAX_VARIABLES
        cmbVariable.AddItem I & ". " & Variables(I)
    Next
    cmbVariable.ListIndex = 0
    
    ' Items
    cmbChangeItemIndex.Clear
    
    For I = 1 To MAX_ITEMS
        cmbChangeItemIndex.AddItem Trim$(Item(I).Name)
    Next
    cmbChangeItemIndex.ListIndex = 0
    
    ' Spells
    cmbChangeSkills.Clear
    
    For I = 1 To MAX_SPELLS
        cmbChangeSkills.AddItem Trim$(Spell(I).Name)
    Next
    cmbChangeSkills.ListIndex = 0
    
    ' Classes
    cmbChangeClass.Clear
    
    For I = 1 To MAX_CLASSES
        cmbChangeClass.AddItem I & ": " & Trim$(Class(I).Name)
    Next
    cmbChangeClass.ListIndex = 0
    
    ' Animations
    cmbPlayAnim.Clear
    
    For I = 1 To MAX_ANIMATIONS
        cmbPlayAnim.AddItem I & ": " & Trim$(Animation(I).Name)
    Next
    cmbPlayAnim.ListIndex = 0
    
    ' BGM
    cmbPlayBGM.Clear
    
    For I = 1 To UBound(MusicCache)
        cmbPlayBGM.AddItem (MusicCache(I))
    Next
    cmbPlayBGM.ListIndex = 0
    
    ' Sounds
    frmEditor_Events.cmbPlaySound.Clear

    For I = 1 To UBound(SoundCache)
        frmEditor_Events.cmbPlaySound.AddItem SoundCache(I)
    Next
    frmEditor_Events.cmbPlaySound.ListIndex = 0
    
    ' Shops
    cmbOpenShop.Clear
    
    For I = 1 To MAX_SHOPS
        cmbOpenShop.AddItem I & ". " & Trim$(Shop(I).Name)
    Next
    cmbOpenShop.ListIndex = 0
    
    ' NPCs
    cmbSpawnNPC.Clear
    
    For I = 1 To MAX_MAP_NPCS
        If Map.NPC(I) > 0 Then
            cmbSpawnNPC.AddItem I & ". " & Trim$(NPC(Map.NPC(I)).Name)
        Else
            cmbSpawnNPC.AddItem I & ". "
        End If
    Next
    cmbSpawnNPC.ListIndex = 0
    
    ' Fogs
    ScrlFogData(0).max = NumFogs
End Sub

Private Sub Form_Load()
    scrlFace.max = NumFaces
    scrlFace2.max = NumFaces
    scrlChangeSprite.max = NumCharacters
    scrlChangeLevel.max = MAX_LEVEL
    txtName.MaxLength = NAME_LENGTH
    txtSearch.MaxLength = NAME_LENGTH
    txtLabelName.MaxLength = NAME_LENGTH
    txtGotoLabel.MaxLength = NAME_LENGTH
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmAdmin.chkEditor(EDITOR_EVENTS).Value = False
    BringWindowToTop (frmAdmin.hWnd)
    Editor = 0
End Sub

Private Sub lstCommands_Click()
    curCommand = lstCommands.ListIndex + 1
End Sub

Sub AddMoveRouteCommand(Index As Integer)
    Dim I As Long, X As Long, Z As Long
    
    Index = Index + 1
    
    If lstMoveRoute.ListIndex > -1 Then
        I = lstMoveRoute.ListIndex + 1
        TempMoveRouteCount = TempMoveRouteCount + 1
        ReDim Preserve TempMoveRoute(TempMoveRouteCount)
        
        For X = TempMoveRouteCount - 1 To I Step -1
            TempMoveRoute(X + 1) = TempMoveRoute(X)
        Next
        
        TempMoveRoute(I).Index = Index
        
        ' If set graphic then
        If Index = 43 Then
            TempMoveRoute(I).Data1 = cmbGraphic.ListIndex
            TempMoveRoute(I).Data2 = scrlGraphic.Value
            TempMoveRoute(I).Data3 = GraphicSelX
            TempMoveRoute(I).Data4 = GraphicSelX2
            TempMoveRoute(I).Data5 = GraphicSelY
            TempMoveRoute(I).Data6 = GraphicSelY2
        End If
        PopulateMoveRouteList
    Else
        TempMoveRouteCount = TempMoveRouteCount + 1
        ReDim Preserve TempMoveRoute(TempMoveRouteCount)
        TempMoveRoute(TempMoveRouteCount).Index = Index
        PopulateMoveRouteList
        
        ' If set graphic then
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
    Dim I As Long
    
    Index = Index + 1
    
    If Index > 0 And Index <= TempMoveRouteCount Then
        For I = Index + 1 To TempMoveRouteCount
            TempMoveRoute(I - 1) = TempMoveRoute(I)
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
        ' Remove move route command lol
        cmdDeleteCommand_Click
    End If
End Sub

Private Sub lstIndex_Click()
    EventEditorInit EditorIndex
End Sub

Private Sub lstMoveRoute_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then
        ' Remove move route command lol
        If lstMoveRoute.ListIndex > -1 Then
            Call RemoveMoveRouteCommand(lstMoveRoute.ListIndex)
        End If
    End If
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

Private Sub optChatBubbleTarget_Click(Index As Integer)
    Dim I As Long

    If Index = 0 Then
        cmbChatBubbleTarget.Visible = False
    ElseIf Index = 1 Then
        cmbChatBubbleTarget.Visible = True
        cmbChatBubbleTarget.Clear
        For I = 1 To MAX_MAP_NPCS
            If Map.NPC(I) <= 0 Then
                cmbChatBubbleTarget.AddItem CStr(I) & ". "
            Else
                cmbChatBubbleTarget.AddItem CStr(I) & ". " & Trim$(NPC(Map.NPC(I)).Name)
            End If
        Next
        cmbChatBubbleTarget.ListIndex = 0
    ElseIf Index = 2 Then
        cmbChatBubbleTarget.Visible = True
        cmbChatBubbleTarget.Clear
        For I = 1 To Map.EventCount
            cmbChatBubbleTarget.AddItem CStr(I) & ". " & Trim$(Map.events(I).Name)
        Next
        cmbChatBubbleTarget.ListIndex = 0
    End If
End Sub

Private Sub optCondition_Index_Click(Index As Integer)
Dim I As Long, X As Long
    For I = 0 To 6
        If optCondition_Index(I).Value = True Then X = I
    Next
    ClearConditionFrame
    Select Case X
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
Dim I As Long
    cmbCondition_PlayerVarIndex.Enabled = False
    cmbCondition_PlayerVarIndex.Clear
    For I = 1 To MAX_VARIABLES
        cmbCondition_PlayerVarIndex.AddItem I & ". " & Variables(I)
    Next
    cmbCondition_PlayerVarIndex.ListIndex = 0
    
    cmbCondition_PlayerVarCompare.ListIndex = 0
    cmbCondition_PlayerVarCompare.Enabled = False
    
    txtCondition_PlayerVarCondition.Enabled = False
    txtCondition_PlayerVarCondition.text = "0"
    
    cmbCondition_PlayerSwitch.Enabled = False
    cmbCondition_PlayerSwitch.Clear
    For I = 1 To MAX_SWITCHES
        cmbCondition_PlayerSwitch.AddItem I & ". " & Switches(I)
    Next
    cmbCondition_PlayerSwitch.ListIndex = 0
    
    cmbCondtion_PlayerSwitchCondition.Enabled = False
    cmbCondtion_PlayerSwitchCondition.ListIndex = 0
    
    cmbCondition_HasItem.Enabled = False
    cmbCondition_HasItem.Clear
    For I = 1 To MAX_ITEMS
        cmbCondition_HasItem.AddItem I & ". " & Trim$(Item(I).Name)
    Next
    cmbCondition_HasItem.ListIndex = 0
    
    cmbCondition_ClassIs.Enabled = False
    cmbCondition_ClassIs.Clear
    For I = 1 To MAX_CLASSES
        cmbCondition_ClassIs.AddItem I & ". " & CStr(Class(I).Name)
    Next
    cmbCondition_ClassIs.ListIndex = 0
    
    cmbCondition_LearntSkill.Enabled = False
    cmbCondition_LearntSkill.Clear
    For I = 1 To MAX_SPELLS
        cmbCondition_LearntSkill.AddItem I & ". " & Trim$(Spell(I).Name)
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

Private Sub optVariableAction_Click(Index As Integer)
    Dim I As Long
    For I = 0 To 3
        If optVariableAction(I).Value = True Then
            Exit For
        End If
    Next
    txtVariableData(0).Enabled = False
    txtVariableData(0).text = "0"
    txtVariableData(1).Enabled = False
    txtVariableData(1).text = "0"
    txtVariableData(2).Enabled = False
    txtVariableData(2).text = "0"
    txtVariableData(3).Enabled = False
    txtVariableData(3).text = "0"
    txtVariableData(4).Enabled = False
    txtVariableData(4).text = "0"
    Select Case I
        Case 0
            txtVariableData(0).Enabled = True
        Case 1
            txtVariableData(1).Enabled = True
        Case 2
            txtVariableData(2).Enabled = True
        Case 3
            txtVariableData(3).Enabled = True
            txtVariableData(4).Enabled = True
    End Select
End Sub

Private Sub picGraphic_Click()
    fraGraphic.Width = 841
    fraGraphic.Height = 585
    fraGraphic.Visible = True
    GraphicSelType = 0
End Sub

Private Sub picGraphicSel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim I As Long
    If frmEditor_Events.cmbGraphic.ListIndex = 2 Then
        'Tileset... hard one....
        If ShiftDown Then
            If GraphicSelX > -1 And GraphicSelY > -1 Then
                If CLng(X + frmEditor_Events.hScrlGraphicSel.Value) / 32 > GraphicSelX And CLng(Y + frmEditor_Events.vScrlGraphicSel.Value) / 32 > GraphicSelY Then
                    GraphicSelX2 = CLng(X + frmEditor_Events.hScrlGraphicSel.Value) / 32
                    GraphicSelY2 = CLng(Y + frmEditor_Events.vScrlGraphicSel.Value) / 32
                End If
            End If
        Else
            GraphicSelX = CLng(X + frmEditor_Events.hScrlGraphicSel.Value) \ 32
            GraphicSelY = CLng(Y + frmEditor_Events.vScrlGraphicSel.Value) \ 32
            GraphicSelX2 = 0
            GraphicSelY2 = 0
        End If
    ElseIf frmEditor_Events.cmbGraphic.ListIndex = 1 Then
        GraphicSelX = CLng(X + frmEditor_Events.hScrlGraphicSel.Value)
        GraphicSelY = CLng(Y + frmEditor_Events.vScrlGraphicSel.Value)
        GraphicSelX2 = 0
        GraphicSelY2 = 0
        
        If frmEditor_Events.scrlGraphic.Value <= 0 Or frmEditor_Events.scrlGraphic.Value > NumCharacters Then Exit Sub
        
        
        For I = 0 To 3
            If GraphicSelX >= ((Tex_Character(frmEditor_Events.scrlGraphic.Value).Width / 4) * I) And GraphicSelX < ((Tex_Character(frmEditor_Events.scrlGraphic.Value).Width / 4) * (I + 1)) Then
                GraphicSelX = I
            End If
        Next
        
        For I = 0 To 3
            If GraphicSelY >= ((Tex_Character(frmEditor_Events.scrlGraphic.Value).Height / 4) * I) And GraphicSelY < ((Tex_Character(frmEditor_Events.scrlGraphic.Value).Height / 4) * (I + 1)) Then
                GraphicSelY = I
            End If
        Next
        
        
    End If
End Sub

Private Sub scrlGraphic_Click()
    lblGraphic.Caption = "Graphic: " & scrlGraphic.Value
    tmpEvent.Pages(curPageNum).Graphic = scrlGraphic.Value
    
    If tmpEvent.Pages(curPageNum).GraphicType = 1 Then
        If frmEditor_Events.scrlGraphic.Value <= 0 Or frmEditor_Events.scrlGraphic.Value > NumCharacters Then Exit Sub
   
        If Tex_Character(frmEditor_Events.scrlGraphic.Value).Width > 793 Then
            frmEditor_Events.hScrlGraphicSel.Visible = True
            frmEditor_Events.hScrlGraphicSel.Value = 0
            frmEditor_Events.hScrlGraphicSel.max = Tex_Character(frmEditor_Events.scrlGraphic.Value).Width - 800
        Else
            frmEditor_Events.hScrlGraphicSel.Visible = False
        End If
                    
        If Tex_Character(frmEditor_Events.scrlGraphic.Value).Height > 472 Then
            frmEditor_Events.vScrlGraphicSel.Visible = True
            frmEditor_Events.vScrlGraphicSel.Value = 0
            frmEditor_Events.vScrlGraphicSel.max = Tex_Character(frmEditor_Events.scrlGraphic.Value).Height - 512
        Else
            frmEditor_Events.vScrlGraphicSel.Visible = False
        End If
    ElseIf tmpEvent.Pages(curPageNum).GraphicType = 2 Then
        If frmEditor_Events.scrlGraphic.Value <= 0 Or frmEditor_Events.scrlGraphic.Value > NumTileSets Then Exit Sub
                    
        If Tex_Tileset(frmEditor_Events.scrlGraphic.Value).Width > 793 Then
            frmEditor_Events.hScrlGraphicSel.Visible = True
            frmEditor_Events.hScrlGraphicSel.Value = 0
            frmEditor_Events.hScrlGraphicSel.max = Tex_Tileset(frmEditor_Events.scrlGraphic.Value).Width - 800
        Else
            frmEditor_Events.hScrlGraphicSel.Visible = False
        End If
                    
        If Tex_Tileset(frmEditor_Events.scrlGraphic.Value).Height > 472 Then
            frmEditor_Events.vScrlGraphicSel.Visible = True
            frmEditor_Events.vScrlGraphicSel.Value = 0
            frmEditor_Events.vScrlGraphicSel.max = Tex_Tileset(frmEditor_Events.scrlGraphic.Value).Height - 512
        Else
            frmEditor_Events.vScrlGraphicSel.Visible = False
        End If
    End If
End Sub

Private Sub scrlAddText_Colour_Click()
    frmEditor_Events.lblAddText_Colour.Caption = "Color: " & GetColorName(frmEditor_Events.scrlAddText_Colour.Value)
End Sub

Private Sub scrlAddText_Colour_Change()
    frmEditor_Events.lblAddText_Colour.Caption = "Color: " & GetColorName(frmEditor_Events.scrlAddText_Colour.Value)
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

Private Sub scrlFace_Change()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lblFace.Caption = "Face: " & scrlFace.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlFace_Change", "frmEditor_Events", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlFace2_Change()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lblFace2.Caption = "Face: " & scrlFace2.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlFace2_Change", "frmEditor_Events", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub ScrlFogData_Change(Index As Integer)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Select Case Index
        Case 0
            If ScrlFogData(0).Value = 0 Then
                lblFogData(0).Caption = "None."
            Else
                lblFogData(0).Caption = "Fog: " & ScrlFogData(0).Value
            End If
        Case 1
            lblFogData(1).Caption = "Fog Speed: " & ScrlFogData(1).Value
        Case 2
            lblFogData(2).Caption = "Fog Opacity: " & ScrlFogData(2).Value
    End Select
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "ScrlFogData_Change(" & CStr(Index) & ")", "frmEditor_Events", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlGiveExp_Change()
    lblGiveExp.Caption = "Experience: " & scrlGiveExp.Value
End Sub

Private Sub scrlGraphic_Change()
    If scrlGraphic.Value = 0 Then
        lblGraphic.Caption = "Number: None"
    Else
        lblGraphic.Caption = "Number: " & scrlGraphic.Value
    End If
    Call cmbGraphic_Click
End Sub

Private Sub scrlMapTintData_Change(Index As Integer)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Select Case Index
        Case 0
            lblMapTintData(0).Caption = "Red: " & scrlMapTintData(0).Value
        Case 1
            lblMapTintData(1).Caption = "Green: " & scrlMapTintData(1).Value
        Case 2
            lblMapTintData(2).Caption = "Blue: " & scrlMapTintData(2).Value
        Case 3
            lblMapTintData(3).Caption = "Alpha: " & scrlMapTintData(3).Value
    End Select
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "ScrlMapTintData_Change(" & CStr(Index) & ")", "frmEditor_MapProperties", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlPlayAnimTileX_Change()
    lblPlayAnimX.Caption = "Map Tile X: " & scrlPlayAnimTileX.Value
End Sub

Private Sub scrlPlayAnimTileY_Change()
    lblPlayAnimY.Caption = "Map Tile Y: " & scrlPlayAnimTileY.Value
End Sub

Private Sub scrlWaitAmount_Change()
    lblWaitAmount.Caption = "Wait: " & scrlWaitAmount.Value & " ms"
End Sub

Private Sub scrlWeatherIntensity_Change()
    lblWeatherIntensity.Caption = "Intensity: " & scrlWeatherIntensity.Value
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
    Dim I As Long

    For I = 1 To 2
        picCommands(I).Visible = False
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
    tmpEvent.Pages(curPageNum).VariableCondition = val(Trim$(txtPlayerVariable.text))
End Sub

Private Sub txtSearch_Change()
    Dim I As Long, Find As String
    
    For I = 0 To lstIndex.ListCount - 1
        Find = Trim$(I + 1 & ": " & txtSearch.text)
        
        ' Make sure we dont try to check a name that's too small
        If Len(lstIndex.List(I)) = Len(Find) Then
            If UCase$(Mid$(Trim$(lstIndex.List(I)), 1, Len(Find))) = UCase$(Find) Then
                lstIndex.ListIndex = I
                Exit For
            End If
        End If
    Next
End Sub

Private Sub txtSearch_GotFocus()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    txtSearch.SelStart = Len(txtSearch)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "txtSearch_GotFocus", "frmEditor_Events", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Dim buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If fraDialogue.Visible Then Exit Sub
    
    If KeyAscii = vbKeyReturn Then
        cmdOK_Click
        KeyAscii = 0
    ElseIf KeyAscii = vbKeyEscape Then
        cmdCancel_Click
        KeyAscii = 0
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "Form_KeyPress", "frmEditor_Events", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub
