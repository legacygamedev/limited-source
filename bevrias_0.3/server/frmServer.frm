VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmServer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bevrias Server"
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   10575
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmServer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   289
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   705
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer4 
      Interval        =   10
      Left            =   8040
      Top             =   4320
   End
   Begin VB.Timer Timer3 
      Interval        =   100
      Left            =   7560
      Top             =   4320
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   7080
      Top             =   4320
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4290
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   10320
      _ExtentX        =   18203
      _ExtentY        =   7567
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   344
      TabMaxWidth     =   2646
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Chat"
      TabPicture(0)   =   "frmServer.frx":038A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label6"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "SSTab2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "picCMsg"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "CustomMsg(5)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "CustomMsg(4)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "CustomMsg(3)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "CustomMsg(2)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "CustomMsg(1)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Frame5"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "CustomMsg(0)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Say(0)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Say(1)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Say(2)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Say(3)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Say(4)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Say(5)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Command29"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).ControlCount=   17
      TabCaption(1)   =   "Players"
      TabPicture(1)   =   "frmServer.frx":03A6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Timer5"
      Tab(1).Control(1)=   "Picture1"
      Tab(1).Control(2)=   "Command45"
      Tab(1).Control(3)=   "Command3"
      Tab(1).Control(4)=   "Command24"
      Tab(1).Control(5)=   "Command23"
      Tab(1).Control(6)=   "Command22"
      Tab(1).Control(7)=   "Command21"
      Tab(1).Control(8)=   "Command20"
      Tab(1).Control(9)=   "Command19"
      Tab(1).Control(10)=   "Command18"
      Tab(1).Control(11)=   "Command17"
      Tab(1).Control(12)=   "Command16"
      Tab(1).Control(13)=   "Command15"
      Tab(1).Control(14)=   "Command14"
      Tab(1).Control(15)=   "Command13"
      Tab(1).Control(16)=   "Command66"
      Tab(1).Control(17)=   "Check1"
      Tab(1).Control(18)=   "picReason"
      Tab(1).Control(19)=   "picStats"
      Tab(1).Control(20)=   "picJail"
      Tab(1).Control(21)=   "lvUsers"
      Tab(1).Control(22)=   "TPO"
      Tab(1).ControlCount=   23
      TabCaption(2)   =   "Control Panel"
      TabPicture(2)   =   "frmServer.frx":03C2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Timer1"
      Tab(2).Control(1)=   "tmrChatLogs"
      Tab(2).Control(2)=   "tmrPlayerSave"
      Tab(2).Control(3)=   "tmrSpawnMapItems"
      Tab(2).Control(4)=   "tmrGameAI"
      Tab(2).Control(5)=   "tmrShutdown"
      Tab(2).Control(6)=   "PlayerTimer"
      Tab(2).Control(7)=   "picExp"
      Tab(2).Control(8)=   "picWeather"
      Tab(2).Control(9)=   "picWarp"
      Tab(2).Control(10)=   "picMap"
      Tab(2).Control(11)=   "Frame6"
      Tab(2).Control(12)=   "Frame25"
      Tab(2).Control(13)=   "Frame2"
      Tab(2).Control(14)=   "Frame9"
      Tab(2).Control(15)=   "Socket(0)"
      Tab(2).Control(16)=   "lblPort"
      Tab(2).Control(17)=   "lblIP"
      Tab(2).ControlCount=   18
      TabCaption(3)   =   "Help"
      TabPicture(3)   =   "frmServer.frx":03DE
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lstTopics"
      Tab(3).Control(1)=   "TopicTitle"
      Tab(3).Control(2)=   "CharInfo(21)"
      Tab(3).Control(3)=   "CharInfo(22)"
      Tab(3).Control(4)=   "CharInfo(23)"
      Tab(3).ControlCount=   5
      TabCaption(4)   =   "Options"
      TabPicture(4)   =   "frmServer.frx":03FA
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame1"
      Tab(4).Control(1)=   "Frame10"
      Tab(4).Control(2)=   "Data"
      Tab(4).Control(3)=   "Frame3"
      Tab(4).ControlCount=   4
      Begin VB.Timer Timer5 
         Interval        =   50
         Left            =   -71040
         Top             =   3840
      End
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   -66240
         Top             =   22
      End
      Begin VB.Timer tmrChatLogs 
         Interval        =   1000
         Left            =   -65760
         Top             =   0
      End
      Begin VB.Timer tmrPlayerSave 
         Interval        =   60000
         Left            =   -65520
         Top             =   240
      End
      Begin VB.Timer tmrSpawnMapItems 
         Interval        =   1000
         Left            =   -66720
         Top             =   0
      End
      Begin VB.Timer tmrGameAI 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   -67200
         Top             =   0
      End
      Begin VB.Timer tmrShutdown 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   -67680
         Top             =   0
      End
      Begin VB.Timer PlayerTimer 
         Enabled         =   0   'False
         Interval        =   5000
         Left            =   -65280
         Top             =   120
      End
      Begin VB.CommandButton Command29 
         Caption         =   "Server Chat Name"
         Height          =   255
         Left            =   8640
         TabIndex        =   227
         Top             =   3830
         Width           =   1455
      End
      Begin VB.PictureBox picExp 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1335
         Left            =   -68040
         ScaleHeight     =   87
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   215
         TabIndex        =   213
         Top             =   227
         Width           =   3255
         Visible         =   0   'False
         Begin VB.CommandButton Command39 
            Caption         =   "Cansel"
            Height          =   255
            Left            =   1560
            TabIndex        =   220
            Top             =   960
            Width           =   1575
         End
         Begin VB.CommandButton Command40 
            Caption         =   "Execute"
            Height          =   255
            Left            =   1560
            TabIndex        =   215
            Top             =   720
            Width           =   1575
         End
         Begin VB.TextBox txtExp 
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
            TabIndex        =   214
            Top             =   360
            Width           =   2955
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Experience:"
            Height          =   195
            Left            =   120
            TabIndex        =   216
            Top             =   120
            Width           =   855
         End
      End
      Begin VB.PictureBox picWeather 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2055
         Left            =   -68040
         ScaleHeight     =   135
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   215
         TabIndex        =   206
         Top             =   227
         Width           =   3255
         Visible         =   0   'False
         Begin VB.CommandButton Command65 
            Caption         =   "Snow"
            Height          =   255
            Left            =   1680
            TabIndex        =   219
            Top             =   1320
            Width           =   1335
         End
         Begin VB.CommandButton Command64 
            Caption         =   "Rain"
            Height          =   255
            Left            =   240
            TabIndex        =   218
            Top             =   1320
            Width           =   1335
         End
         Begin VB.CommandButton Command61 
            Caption         =   "Cancel"
            Height          =   255
            Left            =   1680
            TabIndex        =   210
            Top             =   1680
            Width           =   1335
         End
         Begin VB.HScrollBar scrlRainIntensity 
            Height          =   255
            Left            =   120
            Max             =   50
            Min             =   1
            TabIndex        =   209
            Top             =   360
            Value           =   25
            Width           =   2895
         End
         Begin VB.CommandButton Command62 
            Caption         =   "None"
            Height          =   255
            Left            =   240
            TabIndex        =   208
            Top             =   1080
            Width           =   1335
         End
         Begin VB.CommandButton Command63 
            Caption         =   "Thunder"
            Height          =   255
            Left            =   1680
            TabIndex        =   207
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label lblRainIntensity 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Intensity: 25"
            Height          =   195
            Left            =   120
            TabIndex        =   212
            Top             =   120
            Width           =   930
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Current Weather: None"
            Height          =   195
            Left            =   120
            TabIndex        =   211
            Top             =   720
            Width           =   1710
         End
      End
      Begin VB.PictureBox picWarp 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2535
         Left            =   -68160
         ScaleHeight     =   167
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   223
         TabIndex        =   198
         Top             =   227
         Width           =   3375
         Visible         =   0   'False
         Begin VB.CommandButton Command38 
            Caption         =   "Cancel"
            Height          =   255
            Left            =   1680
            TabIndex        =   217
            Top             =   2160
            Width           =   1575
         End
         Begin VB.CommandButton Command37 
            Caption         =   "Warp"
            Height          =   255
            Left            =   1680
            TabIndex        =   202
            Top             =   1920
            Width           =   1575
         End
         Begin VB.HScrollBar scrlMM 
            Height          =   255
            Left            =   120
            Min             =   1
            TabIndex        =   201
            Top             =   360
            Value           =   1
            Width           =   3135
         End
         Begin VB.HScrollBar scrlMX 
            Height          =   255
            Left            =   120
            TabIndex        =   200
            Top             =   960
            Width           =   3135
         End
         Begin VB.HScrollBar scrlMY 
            Height          =   255
            Left            =   120
            TabIndex        =   199
            Top             =   1560
            Width           =   3135
         End
         Begin VB.Label lblMM 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Map: 1"
            Height          =   195
            Left            =   120
            TabIndex        =   205
            Top             =   120
            Width           =   495
         End
         Begin VB.Label lblMX 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "X: 0"
            Height          =   195
            Left            =   120
            TabIndex        =   204
            Top             =   720
            Width           =   285
         End
         Begin VB.Label lblMY 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Y: 0"
            Height          =   195
            Left            =   120
            TabIndex        =   203
            Top             =   1320
            Width           =   285
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Calculator"
         Height          =   2415
         Left            =   -70200
         TabIndex        =   193
         Top             =   1667
         Width           =   5295
         Begin VB.Label Label9 
            Caption         =   "Under Development..."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2040
            TabIndex        =   194
            Top             =   1080
            Width           =   1815
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   3015
         Left            =   -71040
         ScaleHeight     =   2985
         ScaleWidth      =   3945
         TabIndex        =   175
         Top             =   600
         Width           =   3975
         Visible         =   0   'False
         Begin VB.TextBox Text13 
            Height          =   285
            Left            =   120
            TabIndex        =   231
            Top             =   2280
            Width           =   2055
         End
         Begin VB.CommandButton Command46 
            Caption         =   "Change Guild"
            Height          =   255
            Left            =   2280
            TabIndex        =   230
            Top             =   2280
            Width           =   1575
         End
         Begin VB.TextBox Text5 
            Height          =   285
            Left            =   120
            TabIndex        =   229
            Top             =   2040
            Width           =   2055
         End
         Begin VB.CommandButton Command44 
            Caption         =   "Change Password"
            Height          =   255
            Left            =   2280
            TabIndex        =   228
            Top             =   2040
            Width           =   1575
         End
         Begin VB.CommandButton Command25 
            Caption         =   "*Send Changes to The Client*"
            Height          =   255
            Left            =   1080
            TabIndex        =   195
            Top             =   2640
            Width           =   2775
         End
         Begin VB.CommandButton Command118 
            Caption         =   "Change Stat Points"
            Height          =   255
            Left            =   2280
            TabIndex        =   192
            Top             =   1800
            Width           =   1575
         End
         Begin VB.TextBox Text12 
            Height          =   285
            Left            =   120
            TabIndex        =   191
            Top             =   1800
            Width           =   2055
         End
         Begin VB.CommandButton Command117 
            Caption         =   "Change Wisdom"
            Height          =   255
            Left            =   2280
            TabIndex        =   190
            Top             =   1560
            Width           =   1575
         End
         Begin VB.TextBox Text11 
            Height          =   285
            Left            =   120
            TabIndex        =   189
            Top             =   1560
            Width           =   2055
         End
         Begin VB.CommandButton Command116 
            Caption         =   "Change Agility"
            Height          =   255
            Left            =   2280
            TabIndex        =   188
            Top             =   1320
            Width           =   1575
         End
         Begin VB.TextBox Text10 
            Height          =   285
            Left            =   120
            TabIndex        =   187
            Top             =   1320
            Width           =   2055
         End
         Begin VB.CommandButton Command115 
            Caption         =   "Change Defence"
            Height          =   255
            Left            =   2280
            TabIndex        =   186
            Top             =   1080
            Width           =   1575
         End
         Begin VB.TextBox Text9 
            Height          =   285
            Left            =   120
            TabIndex        =   185
            Top             =   1080
            Width           =   2055
         End
         Begin VB.CommandButton Command96 
            Caption         =   "Change Strength"
            Height          =   255
            Left            =   2280
            TabIndex        =   184
            Top             =   840
            Width           =   1575
         End
         Begin VB.TextBox Text8 
            Height          =   285
            Left            =   120
            TabIndex        =   183
            Top             =   840
            Width           =   2055
         End
         Begin VB.CommandButton Command95 
            Caption         =   "Change Access"
            Height          =   255
            Left            =   2280
            TabIndex        =   182
            Top             =   600
            Width           =   1575
         End
         Begin VB.TextBox Text7 
            Height          =   285
            Left            =   120
            TabIndex        =   181
            Top             =   600
            Width           =   2055
         End
         Begin VB.CommandButton Command94 
            Caption         =   "Change Level"
            Height          =   255
            Left            =   2280
            TabIndex        =   180
            Top             =   360
            Width           =   1575
         End
         Begin VB.TextBox Text6 
            Height          =   285
            Left            =   120
            TabIndex        =   179
            Top             =   360
            Width           =   2055
         End
         Begin VB.CommandButton Command93 
            Caption         =   "Change Name"
            Height          =   255
            Left            =   2280
            TabIndex        =   178
            Top             =   120
            Width           =   1575
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   120
            TabIndex        =   177
            Top             =   120
            Width           =   2055
         End
         Begin VB.CommandButton Command91 
            Caption         =   "Close"
            Height          =   255
            Left            =   120
            TabIndex        =   176
            Top             =   2640
            Width           =   975
         End
      End
      Begin VB.PictureBox picMap 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3375
         Left            =   -68160
         ScaleHeight     =   223
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   223
         TabIndex        =   158
         Top             =   227
         Width           =   3375
         Visible         =   0   'False
         Begin VB.CommandButton Command41 
            Caption         =   "Cancel"
            Height          =   255
            Left            =   1680
            TabIndex        =   160
            Top             =   3000
            Width           =   1575
         End
         Begin VB.ListBox lstNPC 
            Height          =   2400
            Left            =   1680
            TabIndex        =   159
            Top             =   480
            Width           =   1575
         End
         Begin VB.Label MapInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Map"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   174
            Top             =   120
            Width           =   300
         End
         Begin VB.Label MapInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Revision:"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   173
            Top             =   360
            Width           =   660
         End
         Begin VB.Label MapInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Moral:"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   172
            Top             =   600
            Width           =   450
         End
         Begin VB.Label MapInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Up:"
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   171
            Top             =   840
            Width           =   255
         End
         Begin VB.Label MapInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Down:"
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   170
            Top             =   1080
            Width           =   465
         End
         Begin VB.Label MapInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Left:"
            Height          =   195
            Index           =   5
            Left            =   120
            TabIndex        =   169
            Top             =   1320
            Width           =   345
         End
         Begin VB.Label MapInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Right:"
            Height          =   195
            Index           =   6
            Left            =   120
            TabIndex        =   168
            Top             =   1560
            Width           =   435
         End
         Begin VB.Label MapInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Music:"
            Height          =   195
            Index           =   7
            Left            =   120
            TabIndex        =   167
            Top             =   1800
            Width           =   450
         End
         Begin VB.Label MapInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "BootMap:"
            Height          =   195
            Index           =   8
            Left            =   120
            TabIndex        =   166
            Top             =   2040
            Width           =   690
         End
         Begin VB.Label MapInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "BootX:"
            Height          =   195
            Index           =   9
            Left            =   120
            TabIndex        =   165
            Top             =   2280
            Width           =   480
         End
         Begin VB.Label MapInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "BootY:"
            Height          =   195
            Index           =   10
            Left            =   120
            TabIndex        =   164
            Top             =   2520
            Width           =   480
         End
         Begin VB.Label MapInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Shop:"
            Height          =   195
            Index           =   11
            Left            =   120
            TabIndex        =   163
            Top             =   2760
            Width           =   420
         End
         Begin VB.Label MapInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Indoors:"
            Height          =   195
            Index           =   12
            Left            =   120
            TabIndex        =   162
            Top             =   3000
            Width           =   615
         End
         Begin VB.Label MapInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "NPCs"
            Height          =   195
            Index           =   13
            Left            =   1680
            TabIndex        =   161
            Top             =   285
            Width           =   375
         End
      End
      Begin VB.CommandButton Say 
         Caption         =   "Say"
         Height          =   255
         Index           =   5
         Left            =   9600
         TabIndex        =   155
         Top             =   3587
         Width           =   495
      End
      Begin VB.CommandButton Say 
         Caption         =   "Say"
         Height          =   255
         Index           =   4
         Left            =   9600
         TabIndex        =   154
         Top             =   2987
         Width           =   495
      End
      Begin VB.CommandButton Say 
         Caption         =   "Say"
         Height          =   255
         Index           =   3
         Left            =   9600
         TabIndex        =   153
         Top             =   2387
         Width           =   495
      End
      Begin VB.CommandButton Say 
         Caption         =   "Say"
         Height          =   255
         Index           =   2
         Left            =   9600
         TabIndex        =   152
         Top             =   1787
         Width           =   495
      End
      Begin VB.CommandButton Say 
         Caption         =   "Say"
         Height          =   255
         Index           =   1
         Left            =   9600
         TabIndex        =   151
         Top             =   1187
         Width           =   495
      End
      Begin VB.CommandButton Say 
         Caption         =   "Say"
         Height          =   255
         Index           =   0
         Left            =   9600
         TabIndex        =   150
         Top             =   587
         Width           =   495
      End
      Begin VB.CommandButton Command45 
         Caption         =   "Warp"
         Height          =   255
         Left            =   -66600
         TabIndex        =   149
         Top             =   3587
         Width           =   1575
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Heal"
         Height          =   255
         Left            =   -66600
         TabIndex        =   148
         Top             =   3347
         Width           =   1575
      End
      Begin VB.CommandButton Command24 
         Caption         =   "Kill"
         Height          =   255
         Left            =   -66600
         TabIndex        =   147
         Top             =   3107
         Width           =   1575
      End
      Begin VB.CommandButton Command23 
         Caption         =   "UnMute"
         Height          =   255
         Left            =   -66600
         TabIndex        =   146
         Top             =   2867
         Width           =   1575
      End
      Begin VB.CommandButton Command22 
         Caption         =   "Mute"
         Height          =   255
         Left            =   -66600
         TabIndex        =   145
         Top             =   2627
         Width           =   1575
      End
      Begin VB.CommandButton Command21 
         Caption         =   "Message (PM)"
         Height          =   255
         Left            =   -66600
         TabIndex        =   144
         Top             =   2387
         Width           =   1575
      End
      Begin VB.CommandButton Command20 
         Caption         =   "Change Info"
         Height          =   255
         Left            =   -66600
         TabIndex        =   143
         Top             =   2147
         Width           =   1575
      End
      Begin VB.CommandButton Command19 
         Caption         =   "View Info"
         Height          =   255
         Left            =   -66600
         TabIndex        =   142
         Top             =   1907
         Width           =   1575
      End
      Begin VB.CommandButton Command18 
         Caption         =   "Jail (Reason)"
         Height          =   255
         Left            =   -66600
         TabIndex        =   141
         Top             =   1667
         Width           =   1575
      End
      Begin VB.CommandButton Command17 
         Caption         =   "Jail"
         Height          =   255
         Left            =   -66600
         TabIndex        =   140
         Top             =   1427
         Width           =   1575
      End
      Begin VB.CommandButton Command16 
         Caption         =   "Ban (Reason)"
         Height          =   255
         Left            =   -66600
         TabIndex        =   139
         Top             =   1187
         Width           =   1575
      End
      Begin VB.CommandButton Command15 
         Caption         =   "Ban"
         Height          =   255
         Left            =   -66600
         TabIndex        =   138
         Top             =   947
         Width           =   1575
      End
      Begin VB.CommandButton Command14 
         Caption         =   "Kick (Reason)"
         Height          =   255
         Left            =   -66600
         TabIndex        =   137
         Top             =   707
         Width           =   1575
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Kick"
         Height          =   255
         Left            =   -66600
         TabIndex        =   136
         Top             =   467
         Width           =   1575
      End
      Begin VB.Frame Frame6 
         Caption         =   "Commands"
         Height          =   2535
         Left            =   -71520
         TabIndex        =   125
         Top             =   347
         Width           =   2535
         Begin VB.CommandButton Command27 
            Caption         =   "Mass Stamina"
            Height          =   255
            Left            =   120
            TabIndex        =   197
            Top             =   2160
            Width           =   2295
         End
         Begin VB.CommandButton Command26 
            Caption         =   "Mass Mana"
            Height          =   255
            Left            =   120
            TabIndex        =   196
            Top             =   1920
            Width           =   2295
         End
         Begin VB.CommandButton Command77 
            Caption         =   "Mass Save"
            Height          =   255
            Left            =   120
            TabIndex        =   132
            Top             =   1680
            Width           =   2295
         End
         Begin VB.CommandButton Command9 
            Caption         =   "Mass Kick"
            Height          =   255
            Left            =   120
            TabIndex        =   131
            Top             =   1440
            Width           =   2295
         End
         Begin VB.CommandButton Command31 
            Caption         =   "Mass Kill"
            Height          =   255
            Left            =   120
            TabIndex        =   130
            Top             =   1200
            Width           =   2295
         End
         Begin VB.CommandButton Command12 
            Caption         =   "Mass Heal"
            Height          =   255
            Left            =   120
            TabIndex        =   129
            Top             =   960
            Width           =   2295
         End
         Begin VB.CommandButton Command32 
            Caption         =   "Mass Warp"
            Height          =   255
            Left            =   120
            TabIndex        =   128
            Top             =   720
            Width           =   2295
         End
         Begin VB.CommandButton Command33 
            Caption         =   "Mass Experience"
            Height          =   255
            Left            =   120
            TabIndex        =   127
            Top             =   480
            Width           =   2295
         End
         Begin VB.CommandButton Command34 
            Caption         =   "Mass Level"
            Height          =   255
            Left            =   120
            TabIndex        =   126
            Top             =   240
            Width           =   2295
         End
      End
      Begin VB.Frame Frame25 
         Caption         =   "Options"
         Height          =   2535
         Left            =   -74880
         TabIndex        =   119
         Top             =   360
         Width           =   3255
         Begin VB.CommandButton Command28 
            Caption         =   "Dropping Item"
            Height          =   255
            Left            =   120
            TabIndex        =   221
            Top             =   2160
            Width           =   3015
         End
         Begin VB.CommandButton Command67 
            Caption         =   "Quest System"
            Height          =   255
            Left            =   120
            TabIndex        =   135
            Top             =   1920
            Width           =   3015
         End
         Begin VB.CommandButton Command88 
            Caption         =   "Party Range Level"
            Height          =   255
            Left            =   120
            TabIndex        =   134
            Top             =   1680
            Width           =   3015
         End
         Begin VB.CommandButton Command69 
            Caption         =   "PK Level"
            Height          =   255
            Left            =   120
            TabIndex        =   133
            Top             =   1440
            Width           =   3015
         End
         Begin VB.CommandButton Command80 
            Caption         =   "Save Objects"
            Height          =   255
            Left            =   120
            TabIndex        =   124
            Top             =   1200
            Width           =   3015
         End
         Begin VB.CommandButton Command87 
            Caption         =   "Character Size"
            Height          =   255
            Left            =   120
            TabIndex        =   123
            Top             =   960
            Width           =   3015
         End
         Begin VB.CommandButton Command79 
            Caption         =   "Minimum Damage"
            Height          =   255
            Left            =   120
            TabIndex        =   122
            Top             =   720
            Width           =   3015
         End
         Begin VB.CommandButton Command76 
            Caption         =   "HP, MP and SP Regen"
            Height          =   255
            Left            =   120
            TabIndex        =   121
            Top             =   480
            Width           =   3015
         End
         Begin VB.CommandButton Command78 
            Caption         =   "Experience Given Away"
            Height          =   255
            Left            =   120
            TabIndex        =   120
            Top             =   240
            Width           =   3015
         End
      End
      Begin VB.CommandButton Command66 
         Caption         =   "Refresh"
         Height          =   255
         Left            =   -69604
         TabIndex        =   107
         Top             =   3732
         Width           =   1575
      End
      Begin VB.CommandButton CustomMsg 
         Caption         =   "Custom Msg"
         Height          =   255
         Index           =   0
         Left            =   8636
         TabIndex        =   106
         Top             =   347
         Width           =   1455
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Gridlines"
         Height          =   255
         Left            =   -67924
         TabIndex        =   105
         Top             =   3732
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.Frame Frame2 
         Caption         =   "Server"
         Height          =   1575
         Left            =   -68040
         TabIndex        =   95
         Top             =   2387
         Width           =   3255
         Begin VB.CommandButton Command1 
            Caption         =   "Shutdown"
            Height          =   255
            Left            =   2040
            TabIndex        =   103
            Top             =   1080
            Width           =   1095
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Exit"
            Height          =   255
            Left            =   2040
            TabIndex        =   102
            Top             =   840
            Width           =   1095
         End
         Begin VB.CheckBox GMOnly 
            Caption         =   "GMs Only"
            Height          =   255
            Left            =   120
            TabIndex        =   101
            Top             =   360
            Width           =   975
         End
         Begin VB.CheckBox Closed 
            Caption         =   "Closed"
            Height          =   255
            Left            =   120
            TabIndex        =   100
            Top             =   600
            Width           =   855
         End
         Begin VB.CheckBox mnuServerLog 
            Caption         =   "Server Log"
            Height          =   255
            Left            =   120
            TabIndex        =   99
            Top             =   960
            Value           =   1  'Checked
            Width           =   1095
         End
         Begin VB.CommandButton Command58 
            Caption         =   "Day/Night"
            Height          =   255
            Left            =   2040
            TabIndex        =   98
            Top             =   600
            Width           =   1095
         End
         Begin VB.CommandButton Command59 
            Caption         =   "Weather"
            Height          =   255
            Left            =   2040
            TabIndex        =   97
            Top             =   360
            Width           =   1095
         End
         Begin VB.CheckBox chkChat 
            Caption         =   "Save Logs"
            Height          =   255
            Left            =   120
            TabIndex        =   96
            Top             =   1200
            Value           =   1  'Checked
            Width           =   1095
         End
         Begin VB.Label ShutdownTime 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Shutdown: Not Active"
            Height          =   195
            Left            =   1560
            TabIndex        =   104
            Top             =   1320
            Width           =   1575
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Chat Options"
         Height          =   615
         Left            =   236
         TabIndex        =   87
         Top             =   3467
         Width           =   6975
         Begin VB.CheckBox chkBC 
            Caption         =   "Broadcast"
            Height          =   255
            Left            =   120
            TabIndex        =   94
            Top             =   240
            Value           =   1  'Checked
            Width           =   1095
         End
         Begin VB.CheckBox chkE 
            Caption         =   "Emote"
            Height          =   255
            Left            =   1200
            TabIndex        =   93
            Top             =   240
            Value           =   1  'Checked
            Width           =   855
         End
         Begin VB.CheckBox chkM 
            Caption         =   "Map"
            Height          =   255
            Left            =   2040
            TabIndex        =   92
            Top             =   240
            Value           =   1  'Checked
            Width           =   615
         End
         Begin VB.CheckBox chkP 
            Caption         =   "Private"
            Height          =   255
            Left            =   2760
            TabIndex        =   91
            Top             =   240
            Value           =   1  'Checked
            Width           =   855
         End
         Begin VB.CheckBox chkG 
            Caption         =   "Global"
            Height          =   255
            Left            =   3720
            TabIndex        =   90
            Top             =   240
            Value           =   1  'Checked
            Width           =   735
         End
         Begin VB.CheckBox chkA 
            Caption         =   "Admin"
            Height          =   255
            Left            =   4560
            TabIndex        =   89
            Top             =   240
            Value           =   1  'Checked
            Width           =   735
         End
         Begin VB.CommandButton Command60 
            Caption         =   "Save Logs"
            Height          =   255
            Left            =   5400
            TabIndex        =   88
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.CommandButton CustomMsg 
         Caption         =   "Custom Msg"
         Height          =   255
         Index           =   1
         Left            =   8636
         TabIndex        =   86
         Top             =   950
         Width           =   1455
      End
      Begin VB.CommandButton CustomMsg 
         Caption         =   "Custom Msg"
         Height          =   255
         Index           =   2
         Left            =   8636
         TabIndex        =   85
         Top             =   1547
         Width           =   1455
      End
      Begin VB.CommandButton CustomMsg 
         Caption         =   "Custom Msg"
         Height          =   255
         Index           =   3
         Left            =   8636
         TabIndex        =   84
         Top             =   2147
         Width           =   1455
      End
      Begin VB.CommandButton CustomMsg 
         Caption         =   "Custom Msg"
         Height          =   255
         Index           =   4
         Left            =   8636
         TabIndex        =   83
         Top             =   2747
         Width           =   1455
      End
      Begin VB.CommandButton CustomMsg 
         Caption         =   "Custom Msg"
         Height          =   255
         Index           =   5
         Left            =   8636
         TabIndex        =   82
         Top             =   3347
         Width           =   1455
      End
      Begin VB.PictureBox picReason 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   -70440
         ScaleHeight     =   47
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   223
         TabIndex        =   77
         Top             =   467
         Width           =   3375
         Visible         =   0   'False
         Begin VB.TextBox txtReason 
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
            TabIndex        =   80
            Top             =   360
            Width           =   3075
         End
         Begin VB.CommandButton Command7 
            Caption         =   "Caption"
            Height          =   255
            Left            =   1680
            TabIndex        =   79
            Top             =   720
            Width           =   1575
         End
         Begin VB.CommandButton Command6 
            Caption         =   "Cancel"
            Height          =   255
            Left            =   1680
            TabIndex        =   78
            Top             =   960
            Width           =   1575
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Reason:"
            Height          =   195
            Left            =   120
            TabIndex        =   81
            Top             =   120
            Width           =   600
         End
      End
      Begin VB.PictureBox picStats 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3255
         Left            =   -74880
         ScaleHeight     =   215
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   311
         TabIndex        =   54
         Top             =   467
         Width           =   4695
         Visible         =   0   'False
         Begin VB.CommandButton Command8 
            Caption         =   "Cancel"
            Height          =   255
            Left            =   3000
            TabIndex        =   55
            Top             =   2880
            Width           =   1575
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Account:"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   76
            Top             =   120
            Width           =   645
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Character:"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   75
            Top             =   360
            Width           =   780
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Level:"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   74
            Top             =   600
            Width           =   435
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "HP: /"
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   73
            Top             =   840
            Width           =   360
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "MP: /"
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   72
            Top             =   1080
            Width           =   375
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "SP: /"
            Height          =   195
            Index           =   5
            Left            =   120
            TabIndex        =   71
            Top             =   1320
            Width           =   345
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "EXP: /"
            Height          =   195
            Index           =   6
            Left            =   120
            TabIndex        =   70
            Top             =   1560
            Width           =   435
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Access:"
            Height          =   195
            Index           =   7
            Left            =   120
            TabIndex        =   69
            Top             =   1800
            Width           =   555
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PK:"
            Height          =   195
            Index           =   8
            Left            =   120
            TabIndex        =   68
            Top             =   2040
            Width           =   240
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Class:"
            Height          =   195
            Index           =   9
            Left            =   120
            TabIndex        =   67
            Top             =   2280
            Width           =   435
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sprite:"
            Height          =   195
            Index           =   10
            Left            =   120
            TabIndex        =   66
            Top             =   2520
            Width           =   480
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sex:"
            Height          =   195
            Index           =   11
            Left            =   120
            TabIndex        =   65
            Top             =   2760
            Width           =   330
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Map:"
            Height          =   195
            Index           =   12
            Left            =   120
            TabIndex        =   64
            Top             =   3000
            Width           =   360
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Guild:"
            Height          =   195
            Index           =   13
            Left            =   2400
            TabIndex        =   63
            Top             =   120
            Width           =   405
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Guild Access:"
            Height          =   195
            Index           =   14
            Left            =   2400
            TabIndex        =   62
            Top             =   360
            Width           =   945
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Str:"
            Height          =   195
            Index           =   15
            Left            =   2400
            TabIndex        =   61
            Top             =   600
            Width           =   270
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Def:"
            Height          =   195
            Index           =   16
            Left            =   2400
            TabIndex        =   60
            Top             =   840
            Width           =   315
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Agility:"
            Height          =   195
            Index           =   17
            Left            =   2400
            TabIndex        =   59
            Top             =   1080
            Width           =   495
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Wisdom:"
            Height          =   195
            Index           =   18
            Left            =   2400
            TabIndex        =   58
            Top             =   1320
            Width           =   615
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Points:"
            Height          =   195
            Index           =   19
            Left            =   2400
            TabIndex        =   57
            Top             =   1560
            Width           =   495
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Index:"
            Height          =   195
            Index           =   20
            Left            =   2400
            TabIndex        =   56
            Top             =   1800
            Width           =   480
         End
      End
      Begin VB.PictureBox picJail 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2535
         Left            =   -70440
         ScaleHeight     =   167
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   223
         TabIndex        =   46
         Top             =   1187
         Width           =   3375
         Visible         =   0   'False
         Begin VB.CommandButton Command11 
            Caption         =   "Cancel"
            Height          =   255
            Left            =   1680
            TabIndex        =   157
            Top             =   2160
            Width           =   1575
         End
         Begin VB.CommandButton Command10 
            Caption         =   "Jail"
            Height          =   255
            Left            =   1680
            TabIndex        =   50
            Top             =   1920
            Width           =   1575
         End
         Begin VB.HScrollBar scrlMap 
            Height          =   255
            Left            =   120
            Min             =   1
            TabIndex        =   49
            Top             =   360
            Value           =   1
            Width           =   3135
         End
         Begin VB.HScrollBar scrlX 
            Height          =   255
            Left            =   120
            TabIndex        =   48
            Top             =   960
            Width           =   3135
         End
         Begin VB.HScrollBar scrlY 
            Height          =   255
            Left            =   120
            TabIndex        =   47
            Top             =   1560
            Width           =   3135
         End
         Begin VB.Label txtMap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Map: 1"
            Height          =   195
            Left            =   120
            TabIndex        =   53
            Top             =   120
            Width           =   495
         End
         Begin VB.Label txtX 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "X: 0"
            Height          =   195
            Left            =   120
            TabIndex        =   52
            Top             =   720
            Width           =   285
         End
         Begin VB.Label txtY 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Y: 0"
            Height          =   195
            Left            =   120
            TabIndex        =   51
            Top             =   1320
            Width           =   285
         End
      End
      Begin VB.ListBox lstTopics 
         Height          =   2790
         ItemData        =   "frmServer.frx":0416
         Left            =   -74764
         List            =   "frmServer.frx":0418
         TabIndex        =   45
         Top             =   564
         Width           =   2175
      End
      Begin VB.Frame TopicTitle 
         Caption         =   "Topic Title"
         Height          =   3375
         Left            =   -72484
         TabIndex        =   43
         Top             =   444
         Width           =   7575
         Begin VB.TextBox txtTopic 
            Height          =   3015
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   44
            Top             =   240
            Width           =   7335
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Map List"
         Height          =   2055
         Left            =   -68884
         TabIndex        =   38
         Top             =   349
         Width           =   4095
         Begin VB.ListBox MapList 
            Height          =   1230
            Left            =   120
            TabIndex        =   41
            Top             =   320
            Width           =   3855
         End
         Begin VB.CommandButton Command35 
            Caption         =   "Refresh"
            Height          =   255
            Left            =   120
            TabIndex        =   40
            Top             =   1680
            Width           =   1095
         End
         Begin VB.CommandButton Command36 
            Caption         =   "Map Info"
            Height          =   255
            Left            =   1200
            TabIndex        =   39
            Top             =   1680
            Width           =   1095
         End
      End
      Begin VB.PictureBox picCMsg 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1935
         Left            =   5040
         ScaleHeight     =   1905
         ScaleWidth      =   3345
         TabIndex        =   23
         Top             =   840
         Width           =   3375
         Visible         =   0   'False
         Begin VB.CommandButton Command5 
            Caption         =   "Cancel"
            Height          =   255
            Left            =   1680
            TabIndex        =   156
            Top             =   1560
            Width           =   1575
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Save"
            Height          =   255
            Left            =   1680
            TabIndex        =   26
            Top             =   1320
            Width           =   1575
         End
         Begin VB.TextBox txtTitle 
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
            MaxLength       =   13
            TabIndex        =   25
            Top             =   360
            Width           =   3075
         End
         Begin VB.TextBox txtMsg 
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
            TabIndex        =   24
            Top             =   960
            Width           =   3075
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Title:"
            Height          =   195
            Left            =   120
            TabIndex        =   28
            Top             =   120
            Width           =   360
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Message:"
            Height          =   195
            Left            =   120
            TabIndex        =   27
            Top             =   720
            Width           =   690
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Accounts"
         Height          =   1455
         Left            =   -70200
         TabIndex        =   19
         Top             =   227
         Width           =   1455
         Begin VB.CommandButton Command97 
            Caption         =   "Admin Account"
            Height          =   255
            Left            =   120
            TabIndex        =   117
            Top             =   1080
            Width           =   1215
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   120
            TabIndex        =   21
            Top             =   480
            Width           =   1215
         End
         Begin VB.CommandButton Command56 
            Caption         =   "Open Account"
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "Account Name:"
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Frame Data 
         Caption         =   "Data"
         Height          =   3855
         Left            =   -74882
         TabIndex        =   3
         Top             =   229
         Width           =   4575
         Begin VB.CommandButton Command48 
            Caption         =   "Data"
            Height          =   195
            Left            =   3960
            TabIndex        =   4
            Top             =   3600
            Width           =   495
         End
         Begin VB.Label gamename 
            BackStyle       =   0  'Transparent
            Caption         =   "1"
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   240
            Width           =   4335
         End
         Begin VB.Label maxplayers 
            BackStyle       =   0  'Transparent
            Caption         =   "2"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   480
            Width           =   4335
         End
         Begin VB.Label maxitems 
            BackStyle       =   0  'Transparent
            Caption         =   "3"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   720
            Width           =   4335
         End
         Begin VB.Label maxnpcs 
            BackStyle       =   0  'Transparent
            Caption         =   "4"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   960
            Width           =   4335
         End
         Begin VB.Label maxshops 
            BackStyle       =   0  'Transparent
            Caption         =   "5"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   1200
            Width           =   4335
         End
         Begin VB.Label maxspells 
            BackStyle       =   0  'Transparent
            Caption         =   "6"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   1440
            Width           =   4335
         End
         Begin VB.Label maxmaps 
            BackStyle       =   0  'Transparent
            Caption         =   "7"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   1680
            Width           =   4335
         End
         Begin VB.Label maxmapitems 
            BackStyle       =   0  'Transparent
            Caption         =   "8"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   1920
            Width           =   4335
         End
         Begin VB.Label maxguilds 
            BackStyle       =   0  'Transparent
            Caption         =   "9"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   2160
            Width           =   4335
         End
         Begin VB.Label maxguildmembers 
            BackStyle       =   0  'Transparent
            Caption         =   "10"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   2400
            Width           =   4335
         End
         Begin VB.Label maxemoticons 
            BackStyle       =   0  'Transparent
            Caption         =   "11"
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   2640
            Width           =   4335
         End
         Begin VB.Label maxlevel 
            BackStyle       =   0  'Transparent
            Caption         =   "12"
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   2880
            Width           =   4335
         End
         Begin VB.Label maxpartymembers 
            BackStyle       =   0  'Transparent
            Caption         =   "14"
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   3360
            Width           =   4335
         End
         Begin VB.Label Scripting 
            BackStyle       =   0  'Transparent
            Caption         =   "13"
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   3120
            Width           =   4335
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Classes"
         Height          =   1455
         Left            =   -68640
         TabIndex        =   2
         Top             =   227
         Width           =   1455
         Begin VB.TextBox Text3 
            Height          =   285
            Left            =   120
            TabIndex        =   118
            Top             =   480
            Width           =   495
         End
         Begin VB.CommandButton Command30 
            Caption         =   "Open"
            Height          =   273
            Left            =   600
            TabIndex        =   116
            Top             =   480
            Width           =   735
         End
         Begin VB.CommandButton Command73 
            Caption         =   "Default Class"
            Height          =   255
            Left            =   120
            TabIndex        =   1
            Top             =   1080
            Width           =   1215
         End
         Begin VB.CommandButton Command75 
            Caption         =   "Class Info"
            Height          =   255
            Left            =   120
            TabIndex        =   42
            Top             =   840
            Width           =   1215
         End
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   2985
         Left            =   118
         TabIndex        =   29
         Top             =   349
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   5265
         _Version        =   393216
         Style           =   1
         Tabs            =   7
         TabsPerRow      =   7
         TabHeight       =   353
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Main"
         TabPicture(0)   =   "frmServer.frx":041A
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "txtChat"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "txtText(0)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Picture2"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).ControlCount=   3
         TabCaption(1)   =   "Broadcast"
         TabPicture(1)   =   "frmServer.frx":0436
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "txtText(1)"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Global"
         TabPicture(2)   =   "frmServer.frx":0452
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "txtText(2)"
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "Map"
         TabPicture(3)   =   "frmServer.frx":046E
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "txtText(3)"
         Tab(3).ControlCount=   1
         TabCaption(4)   =   "Private"
         TabPicture(4)   =   "frmServer.frx":048A
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "txtText(4)"
         Tab(4).ControlCount=   1
         TabCaption(5)   =   "Admin"
         TabPicture(5)   =   "frmServer.frx":04A6
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "txtText(5)"
         Tab(5).ControlCount=   1
         TabCaption(6)   =   "Emote"
         TabPicture(6)   =   "frmServer.frx":04C2
         Tab(6).ControlEnabled=   0   'False
         Tab(6).Control(0)=   "txtText(6)"
         Tab(6).ControlCount=   1
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   1335
            Left            =   1440
            ScaleHeight     =   1305
            ScaleWidth      =   3345
            TabIndex        =   222
            Top             =   840
            Width           =   3375
            Visible         =   0   'False
            Begin VB.CommandButton Command42 
               Caption         =   "Cancel"
               Height          =   255
               Left            =   1680
               TabIndex        =   226
               Top             =   960
               Width           =   1575
            End
            Begin VB.CommandButton Command43 
               Caption         =   "Save"
               Height          =   255
               Left            =   1680
               TabIndex        =   225
               Top             =   720
               Width           =   1575
            End
            Begin VB.TextBox Text4 
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
               MaxLength       =   13
               TabIndex        =   223
               Top             =   360
               Width           =   3075
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Server Chat Name:"
               Height          =   195
               Left            =   120
               TabIndex        =   224
               Top             =   120
               Width           =   1380
            End
         End
         Begin VB.TextBox txtText 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2250
            Index           =   0
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   37
            Top             =   360
            Width           =   8115
         End
         Begin VB.TextBox txtChat 
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
            TabIndex        =   36
            Top             =   2640
            Width           =   8115
         End
         Begin VB.TextBox txtText 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2490
            Index           =   1
            Left            =   -74880
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   35
            Top             =   360
            Width           =   8115
         End
         Begin VB.TextBox txtText 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2490
            Index           =   2
            Left            =   -74880
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   34
            Top             =   360
            Width           =   8115
         End
         Begin VB.TextBox txtText 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2490
            Index           =   3
            Left            =   -74880
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   33
            Top             =   360
            Width           =   8115
         End
         Begin VB.TextBox txtText 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2490
            Index           =   4
            Left            =   -74880
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   32
            Top             =   360
            Width           =   8115
         End
         Begin VB.TextBox txtText 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2490
            Index           =   5
            Left            =   -74880
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   31
            Top             =   360
            Width           =   8115
         End
         Begin VB.TextBox txtText 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2490
            Index           =   6
            Left            =   -74880
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   30
            Top             =   360
            Width           =   8115
         End
      End
      Begin MSComctlLib.ListView lvUsers 
         Height          =   3255
         Left            =   -74880
         TabIndex        =   108
         Top             =   467
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   5741
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Index"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Account"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Character"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Level"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Sprite"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Access"
            Object.Width           =   1235
         EndProperty
      End
      Begin MSWinsockLib.Winsock Socket 
         Index           =   0
         Left            =   -68760
         Top             =   120
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.Label TPO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Players Online:"
         Height          =   195
         Left            =   -74764
         TabIndex        =   115
         Top             =   3744
         Width           =   1485
      End
      Begin VB.Label lblPort 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Port:"
         Height          =   195
         Left            =   -74880
         TabIndex        =   114
         Top             =   3827
         Width           =   360
      End
      Begin VB.Label lblIP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ip Address:"
         Height          =   195
         Left            =   -74880
         TabIndex        =   113
         Top             =   3587
         Width           =   840
      End
      Begin VB.Label CharInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Topics:"
         Height          =   195
         Index           =   21
         Left            =   -74764
         TabIndex        =   112
         Top             =   324
         Width           =   510
      End
      Begin VB.Label CharInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "For More Information Go To:"
         Height          =   195
         Index           =   22
         Left            =   -74764
         TabIndex        =   111
         Top             =   3444
         Width           =   2055
      End
      Begin VB.Label CharInfo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "www.Bevrias.com"
         Height          =   195
         Index           =   23
         Left            =   -74445
         TabIndex        =   110
         Top             =   3707
         Width           =   1305
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Chat Log Save In"
         Height          =   195
         Left            =   7320
         TabIndex        =   109
         Top             =   3610
         Width           =   1245
      End
   End
   Begin VB.Menu stFile 
      Caption         =   "File"
      Begin VB.Menu stBugFixes 
         Caption         =   "Bug Fixes"
      End
      Begin VB.Menu stShutdown 
         Caption         =   "Shutdown"
      End
      Begin VB.Menu stExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu stEditors 
      Caption         =   "Editors"
      Begin VB.Menu stDataEditor 
         Caption         =   "Data Editor"
      End
      Begin VB.Menu stQuickMsgEditor 
         Caption         =   "Quick Msg Editor"
      End
      Begin VB.Menu stMOTDEditor 
         Caption         =   "MOTD Editor"
      End
   End
   Begin VB.Menu stReloads 
      Caption         =   "Reloads"
      Begin VB.Menu stSpells 
         Caption         =   "Spells"
      End
      Begin VB.Menu stShops 
         Caption         =   "Shops"
      End
      Begin VB.Menu stNPCS 
         Caption         =   "NPCS"
      End
      Begin VB.Menu stItems 
         Caption         =   "Items"
      End
      Begin VB.Menu stMaps 
         Caption         =   "Maps"
      End
      Begin VB.Menu stClasses 
         Caption         =   "Classes"
      End
      Begin VB.Menu stMain 
         Caption         =   "Main"
      End
      Begin VB.Menu stMOTD 
         Caption         =   "MOTD"
      End
      Begin VB.Menu stArrows 
         Caption         =   "Arrows"
      End
      Begin VB.Menu stConstants 
         Caption         =   "Constants"
      End
      Begin VB.Menu stReloadAll 
         Caption         =   "Reload All"
      End
   End
   Begin VB.Menu stScripting 
      Caption         =   "Scripting"
      Begin VB.Menu stTurnOn 
         Caption         =   "Turn On"
      End
      Begin VB.Menu stTurnOff 
         Caption         =   "Turn Off"
      End
      Begin VB.Menu stMaintxt 
         Caption         =   "Main"
      End
      Begin VB.Menu stEmpty 
         Caption         =   "-------------------"
      End
      Begin VB.Menu stScriptingGuide 
         Caption         =   "Scripting Guide"
      End
      Begin VB.Menu stCommand 
         Caption         =   "Command"
      End
   End
   Begin VB.Menu stChange 
      Caption         =   "Change"
      Begin VB.Menu stChangeArrows 
         Caption         =   "Arrows"
      End
      Begin VB.Menu stChangeCMessages 
         Caption         =   "CMessages"
      End
      Begin VB.Menu stChangeEmoticons 
         Caption         =   "Emoticons"
      End
      Begin VB.Menu stChangeExperience 
         Caption         =   "Experience"
      End
      Begin VB.Menu stChangeMOTD 
         Caption         =   "MOTD"
      End
      Begin VB.Menu stChangeStats 
         Caption         =   "Stats"
      End
      Begin VB.Menu stChangeWishes 
         Caption         =   "Wishes"
      End
   End
   Begin VB.Menu stCheck 
      Caption         =   "Check"
      Begin VB.Menu stCheckAdminLogs 
         Caption         =   "Admin Logs"
      End
      Begin VB.Menu stCheckPlayerlogs 
         Caption         =   "Player Logs"
      End
      Begin VB.Menu stCheckAdmin 
         Caption         =   "Admin"
      End
      Begin VB.Menu stCheckPlayer 
         Caption         =   "Player"
      End
      Begin VB.Menu stCheckCharList 
         Caption         =   "Char List"
      End
      Begin VB.Menu stCheckBanList 
         Caption         =   "Ban List"
      End
   End
   Begin VB.Menu stIPInfo 
      Caption         =   "IP Info"
      Begin VB.Menu stwhatismyip 
         Caption         =   "whatismyip.com"
      End
      Begin VB.Menu stipchicken 
         Caption         =   "ipchicken.com"
      End
      Begin VB.Menu stshowmyip 
         Caption         =   "showmyip.com"
      End
      Begin VB.Menu stcheckip 
         Caption         =   "checkip.dyndns.org"
      End
      Begin VB.Menu stsupport 
         Caption         =   "support.netcentral.co.uk"
      End
   End
   Begin VB.Menu stAbout 
      Caption         =   "About"
      Begin VB.Menu stWhos 
         Caption         =   "Whos The Creator?"
      End
      Begin VB.Menu stCredits 
         Caption         =   "Credits"
      End
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CM As Long
Dim num As Long

Private Sub Command115_Click()
Dim index As Long
index = lvUsers.ListItems(lvUsers.SelectedItem.index).text
Call SetPlayerDEF(index, Text9.text)
CharInfo(16).Caption = "Def: " & Text9.text
End Sub

Private Sub Command116_Click()
Dim index As Long
index = lvUsers.ListItems(lvUsers.SelectedItem.index).text
Call SetPlayerSPEED(index, Text10.text)
CharInfo(17).Caption = "Agility: " & Text10.text
End Sub

Private Sub Command117_Click()
Dim index As Long
index = lvUsers.ListItems(lvUsers.SelectedItem.index).text
Call SetPlayerMAGI(index, Text11.text)
CharInfo(18).Caption = "Wisdom: " & Text11.text
End Sub

Private Sub Command118_Click()
Dim index As Long
index = lvUsers.ListItems(lvUsers.SelectedItem.index).text
Call SetPlayerMAGI(index, Text12.text)
CharInfo(19).Caption = "Points: " & Text12.text
End Sub

Private Sub Command20_Click()
Timer4.Enabled = True
Dim index As Long
If lvUsers.ListItems.Count = 0 Then Exit Sub
index = lvUsers.ListItems(lvUsers.SelectedItem.index).text
If IsPlaying(index) = False Then Exit Sub
Picture1.Visible = True
End Sub

Private Sub Command26_Click()
Dim index As Long

For index = 1 To MAX_PLAYERS
    If IsPlaying(index) = True Then
        Call SetPlayerMP(index, GetPlayerMaxMP(index))
        Call SendMP(index)
        Call PlayerMsg(index, "You have gained more Mana from the server!", BrightGreen)
    End If
Next index
End Sub

Private Sub Command27_Click()
Dim index As Long

For index = 1 To MAX_PLAYERS
    If IsPlaying(index) = True Then
        Call SetPlayerSP(index, GetPlayerMaxSP(index))
        Call SendSP(index)
        Call PlayerMsg(index, "You have gained more Stamina from the server!", BrightGreen)
    End If
Next index
End Sub

Private Sub Command68_Click()
    AFileName = "Emoticons.ini"
    Unload frmEditor
    frmEditor.Show
End Sub

Private Sub Command28_Click()
frmDroppingItem.Visible = True
End Sub

Private Sub Command29_Click()
Picture2.Visible = True
End Sub

Private Sub Command42_Click()
Picture2.Visible = False
End Sub

Private Sub Command43_Click()
PutVar App.Path & "\Data.ini", "ADDED", "ServerChatName", Text4.text
End Sub

Private Sub Command44_Click()
Dim index As Long
index = lvUsers.ListItems(lvUsers.SelectedItem.index).text
Call SetPlayerPassword(index, Text5.text)
End Sub

Private Sub Command46_Click()
Dim index As Long
index = lvUsers.ListItems(lvUsers.SelectedItem.index).text
Call SetPlayerGuild(index, Text13.text)
CharInfo(13).Caption = "Guild: " & Text13.text
End Sub

Private Sub Command67_Click()
frmQuest.Visible = True
End Sub

Private Sub Command69_Click()
frmPKLevel.Visible = True
End Sub

Private Sub Command72_Click()
    AFileName = "/scripts/readme.txt"
    Unload frmEditor
    frmEditor.Show
End Sub

Private Sub Command73_Click()
    AFileName = "/Accounts/.ini"
    Unload frmEditor
    frmEditor.Show
End Sub

Private Sub Command75_Click()
    AFileName = "/Classes/info.ini"
    Unload frmEditor
    frmEditor.Show
End Sub

Private Sub Command77_Click()
Call SaveAllPlayersOnline
End Sub

Private Sub Command78_Click()
frmExpToAttacker.Visible = True
End Sub
Private Sub Command76_Click()
frmOptions.Visible = True
End Sub
Private Sub Command79_Click()
frmMinimumDmg.Visible = True
End Sub

Private Sub Command80_Click()
frmSave.Visible = True
End Sub

Private Sub Command87_Click()
frmCharacterSize.Visible = True
End Sub

Private Sub Command93_Click()
Dim index As Long
index = lvUsers.ListItems(lvUsers.SelectedItem.index).text
Call SetPlayerName(index, Text2.text)
CharInfo(1).Caption = "Character: " & Text2.text
End Sub

Private Sub Command88_Click()
frmPartyRangeLevel.Visible = True
End Sub

Private Sub Command95_Click()
Dim index As Long
index = lvUsers.ListItems(lvUsers.SelectedItem.index).text
Call SetPlayerAccess(index, Text7.text)
CharInfo(7).Caption = "Access: " & Text7.text
End Sub

Private Sub Command96_Click()
Dim index As Long
index = lvUsers.ListItems(lvUsers.SelectedItem.index).text
Call SetPlayerSTR(index, Text8.text)
CharInfo(15).Caption = "Str: " & Text8.text
End Sub

Private Sub Command90_Click()
    AFileName = "Experience.ini"
    Unload frmEditor
    frmEditor.Show
End Sub

Private Sub Command91_Click()
Picture1.Visible = False
End Sub

Private Sub Command94_Click()
Dim index As Long
index = lvUsers.ListItems(lvUsers.SelectedItem.index).text
Call SetPlayerLevel(index, Text6.text)
CharInfo(2).Caption = "Level: " & Text6.text
End Sub

Private Sub Command97_Click()
    AFileName = "/Accounts/Admin.ini"
    Unload frmEditor
    frmEditor.Show
End Sub

Private Sub Form_Load()
    GAME_NAME = Trim(GetVar(App.Path & "\Data.ini", "CONFIG", "GameName"))
    MAX_PLAYERS = GetVar(App.Path & "\Data.ini", "MAX", "MAX_PLAYERS")
    MAX_ITEMS = GetVar(App.Path & "\Data.ini", "MAX", "MAX_ITEMS")
    MAX_NPCS = GetVar(App.Path & "\Data.ini", "MAX", "MAX_NPCS")
    MAX_SHOPS = GetVar(App.Path & "\Data.ini", "MAX", "MAX_SHOPS")
    MAX_SPELLS = GetVar(App.Path & "\Data.ini", "MAX", "MAX_SPELLS")
    MAX_MAPS = GetVar(App.Path & "\Data.ini", "MAX", "MAX_MAPS")
    MAX_MAP_ITEMS = GetVar(App.Path & "\Data.ini", "MAX", "MAX_MAP_ITEMS")
    MAX_GUILDS = GetVar(App.Path & "\Data.ini", "MAX", "MAX_GUILDS")
    MAX_GUILD_MEMBERS = GetVar(App.Path & "\Data.ini", "MAX", "MAX_GUILD_MEMBERS")
    MAX_EMOTICONS = GetVar(App.Path & "\Data.ini", "MAX", "MAX_EMOTICONS")
    MAX_LEVEL = GetVar(App.Path & "\Data.ini", "MAX", "MAX_LEVEL")
    Scripting = GetVar(App.Path & "\Data.ini", "CONFIG", "Scripting")
    MAX_PARTY_MEMBERS = GetVar(App.Path & "\Data.ini", "MAX", "MAX_PARTY_MEMBERS")
    gamename.Caption = "Game Name: " & GAME_NAME
    maxplayers.Caption = "Max Players: " & MAX_PLAYERS
    maxitems.Caption = "Max Items: " & MAX_ITEMS
    maxnpcs.Caption = "Max NPCS: " & MAX_NPCS
    maxshops.Caption = "Max Shops: " & MAX_SHOPS
    maxspells.Caption = "Max Spells: " & MAX_SPELLS
    maxmaps.Caption = "Max Maps: " & MAX_MAPS
    maxmapitems.Caption = "Max Map Items: " & MAX_MAP_ITEMS
    maxguilds.Caption = "Max Guilds: " & MAX_GUILDS
    maxguildmembers.Caption = "Max Guild Members: " & MAX_GUILD_MEMBERS
    maxemoticons.Caption = "Max Emoticons: " & MAX_EMOTICONS
    maxlevel.Caption = "Max Level: " & MAX_LEVEL
    Scripting.Caption = "Scripting: " & Scripting
    maxpartymembers.Caption = "Max Party Members: " & MAX_PARTY_MEMBERS
    Text4.text = GetVar(App.Path & "\Data.ini", "ADDED", "ServerChatName")
End Sub
Private Sub Check1_Click()
    If Check1.value = Checked Then
        lvUsers.GridLines = True
    Else
        lvUsers.GridLines = False
    End If
End Sub
Private Sub Command1_Click()
If tmrShutdown.Enabled = False Then
    tmrShutdown.Enabled = True
End If
End Sub

Private Sub Command10_Click()
Dim index As Long

index = lvUsers.ListItems(lvUsers.SelectedItem.index).text

If Command10.Caption = "Warp" Then
    If index > 0 Then
        If IsPlaying(index) Then
            Call PlayerMsg(index, "You have been warp by the server to Map:" & scrlMap.value & " X:" & scrlX.value & " Y:" & scrlY.value, White)
            Call PlayerWarp(index, scrlMap.value, scrlX.value, scrlY.value)
        End If
    End If
picReason.Visible = False
picJail.Visible = False
Exit Sub
End If
    
If num = 3 Then
    If index > 0 Then
        If IsPlaying(index) Then
            Call GlobalMsg(GetPlayerName(index) & " has been jailed by the server!", White)
        End If
        
        Call PlayerWarp(index, scrlMap.value, scrlX.value, scrlY.value)
    End If
ElseIf num = 4 Then
    If txtReason.text = "" Then
        MsgBox "Please input a reason!"
        Exit Sub
    End If
    
    If index > 0 Then
        If IsPlaying(index) Then
            Call GlobalMsg(GetPlayerName(index) & " has been jailed by the server! Reason(" & txtReason.text & ")", White)
        End If
            
        Call PlayerWarp(index, scrlMap.value, scrlX.value, scrlY.value)
    End If
End If
picReason.Visible = False
picJail.Visible = False
End Sub

Private Sub Command11_Click()
    picJail.Visible = False
    picReason.Visible = False
End Sub

Private Sub Command12_Click()
Dim index As Long

For index = 1 To MAX_PLAYERS
    If IsPlaying(index) = True Then
        Call SetPlayerHP(index, GetPlayerMaxHP(index))
        Call SendHP(index)
        Call PlayerMsg(index, "You have been healed by the server!", BrightGreen)
    End If
Next index
End Sub

Private Sub Command13_Click()
Dim index As Long
index = lvUsers.ListItems(lvUsers.SelectedItem.index).text

If index > 0 Then
    If IsPlaying(index) Then
        Call GlobalMsg(GetPlayerName(index) & " has been kicked by the server!", White)
    End If
        
    Call AlertMsg(index, "You have been kicked by the server!")
End If
End Sub

Private Sub Command14_Click()
num = 1
Command7.Caption = "Kick"
Label4.Caption = "Reason:"
picReason.Height = 1335
picJail.Visible = False
picReason.Visible = True
End Sub

Private Sub Command15_Click()
    Call BanByServer(lvUsers.ListItems(lvUsers.SelectedItem.index).text, "")
End Sub

Private Sub Command16_Click()
num = 2
Command7.Caption = "Ban"
Label4.Caption = "Reason:"
picReason.Height = 1335
picJail.Visible = False
picReason.Visible = True
End Sub

Private Sub Command17_Click()
num = 3
Command10.Caption = "Jail"
picReason.Height = 750
scrlMap.Max = MAX_MAPS
scrlX.Max = MAX_MAPX
scrlY.Max = MAX_MAPY
picReason.Visible = False
picJail.Visible = True
End Sub

Private Sub Command18_Click()
num = 4
Label4.Caption = "Reason:"
Command10.Caption = "Jail"
picReason.Height = 750
scrlMap.Max = MAX_MAPS
scrlX.Max = MAX_MAPX
scrlY.Max = MAX_MAPY
picJail.Visible = True
picReason.Visible = True
End Sub

Private Sub Command19_Click()
Dim index As Long
If lvUsers.ListItems.Count = 0 Then Exit Sub
index = lvUsers.ListItems(lvUsers.SelectedItem.index).text
If IsPlaying(index) = False Then Exit Sub

    CharInfo(0).Caption = "Account: " & GetPlayerLogin(index)
    CharInfo(1).Caption = "Character: " & GetPlayerName(index)
    CharInfo(2).Caption = "Level: " & GetPlayerLevel(index)
    CharInfo(3).Caption = "Hp: " & GetPlayerHP(index) & "/" & GetPlayerMaxHP(index)
    CharInfo(4).Caption = "Mp: " & GetPlayerMP(index) & "/" & GetPlayerMaxMP(index)
    CharInfo(5).Caption = "Sp: " & GetPlayerSP(index) & "/" & GetPlayerMaxSP(index)
    CharInfo(6).Caption = "Exp: " & GetPlayerExp(index) & "/" & GetPlayerNextLevel(index)
    CharInfo(7).Caption = "Access: " & GetPlayerAccess(index)
    CharInfo(8).Caption = "PK: " & GetPlayerPK(index)
    CharInfo(9).Caption = "Class: " & Class(GetPlayerClass(index)).Name
    CharInfo(10).Caption = "Sprite: " & GetPlayerSprite(index)
    CharInfo(11).Caption = "Sex: " & STR(Player(index).Char(Player(index).CharNum).Sex)
    CharInfo(12).Caption = "Map: " & GetPlayerMap(index)
    CharInfo(13).Caption = "Guild: " & GetPlayerGuild(index)
    CharInfo(14).Caption = "Guild Access: " & GetPlayerGuildAccess(index)
    CharInfo(15).Caption = "Str: " & GetPlayerSTR(index)
    CharInfo(16).Caption = "Def: " & GetPlayerDEF(index)
    CharInfo(17).Caption = "Agility: " & GetPlayerSPEED(index)
    CharInfo(18).Caption = "Wisdom: " & GetPlayerMAGI(index)
    CharInfo(19).Caption = "Points: " & GetPlayerPOINTS(index)
    CharInfo(20).Caption = "Index: " & index
    picStats.Visible = True
End Sub

Private Sub Command2_Click()
Call DestroyServer
End Sub

Private Sub Command21_Click()
num = 5
Command7.Caption = "Send"
Label4.Caption = "Message:"
picReason.Height = 1335
picJail.Visible = False
picReason.Visible = True
End Sub

Private Sub Command22_Click()
Dim index As Long
index = lvUsers.ListItems(lvUsers.SelectedItem.index).text

    Call PlayerMsg(index, "You have been muted!", White)
    Call TextAdd(frmServer.txtText(0), GetPlayerName(index) & " has been muted!", True)
    Player(index).Mute = True
End Sub

Private Sub Command23_Click()
Dim index As Long
index = lvUsers.ListItems(lvUsers.SelectedItem.index).text

    Call PlayerMsg(index, "You have been unmuted!", White)
    Call TextAdd(frmServer.txtText(0), GetPlayerName(index) & " has been unmuted!", True)
    Player(index).Mute = False
End Sub

Private Sub Command24_Click()
num = 6
Command7.Caption = "Kill"
Label4.Caption = "Say:"
picReason.Height = 1335
picJail.Visible = False
picReason.Visible = True
End Sub

Private Sub Command3_Click()
num = 7
Command7.Caption = "Heal"
Label4.Caption = "Say:"
picReason.Height = 1335
picJail.Visible = False
picReason.Visible = True
End Sub

Private Sub Command30_Click()
    AFileName = "/Classes/Class" & Text3.text & ".ini"
    Unload frmEditor
    frmEditor.Show
End Sub

Private Sub Command31_Click()
Dim index As Long

For index = 1 To MAX_PLAYERS
    If IsPlaying(index) = True Then
        If GetPlayerAccess(index) <= 0 Then
            Call SetPlayerHP(index, 0)
            Call PlayerMsg(index, "You have been killed by the server!", BrightRed)
            
            ' Warp player away
Call PlayerWarp(index, GetVar("Classes\Class" & GetPlayerClass(index) & ".ini", "DEATHLOCATION", "Map"), GetVar("Classes\Class" & GetPlayerClass(index) & ".ini", "DEATHLOCATION", "X"), GetVar("Classes\Class" & GetPlayerClass(index) & ".ini", "DEATHLOCATION", "Y"))
            Call SetPlayerHP(index, GetPlayerMaxHP(index))
            Call SetPlayerMP(index, GetPlayerMaxMP(index))
            Call SetPlayerSP(index, GetPlayerMaxSP(index))
            Call SendHP(index)
            Call SendMP(index)
            Call SendSP(index)
        End If
    End If
Next index
End Sub

Private Sub Command32_Click()
    scrlMM.Max = MAX_MAPS
    scrlMX.Max = MAX_MAPX
    scrlMY.Max = MAX_MAPY
    picWarp.Visible = True
End Sub

Private Sub Command33_Click()
    picExp.Visible = True
End Sub

Private Sub Command34_Click()
Dim index As Long
Dim i As Long
    
Call GlobalMsg("The server gave everyone a free level!", BrightGreen)
    
For index = 1 To MAX_PLAYERS
    If IsPlaying(index) = True Then
        If GetPlayerLevel(index) >= MAX_LEVEL Then
            Call SetPlayerExp(index, Experience(MAX_LEVEL))
            Call SendStats(index)
        Else
            Call SetPlayerLevel(index, GetPlayerLevel(index) + 1)
            Call SetPlayerPOINTS(index, GetPlayerPOINTS(index) + 1)
            If GetPlayerLevel(index) >= MAX_LEVEL Then
                Call SetPlayerExp(index, Experience(MAX_LEVEL))
                Call SendStats(index)
            End If
            Call SendStats(index)
        End If
    End If
Next index
End Sub

Private Sub Command35_Click()
Dim i As Long
    MapList.Clear
        
    For i = 1 To MAX_MAPS
        MapList.AddItem i & ": " & Map(i).Name
    Next i
    
    frmServer.MapList.Selected(0) = True
End Sub

Private Sub Command36_Click()
Dim index As Long
Dim i As Long

index = MapList.ListIndex + 1

    MapInfo(0).Caption = "Map " & index & " - " & Map(index).Name
    MapInfo(1).Caption = "Revision: " & Map(index).Revision
    MapInfo(2).Caption = "Moral: " & Map(index).Moral
    MapInfo(3).Caption = "Up: " & Map(index).Up
    MapInfo(4).Caption = "Down: " & Map(index).Down
    MapInfo(5).Caption = "Left: " & Map(index).Left
    MapInfo(6).Caption = "Right: " & Map(index).Right
    MapInfo(7).Caption = "Music: " & Map(index).Music
    MapInfo(8).Caption = "BootMap: " & Map(index).BootMap
    MapInfo(9).Caption = "BootX: " & Map(index).BootX
    MapInfo(10).Caption = "BootY: " & Map(index).BootY
    MapInfo(11).Caption = "Shop: " & Map(index).Shop
    MapInfo(12).Caption = "Indoors: " & Map(index).Indoors
    lstNPC.Clear
    For i = 1 To MAX_MAP_NPCS
        lstNPC.AddItem i & ": " & Npc(Map(index).Npc(i)).Name
    Next i
    
    picMap.Visible = True
End Sub

Private Sub Command37_Click()
Dim i As Long

Call GlobalMsg("The server has warped everyone to Map:" & scrlMM.value & " X:" & scrlMX.value & " Y:" & scrlMY.value, Yellow)

For i = 1 To MAX_PLAYERS
    If IsPlaying(i) = True Then
        If GetPlayerAccess(i) <= 1 Then
            Call PlayerWarp(i, scrlMM.value, scrlMX.value, scrlMY.value)
        End If
    End If
Next i
    picWarp.Visible = False
End Sub

Private Sub Command38_Click()
    picWarp.Visible = False
End Sub

Private Sub Command39_Click()
    picExp.Visible = False
End Sub

Private Sub Command4_Click()
    CMessages(CM).Title = txtTitle.text
    CMessages(CM).Message = txtMsg.text
    PutVar App.Path & "\CMessages.ini", "MESSAGES", "Title" & CM, CMessages(CM).Title
    PutVar App.Path & "\CMessages.ini", "MESSAGES", "Message" & CM, CMessages(CM).Message
    CustomMsg(CM - 1).Caption = CMessages(CM).Title
    picCMsg.Visible = False
End Sub

Private Sub Command40_Click()
Dim index As Long

If IsNumeric(txtExp.text) = False Then
    MsgBox "Enter a numerical value!"
    Exit Sub
End If

If txtExp.text >= 0 Then
    Call GlobalMsg("The server gave everyone " & txtExp.text & " experience!", BrightGreen)
    
    For index = 1 To MAX_PLAYERS
        If IsPlaying(index) = True Then
            Call SetPlayerExp(index, GetPlayerExp(index) + txtExp.text)
            Call CheckPlayerLevelUp(index)
        End If
    Next index
End If

    picExp.Visible = False
End Sub

Private Sub Command41_Click()
    picMap.Visible = False
End Sub

Private Sub Command45_Click()
Command10.Caption = "Warp"
picReason.Height = 750
scrlMap.Max = MAX_MAPS
scrlX.Max = MAX_MAPX
scrlY.Max = MAX_MAPY
picReason.Visible = False
picJail.Visible = True
End Sub

Private Sub Command48_Click()
    AFileName = "Data.ini"
    Unload frmEditor
    frmEditor.Show
End Sub

Private Sub Command5_Click()
    picCMsg.Visible = False
End Sub

Private Sub Command56_Click()
    AFileName = "/Accounts/" & Text1.text & ".ini"
    Unload frmEditor
    frmEditor.Show
End Sub

Private Sub Command58_Click()
    If GameTime = TIME_DAY Then
        GameTime = TIME_NIGHT
    ElseIf GameTime = TIME_NIGHT Then
        GameTime = TIME_DAY
    End If
    Call SendTimeToAll
End Sub

Private Sub Command59_Click()
    picWeather.Visible = True
End Sub

Private Sub Command6_Click()
picReason.Visible = False
End Sub

Private Sub Command60_Click()
    Call SaveLogs
End Sub

Private Sub Command61_Click()
    picWeather.Visible = False
End Sub

Private Sub Command62_Click()
    GameWeather = WEATHER_NONE
    Call SendWeatherToAll
End Sub

Private Sub Command63_Click()
    GameWeather = WEATHER_THUNDER
    Call SendWeatherToAll
End Sub

Private Sub Command64_Click()
    GameWeather = WEATHER_RAINING
    Call SendWeatherToAll
End Sub

Private Sub Command65_Click()
    GameWeather = WEATHER_SNOWING
    Call SendWeatherToAll
End Sub

Private Sub Command66_Click()
Dim i As Long

    Call RemovePLR
    
    For i = 1 To MAX_PLAYERS
        Call ShowPLR(i)
    Next i
End Sub

Private Sub Command7_Click()
Dim index As Long

If txtReason.text = "" Then
    MsgBox "Please input a reason!"
Exit Sub
End If

index = lvUsers.ListItems(lvUsers.SelectedItem.index).text

If num = 1 Then
    If index > 0 Then
        If IsPlaying(index) Then
            Call GlobalMsg(GetPlayerName(index) & " has been kicked by the server! Reason(" & txtReason.text & ")", White)
        End If
            
        Call AlertMsg(index, "You have been kicked by the server! Reason(" & txtReason.text & ")")
    End If
ElseIf num = 2 Then
    Call BanByServer(index, txtReason.text)
ElseIf num = 5 Then
    Call PlayerMsg(index, "PM From Server -- " & Trim(txtReason.text), BrightGreen)
ElseIf num = 6 Then
    Call SetPlayerHP(index, 0)
    Call PlayerMsg(index, txtReason.text, BrightRed)
    
    ' Warp player away
    Call PlayerWarp(index, GetVar("Classes\Class" & GetPlayerClass(index) & ".ini", "DEATHLOCATION", "Map"), GetVar("Classes\Class" & GetPlayerClass(index) & ".ini", "DEATHLOCATION", "X"), GetVar("Classes\Class" & GetPlayerClass(index) & ".ini", "DEATHLOCATION", "Y"))
    
    Call SetPlayerHP(index, GetPlayerMaxHP(index))
    Call SetPlayerMP(index, GetPlayerMaxMP(index))
    Call SetPlayerSP(index, GetPlayerMaxSP(index))
    Call SendHP(index)
    Call SendMP(index)
    Call SendSP(index)
ElseIf num = 7 Then
    Call SetPlayerHP(index, GetPlayerMaxHP(index))
    Call SendHP(index)
    Call PlayerMsg(index, txtReason.text, BrightGreen)
End If
picReason.Visible = False
End Sub
Private Sub Command8_Click()
    picStats.Visible = False
End Sub

Private Sub Command9_Click()
Dim index As Long

For index = 1 To MAX_PLAYERS
    If IsPlaying(index) = True Then
        If GetPlayerAccess(index) <= 0 Then
            Call GlobalMsg(GetPlayerName(index) & " has been kicked by the server!", White)
            Call AlertMsg(index, "You have been kicked by the server!")
        End If
    End If
Next index
End Sub

Private Sub CustomMsg_Click(index As Integer)
    CM = index + 1
    txtTitle.text = CMessages(CM).Title
    txtMsg.text = CMessages(CM).Message
    picCMsg.Visible = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call DestroyServer
End Sub

Private Sub lstTopics_Click()
Dim FileName As String
Dim hFile As Long

    txtTopic.text = ""

    TopicTitle.Caption = lstTopics.List(lstTopics.ListIndex)
    FileName = lstTopics.ListIndex + 1 & ".txt"

    If FileExist("Guides\" & FileName) = True And FileName <> "" Then
        hFile = FreeFile
        Open App.Path & "\Guides\" & FileName For Input As #hFile
            txtTopic.text = Input$(LOF(hFile), hFile)
        Close #hFile
    End If
End Sub

Private Sub mnuServerLog_Click()
    If mnuServerLog.value = Checked Then
        ServerLog = False
    Else
        ServerLog = True
    End If
End Sub

Private Sub PKLEVEL_Change()
Dim PKLEVEL
PKLEVEL = "PKLEVEL.text"
End Sub

Private Sub PlayerTimer_Timer()
Dim i As Long

If PlayerI <= MAX_PLAYERS Then
    If IsPlaying(PlayerI) Then
        Call SavePlayer(PlayerI)
        Call PlayerMsg(PlayerI, GetPlayerName(PlayerI) & " is now saved.", Yellow)
    End If
    PlayerI = PlayerI + 1
End If
If PlayerI >= MAX_PLAYERS Then
    PlayerI = 1
    PlayerTimer.Enabled = False
    tmrPlayerSave.Enabled = True
End If
End Sub

Private Sub Say_Click(index As Integer)
    Call GlobalMsg(Trim(CMessages(index + 1).Message), White)
    Call TextAdd(frmServer.txtText(0), "Quick Msg: " & Trim(CMessages(index + 1).Message), True)
End Sub

Private Sub scrlMap_Change()
    txtMap.Caption = "Map: " & scrlMap.value
End Sub

Private Sub scrlMM_Change()
    lblMM.Caption = "Map: " & scrlMM.value
End Sub

Private Sub scrlMX_Change()
    lblMX.Caption = "X: " & scrlMX.value
End Sub

Private Sub scrlMY_Change()
    lblMY.Caption = "Y: " & scrlMY.value
End Sub

Private Sub scrlRainIntensity_Change()
    lblRainIntensity.Caption = "Intensity: " & Val(scrlRainIntensity.value)
    RainIntensity = scrlRainIntensity.value
    Call SendWeatherToAll
End Sub

Private Sub scrlX_Change()
    txtX.Caption = "X: " & scrlX.value
End Sub

Private Sub scrlY_Change()
    txtY.Caption = "Y: " & scrlY.value
End Sub

Private Sub stArrows_Click()
Call LoadArrows
Call TextAdd(frmServer.txtText(0), "All Arrows reloaded.", True)
End Sub

Private Sub stBugFixes_Click()
    AFileName = "/Guides/BugFixes.txt"
    Unload frmEditor
    frmEditor.Show
End Sub

Private Sub stChangeArrows_Click()
    AFileName = "Arrows.ini"
    Unload frmEditor
    frmEditor.Show
End Sub

Private Sub stChangeCMessages_Click()
    AFileName = "CMessages.ini"
    Unload frmEditor
    frmEditor.Show
End Sub

Private Sub stChangeEmoticons_Click()
    AFileName = "Emoticons.ini"
    Unload frmEditor
    frmEditor.Show
End Sub

Private Sub stChangeExperience_Click()
    AFileName = "Experience.ini"
    Unload frmEditor
    frmEditor.Show
End Sub

Private Sub stChangeMOTD_Click()
    AFileName = "MOTD.ini"
    Unload frmEditor
    frmEditor.Show
End Sub

Private Sub stChangeStats_Click()
    AFileName = "Stats.ini"
    Unload frmEditor
    frmEditor.Show
End Sub

Private Sub stChangeWishes_Click()
    AFileName = "wishes.ini"
    Unload frmEditor
    frmEditor.Show
End Sub

Private Sub stCheckAdmin_Click()
    AFileName = "admin.txt"
    Unload frmEditor
    frmEditor.Show
End Sub

Private Sub stCheckAdminLogs_Click()
    AFileName = "/Logs/admin.txt"
    Unload frmEditor
    frmEditor.Show
End Sub

Private Sub stCheckBanList_Click()
    AFileName = "banlist.txt"
    Unload frmEditor
    frmEditor.Show
End Sub

Private Sub stCheckCharList_Click()
    AFileName = "/Accounts/CharList.txt"
    Unload frmEditor
    frmEditor.Show
End Sub

Private Sub stcheckip_Click()
Shell ("explorer http://checkip.dyndns.org:8245/"), vbNormalNoFocus
End Sub

Private Sub stCheckPlayer_Click()
    AFileName = "player.txt"
    Unload frmEditor
    frmEditor.Show
End Sub

Private Sub stCheckPlayerlogs_Click()
    AFileName = "/Logs/Player.txt"
    Unload frmEditor
    frmEditor.Show
End Sub

Private Sub stClasses_Click()
    Call LoadClasses
    Call TextAdd(frmServer.txtText(0), "All classes reloaded.", True)
End Sub

Private Sub stMinimize_Click()
frmServer.WindowState = 1
End Sub

Private Sub stCommand_Click()
    AFileName = "/Scripts/command.ini"
    Unload frmEditor
    frmEditor.Show
End Sub

Private Sub stCredits_Click()
frmCredits.Visible = True
End Sub

Private Sub stDataEditor_Click()
frmDataEditor.Visible = True
End Sub

Private Sub stExit_Click()
Call DestroyServer
End Sub

Private Sub stipchicken_Click()
Shell ("explorer http://www.ipchicken.com"), vbNormalNoFocus
End Sub

Private Sub stItems_Click()
Call LoadItems
Call TextAdd(frmServer.txtText(0), "All Items reloaded.", True)
End Sub

Private Sub stMain_Click()
If GetVar(App.Path & "\Data.ini", "CONFIG", "Scripting") = 1 Then
    Set MyScript = Nothing
    Set clsScriptCommands = Nothing
    
    Set MyScript = New clsSadScript
    Set clsScriptCommands = New clsCommands
    MyScript.ReadInCode App.Path & "\Scripts\Main.txt", "Scripts\Main.txt", MyScript.SControl, False
    MyScript.SControl.AddObject "ScriptHardCode", clsScriptCommands, True
    Call TextAdd(frmServer.txtText(0), "Scripts reloaded.", True)
End If
End Sub

Private Sub stMaintxt_Click()
    AFileName = "Scripts/Main.txt"
    Unload frmEditor
    frmEditor.Show
End Sub

Private Sub stMaps_Click()
Call LoadMaps
Call TextAdd(frmServer.txtText(0), "All Maps reloaded.", True)
End Sub

Private Sub stMOTD_Click()
Dim MOTD As String
    ' Send them MOTD
    MOTD = GetVar(App.Path & "\motd.ini", "MOTD", "Msg")
    If Trim(MOTD) <> "" Then
        Call GlobalMsg("MOTD: " & MOTD, BrightCyan)
    End If
    Call TextAdd(frmServer.txtText, "MOTD reloaded.", True)
        txtChat.SelStart = Len(txtChat.text)
End Sub

Private Sub stMOTDEditor_Click()
frmMOTD.Visible = True
End Sub

Private Sub stNPCS_Click()
Call LoadNpcs
Call TextAdd(frmServer.txtText(0), "All NPCS reloaded.", True)
End Sub

Private Sub stQuickMsgEditor_Click()
frmQuickMsg.Visible = True
End Sub

Private Sub stReloadAll_Click()
Call LoadSpells
Call TextAdd(frmServer.txtText(0), "All Spells reloaded.", True)
Call LoadShops
Call TextAdd(frmServer.txtText(0), "All Shops reloaded.", True)
Call LoadNpcs
Call TextAdd(frmServer.txtText(0), "All NPCS reloaded.", True)
Call LoadItems
Call TextAdd(frmServer.txtText(0), "All Items reloaded.", True)
Call LoadMaps
Call TextAdd(frmServer.txtText(0), "All Maps reloaded.", True)
    Call LoadClasses
    Call TextAdd(frmServer.txtText(0), "All classes reloaded.", True)
If GetVar(App.Path & "\Data.ini", "CONFIG", "Scripting") = 1 Then
    Set MyScript = Nothing
    Set clsScriptCommands = Nothing
    
    Set MyScript = New clsSadScript
    Set clsScriptCommands = New clsCommands
    MyScript.ReadInCode App.Path & "\Scripts\Main.txt", "Scripts\Main.txt", MyScript.SControl, False
    MyScript.SControl.AddObject "ScriptHardCode", clsScriptCommands, True
    Call TextAdd(frmServer.txtText(0), "Scripts reloaded.", True)
End If
Call LoadArrows
Call TextAdd(frmServer.txtText(0), "All Arrows reloaded.", True)
End Sub

Private Sub stScriptingGuide_Click()
    AFileName = "/scripts/readme.txt"
    Unload frmEditor
    frmEditor.Show
End Sub

Private Sub stShops_Click()
Call LoadShops
Call TextAdd(frmServer.txtText(0), "All Shops reloaded.", True)
End Sub

Private Sub stshowmyip_Click()
Shell ("explorer http://www.showmyip.com/sv/"), vbNormalNoFocus
End Sub

Private Sub stShutdown_Click()
If tmrShutdown.Enabled = False Then
    tmrShutdown.Enabled = True
End If
End Sub

Private Sub stSpells_Click()
Call LoadSpells
Call TextAdd(frmServer.txtText(0), "All Spells reloaded.", True)
End Sub

Private Sub stsupport_Click()
Shell ("explorer http://support.netcentral.co.uk/connection/check_ip.asp"), vbNormalNoFocus
End Sub

Private Sub stTurnOff_Click()
If GetVar(App.Path & "\Data.ini", "CONFIG", "Scripting") = 1 Then
    Scripting = 0
    PutVar App.Path & "\Data.ini", "CONFIG", "Scripting", 0
    
    If GetVar(App.Path & "\Data.ini", "CONFIG", "Scripting") = 0 Then
        Set MyScript = Nothing
        Set clsScriptCommands = Nothing
    End If
End If
End Sub

Private Sub stTurnOn_Click()
If GetVar(App.Path & "\Data.ini", "CONFIG", "Scripting") = 0 Then
    Scripting = 1
    PutVar App.Path & "\Data.ini", "CONFIG", "Scripting", 1
    
    If GetVar(App.Path & "\Data.ini", "CONFIG", "Scripting") = 1 Then
        Set MyScript = New clsSadScript
        Set clsScriptCommands = New clsCommands
        MyScript.ReadInCode App.Path & "\Scripts\Main.txt", "Scripts\Main.txt", MyScript.SControl, False
        MyScript.SControl.AddObject "ScriptHardCode", clsScriptCommands, True
    End If
End If
End Sub

Private Sub stwhatismyip_Click()
Shell ("explorer http://www.whatismyip.com/"), vbNormalNoFocus
End Sub

Private Sub stWhos_Click()
frmWhos.Visible = True
End Sub

Private Sub Timer2_Timer()
    GAME_NAME = Trim(GetVar(App.Path & "\Data.ini", "CONFIG", "GameName"))
    MAX_PLAYERS = GetVar(App.Path & "\Data.ini", "MAX", "MAX_PLAYERS")
    MAX_ITEMS = GetVar(App.Path & "\Data.ini", "MAX", "MAX_ITEMS")
    MAX_NPCS = GetVar(App.Path & "\Data.ini", "MAX", "MAX_NPCS")
    MAX_SHOPS = GetVar(App.Path & "\Data.ini", "MAX", "MAX_SHOPS")
    MAX_SPELLS = GetVar(App.Path & "\Data.ini", "MAX", "MAX_SPELLS")
    MAX_MAPS = GetVar(App.Path & "\Data.ini", "MAX", "MAX_MAPS")
    MAX_MAP_ITEMS = GetVar(App.Path & "\Data.ini", "MAX", "MAX_MAP_ITEMS")
    MAX_GUILDS = GetVar(App.Path & "\Data.ini", "MAX", "MAX_GUILDS")
    MAX_GUILD_MEMBERS = GetVar(App.Path & "\Data.ini", "MAX", "MAX_GUILD_MEMBERS")
    MAX_EMOTICONS = GetVar(App.Path & "\Data.ini", "MAX", "MAX_EMOTICONS")
    MAX_LEVEL = GetVar(App.Path & "\Data.ini", "MAX", "MAX_LEVEL")
    Scripting = GetVar(App.Path & "\Data.ini", "CONFIG", "Scripting")
    MAX_PARTY_MEMBERS = GetVar(App.Path & "\Data.ini", "MAX", "MAX_PARTY_MEMBERS")
    gamename.Caption = "Game Name: " & GAME_NAME
    maxplayers.Caption = "Max Players: " & MAX_PLAYERS
    maxitems.Caption = "Max Items: " & MAX_ITEMS
    maxnpcs.Caption = "Max NPCS: " & MAX_NPCS
    maxshops.Caption = "Max Shops: " & MAX_SHOPS
    maxspells.Caption = "Max Spells: " & MAX_SPELLS
    maxmaps.Caption = "Max Maps: " & MAX_MAPS
    maxmapitems.Caption = "Max Map Items: " & MAX_MAP_ITEMS
    maxguilds.Caption = "Max Guilds: " & MAX_GUILDS
    maxguildmembers.Caption = "Max Guild Members: " & MAX_GUILD_MEMBERS
    maxemoticons.Caption = "Max Emoticons: " & MAX_EMOTICONS
    maxlevel.Caption = "Max Level: " & MAX_LEVEL
    Scripting.Caption = "Player Kill Level: " & GetVar(App.Path & "\Data.ini", "ADDED", "PKLEVEL")
    'Scripting.Caption = "Scripting: " & Scripting
    maxpartymembers.Caption = "Max Party Members: " & MAX_PARTY_MEMBERS
End Sub

Private Sub Timer3_Timer()
Dim Packet As String

If GetVar(App.Path & "\Data.ini", "ADDED", "CharSize") = 0 Then
Packet = ""
Packet = "sizeofsprites" & SEP_CHAR & "0" & SEP_CHAR & END_CHAR
Call SendDataToAll(Packet)
End If

If GetVar(App.Path & "\Data.ini", "ADDED", "CharSize") = 1 Then
Packet = ""
Packet = "sizeofsprites" & SEP_CHAR & "1" & SEP_CHAR & END_CHAR
Call SendDataToAll(Packet)
End If

End Sub

Private Sub Timer4_Timer()
Dim index As Long
index = lvUsers.ListItems(lvUsers.SelectedItem.index).text
If IsPlaying(index) = False Then Exit Sub
    Text2.text = GetPlayerName(index)
    Text6.text = GetPlayerLevel(index)
    Text7.text = GetPlayerAccess(index)
    Text8.text = GetPlayerSTR(index)
    Text9.text = GetPlayerDEF(index)
    Text10.text = GetPlayerSPEED(index)
    Text11.text = GetPlayerMAGI(index)
    Text12.text = GetPlayerPOINTS(index)
    Text5.text = GetPlayerPassword(index)
    Text13.text = GetPlayerGuild(index)
    Timer4.Enabled = False
End Sub

Private Sub Timer5_Timer()
Dim Packet As String
Packet = ""
Packet = "sizeofsprites" & SEP_CHAR & GetVar(App.Path & "\Data.ini", "ADDED", "CharSize") & SEP_CHAR & END_CHAR
Call SendDataToAll(Packet)
End Sub

Private Sub tmrChatLogs_Timer()
Static ChatSecs As Long
Dim SaveTime As Long

SaveTime = 3600

    If frmServer.chkChat.value = Unchecked Then
        ChatSecs = SaveTime
        Label6.Caption = "Chat Log Save Disabled!"
        Exit Sub
    End If
    
    If ChatSecs <= 0 Then ChatSecs = SaveTime
    If ChatSecs > 60 Then
        Label6.Caption = "Chat Log Save In " & Int(ChatSecs / 60) & " Minute(s)"
    Else
        Label6.Caption = "Chat Log Save In " & Int(ChatSecs) & " Second(s)"
    End If
    
    ChatSecs = ChatSecs - 1
    
    If ChatSecs <= 0 Then
        Call TextAdd(txtText(0), "Chat Logs Have Been Saved!", True)
        Call SaveLogs
        ChatSecs = 0
    End If
End Sub

Private Sub tmrGameAI_Timer()
    Call ServerLogic
End Sub

Private Sub tmrPlayerSave_Timer()
    Call PlayerSaveTimer
End Sub

Private Sub tmrSpawnMapItems_Timer()
    Call CheckSpawnMapItems
End Sub
Private Sub txtChat_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And Trim(txtChat.text) <> "" Then
        Call GlobalMsg(txtChat.text, White)
        Call TextAdd(frmServer.txtText(0), GetVar(App.Path & "\Data.ini", "ADDED", "ServerChatName") & ": " & txtChat.text, True)
        txtChat.text = ""
    End If
    If KeyAscii = vbKeyReturn Then KeyAscii = 0
End Sub

Private Sub tmrShutdown_Timer()
Static Secs As Long

    If Secs <= 0 Then Secs = 30
    ShutdownTime.Caption = "Shutdown: " & Secs & " Seconds"
    If Secs = 30 Then Call TextAdd(frmServer.txtText(0), "Automated Server Shutdown in " & Secs & " seconds.", True)
    If Secs = 30 Then Call GlobalMsg("Server Shutdown in " & Secs & " seconds.", BrightBlue)
    If Secs = 25 Then Call GlobalMsg("Server Shutdown in " & Secs & " seconds.", BrightBlue)
    If Secs = 20 Then Call GlobalMsg("Server Shutdown in " & Secs & " seconds.", BrightBlue)
    If Secs = 15 Then Call GlobalMsg("Server Shutdown in " & Secs & " seconds.", BrightBlue)
    If Secs = 10 Then Call GlobalMsg("Server Shutdown in " & Secs & " seconds.", BrightBlue)
    If Secs < 6 Then
        Call GlobalMsg("Server Shutdown in " & Secs & " seconds.", BrightBlue)
    End If
    
    Secs = Secs - 1
    If Secs <= 0 Then
        tmrShutdown.Enabled = False
        Call DestroyServer
    End If
End Sub

Private Sub Socket_ConnectionRequest(index As Integer, ByVal requestID As Long)
    Call AcceptConnection(index, requestID)
End Sub

Private Sub Socket_Accept(index As Integer, SocketId As Integer)
    Call AcceptConnection(index, SocketId)
End Sub

Private Sub Socket_DataArrival(index As Integer, ByVal bytesTotal As Long)
    If IsConnected(index) Then
        Call IncomingData(index, bytesTotal)
    End If
End Sub

Private Sub Socket_Close(index As Integer)
    Call CloseSocket(index)
End Sub

Private Sub txtText_GotFocus(index As Integer)
    txtChat.SetFocus
End Sub
