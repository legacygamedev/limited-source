VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCN.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL3N.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT3N.OCX"
Begin VB.Form frmServer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Konfuze Server"
   ClientHeight    =   4950
   ClientLeft      =   180
   ClientTop       =   720
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
   ScaleHeight     =   330
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   705
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   8281
      _Version        =   393216
      Tabs            =   4
      Tab             =   2
      TabsPerRow      =   4
      TabHeight       =   370
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
      TabPicture(0)   =   "frmServer.frx":0FC2
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "tmrChatLogs"
      Tab(0).Control(1)=   "picCMsg"
      Tab(0).Control(2)=   "SSTab2"
      Tab(0).Control(3)=   "Say(5)"
      Tab(0).Control(4)=   "Say(4)"
      Tab(0).Control(5)=   "Say(3)"
      Tab(0).Control(6)=   "Say(2)"
      Tab(0).Control(7)=   "Say(1)"
      Tab(0).Control(8)=   "CustomMsg(5)"
      Tab(0).Control(9)=   "CustomMsg(4)"
      Tab(0).Control(10)=   "CustomMsg(3)"
      Tab(0).Control(11)=   "CustomMsg(2)"
      Tab(0).Control(12)=   "CustomMsg(1)"
      Tab(0).Control(13)=   "Frame5"
      Tab(0).Control(14)=   "Say(0)"
      Tab(0).Control(15)=   "CustomMsg(0)"
      Tab(0).Control(16)=   "Label6"
      Tab(0).ControlCount=   17
      TabCaption(1)   =   "Players"
      TabPicture(1)   =   "frmServer.frx":0FDE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Command45"
      Tab(1).Control(1)=   "picJail"
      Tab(1).Control(2)=   "picStats"
      Tab(1).Control(3)=   "picReason"
      Tab(1).Control(4)=   "Command3"
      Tab(1).Control(5)=   "Command24"
      Tab(1).Control(6)=   "Command23"
      Tab(1).Control(7)=   "Command22"
      Tab(1).Control(8)=   "Command21"
      Tab(1).Control(9)=   "Command20"
      Tab(1).Control(10)=   "Command19"
      Tab(1).Control(11)=   "Command18"
      Tab(1).Control(12)=   "Command17"
      Tab(1).Control(13)=   "Command16"
      Tab(1).Control(14)=   "Command15"
      Tab(1).Control(15)=   "Command14"
      Tab(1).Control(16)=   "Command13"
      Tab(1).Control(17)=   "Check1"
      Tab(1).Control(18)=   "Command66"
      Tab(1).Control(19)=   "lvUsers"
      Tab(1).Control(20)=   "TPO"
      Tab(1).ControlCount=   21
      TabCaption(2)   =   "Control Panel"
      TabPicture(2)   =   "frmServer.frx":0FFA
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "lblPort"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "lblIP"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label7"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Frame7"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Socket(0)"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Frame1"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Frame2"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "Frame3"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "Frame4"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "PlayerTimer"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "tmrShutdown"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "tmrGameAI"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "tmrSpawnMapItems"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "tmrPlayerSave"
      Tab(2).Control(13).Enabled=   0   'False
      Tab(2).Control(14)=   "Frame6"
      Tab(2).Control(14).Enabled=   0   'False
      Tab(2).Control(15)=   "picWarp"
      Tab(2).Control(15).Enabled=   0   'False
      Tab(2).Control(16)=   "Frame9"
      Tab(2).Control(16).Enabled=   0   'False
      Tab(2).Control(17)=   "picExp"
      Tab(2).Control(17).Enabled=   0   'False
      Tab(2).Control(18)=   "picMap"
      Tab(2).Control(18).Enabled=   0   'False
      Tab(2).Control(19)=   "picWeather"
      Tab(2).Control(19).Enabled=   0   'False
      Tab(2).Control(20)=   "Timer1"
      Tab(2).Control(20).Enabled=   0   'False
      Tab(2).Control(21)=   "Time"
      Tab(2).Control(21).Enabled=   0   'False
      Tab(2).ControlCount=   22
      TabCaption(3)   =   "Help"
      TabPicture(3)   =   "frmServer.frx":1016
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "TopicTitle"
      Tab(3).Control(1)=   "lstTopics"
      Tab(3).Control(2)=   "CharInfo(23)"
      Tab(3).Control(3)=   "CharInfo(22)"
      Tab(3).Control(4)=   "CharInfo(21)"
      Tab(3).ControlCount=   5
      Begin VB.Frame Time 
         Caption         =   "Time"
         Height          =   735
         Left            =   240
         TabIndex        =   188
         Top             =   3840
         Width           =   9975
         Begin VB.CommandButton Command69 
            Caption         =   "Disable Time"
            Height          =   255
            Left            =   7320
            TabIndex        =   196
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton Command68 
            Caption         =   "Change Speed"
            Height          =   255
            Left            =   5880
            TabIndex        =   193
            Top             =   240
            Width           =   1335
         End
         Begin VB.TextBox GameTimeSpeed 
            Height          =   285
            Left            =   5280
            TabIndex        =   192
            Text            =   "1"
            Top             =   240
            Width           =   495
         End
         Begin VB.CommandButton Command67 
            Caption         =   "Randomize!"
            Height          =   255
            Left            =   2520
            TabIndex        =   190
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            Caption         =   "0:00:00"
            Height          =   255
            Left            =   8280
            TabIndex        =   195
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label Label10 
            Alignment       =   2  'Center
            Caption         =   "Time until (...):"
            Height          =   255
            Left            =   8280
            TabIndex        =   194
            Top             =   120
            Width           =   1575
         End
         Begin VB.Label Label9 
            Alignment       =   2  'Center
            Caption         =   "Game Time Speed:"
            Height          =   255
            Left            =   3720
            TabIndex        =   191
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label8 
            Height          =   255
            Left            =   240
            TabIndex        =   189
            Top             =   240
            Width           =   2175
         End
      End
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   6960
         Top             =   0
      End
      Begin VB.Timer tmrChatLogs 
         Interval        =   1000
         Left            =   -65160
         Top             =   360
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
         Left            =   6960
         ScaleHeight     =   135
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   215
         TabIndex        =   176
         Top             =   1680
         Visible         =   0   'False
         Width           =   3255
         Begin VB.CommandButton Command65 
            Caption         =   "Snow"
            Height          =   255
            Left            =   1680
            TabIndex        =   184
            Top             =   1320
            Width           =   1335
         End
         Begin VB.CommandButton Command64 
            Caption         =   "Rain"
            Height          =   255
            Left            =   240
            TabIndex        =   183
            Top             =   1320
            Width           =   1335
         End
         Begin VB.CommandButton Command63 
            Caption         =   "Thunder"
            Height          =   255
            Left            =   1680
            TabIndex        =   182
            Top             =   1080
            Width           =   1335
         End
         Begin VB.CommandButton Command62 
            Caption         =   "None"
            Height          =   255
            Left            =   240
            TabIndex        =   181
            Top             =   1080
            Width           =   1335
         End
         Begin VB.HScrollBar scrlRainIntensity 
            Height          =   255
            Left            =   120
            Max             =   50
            Min             =   1
            TabIndex        =   179
            Top             =   360
            Value           =   25
            Width           =   2895
         End
         Begin VB.CommandButton Command61 
            Caption         =   "Cancel"
            Height          =   255
            Left            =   1560
            TabIndex        =   177
            Top             =   1680
            Width           =   1575
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Current Weather: None"
            Height          =   195
            Left            =   120
            TabIndex        =   180
            Top             =   720
            Width           =   1710
         End
         Begin VB.Label lblRainIntensity 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Intensity: 25"
            Height          =   195
            Left            =   120
            TabIndex        =   178
            Top             =   120
            Width           =   930
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
         Left            =   -69840
         ScaleHeight     =   1905
         ScaleWidth      =   3345
         TabIndex        =   50
         Top             =   360
         Visible         =   0   'False
         Width           =   3375
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
            TabIndex        =   56
            Top             =   960
            Width           =   3075
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
            TabIndex        =   55
            Top             =   360
            Width           =   3075
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Cancel"
            Height          =   255
            Left            =   1680
            TabIndex        =   52
            Top             =   1560
            Width           =   1575
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Save"
            Height          =   255
            Left            =   1680
            TabIndex        =   51
            Top             =   1320
            Width           =   1575
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Message:"
            Height          =   195
            Left            =   120
            TabIndex        =   54
            Top             =   720
            Width           =   690
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Title:"
            Height          =   195
            Left            =   120
            TabIndex        =   53
            Top             =   120
            Width           =   360
         End
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   3015
         Left            =   -74880
         TabIndex        =   165
         Top             =   360
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   5318
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
         TabPicture(0)   =   "frmServer.frx":1032
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "txtText(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "txtChat"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Broadcast"
         TabPicture(1)   =   "frmServer.frx":104E
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "txtText(1)"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Global"
         TabPicture(2)   =   "frmServer.frx":106A
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "txtText(2)"
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "Map"
         TabPicture(3)   =   "frmServer.frx":1086
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "txtText(3)"
         Tab(3).ControlCount=   1
         TabCaption(4)   =   "Private"
         TabPicture(4)   =   "frmServer.frx":10A2
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "txtText(4)"
         Tab(4).ControlCount=   1
         TabCaption(5)   =   "Admin"
         TabPicture(5)   =   "frmServer.frx":10BE
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "txtText(5)"
         Tab(5).ControlCount=   1
         TabCaption(6)   =   "Emote"
         TabPicture(6)   =   "frmServer.frx":10DA
         Tab(6).ControlEnabled=   0   'False
         Tab(6).Control(0)=   "txtText(6)"
         Tab(6).ControlCount=   1
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
            TabIndex        =   173
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
            TabIndex        =   172
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
            TabIndex        =   171
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
            TabIndex        =   170
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
            TabIndex        =   169
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
            Index           =   1
            Left            =   -74880
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   168
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
            TabIndex        =   167
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
            Height          =   2250
            Index           =   0
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   166
            Top             =   360
            Width           =   8115
         End
      End
      Begin VB.CommandButton Command45 
         Caption         =   "Warp"
         Height          =   255
         Left            =   -66600
         TabIndex        =   160
         Top             =   3600
         Width           =   1575
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
         Left            =   3600
         ScaleHeight     =   223
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   223
         TabIndex        =   143
         Top             =   360
         Visible         =   0   'False
         Width           =   3375
         Begin VB.ListBox lstNPC 
            Height          =   2400
            Left            =   1680
            TabIndex        =   158
            Top             =   480
            Width           =   1575
         End
         Begin VB.CommandButton Command41 
            Caption         =   "Cancel"
            Height          =   255
            Left            =   1680
            TabIndex        =   144
            Top             =   3000
            Width           =   1575
         End
         Begin VB.Label MapInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "NPCs"
            Height          =   195
            Index           =   13
            Left            =   1680
            TabIndex        =   159
            Top             =   285
            Width           =   375
         End
         Begin VB.Label MapInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Indoors:"
            Height          =   195
            Index           =   12
            Left            =   120
            TabIndex        =   157
            Top             =   3000
            Width           =   615
         End
         Begin VB.Label MapInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Shop:"
            Height          =   195
            Index           =   11
            Left            =   120
            TabIndex        =   156
            Top             =   2760
            Width           =   420
         End
         Begin VB.Label MapInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "BootY:"
            Height          =   195
            Index           =   10
            Left            =   120
            TabIndex        =   155
            Top             =   2520
            Width           =   480
         End
         Begin VB.Label MapInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "BootX:"
            Height          =   195
            Index           =   9
            Left            =   120
            TabIndex        =   154
            Top             =   2280
            Width           =   480
         End
         Begin VB.Label MapInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "BootMap:"
            Height          =   195
            Index           =   8
            Left            =   120
            TabIndex        =   153
            Top             =   2040
            Width           =   690
         End
         Begin VB.Label MapInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Music:"
            Height          =   195
            Index           =   7
            Left            =   120
            TabIndex        =   152
            Top             =   1800
            Width           =   450
         End
         Begin VB.Label MapInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Right:"
            Height          =   195
            Index           =   6
            Left            =   120
            TabIndex        =   151
            Top             =   1560
            Width           =   435
         End
         Begin VB.Label MapInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Left:"
            Height          =   195
            Index           =   5
            Left            =   120
            TabIndex        =   150
            Top             =   1320
            Width           =   345
         End
         Begin VB.Label MapInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Down:"
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   149
            Top             =   1080
            Width           =   465
         End
         Begin VB.Label MapInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Up:"
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   148
            Top             =   840
            Width           =   255
         End
         Begin VB.Label MapInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Moral:"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   147
            Top             =   600
            Width           =   450
         End
         Begin VB.Label MapInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Revision:"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   146
            Top             =   360
            Width           =   660
         End
         Begin VB.Label MapInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Map"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   145
            Top             =   120
            Width           =   300
         End
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
         Left            =   6960
         ScaleHeight     =   87
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   215
         TabIndex        =   136
         Top             =   360
         Visible         =   0   'False
         Width           =   3255
         Begin VB.CommandButton Command39 
            Caption         =   "Cancel"
            Height          =   255
            Left            =   1560
            TabIndex        =   140
            Top             =   960
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
            TabIndex        =   138
            Top             =   360
            Width           =   2955
         End
         Begin VB.CommandButton Command40 
            Caption         =   "Execute"
            Height          =   255
            Left            =   1560
            TabIndex        =   137
            Top             =   720
            Width           =   1575
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Experience:"
            Height          =   195
            Left            =   120
            TabIndex        =   139
            Top             =   120
            Width           =   855
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Map List"
         Height          =   1815
         Left            =   6000
         TabIndex        =   141
         Top             =   480
         Width           =   4095
         Begin VB.ListBox MapList 
            Height          =   1425
            Left            =   120
            TabIndex        =   142
            Top             =   240
            Width           =   3855
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
         Left            =   240
         ScaleHeight     =   167
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   223
         TabIndex        =   109
         Top             =   360
         Visible         =   0   'False
         Width           =   3375
         Begin VB.CommandButton Command38 
            Caption         =   "Cancel"
            Height          =   255
            Left            =   1680
            TabIndex        =   119
            Top             =   2160
            Width           =   1575
         End
         Begin VB.HScrollBar scrlMY 
            Height          =   255
            Left            =   120
            TabIndex        =   113
            Top             =   1560
            Width           =   3135
         End
         Begin VB.HScrollBar scrlMX 
            Height          =   255
            Left            =   120
            TabIndex        =   112
            Top             =   960
            Width           =   3135
         End
         Begin VB.HScrollBar scrlMM 
            Height          =   255
            Left            =   120
            Min             =   1
            TabIndex        =   111
            Top             =   360
            Value           =   1
            Width           =   3135
         End
         Begin VB.CommandButton Command37 
            Caption         =   "Warp"
            Height          =   255
            Left            =   1680
            TabIndex        =   110
            Top             =   1920
            Width           =   1575
         End
         Begin VB.Label lblMY 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Y: 0"
            Height          =   195
            Left            =   120
            TabIndex        =   116
            Top             =   1320
            Width           =   285
         End
         Begin VB.Label lblMX 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "X: 0"
            Height          =   195
            Left            =   120
            TabIndex        =   115
            Top             =   720
            Width           =   285
         End
         Begin VB.Label lblMM 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Map: 1"
            Height          =   195
            Left            =   120
            TabIndex        =   114
            Top             =   120
            Width           =   495
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Commands"
         Height          =   2415
         Left            =   2160
         TabIndex        =   100
         Top             =   480
         Width           =   1815
         Begin VB.CommandButton Command36 
            Caption         =   "Map Info"
            Height          =   255
            Left            =   120
            TabIndex        =   108
            Top             =   2040
            Width           =   1575
         End
         Begin VB.CommandButton Command35 
            Caption         =   "Map List"
            Height          =   255
            Left            =   120
            TabIndex        =   107
            Top             =   1800
            Width           =   1575
         End
         Begin VB.CommandButton Command34 
            Caption         =   "Mass Level"
            Height          =   255
            Left            =   120
            TabIndex        =   106
            Top             =   1560
            Width           =   1575
         End
         Begin VB.CommandButton Command33 
            Caption         =   "Mass Experience"
            Height          =   255
            Left            =   120
            TabIndex        =   105
            Top             =   1320
            Width           =   1575
         End
         Begin VB.CommandButton Command32 
            Caption         =   "Mass Warp"
            Height          =   255
            Left            =   120
            TabIndex        =   104
            Top             =   1080
            Width           =   1575
         End
         Begin VB.CommandButton Command12 
            Caption         =   "Mass Heal"
            Height          =   255
            Left            =   120
            TabIndex        =   103
            Top             =   840
            Width           =   1575
         End
         Begin VB.CommandButton Command31 
            Caption         =   "Mass Kill"
            Height          =   255
            Left            =   120
            TabIndex        =   102
            Top             =   600
            Width           =   1575
         End
         Begin VB.CommandButton Command9 
            Caption         =   "Mass Kick"
            Height          =   255
            Left            =   120
            TabIndex        =   101
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.Frame TopicTitle 
         Caption         =   "Topic Title"
         Height          =   3375
         Left            =   -72480
         TabIndex        =   96
         Top             =   480
         Width           =   7575
         Begin VB.TextBox txtTopic 
            Height          =   3015
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   97
            Top             =   240
            Width           =   7335
         End
      End
      Begin VB.ListBox lstTopics 
         Height          =   2790
         Left            =   -74760
         TabIndex        =   94
         Top             =   600
         Width           =   2175
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
         Left            =   -70080
         ScaleHeight     =   167
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   223
         TabIndex        =   86
         Top             =   1200
         Visible         =   0   'False
         Width           =   3375
         Begin VB.CommandButton Command11 
            Caption         =   "Cancel"
            Height          =   255
            Left            =   1680
            TabIndex        =   118
            Top             =   2160
            Width           =   1575
         End
         Begin VB.HScrollBar scrlY 
            Height          =   255
            Left            =   120
            TabIndex        =   90
            Top             =   1560
            Width           =   3135
         End
         Begin VB.HScrollBar scrlX 
            Height          =   255
            Left            =   120
            TabIndex        =   89
            Top             =   960
            Width           =   3135
         End
         Begin VB.HScrollBar scrlMap 
            Height          =   255
            Left            =   120
            Min             =   1
            TabIndex        =   88
            Top             =   360
            Value           =   1
            Width           =   3135
         End
         Begin VB.CommandButton Command10 
            Caption         =   "Jail"
            Height          =   255
            Left            =   1680
            TabIndex        =   87
            Top             =   1920
            Width           =   1575
         End
         Begin VB.Label txtY 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Y: 0"
            Height          =   195
            Left            =   120
            TabIndex        =   93
            Top             =   1320
            Width           =   285
         End
         Begin VB.Label txtX 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "X: 0"
            Height          =   195
            Left            =   120
            TabIndex        =   92
            Top             =   720
            Width           =   285
         End
         Begin VB.Label txtMap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Map: 1"
            Height          =   195
            Left            =   120
            TabIndex        =   91
            Top             =   120
            Width           =   495
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
         Left            =   -74760
         ScaleHeight     =   215
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   311
         TabIndex        =   63
         Top             =   480
         Visible         =   0   'False
         Width           =   4695
         Begin VB.CommandButton Command8 
            Caption         =   "Cancel"
            Height          =   255
            Left            =   3000
            TabIndex        =   64
            Top             =   2880
            Width           =   1575
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Index:"
            Height          =   195
            Index           =   20
            Left            =   2400
            TabIndex        =   85
            Top             =   1800
            Width           =   480
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Points:"
            Height          =   195
            Index           =   19
            Left            =   2400
            TabIndex        =   84
            Top             =   1560
            Width           =   495
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Magi:"
            Height          =   195
            Index           =   18
            Left            =   2400
            TabIndex        =   83
            Top             =   1320
            Width           =   390
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Speed:"
            Height          =   195
            Index           =   17
            Left            =   2400
            TabIndex        =   82
            Top             =   1080
            Width           =   510
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Def:"
            Height          =   195
            Index           =   16
            Left            =   2400
            TabIndex        =   81
            Top             =   840
            Width           =   315
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Str:"
            Height          =   195
            Index           =   15
            Left            =   2400
            TabIndex        =   80
            Top             =   600
            Width           =   270
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Guild Access:"
            Height          =   195
            Index           =   14
            Left            =   2400
            TabIndex        =   79
            Top             =   360
            Width           =   945
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Guild:"
            Height          =   195
            Index           =   13
            Left            =   2400
            TabIndex        =   78
            Top             =   120
            Width           =   405
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Map:"
            Height          =   195
            Index           =   12
            Left            =   120
            TabIndex        =   77
            Top             =   3000
            Width           =   360
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sex:"
            Height          =   195
            Index           =   11
            Left            =   120
            TabIndex        =   76
            Top             =   2760
            Width           =   330
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sprite:"
            Height          =   195
            Index           =   10
            Left            =   120
            TabIndex        =   75
            Top             =   2520
            Width           =   480
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Class:"
            Height          =   195
            Index           =   9
            Left            =   120
            TabIndex        =   74
            Top             =   2280
            Width           =   435
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PK:"
            Height          =   195
            Index           =   8
            Left            =   120
            TabIndex        =   73
            Top             =   2040
            Width           =   240
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Access:"
            Height          =   195
            Index           =   7
            Left            =   120
            TabIndex        =   72
            Top             =   1800
            Width           =   555
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "EXP: /"
            Height          =   195
            Index           =   6
            Left            =   120
            TabIndex        =   71
            Top             =   1560
            Width           =   435
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "SP: /"
            Height          =   195
            Index           =   5
            Left            =   120
            TabIndex        =   70
            Top             =   1320
            Width           =   345
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "MP: /"
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   69
            Top             =   1080
            Width           =   375
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "HP: /"
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   68
            Top             =   840
            Width           =   360
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Level:"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   67
            Top             =   600
            Width           =   435
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Character:"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   66
            Top             =   360
            Width           =   780
         End
         Begin VB.Label CharInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Account:"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   65
            Top             =   120
            Width           =   645
         End
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
         Left            =   -70080
         ScaleHeight     =   47
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   223
         TabIndex        =   59
         Top             =   480
         Visible         =   0   'False
         Width           =   3375
         Begin VB.CommandButton Command6 
            Caption         =   "Cancel"
            Height          =   255
            Left            =   1680
            TabIndex        =   117
            Top             =   960
            Width           =   1575
         End
         Begin VB.CommandButton Command7 
            Caption         =   "Caption"
            Height          =   255
            Left            =   1680
            TabIndex        =   61
            Top             =   720
            Width           =   1575
         End
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
            TabIndex        =   60
            Top             =   360
            Width           =   3075
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Reason:"
            Height          =   195
            Left            =   120
            TabIndex        =   62
            Top             =   120
            Width           =   600
         End
      End
      Begin VB.CommandButton Say 
         Caption         =   "Say"
         Height          =   255
         Index           =   5
         Left            =   -65400
         TabIndex        =   49
         Top             =   3600
         Width           =   495
      End
      Begin VB.CommandButton Say 
         Caption         =   "Say"
         Height          =   255
         Index           =   4
         Left            =   -65400
         TabIndex        =   48
         Top             =   3000
         Width           =   495
      End
      Begin VB.CommandButton Say 
         Caption         =   "Say"
         Height          =   255
         Index           =   3
         Left            =   -65400
         TabIndex        =   47
         Top             =   2400
         Width           =   495
      End
      Begin VB.CommandButton Say 
         Caption         =   "Say"
         Height          =   255
         Index           =   2
         Left            =   -65400
         TabIndex        =   46
         Top             =   1800
         Width           =   495
      End
      Begin VB.CommandButton Say 
         Caption         =   "Say"
         Height          =   255
         Index           =   1
         Left            =   -65400
         TabIndex        =   45
         Top             =   1200
         Width           =   495
      End
      Begin VB.CommandButton CustomMsg 
         Caption         =   "Custom Msg"
         Height          =   255
         Index           =   5
         Left            =   -66360
         TabIndex        =   44
         Top             =   3360
         Width           =   1455
      End
      Begin VB.CommandButton CustomMsg 
         Caption         =   "Custom Msg"
         Height          =   255
         Index           =   4
         Left            =   -66360
         TabIndex        =   43
         Top             =   2760
         Width           =   1455
      End
      Begin VB.CommandButton CustomMsg 
         Caption         =   "Custom Msg"
         Height          =   255
         Index           =   3
         Left            =   -66360
         TabIndex        =   42
         Top             =   2160
         Width           =   1455
      End
      Begin VB.CommandButton CustomMsg 
         Caption         =   "Custom Msg"
         Height          =   255
         Index           =   2
         Left            =   -66360
         TabIndex        =   41
         Top             =   1560
         Width           =   1455
      End
      Begin VB.CommandButton CustomMsg 
         Caption         =   "Custom Msg"
         Height          =   255
         Index           =   1
         Left            =   -66360
         TabIndex        =   40
         Top             =   960
         Width           =   1455
      End
      Begin VB.Frame Frame5 
         Caption         =   "Chat Options"
         Height          =   615
         Left            =   -74760
         TabIndex        =   34
         Top             =   3360
         Width           =   6975
         Begin VB.CommandButton Command60 
            Caption         =   "Save Logs"
            Height          =   255
            Left            =   5400
            TabIndex        =   174
            Top             =   240
            Width           =   1455
         End
         Begin VB.CheckBox chkA 
            Caption         =   "Admin"
            Height          =   255
            Left            =   4560
            TabIndex        =   58
            Top             =   240
            Value           =   1  'Checked
            Width           =   735
         End
         Begin VB.CheckBox chkG 
            Caption         =   "Global"
            Height          =   255
            Left            =   3720
            TabIndex        =   57
            Top             =   240
            Value           =   1  'Checked
            Width           =   735
         End
         Begin VB.CheckBox chkP 
            Caption         =   "Private"
            Height          =   255
            Left            =   2760
            TabIndex        =   38
            Top             =   240
            Value           =   1  'Checked
            Width           =   855
         End
         Begin VB.CheckBox chkM 
            Caption         =   "Map"
            Height          =   255
            Left            =   2040
            TabIndex        =   37
            Top             =   240
            Value           =   1  'Checked
            Width           =   615
         End
         Begin VB.CheckBox chkE 
            Caption         =   "Emote"
            Height          =   255
            Left            =   1200
            TabIndex        =   36
            Top             =   240
            Value           =   1  'Checked
            Width           =   855
         End
         Begin VB.CheckBox chkBC 
            Caption         =   "Broadcast"
            Height          =   255
            Left            =   120
            TabIndex        =   35
            Top             =   240
            Value           =   1  'Checked
            Width           =   1095
         End
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Heal"
         Height          =   255
         Left            =   -66600
         TabIndex        =   33
         Top             =   3360
         Width           =   1575
      End
      Begin VB.Timer tmrPlayerSave 
         Interval        =   60000
         Left            =   7440
         Top             =   0
      End
      Begin VB.Timer tmrSpawnMapItems 
         Interval        =   1000
         Left            =   9360
         Top             =   0
      End
      Begin VB.Timer tmrGameAI 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   8880
         Top             =   0
      End
      Begin VB.Timer tmrShutdown 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   8400
         Top             =   0
      End
      Begin VB.Timer PlayerTimer 
         Enabled         =   0   'False
         Interval        =   5000
         Left            =   7920
         Top             =   0
      End
      Begin VB.Frame Frame4 
         Caption         =   "Editing"
         Height          =   3375
         Left            =   4080
         TabIndex        =   29
         Top             =   480
         Width           =   1815
         Begin VB.CommandButton Command57 
            Caption         =   "Account Editor*"
            Height          =   255
            Left            =   120
            TabIndex        =   135
            Top             =   3000
            Width           =   1575
         End
         Begin VB.CommandButton Command56 
            Caption         =   "Quick Msg Editor*"
            Height          =   255
            Left            =   120
            TabIndex        =   134
            Top             =   2760
            Width           =   1575
         End
         Begin VB.CommandButton Command55 
            Caption         =   "MOTD Editor*"
            Height          =   255
            Left            =   120
            TabIndex        =   133
            Top             =   2520
            Width           =   1575
         End
         Begin VB.CommandButton Command54 
            Caption         =   "Experience Editor*"
            Height          =   255
            Left            =   120
            TabIndex        =   132
            Top             =   2280
            Width           =   1575
         End
         Begin VB.CommandButton Command53 
            Caption         =   "Data Editor*"
            Height          =   255
            Left            =   120
            TabIndex        =   131
            Top             =   2040
            Width           =   1575
         End
         Begin VB.CommandButton Command52 
            Caption         =   "Emoticon Editor*"
            Height          =   255
            Left            =   120
            TabIndex        =   130
            Top             =   1800
            Width           =   1575
         End
         Begin VB.CommandButton Command51 
            Caption         =   "Class Editor*"
            Height          =   255
            Left            =   120
            TabIndex        =   129
            Top             =   1560
            Width           =   1575
         End
         Begin VB.CommandButton Command50 
            Caption         =   "Spell Editor*"
            Height          =   255
            Left            =   120
            TabIndex        =   128
            Top             =   1320
            Width           =   1575
         End
         Begin VB.CommandButton Command49 
            Caption         =   "Shop Editor*"
            Height          =   255
            Left            =   120
            TabIndex        =   127
            Top             =   1080
            Width           =   1575
         End
         Begin VB.CommandButton Command48 
            Caption         =   "NPC Editor*"
            Height          =   255
            Left            =   120
            TabIndex        =   126
            Top             =   840
            Width           =   1575
         End
         Begin VB.CommandButton Command47 
            Caption         =   "Item Editor*"
            Height          =   255
            Left            =   120
            TabIndex        =   125
            Top             =   600
            Width           =   1575
         End
         Begin VB.CommandButton Command46 
            Caption         =   "Map Editor*"
            Height          =   255
            Left            =   120
            TabIndex        =   124
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Classes"
         Height          =   975
         Left            =   2160
         TabIndex        =   26
         Top             =   2880
         Width           =   1815
         Begin VB.CommandButton Command30 
            Caption         =   "Edit"
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   600
            Width           =   1575
         End
         Begin VB.CommandButton Command29 
            Caption         =   "Reload"
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Server"
         Height          =   1575
         Left            =   6480
         TabIndex        =   25
         Top             =   2280
         Width           =   3135
         Begin VB.CheckBox chkChat 
            Caption         =   "Save Logs"
            Height          =   255
            Left            =   120
            TabIndex        =   175
            Top             =   720
            Value           =   1  'Checked
            Width           =   1095
         End
         Begin VB.CommandButton Command59 
            Caption         =   "Weather"
            Height          =   255
            Left            =   1560
            TabIndex        =   164
            Top             =   1200
            Width           =   1455
         End
         Begin VB.CommandButton Command58 
            Caption         =   "Weirdify!"
            Height          =   255
            Left            =   120
            TabIndex        =   163
            Top             =   1200
            Width           =   1455
         End
         Begin VB.CheckBox mnuServerLog 
            Caption         =   "Server Log"
            Height          =   255
            Left            =   120
            TabIndex        =   162
            Top             =   960
            Value           =   1  'Checked
            Width           =   1095
         End
         Begin VB.CheckBox Closed 
            Caption         =   "Closed"
            Height          =   255
            Left            =   120
            TabIndex        =   161
            Top             =   480
            Width           =   855
         End
         Begin VB.CheckBox GMOnly 
            Caption         =   "GMs Only"
            Height          =   255
            Left            =   120
            TabIndex        =   39
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Exit"
            Height          =   255
            Left            =   1560
            TabIndex        =   31
            Top             =   600
            Width           =   1455
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Shutdown"
            Height          =   255
            Left            =   1560
            TabIndex        =   30
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label ShutdownTime 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Shutdown: Not Active"
            Height          =   195
            Left            =   1440
            TabIndex        =   32
            Top             =   960
            Width           =   1575
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Scripts"
         Height          =   1455
         Left            =   240
         TabIndex        =   20
         Top             =   2400
         Width           =   1815
         Begin VB.CommandButton Command28 
            Caption         =   "Edit"
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   1080
            Width           =   1575
         End
         Begin VB.CommandButton Command27 
            Caption         =   "Turn Off"
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   840
            Width           =   1575
         End
         Begin VB.CommandButton Command26 
            Caption         =   "Turn On"
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   600
            Width           =   1575
         End
         Begin VB.CommandButton Command25 
            Caption         =   "Reload"
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.CommandButton Command24 
         Caption         =   "Kill"
         Height          =   255
         Left            =   -66600
         TabIndex        =   19
         Top             =   3120
         Width           =   1575
      End
      Begin VB.CommandButton Command23 
         Caption         =   "UnMute"
         Height          =   255
         Left            =   -66600
         TabIndex        =   18
         Top             =   2880
         Width           =   1575
      End
      Begin VB.CommandButton Command22 
         Caption         =   "Mute"
         Height          =   255
         Left            =   -66600
         TabIndex        =   17
         Top             =   2640
         Width           =   1575
      End
      Begin VB.CommandButton Command21 
         Caption         =   "Message (PM)"
         Height          =   255
         Left            =   -66600
         TabIndex        =   16
         Top             =   2400
         Width           =   1575
      End
      Begin VB.CommandButton Command20 
         Caption         =   "Change Info"
         Enabled         =   0   'False
         Height          =   255
         Left            =   -66600
         TabIndex        =   15
         Top             =   2160
         Width           =   1575
      End
      Begin VB.CommandButton Command19 
         Caption         =   "View Info"
         Height          =   255
         Left            =   -66600
         TabIndex        =   14
         Top             =   1920
         Width           =   1575
      End
      Begin VB.CommandButton Command18 
         Caption         =   "Jail (Reason)"
         Height          =   255
         Left            =   -66600
         TabIndex        =   13
         Top             =   1680
         Width           =   1575
      End
      Begin VB.CommandButton Command17 
         Caption         =   "Jail"
         Height          =   255
         Left            =   -66600
         TabIndex        =   12
         Top             =   1440
         Width           =   1575
      End
      Begin VB.CommandButton Command16 
         Caption         =   "Ban (Reason)"
         Height          =   255
         Left            =   -66600
         TabIndex        =   11
         Top             =   1200
         Width           =   1575
      End
      Begin VB.CommandButton Command15 
         Caption         =   "Ban"
         Height          =   255
         Left            =   -66600
         TabIndex        =   10
         Top             =   960
         Width           =   1575
      End
      Begin VB.CommandButton Command14 
         Caption         =   "Kick (Reason)"
         Height          =   255
         Left            =   -66600
         TabIndex        =   9
         Top             =   720
         Width           =   1575
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Kick"
         Height          =   255
         Left            =   -66600
         TabIndex        =   8
         Top             =   480
         Width           =   1575
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Gridlines"
         Height          =   255
         Left            =   -67920
         TabIndex        =   4
         Top             =   3720
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.CommandButton Say 
         Caption         =   "Say"
         Height          =   255
         Index           =   0
         Left            =   -65400
         TabIndex        =   2
         Top             =   600
         Width           =   495
      End
      Begin VB.CommandButton CustomMsg 
         Caption         =   "Custom Msg"
         Height          =   255
         Index           =   0
         Left            =   -66360
         TabIndex        =   1
         Top             =   360
         Width           =   1455
      End
      Begin MSWinsockLib.Winsock Socket 
         Index           =   0
         Left            =   9840
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.Frame Frame7 
         Caption         =   "Text Files"
         Height          =   1215
         Left            =   240
         TabIndex        =   120
         Top             =   1200
         Width           =   1815
         Begin VB.CommandButton Command44 
            Caption         =   "Player.txt"
            Height          =   255
            Left            =   120
            TabIndex        =   123
            Top             =   840
            Width           =   1575
         End
         Begin VB.CommandButton Command43 
            Caption         =   "BanList.txt"
            Height          =   255
            Left            =   120
            TabIndex        =   122
            Top             =   600
            Width           =   1575
         End
         Begin VB.CommandButton Command42 
            Caption         =   "Admin.txt"
            Height          =   255
            Left            =   120
            TabIndex        =   121
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.CommandButton Command66 
         Caption         =   "Refresh"
         Height          =   255
         Left            =   -69600
         TabIndex        =   186
         Top             =   3720
         Width           =   1575
      End
      Begin MSComctlLib.ListView lvUsers 
         Height          =   3255
         Left            =   -74760
         TabIndex        =   3
         Top             =   480
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
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Click Here To Check Ip"
         Height          =   195
         Left            =   120
         TabIndex        =   187
         Top             =   840
         Width           =   1605
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Chat Log Save In"
         Height          =   195
         Left            =   -67680
         TabIndex        =   185
         Top             =   3720
         Width           =   1245
      End
      Begin VB.Label CharInfo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "www.konfuze.com"
         Height          =   195
         Index           =   23
         Left            =   -74760
         TabIndex        =   99
         Top             =   3720
         Width           =   2055
      End
      Begin VB.Label CharInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "For More Information Go To:"
         Height          =   195
         Index           =   22
         Left            =   -74760
         TabIndex        =   98
         Top             =   3480
         Width           =   2055
      End
      Begin VB.Label CharInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Topics:"
         Height          =   195
         Index           =   21
         Left            =   -74760
         TabIndex        =   95
         Top             =   360
         Width           =   510
      End
      Begin VB.Label lblIP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ip Address:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   840
      End
      Begin VB.Label lblPort 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Port:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   360
      End
      Begin VB.Label TPO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Players Online:"
         Height          =   195
         Left            =   -74760
         TabIndex        =   5
         Top             =   3780
         Width           =   1485
      End
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Seconds As Long
Dim Minutes As Integer
Dim CM As Long
Dim num As Long

Private Sub Command67_Click()
Hours = Rand(1, 24)
Minutes = Rand(0, 59)
Seconds = Rand(0, 59)
End Sub

Private Sub Command68_Click()
If IsNumeric(GameTimeSpeed.text) = False Then
    MsgBox "Enter a numerical value!"
    Exit Sub
End If

If GameTimeSpeed.text > 59 Then
    MsgBox "Enter a number less than 60"
    Exit Sub
End If

Gamespeed = GameTimeSpeed.text

End Sub

Private Sub Command69_Click()
    If TimeDisable = False Then
        Gamespeed = 0
        GameTimeSpeed.text = 0
        TimeDisable = True
        Timer1.Enabled = False
        frmServer.Command69.Caption = "Enable Time"
    Else
        Gamespeed = 1
        GameTimeSpeed.text = 1
        TimeDisable = False
        Timer1.Enabled = True
        frmServer.Command69.Caption = "Disable Time"
    End If
    
    Call DisabledTime
End Sub

Private Sub Form_Load()
Hours = Rand(1, 24)
Minutes = Rand(0, 59)
Seconds = Rand(0, 59)
Gamespeed = 1
Wierd = 0
End Sub

Private Sub Check1_Click()
    If Check1.Value = Checked Then
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
            Call PlayerMsg(index, "You have been warp by the server to Map:" & scrlMap.Value & " X:" & scrlX.Value & " Y:" & scrlY.Value, White)
            Call PlayerWarp(index, scrlMap.Value, scrlX.Value, scrlY.Value)
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
        
        Call PlayerWarp(index, scrlMap.Value, scrlX.Value, scrlY.Value)
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
            
        Call PlayerWarp(index, scrlMap.Value, scrlX.Value, scrlY.Value)
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
    CharInfo(17).Caption = "Speed: " & GetPlayerSPEED(index)
    CharInfo(18).Caption = "Magi: " & GetPlayerMAGI(index)
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

Private Sub Command25_Click()
If Scripting = 1 Then
    Set MyScript = Nothing
    Set clsScriptCommands = Nothing
    
    Set MyScript = New clsSadScript
    Set clsScriptCommands = New clsCommands
    MyScript.ReadInCode App.Path & "\Scripts\Main.txt", "Scripts\Main.txt", MyScript.SControl, False
    MyScript.SControl.AddObject "ScriptHardCode", clsScriptCommands, True
    Call TextAdd(frmServer.txtText(0), "Scripts reloaded.", True)
End If
End Sub

Private Sub Command26_Click()

If Scripting = 0 Then
    Scripting = 1
    PutVar App.Path & "\Data.ini", "CONFIG", "Scripting", 1
    
    If Scripting = 1 Then
        Set MyScript = New clsSadScript
        Set clsScriptCommands = New clsCommands
        MyScript.ReadInCode App.Path & "\Scripts\Main.txt", "Scripts\Main.txt", MyScript.SControl, False
        MyScript.SControl.AddObject "ScriptHardCode", clsScriptCommands, True
    End If
End If
End Sub

Private Sub Command27_Click()
If Scripting = 1 Then
    Scripting = 0
    PutVar App.Path & "\Data.ini", "CONFIG", "Scripting", 0
    
    If Scripting = 0 Then
        Set MyScript = Nothing
        Set clsScriptCommands = Nothing
    End If
End If
End Sub

Private Sub Command28_Click()
    AFileName = "Scripts/Main.txt"
    Unload frmEditor
    frmEditor.Show
End Sub

Private Sub Command29_Click()
    Call LoadClasses
    Call TextAdd(frmServer.txtText(0), "All classes reloaded.", True)
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
    AFileName = "Classes\Info.ini"
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
            If Scripting = 1 Then
                MyScript.ExecuteStatement "Scripts\Main.txt", "OnDeath " & index
            Else
                Call PlayerWarp(index, START_MAP, START_X, START_Y)
            End If
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
                                
            i = Int(GetPlayerSPEED(index) / 10)
            If i < 1 Then i = 1
            If i > 3 Then i = 3
                
            Call SetPlayerPOINTS(index, GetPlayerPOINTS(index) + i)
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

Call GlobalMsg("The server has warped everyone to Map:" & scrlMM.Value & " X:" & scrlMX.Value & " Y:" & scrlMY.Value, Yellow)

For i = 1 To MAX_PLAYERS
    If IsPlaying(i) = True Then
        If GetPlayerAccess(i) <= 1 Then
            Call PlayerWarp(i, scrlMM.Value, scrlMX.Value, scrlMY.Value)
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

Private Sub Command42_Click()
    AFileName = "admin.txt"
    Unload frmEditor
    frmEditor.Show
End Sub

Private Sub Command43_Click()
    AFileName = "banlist.txt"
    Unload frmEditor
    frmEditor.Show
End Sub

Private Sub Command44_Click()
    AFileName = "player.txt"
    Unload frmEditor
    frmEditor.Show
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

Private Sub Command5_Click()
    picCMsg.Visible = False
End Sub

Private Sub Command58_Click()
    Call SendWierdToAll
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
    If Scripting = 1 Then
        MyScript.ExecuteStatement "Scripts\Main.txt", "OnDeath " & index
    Else
        Call PlayerWarp(index, START_MAP, START_X, START_Y)
    End If
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

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim lmsg As Long
    
    lmsg = X
    Select Case lmsg
        Case WM_LBUTTONDBLCLK
            frmServer.WindowState = vbNormal
            frmServer.Show
    End Select
End Sub

Private Sub Form_Resize()
    'If frmServer.WindowState = vbMinimized Then
        'frmServer.Hide
    'End If
End Sub

Private Sub Form_Terminate()
    Call DestroyServer
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call DestroyServer
End Sub


Private Sub Label7_Click()
    Shell ("explorer http://www.ipchicken.com"), vbNormalNoFocus
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
    If mnuServerLog.Value = Checked Then
        ServerLog = False
    Else
        ServerLog = True
    End If
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
    txtMap.Caption = "Map: " & scrlMap.Value
End Sub

Private Sub scrlMM_Change()
    lblMM.Caption = "Map: " & scrlMM.Value
End Sub

Private Sub scrlMX_Change()
    lblMX.Caption = "X: " & scrlMX.Value
End Sub

Private Sub scrlMY_Change()
    lblMY.Caption = "Y: " & scrlMY.Value
End Sub

Private Sub scrlRainIntensity_Change()
    lblRainIntensity.Caption = "Intensity: " & Val(scrlRainIntensity.Value)
    RainIntensity = scrlRainIntensity.Value
    Call SendWeatherToAll
End Sub

Private Sub scrlX_Change()
    txtX.Caption = "X: " & scrlX.Value
End Sub

Private Sub scrlY_Change()
    txtY.Caption = "Y: " & scrlY.Value
End Sub
Private Sub Timer1_Timer()

Dim AMorPM As String
Dim TempSeconds As Integer
Dim PrintSeconds As String
Dim PrintSeconds2 As String
Dim PrintMinutes As String
Dim PrintMinutes2 As String
Dim PrintHours As Integer

Seconds = Seconds + Gamespeed

If Seconds > 59 Then
    Minutes = Minutes + 1
    Seconds = Seconds - 60
End If
If Minutes > 59 Then
    Hours = Hours + 1
    Minutes = 0
End If
If Hours > 24 Then
    Hours = 1
End If

If Hours > 12 Then
    AMorPM = "PM"
    PrintHours = Hours - 12
Else
    AMorPM = "AM"
    PrintHours = Hours
End If

If Hours = 24 Then
    AMorPM = "AM"
End If

TempSeconds = Seconds

If Seconds > 9 Then
    PrintSeconds = TempSeconds
Else
    PrintSeconds = "0" & Seconds
End If

If Seconds > 50 Then
    PrintSeconds2 = "0" & 60 - TempSeconds
Else
    PrintSeconds2 = 60 - TempSeconds
End If

If Minutes > 9 Then
    PrintMinutes = Minutes
Else
    PrintMinutes = "0" & Minutes
End If

If Minutes > 50 Then
    PrintMinutes2 = "0" & 60 - Minutes
Else
    PrintMinutes2 = 60 - Minutes
End If

Label8.Caption = "Current Time is " & PrintHours & ":" & PrintMinutes & ":" & PrintSeconds & " " & AMorPM

If Hours > 20 And GameTime = TIME_DAY Then
    GameTime = TIME_NIGHT
    Call SendTimeToAll
    End If
If Hours < 21 And Hours > 6 And GameTime = TIME_NIGHT Then
    GameTime = TIME_DAY
    Call SendTimeToAll
    End If
If Hours < 7 And GameTime = TIME_DAY Then
    GameTime = TIME_NIGHT
    Call SendTimeToAll
    End If
    
If Hours < 21 And Hours > 6 Then
    Label10.Caption = "Time until night:"
    Label11.Caption = 21 - Hours - 1 & ":" & PrintMinutes2 & ":" & PrintSeconds2
Else
    Label10.Caption = "Time until day:"
    If Hours < 7 Then
    Label11.Caption = 7 - Hours - 1 & ":" & PrintMinutes2 & ":" & PrintSeconds2
    Else
    Label11.Caption = 24 - Hours - 1 + 7 & ":" & PrintMinutes2 & ":" & PrintSeconds2
    End If
End If

If Hours > 11 Then
    GameClock = Hours - 12 & ":" & PrintMinutes & ":" & PrintSeconds & " " & AMorPM
Else
    GameClock = Hours & ":" & PrintMinutes & ":" & PrintSeconds & " " & AMorPM
End If

Call SendGameClockToAll

End Sub

Private Sub tmrChatLogs_Timer()
Static ChatSecs As Long
Dim SaveTime As Long

SaveTime = 3600

    If frmServer.chkChat.Value = Unchecked Then
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
        Call TextAdd(frmServer.txtText(0), "Server: " & txtChat.text, True)
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

Public Function Rand(ByVal Low As Long, _
                     ByVal High As Long) As Long
  Rand = Int((High - Low + 1) * Rnd) + Low
End Function

