VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmQuest 
   Caption         =   "Edit Quest"
   ClientHeight    =   6375
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7755
   LinkTopic       =   "Form1"
   ScaleHeight     =   6375
   ScaleWidth      =   7755
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   6135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7545
      _ExtentX        =   13309
      _ExtentY        =   10821
      _Version        =   393216
      Style           =   1
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   353
      TabMaxWidth     =   1587
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Quest Help"
      TabPicture(0)   =   "frmQuest.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Command"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame7"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame8"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "NPC 1"
      TabPicture(1)   =   "frmQuest.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame3"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Frame4"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Frame5"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Frame6"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Frame9"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Frame10"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Frame11"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).ControlCount=   9
      TabCaption(2)   =   "NPC 2"
      TabPicture(2)   =   "frmQuest.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame12"
      Tab(2).Control(1)=   "Frame13"
      Tab(2).Control(2)=   "Frame14"
      Tab(2).Control(3)=   "Frame15"
      Tab(2).Control(4)=   "Frame16"
      Tab(2).Control(5)=   "Frame17"
      Tab(2).Control(6)=   "Frame18"
      Tab(2).Control(7)=   "Frame19"
      Tab(2).Control(8)=   "Frame20"
      Tab(2).ControlCount=   9
      TabCaption(3)   =   "NPC 3"
      TabPicture(3)   =   "frmQuest.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame29"
      Tab(3).Control(1)=   "Frame28"
      Tab(3).Control(2)=   "Frame27"
      Tab(3).Control(3)=   "Frame26"
      Tab(3).Control(4)=   "Frame25"
      Tab(3).Control(5)=   "Frame24"
      Tab(3).Control(6)=   "Frame23"
      Tab(3).Control(7)=   "Frame22"
      Tab(3).Control(8)=   "Frame21"
      Tab(3).ControlCount=   9
      TabCaption(4)   =   "NPC 4"
      TabPicture(4)   =   "frmQuest.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame38"
      Tab(4).Control(1)=   "Frame37"
      Tab(4).Control(2)=   "Frame36"
      Tab(4).Control(3)=   "Frame35"
      Tab(4).Control(4)=   "Frame34"
      Tab(4).Control(5)=   "Frame33"
      Tab(4).Control(6)=   "Frame32"
      Tab(4).Control(7)=   "Frame31"
      Tab(4).Control(8)=   "Frame30"
      Tab(4).ControlCount=   9
      TabCaption(5)   =   "More"
      TabPicture(5)   =   "frmQuest.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Command21"
      Tab(5).ControlCount=   1
      Begin VB.CommandButton Command21 
         Caption         =   "More Quests"
         Height          =   255
         Left            =   -69360
         TabIndex        =   237
         Top             =   5640
         Width           =   1575
      End
      Begin VB.Frame Frame38 
         Caption         =   "Sprite Change if Player had Item"
         Height          =   855
         Left            =   -70440
         TabIndex        =   233
         Top             =   5100
         Width           =   2775
         Begin VB.HScrollBar HScroll36 
            Height          =   255
            Left            =   120
            Max             =   1000
            TabIndex        =   234
            Top             =   480
            Width           =   2535
         End
         Begin VB.Label Label97 
            Caption         =   "0"
            Height          =   255
            Left            =   720
            TabIndex        =   236
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label96 
            Caption         =   "Sprite:"
            Height          =   255
            Left            =   120
            TabIndex        =   235
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.Frame Frame37 
         Caption         =   "NPC Name"
         Height          =   735
         Left            =   -74880
         TabIndex        =   230
         Top             =   300
         Width           =   4335
         Begin VB.CommandButton Command20 
            Caption         =   "Save"
            Height          =   255
            Left            =   3480
            TabIndex        =   232
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox Text24 
            Height          =   285
            Left            =   120
            TabIndex        =   231
            Top             =   240
            Width           =   3255
         End
      End
      Begin VB.Frame Frame36 
         Caption         =   "NPC Say If Right Quest"
         Height          =   735
         Left            =   -74880
         TabIndex        =   227
         Top             =   1020
         Width           =   4335
         Begin VB.CommandButton Command19 
            Caption         =   "Save"
            Height          =   255
            Left            =   3480
            TabIndex        =   229
            Top             =   360
            Width           =   735
         End
         Begin VB.TextBox Text23 
            Height          =   285
            Left            =   120
            TabIndex        =   228
            Top             =   360
            Width           =   3255
         End
      End
      Begin VB.Frame Frame35 
         Caption         =   "If Player Has Item Number"
         Height          =   2775
         Left            =   -74880
         TabIndex        =   211
         Top             =   1740
         Width           =   4335
         Begin VB.CommandButton Command18 
            Caption         =   "Save"
            Height          =   255
            Left            =   3480
            TabIndex        =   217
            Top             =   2280
            Width           =   735
         End
         Begin VB.TextBox Text22 
            Height          =   285
            Left            =   120
            TabIndex        =   216
            Top             =   2280
            Width           =   3255
         End
         Begin VB.HScrollBar HScroll35 
            Height          =   255
            Left            =   120
            Max             =   500
            TabIndex        =   215
            Top             =   360
            Width           =   3495
         End
         Begin VB.HScrollBar HScroll34 
            Height          =   255
            Left            =   960
            Max             =   500
            TabIndex        =   214
            Top             =   1080
            Width           =   2655
         End
         Begin VB.HScrollBar HScroll33 
            Height          =   255
            Left            =   960
            Max             =   150
            TabIndex        =   213
            Top             =   1320
            Width           =   2655
         End
         Begin VB.HScrollBar HScroll32 
            Height          =   255
            Left            =   960
            Max             =   30000
            TabIndex        =   212
            Top             =   1560
            Width           =   2655
         End
         Begin VB.Label Label95 
            Caption         =   "If Player Has It Then He Get:"
            Height          =   255
            Left            =   120
            TabIndex        =   226
            Top             =   720
            Width           =   3015
         End
         Begin VB.Label Label94 
            Caption         =   "ItemNr"
            Height          =   255
            Left            =   120
            TabIndex        =   225
            Top             =   1080
            Width           =   495
         End
         Begin VB.Label Label93 
            Caption         =   "Stat Points"
            Height          =   255
            Left            =   120
            TabIndex        =   224
            Top             =   1350
            Width           =   855
         End
         Begin VB.Label Label92 
            Caption         =   "Experience"
            Height          =   255
            Left            =   120
            TabIndex        =   223
            Top             =   1590
            Width           =   855
         End
         Begin VB.Label Label91 
            Caption         =   "NPC Say If Player Had Item"
            Height          =   255
            Left            =   120
            TabIndex        =   222
            Top             =   2040
            Width           =   3255
         End
         Begin VB.Label Label90 
            Caption         =   "0"
            Height          =   255
            Left            =   3720
            TabIndex        =   221
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label89 
            Caption         =   "0"
            Height          =   255
            Left            =   3720
            TabIndex        =   220
            Top             =   1080
            Width           =   495
         End
         Begin VB.Label Label88 
            Caption         =   "0"
            Height          =   255
            Left            =   3720
            TabIndex        =   219
            Top             =   1320
            Width           =   495
         End
         Begin VB.Label Label87 
            Caption         =   "0"
            Height          =   255
            Left            =   3720
            TabIndex        =   218
            Top             =   1560
            Width           =   495
         End
      End
      Begin VB.Frame Frame34 
         Caption         =   "If Player Didn't Have The Item Number NPC Say"
         Height          =   735
         Left            =   -74880
         TabIndex        =   208
         Top             =   4500
         Width           =   4335
         Begin VB.CommandButton Command17 
            Caption         =   "Save"
            Height          =   255
            Left            =   3480
            TabIndex        =   210
            Top             =   360
            Width           =   735
         End
         Begin VB.TextBox Text21 
            Height          =   285
            Left            =   120
            TabIndex        =   209
            Top             =   360
            Width           =   3255
         End
      End
      Begin VB.Frame Frame33 
         Caption         =   "NPC Say If Wrong Quest"
         Height          =   735
         Left            =   -74880
         TabIndex        =   205
         Top             =   5220
         Width           =   4335
         Begin VB.CommandButton Command16 
            Caption         =   "Save"
            Height          =   255
            Left            =   3480
            TabIndex        =   207
            Top             =   360
            Width           =   735
         End
         Begin VB.TextBox Text20 
            Height          =   285
            Left            =   120
            TabIndex        =   206
            Top             =   360
            Width           =   3255
         End
      End
      Begin VB.Frame Frame32 
         Caption         =   "Keep Item If Player Has it"
         Height          =   975
         Left            =   -70440
         TabIndex        =   202
         Top             =   900
         Width           =   2775
         Begin VB.OptionButton Option8 
            Caption         =   "Don't Keep The Item"
            Height          =   255
            Left            =   120
            TabIndex        =   204
            Top             =   360
            Value           =   -1  'True
            Width           =   2535
         End
         Begin VB.OptionButton Option7 
            Caption         =   "Keep The Item"
            Height          =   255
            Left            =   120
            TabIndex        =   203
            Top             =   600
            Width           =   2535
         End
      End
      Begin VB.Frame Frame31 
         Caption         =   "Stats If Player Had ItemNr"
         Height          =   3255
         Left            =   -70440
         TabIndex        =   186
         Top             =   1860
         Width           =   2775
         Begin VB.CheckBox Check12 
            Caption         =   "Full Health"
            Height          =   255
            Left            =   240
            TabIndex        =   193
            Top             =   360
            Width           =   2415
         End
         Begin VB.CheckBox Check11 
            Caption         =   "Full Mana"
            Height          =   255
            Left            =   240
            TabIndex        =   192
            Top             =   600
            Width           =   2415
         End
         Begin VB.CheckBox Check10 
            Caption         =   "Full Stamina"
            Height          =   255
            Left            =   240
            TabIndex        =   191
            Top             =   840
            Width           =   2415
         End
         Begin VB.HScrollBar HScroll31 
            Height          =   255
            Left            =   240
            Max             =   100
            TabIndex        =   190
            Top             =   1440
            Width           =   2415
         End
         Begin VB.HScrollBar HScroll30 
            Height          =   255
            Left            =   240
            Max             =   100
            TabIndex        =   189
            Top             =   1920
            Width           =   2415
         End
         Begin VB.HScrollBar HScroll29 
            Height          =   255
            Left            =   240
            Max             =   100
            TabIndex        =   188
            Top             =   2400
            Width           =   2415
         End
         Begin VB.HScrollBar HScroll28 
            Height          =   255
            Left            =   240
            Max             =   100
            TabIndex        =   187
            Top             =   2880
            Width           =   2415
         End
         Begin VB.Label Label86 
            Caption         =   "Strength:"
            Height          =   255
            Left            =   240
            TabIndex        =   201
            Top             =   1200
            Width           =   735
         End
         Begin VB.Label Label85 
            Caption         =   "0"
            Height          =   255
            Left            =   960
            TabIndex        =   200
            Top             =   1200
            Width           =   1575
         End
         Begin VB.Label Label84 
            Caption         =   "Defence:"
            Height          =   255
            Left            =   240
            TabIndex        =   199
            Top             =   1680
            Width           =   735
         End
         Begin VB.Label Label83 
            Caption         =   "0"
            Height          =   255
            Left            =   960
            TabIndex        =   198
            Top             =   1680
            Width           =   1575
         End
         Begin VB.Label Label82 
            Caption         =   "Agility:"
            Height          =   255
            Left            =   240
            TabIndex        =   197
            Top             =   2160
            Width           =   735
         End
         Begin VB.Label Label81 
            Caption         =   "0"
            Height          =   255
            Left            =   960
            TabIndex        =   196
            Top             =   2160
            Width           =   1575
         End
         Begin VB.Label Label80 
            Caption         =   "Wisdom:"
            Height          =   255
            Left            =   240
            TabIndex        =   195
            Top             =   2640
            Width           =   735
         End
         Begin VB.Label Label79 
            Caption         =   "0"
            Height          =   255
            Left            =   960
            TabIndex        =   194
            Top             =   2640
            Width           =   1575
         End
      End
      Begin VB.Frame Frame30 
         Caption         =   "Quest Level Requirement"
         Height          =   615
         Left            =   -70440
         TabIndex        =   183
         Top             =   300
         Width           =   2775
         Begin VB.HScrollBar HScroll27 
            Height          =   255
            Left            =   120
            Max             =   15000
            TabIndex        =   184
            Top             =   240
            Width           =   1935
         End
         Begin VB.Label Label78 
            Caption         =   "0"
            Height          =   255
            Left            =   2160
            TabIndex        =   185
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.Frame Frame29 
         Caption         =   "Sprite Change if Player had Item"
         Height          =   855
         Left            =   -70440
         TabIndex        =   179
         Top             =   5115
         Width           =   2775
         Begin VB.HScrollBar HScroll26 
            Height          =   255
            Left            =   120
            Max             =   1000
            TabIndex        =   180
            Top             =   480
            Width           =   2535
         End
         Begin VB.Label Label77 
            Caption         =   "Sprite:"
            Height          =   255
            Left            =   120
            TabIndex        =   182
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label76 
            Caption         =   "0"
            Height          =   255
            Left            =   720
            TabIndex        =   181
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.Frame Frame28 
         Caption         =   "NPC Name"
         Height          =   735
         Left            =   -74880
         TabIndex        =   176
         Top             =   315
         Width           =   4335
         Begin VB.TextBox Text19 
            Height          =   285
            Left            =   120
            TabIndex        =   178
            Top             =   240
            Width           =   3255
         End
         Begin VB.CommandButton Command15 
            Caption         =   "Save"
            Height          =   255
            Left            =   3480
            TabIndex        =   177
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.Frame Frame27 
         Caption         =   "NPC Say If Right Quest"
         Height          =   735
         Left            =   -74880
         TabIndex        =   173
         Top             =   1035
         Width           =   4335
         Begin VB.TextBox Text18 
            Height          =   285
            Left            =   120
            TabIndex        =   175
            Top             =   360
            Width           =   3255
         End
         Begin VB.CommandButton Command14 
            Caption         =   "Save"
            Height          =   255
            Left            =   3480
            TabIndex        =   174
            Top             =   360
            Width           =   735
         End
      End
      Begin VB.Frame Frame26 
         Caption         =   "If Player Has Item Number"
         Height          =   2775
         Left            =   -74880
         TabIndex        =   157
         Top             =   1755
         Width           =   4335
         Begin VB.HScrollBar HScroll25 
            Height          =   255
            Left            =   960
            Max             =   30000
            TabIndex        =   163
            Top             =   1560
            Width           =   2655
         End
         Begin VB.HScrollBar HScroll24 
            Height          =   255
            Left            =   960
            Max             =   150
            TabIndex        =   162
            Top             =   1320
            Width           =   2655
         End
         Begin VB.HScrollBar HScroll23 
            Height          =   255
            Left            =   960
            Max             =   500
            TabIndex        =   161
            Top             =   1080
            Width           =   2655
         End
         Begin VB.HScrollBar HScroll22 
            Height          =   255
            Left            =   120
            Max             =   500
            TabIndex        =   160
            Top             =   360
            Width           =   3495
         End
         Begin VB.TextBox Text17 
            Height          =   285
            Left            =   120
            TabIndex        =   159
            Top             =   2280
            Width           =   3255
         End
         Begin VB.CommandButton Command13 
            Caption         =   "Save"
            Height          =   255
            Left            =   3480
            TabIndex        =   158
            Top             =   2280
            Width           =   735
         End
         Begin VB.Label Label75 
            Caption         =   "0"
            Height          =   255
            Left            =   3720
            TabIndex        =   172
            Top             =   1560
            Width           =   495
         End
         Begin VB.Label Label74 
            Caption         =   "0"
            Height          =   255
            Left            =   3720
            TabIndex        =   171
            Top             =   1320
            Width           =   495
         End
         Begin VB.Label Label73 
            Caption         =   "0"
            Height          =   255
            Left            =   3720
            TabIndex        =   170
            Top             =   1080
            Width           =   495
         End
         Begin VB.Label Label72 
            Caption         =   "0"
            Height          =   255
            Left            =   3720
            TabIndex        =   169
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label71 
            Caption         =   "NPC Say If Player Had Item"
            Height          =   255
            Left            =   120
            TabIndex        =   168
            Top             =   2040
            Width           =   3255
         End
         Begin VB.Label Label70 
            Caption         =   "Experience"
            Height          =   255
            Left            =   120
            TabIndex        =   167
            Top             =   1590
            Width           =   855
         End
         Begin VB.Label Label69 
            Caption         =   "Stat Points"
            Height          =   255
            Left            =   120
            TabIndex        =   166
            Top             =   1350
            Width           =   855
         End
         Begin VB.Label Label68 
            Caption         =   "ItemNr"
            Height          =   255
            Left            =   120
            TabIndex        =   165
            Top             =   1080
            Width           =   495
         End
         Begin VB.Label Label67 
            Caption         =   "If Player Has It Then He Get:"
            Height          =   255
            Left            =   120
            TabIndex        =   164
            Top             =   720
            Width           =   3015
         End
      End
      Begin VB.Frame Frame25 
         Caption         =   "If Player Didn't Have The Item Number NPC Say"
         Height          =   735
         Left            =   -74880
         TabIndex        =   154
         Top             =   4515
         Width           =   4335
         Begin VB.TextBox Text16 
            Height          =   285
            Left            =   120
            TabIndex        =   156
            Top             =   360
            Width           =   3255
         End
         Begin VB.CommandButton Command12 
            Caption         =   "Save"
            Height          =   255
            Left            =   3480
            TabIndex        =   155
            Top             =   360
            Width           =   735
         End
      End
      Begin VB.Frame Frame24 
         Caption         =   "NPC Say If Wrong Quest"
         Height          =   735
         Left            =   -74880
         TabIndex        =   151
         Top             =   5235
         Width           =   4335
         Begin VB.TextBox Text15 
            Height          =   285
            Left            =   120
            TabIndex        =   153
            Top             =   360
            Width           =   3255
         End
         Begin VB.CommandButton Command11 
            Caption         =   "Save"
            Height          =   255
            Left            =   3480
            TabIndex        =   152
            Top             =   360
            Width           =   735
         End
      End
      Begin VB.Frame Frame23 
         Caption         =   "Keep Item If Player Has it"
         Height          =   975
         Left            =   -70440
         TabIndex        =   148
         Top             =   915
         Width           =   2775
         Begin VB.OptionButton Option6 
            Caption         =   "Keep The Item"
            Height          =   255
            Left            =   120
            TabIndex        =   150
            Top             =   600
            Width           =   2535
         End
         Begin VB.OptionButton Option5 
            Caption         =   "Don't Keep The Item"
            Height          =   255
            Left            =   120
            TabIndex        =   149
            Top             =   360
            Value           =   -1  'True
            Width           =   2535
         End
      End
      Begin VB.Frame Frame22 
         Caption         =   "Stats If Player Had ItemNr"
         Height          =   3255
         Left            =   -70440
         TabIndex        =   132
         Top             =   1875
         Width           =   2775
         Begin VB.HScrollBar HScroll21 
            Height          =   255
            Left            =   240
            Max             =   100
            TabIndex        =   139
            Top             =   2880
            Width           =   2415
         End
         Begin VB.HScrollBar HScroll20 
            Height          =   255
            Left            =   240
            Max             =   100
            TabIndex        =   138
            Top             =   2400
            Width           =   2415
         End
         Begin VB.HScrollBar HScroll19 
            Height          =   255
            Left            =   240
            Max             =   100
            TabIndex        =   137
            Top             =   1920
            Width           =   2415
         End
         Begin VB.HScrollBar HScroll18 
            Height          =   255
            Left            =   240
            Max             =   100
            TabIndex        =   136
            Top             =   1440
            Width           =   2415
         End
         Begin VB.CheckBox Check9 
            Caption         =   "Full Stamina"
            Height          =   255
            Left            =   240
            TabIndex        =   135
            Top             =   840
            Width           =   2415
         End
         Begin VB.CheckBox Check8 
            Caption         =   "Full Mana"
            Height          =   255
            Left            =   240
            TabIndex        =   134
            Top             =   600
            Width           =   2415
         End
         Begin VB.CheckBox Check7 
            Caption         =   "Full Health"
            Height          =   255
            Left            =   240
            TabIndex        =   133
            Top             =   360
            Width           =   2415
         End
         Begin VB.Label Label66 
            Caption         =   "0"
            Height          =   255
            Left            =   960
            TabIndex        =   147
            Top             =   2640
            Width           =   1575
         End
         Begin VB.Label Label65 
            Caption         =   "Wisdom:"
            Height          =   255
            Left            =   240
            TabIndex        =   146
            Top             =   2640
            Width           =   735
         End
         Begin VB.Label Label64 
            Caption         =   "0"
            Height          =   255
            Left            =   960
            TabIndex        =   145
            Top             =   2160
            Width           =   1575
         End
         Begin VB.Label Label63 
            Caption         =   "Agility:"
            Height          =   255
            Left            =   240
            TabIndex        =   144
            Top             =   2160
            Width           =   735
         End
         Begin VB.Label Label62 
            Caption         =   "0"
            Height          =   255
            Left            =   960
            TabIndex        =   143
            Top             =   1680
            Width           =   1575
         End
         Begin VB.Label Label61 
            Caption         =   "Defence:"
            Height          =   255
            Left            =   240
            TabIndex        =   142
            Top             =   1680
            Width           =   735
         End
         Begin VB.Label Label60 
            Caption         =   "0"
            Height          =   255
            Left            =   960
            TabIndex        =   141
            Top             =   1200
            Width           =   1575
         End
         Begin VB.Label Label59 
            Caption         =   "Strength:"
            Height          =   255
            Left            =   240
            TabIndex        =   140
            Top             =   1200
            Width           =   735
         End
      End
      Begin VB.Frame Frame21 
         Caption         =   "Quest Level Requirement"
         Height          =   615
         Left            =   -70440
         TabIndex        =   129
         Top             =   315
         Width           =   2775
         Begin VB.HScrollBar HScroll17 
            Height          =   255
            Left            =   120
            Max             =   15000
            TabIndex        =   130
            Top             =   240
            Width           =   1935
         End
         Begin VB.Label Label58 
            Caption         =   "0"
            Height          =   255
            Left            =   2160
            TabIndex        =   131
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.Frame Frame20 
         Caption         =   "Sprite Change if Player had Item"
         Height          =   855
         Left            =   -70440
         TabIndex        =   120
         Top             =   5040
         Width           =   2775
         Begin VB.HScrollBar HScroll16 
            Height          =   255
            Left            =   120
            Max             =   1000
            TabIndex        =   121
            Top             =   480
            Width           =   2535
         End
         Begin VB.Label Label57 
            Caption         =   "0"
            Height          =   255
            Left            =   720
            TabIndex        =   123
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label56 
            Caption         =   "Sprite:"
            Height          =   255
            Left            =   120
            TabIndex        =   122
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.Frame Frame19 
         Caption         =   "NPC Name"
         Height          =   735
         Left            =   -74880
         TabIndex        =   118
         Top             =   240
         Width           =   4335
         Begin VB.CommandButton Command10 
            Caption         =   "Save"
            Height          =   255
            Left            =   3480
            TabIndex        =   124
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox Text14 
            Height          =   285
            Left            =   120
            TabIndex        =   119
            Top             =   240
            Width           =   3255
         End
      End
      Begin VB.Frame Frame18 
         Caption         =   "NPC Say If Right Quest"
         Height          =   735
         Left            =   -74880
         TabIndex        =   116
         Top             =   960
         Width           =   4335
         Begin VB.CommandButton Command3 
            Caption         =   "Save"
            Height          =   255
            Left            =   3480
            TabIndex        =   125
            Top             =   360
            Width           =   735
         End
         Begin VB.TextBox Text13 
            Height          =   285
            Left            =   120
            TabIndex        =   117
            Top             =   360
            Width           =   3255
         End
      End
      Begin VB.Frame Frame17 
         Caption         =   "If Player Has Item Number"
         Height          =   2775
         Left            =   -74880
         TabIndex        =   101
         Top             =   1680
         Width           =   4335
         Begin VB.CommandButton Command4 
            Caption         =   "Save"
            Height          =   255
            Left            =   3480
            TabIndex        =   126
            Top             =   2280
            Width           =   735
         End
         Begin VB.TextBox Text12 
            Height          =   285
            Left            =   120
            TabIndex        =   106
            Top             =   2280
            Width           =   3255
         End
         Begin VB.HScrollBar HScroll15 
            Height          =   255
            Left            =   120
            Max             =   500
            TabIndex        =   105
            Top             =   360
            Width           =   3495
         End
         Begin VB.HScrollBar HScroll14 
            Height          =   255
            Left            =   960
            Max             =   500
            TabIndex        =   104
            Top             =   1080
            Width           =   2655
         End
         Begin VB.HScrollBar HScroll13 
            Height          =   255
            Left            =   960
            Max             =   150
            TabIndex        =   103
            Top             =   1320
            Width           =   2655
         End
         Begin VB.HScrollBar HScroll12 
            Height          =   255
            Left            =   960
            Max             =   30000
            TabIndex        =   102
            Top             =   1560
            Width           =   2655
         End
         Begin VB.Label Label55 
            Caption         =   "If Player Has It Then He Get:"
            Height          =   255
            Left            =   120
            TabIndex        =   115
            Top             =   720
            Width           =   3015
         End
         Begin VB.Label Label54 
            Caption         =   "ItemNr"
            Height          =   255
            Left            =   120
            TabIndex        =   114
            Top             =   1080
            Width           =   495
         End
         Begin VB.Label Label53 
            Caption         =   "Stat Points"
            Height          =   255
            Left            =   120
            TabIndex        =   113
            Top             =   1350
            Width           =   855
         End
         Begin VB.Label Label52 
            Caption         =   "Experience"
            Height          =   255
            Left            =   120
            TabIndex        =   112
            Top             =   1590
            Width           =   855
         End
         Begin VB.Label Label51 
            Caption         =   "NPC Say If Player Had Item"
            Height          =   255
            Left            =   120
            TabIndex        =   111
            Top             =   2040
            Width           =   3255
         End
         Begin VB.Label Label50 
            Caption         =   "0"
            Height          =   255
            Left            =   3720
            TabIndex        =   110
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label49 
            Caption         =   "0"
            Height          =   255
            Left            =   3720
            TabIndex        =   109
            Top             =   1080
            Width           =   495
         End
         Begin VB.Label Label48 
            Caption         =   "0"
            Height          =   255
            Left            =   3720
            TabIndex        =   108
            Top             =   1320
            Width           =   495
         End
         Begin VB.Label Label47 
            Caption         =   "0"
            Height          =   255
            Left            =   3720
            TabIndex        =   107
            Top             =   1560
            Width           =   495
         End
      End
      Begin VB.Frame Frame16 
         Caption         =   "If Player Didn't Have The Item Number NPC Say"
         Height          =   735
         Left            =   -74880
         TabIndex        =   99
         Top             =   4440
         Width           =   4335
         Begin VB.CommandButton Command5 
            Caption         =   "Save"
            Height          =   255
            Left            =   3480
            TabIndex        =   127
            Top             =   360
            Width           =   735
         End
         Begin VB.TextBox Text11 
            Height          =   285
            Left            =   120
            TabIndex        =   100
            Top             =   360
            Width           =   3255
         End
      End
      Begin VB.Frame Frame15 
         Caption         =   "NPC Say If Wrong Quest"
         Height          =   735
         Left            =   -74880
         TabIndex        =   97
         Top             =   5160
         Width           =   4335
         Begin VB.CommandButton Command6 
            Caption         =   "Save"
            Height          =   255
            Left            =   3480
            TabIndex        =   128
            Top             =   360
            Width           =   735
         End
         Begin VB.TextBox Text10 
            Height          =   285
            Left            =   120
            TabIndex        =   98
            Top             =   360
            Width           =   3255
         End
      End
      Begin VB.Frame Frame14 
         Caption         =   "Keep Item If Player Has it"
         Height          =   855
         Left            =   -70440
         TabIndex        =   94
         Top             =   960
         Width           =   2775
         Begin VB.OptionButton Option4 
            Caption         =   "Don't Keep The Item"
            Height          =   255
            Left            =   120
            TabIndex        =   96
            Top             =   240
            Value           =   -1  'True
            Width           =   2535
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Keep The Item"
            Height          =   255
            Left            =   120
            TabIndex        =   95
            Top             =   480
            Width           =   2535
         End
      End
      Begin VB.Frame Frame13 
         Caption         =   "Stats If Player Had ItemNr"
         Height          =   3255
         Left            =   -70440
         TabIndex        =   78
         Top             =   1800
         Width           =   2775
         Begin VB.CheckBox Check6 
            Caption         =   "Full Health"
            Height          =   255
            Left            =   240
            TabIndex        =   85
            Top             =   360
            Width           =   2415
         End
         Begin VB.CheckBox Check5 
            Caption         =   "Full Mana"
            Height          =   255
            Left            =   240
            TabIndex        =   84
            Top             =   600
            Width           =   2415
         End
         Begin VB.CheckBox Check4 
            Caption         =   "Full Stamina"
            Height          =   255
            Left            =   240
            TabIndex        =   83
            Top             =   840
            Width           =   2415
         End
         Begin VB.HScrollBar HScroll11 
            Height          =   255
            Left            =   240
            Max             =   100
            TabIndex        =   82
            Top             =   1440
            Width           =   2415
         End
         Begin VB.HScrollBar HScroll10 
            Height          =   255
            Left            =   240
            Max             =   100
            TabIndex        =   81
            Top             =   1920
            Width           =   2415
         End
         Begin VB.HScrollBar HScroll9 
            Height          =   255
            Left            =   240
            Max             =   100
            TabIndex        =   80
            Top             =   2400
            Width           =   2415
         End
         Begin VB.HScrollBar HScroll8 
            Height          =   255
            Left            =   240
            Max             =   100
            TabIndex        =   79
            Top             =   2880
            Width           =   2415
         End
         Begin VB.Label Label46 
            Caption         =   "Strength:"
            Height          =   255
            Left            =   240
            TabIndex        =   93
            Top             =   1200
            Width           =   735
         End
         Begin VB.Label Label45 
            Caption         =   "0"
            Height          =   255
            Left            =   960
            TabIndex        =   92
            Top             =   1200
            Width           =   1575
         End
         Begin VB.Label Label44 
            Caption         =   "Defence:"
            Height          =   255
            Left            =   240
            TabIndex        =   91
            Top             =   1680
            Width           =   735
         End
         Begin VB.Label Label43 
            Caption         =   "0"
            Height          =   255
            Left            =   960
            TabIndex        =   90
            Top             =   1680
            Width           =   1575
         End
         Begin VB.Label Label42 
            Caption         =   "Agility:"
            Height          =   255
            Left            =   240
            TabIndex        =   89
            Top             =   2160
            Width           =   735
         End
         Begin VB.Label Label41 
            Caption         =   "0"
            Height          =   255
            Left            =   960
            TabIndex        =   88
            Top             =   2160
            Width           =   1575
         End
         Begin VB.Label Label40 
            Caption         =   "Wisdom:"
            Height          =   255
            Left            =   240
            TabIndex        =   87
            Top             =   2640
            Width           =   735
         End
         Begin VB.Label Label39 
            Caption         =   "0"
            Height          =   255
            Left            =   960
            TabIndex        =   86
            Top             =   2640
            Width           =   1575
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "Quest Level Requirement"
         Height          =   615
         Left            =   -70440
         TabIndex        =   75
         Top             =   255
         Width           =   2775
         Begin VB.HScrollBar HScroll7 
            Height          =   255
            Left            =   120
            Max             =   15000
            TabIndex        =   76
            Top             =   240
            Width           =   1935
         End
         Begin VB.Label Label38 
            Caption         =   "0"
            Height          =   255
            Left            =   2160
            TabIndex        =   77
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "Sprite Change if Player had Item"
         Height          =   855
         Left            =   -70440
         TabIndex        =   70
         Top             =   5100
         Width           =   2775
         Begin VB.HScrollBar HScroll6 
            Height          =   255
            Left            =   120
            Max             =   1000
            TabIndex        =   71
            Top             =   480
            Width           =   2535
         End
         Begin VB.Label Label36 
            Caption         =   "Sprite:"
            Height          =   255
            Left            =   120
            TabIndex        =   73
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label35 
            Caption         =   "0"
            Height          =   255
            Left            =   720
            TabIndex        =   72
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Quest Level Requirement"
         Height          =   615
         Left            =   -70440
         TabIndex        =   67
         Top             =   300
         Width           =   2775
         Begin VB.HScrollBar HScroll5 
            Height          =   255
            Left            =   120
            Max             =   15000
            TabIndex        =   68
            Top             =   240
            Width           =   1935
         End
         Begin VB.Label Label34 
            Caption         =   "0"
            Height          =   255
            Left            =   2160
            TabIndex        =   69
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Stats If Player Had ItemNr"
         Height          =   3255
         Left            =   -70440
         TabIndex        =   43
         Top             =   1860
         Width           =   2775
         Begin VB.HScrollBar HScroll4 
            Height          =   255
            Left            =   240
            Max             =   100
            TabIndex        =   56
            Top             =   2880
            Width           =   2415
         End
         Begin VB.HScrollBar HScroll3 
            Height          =   255
            Left            =   240
            Max             =   100
            TabIndex        =   53
            Top             =   2400
            Width           =   2415
         End
         Begin VB.HScrollBar HScroll2 
            Height          =   255
            Left            =   240
            Max             =   100
            TabIndex        =   50
            Top             =   1920
            Width           =   2415
         End
         Begin VB.HScrollBar HScroll1 
            Height          =   255
            Left            =   240
            Max             =   100
            TabIndex        =   47
            Top             =   1440
            Width           =   2415
         End
         Begin VB.CheckBox Check3 
            Caption         =   "Full Stamina"
            Height          =   255
            Left            =   240
            TabIndex        =   46
            Top             =   840
            Width           =   2415
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Full Mana"
            Height          =   255
            Left            =   240
            TabIndex        =   45
            Top             =   600
            Width           =   2415
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Full Health"
            Height          =   255
            Left            =   240
            TabIndex        =   44
            Top             =   360
            Width           =   2415
         End
         Begin VB.Label Label29 
            Caption         =   "0"
            Height          =   255
            Left            =   960
            TabIndex        =   58
            Top             =   2640
            Width           =   1575
         End
         Begin VB.Label Label28 
            Caption         =   "Wisdom:"
            Height          =   255
            Left            =   240
            TabIndex        =   57
            Top             =   2640
            Width           =   735
         End
         Begin VB.Label Label27 
            Caption         =   "0"
            Height          =   255
            Left            =   960
            TabIndex        =   55
            Top             =   2160
            Width           =   1575
         End
         Begin VB.Label Label26 
            Caption         =   "Agility:"
            Height          =   255
            Left            =   240
            TabIndex        =   54
            Top             =   2160
            Width           =   735
         End
         Begin VB.Label Label25 
            Caption         =   "0"
            Height          =   255
            Left            =   960
            TabIndex        =   52
            Top             =   1680
            Width           =   1575
         End
         Begin VB.Label Label24 
            Caption         =   "Defence:"
            Height          =   255
            Left            =   240
            TabIndex        =   51
            Top             =   1680
            Width           =   735
         End
         Begin VB.Label Label23 
            Caption         =   "0"
            Height          =   255
            Left            =   960
            TabIndex        =   49
            Top             =   1200
            Width           =   1575
         End
         Begin VB.Label Label2 
            Caption         =   "Strength:"
            Height          =   255
            Left            =   240
            TabIndex        =   48
            Top             =   1200
            Width           =   735
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Save Quest"
         Height          =   855
         Left            =   120
         TabIndex        =   35
         Top             =   5160
         Width           =   6255
         Begin VB.Label Label19 
            Caption         =   "Make sure you always save the quest by pressing on the save button or on the Save All button."
            Height          =   375
            Left            =   720
            TabIndex        =   37
            Top             =   240
            Width           =   5415
         End
         Begin VB.Label Label18 
            Caption         =   "Save:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   36
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Help Part"
         Height          =   4815
         Left            =   120
         TabIndex        =   22
         Top             =   360
         Width           =   7215
         Begin VB.Label Label37 
            Caption         =   "Leave Values 0 to make it not act."
            Height          =   255
            Left            =   120
            TabIndex        =   74
            Top             =   4440
            Width           =   6975
         End
         Begin VB.Label Label22 
            Caption         =   "If the player had the item number then the NPC will say this."
            Height          =   255
            Left            =   2640
            TabIndex        =   42
            Top             =   2760
            Width           =   4455
         End
         Begin VB.Label Label21 
            Caption         =   "NPC Say If Player Had Item:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   41
            Top             =   2760
            Width           =   2535
         End
         Begin VB.Label Label17 
            Caption         =   "If you are talking to NPC 3 and you didn't talk with NPC 2 first, then the NPC will say this text only."
            Height          =   495
            Left            =   2400
            TabIndex        =   34
            Top             =   3840
            Width           =   4575
         End
         Begin VB.Label Label16 
            Caption         =   "NPC Say If Wrong Quest:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   3840
            Width           =   2175
         End
         Begin VB.Label Label15 
            Caption         =   $"frmQuest.frx":00A8
            Height          =   855
            Left            =   4320
            TabIndex        =   32
            Top             =   3000
            Width           =   2775
         End
         Begin VB.Label Label14 
            Caption         =   "If Player Didn't Have The Item Number NPC Say:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   31
            Top             =   3000
            Width           =   4215
         End
         Begin VB.Label Label13 
            Caption         =   $"frmQuest.frx":013D
            Height          =   855
            Left            =   2760
            TabIndex        =   30
            Top             =   1920
            Width           =   4335
         End
         Begin VB.Label Label12 
            Caption         =   "If Player Has It Then He Get:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   1920
            Width           =   2535
         End
         Begin VB.Label Label11 
            Caption         =   $"frmQuest.frx":022C
            Height          =   615
            Left            =   2520
            TabIndex        =   28
            Top             =   1320
            Width           =   4575
         End
         Begin VB.Label Label10 
            Caption         =   "If Player Has Item Number:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   1320
            Width           =   2295
         End
         Begin VB.Label Label9 
            Caption         =   $"frmQuest.frx":02D4
            Height          =   855
            Left            =   2280
            TabIndex        =   26
            Top             =   480
            Width           =   4815
         End
         Begin VB.Label Label8 
            Caption         =   "NPC Say If Right Quest:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   480
            Width           =   2175
         End
         Begin VB.Label Label7 
            Caption         =   "Examples: Artimus, Whayn, Indy, Jean."
            Height          =   255
            Left            =   1200
            TabIndex        =   24
            Top             =   240
            Width           =   4695
         End
         Begin VB.Label Label6 
            Caption         =   "NPC Name:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.CommandButton Command 
         Caption         =   "Close"
         Height          =   255
         Left            =   6480
         TabIndex        =   21
         Top             =   5760
         Width           =   855
      End
      Begin VB.Frame Frame6 
         Caption         =   "Keep Item If Player Has it"
         Height          =   975
         Left            =   -70440
         TabIndex        =   18
         Top             =   900
         Width           =   2775
         Begin VB.OptionButton Option2 
            Caption         =   "Keep The Item"
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   600
            Width           =   2535
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Don't Keep The Item"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   360
            Value           =   -1  'True
            Width           =   2535
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "NPC Say If Wrong Quest"
         Height          =   735
         Left            =   -74880
         TabIndex        =   15
         Top             =   5220
         Width           =   4335
         Begin VB.CommandButton Command9 
            Caption         =   "Save"
            Height          =   255
            Left            =   3480
            TabIndex        =   17
            Top             =   375
            Width           =   735
         End
         Begin VB.TextBox Text9 
            Height          =   285
            Left            =   120
            TabIndex        =   16
            Top             =   360
            Width           =   3255
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "If Player Didn't Have The Item Number NPC Say"
         Height          =   735
         Left            =   -74880
         TabIndex        =   12
         Top             =   4500
         Width           =   4335
         Begin VB.CommandButton Command8 
            Caption         =   "Save"
            Height          =   255
            Left            =   3480
            TabIndex        =   14
            Top             =   375
            Width           =   735
         End
         Begin VB.TextBox Text8 
            Height          =   285
            Left            =   120
            TabIndex        =   13
            Top             =   360
            Width           =   3255
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "If Player Has Item Number"
         Height          =   2775
         Left            =   -74880
         TabIndex        =   7
         Top             =   1740
         Width           =   4335
         Begin VB.HScrollBar Text6 
            Height          =   255
            Left            =   960
            Max             =   30000
            TabIndex        =   65
            Top             =   1560
            Width           =   2655
         End
         Begin VB.HScrollBar Text5 
            Height          =   255
            Left            =   960
            Max             =   150
            TabIndex        =   63
            Top             =   1320
            Width           =   2655
         End
         Begin VB.HScrollBar Text4 
            Height          =   255
            Left            =   960
            Max             =   500
            TabIndex        =   61
            Top             =   1080
            Width           =   2655
         End
         Begin VB.HScrollBar Text3 
            Height          =   255
            Left            =   120
            Max             =   500
            TabIndex        =   59
            Top             =   360
            Width           =   3495
         End
         Begin VB.TextBox Text7 
            Height          =   285
            Left            =   120
            TabIndex        =   39
            Top             =   2280
            Width           =   3255
         End
         Begin VB.CommandButton Command7 
            Caption         =   "Save"
            Height          =   255
            Left            =   3480
            TabIndex        =   38
            Top             =   2295
            Width           =   735
         End
         Begin VB.Label Label33 
            Caption         =   "0"
            Height          =   255
            Left            =   3720
            TabIndex        =   66
            Top             =   1560
            Width           =   495
         End
         Begin VB.Label Label32 
            Caption         =   "0"
            Height          =   255
            Left            =   3720
            TabIndex        =   64
            Top             =   1320
            Width           =   495
         End
         Begin VB.Label Label31 
            Caption         =   "0"
            Height          =   255
            Left            =   3720
            TabIndex        =   62
            Top             =   1080
            Width           =   495
         End
         Begin VB.Label Label30 
            Caption         =   "0"
            Height          =   255
            Left            =   3720
            TabIndex        =   60
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label20 
            Caption         =   "NPC Say If Player Had Item"
            Height          =   255
            Left            =   120
            TabIndex        =   40
            Top             =   2040
            Width           =   3255
         End
         Begin VB.Label Label5 
            Caption         =   "Experience"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   1590
            Width           =   855
         End
         Begin VB.Label Label4 
            Caption         =   "Stat Points"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   1350
            Width           =   855
         End
         Begin VB.Label Label3 
            Caption         =   "ItemNr"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   1080
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "If Player Has It Then He Get:"
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   720
            Width           =   3015
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "NPC Say If Right Quest"
         Height          =   735
         Left            =   -74880
         TabIndex        =   4
         Top             =   1020
         Width           =   4335
         Begin VB.CommandButton Command2 
            Caption         =   "Save"
            Height          =   255
            Left            =   3480
            TabIndex        =   6
            Top             =   360
            Width           =   735
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   120
            TabIndex        =   5
            Top             =   360
            Width           =   3255
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "NPC Name"
         Height          =   735
         Left            =   -74880
         TabIndex        =   1
         Top             =   300
         Width           =   4335
         Begin VB.CommandButton Command1 
            Caption         =   "Save"
            Height          =   255
            Left            =   3480
            TabIndex        =   3
            Top             =   250
            Width           =   735
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   3255
         End
      End
   End
End
Attribute VB_Name = "frmQuest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Check1_Click()
If Check1.value = 1 Then
Call PutVar(App.Path & "\Quests.ini", "QUEST1", "FullHP", 1)
Else
Call PutVar(App.Path & "\Quests.ini", "QUEST1", "FullHP", 0)
End If
End Sub

Private Sub Check10_Click()
If Check10.value = 1 Then
Call PutVar(App.Path & "\Quests.ini", "QUEST4", "FullSP", 1)
Else
Call PutVar(App.Path & "\Quests.ini", "QUEST4", "FullSP", 0)
End If
End Sub

Private Sub Check11_Click()
If Check11.value = 1 Then
Call PutVar(App.Path & "\Quests.ini", "QUEST4", "FullMP", 1)
Else
Call PutVar(App.Path & "\Quests.ini", "QUEST4", "FullMP", 0)
End If
End Sub

Private Sub Check12_Click()
If Check12.value = 1 Then
Call PutVar(App.Path & "\Quests.ini", "QUEST4", "FullHP", 1)
Else
Call PutVar(App.Path & "\Quests.ini", "QUEST4", "FullHP", 0)
End If
End Sub

Private Sub Check2_Click()
If Check2.value = 1 Then
Call PutVar(App.Path & "\Quests.ini", "QUEST1", "FullMP", 1)
Else
Call PutVar(App.Path & "\Quests.ini", "QUEST1", "FullMP", 0)
End If
End Sub

Private Sub Check3_Click()
If Check3.value = 1 Then
Call PutVar(App.Path & "\Quests.ini", "QUEST1", "FullSP", 1)
Else
Call PutVar(App.Path & "\Quests.ini", "QUEST1", "FullSP", 0)
End If
End Sub

Private Sub Check4_Click()
If Check4.value = 1 Then
Call PutVar(App.Path & "\Quests.ini", "QUEST2", "FullSP", 1)
Else
Call PutVar(App.Path & "\Quests.ini", "QUEST2", "FullSP", 0)
End If
End Sub

Private Sub Check5_Click()
If Check5.value = 1 Then
Call PutVar(App.Path & "\Quests.ini", "QUEST2", "FullMP", 1)
Else
Call PutVar(App.Path & "\Quests.ini", "QUEST2", "FullMP", 0)
End If
End Sub

Private Sub Check6_Click()
If Check6.value = 1 Then
Call PutVar(App.Path & "\Quests.ini", "QUEST2", "FullHP", 1)
Else
Call PutVar(App.Path & "\Quests.ini", "QUEST2", "FullHP", 0)
End If
End Sub

Private Sub Check7_Click()
If Check7.value = 1 Then
Call PutVar(App.Path & "\Quests.ini", "QUEST3", "FullHP", 1)
Else
Call PutVar(App.Path & "\Quests.ini", "QUEST3", "FullHP", 0)
End If
End Sub

Private Sub Check8_Click()
If Check8.value = 1 Then
Call PutVar(App.Path & "\Quests.ini", "QUEST3", "FullMP", 1)
Else
Call PutVar(App.Path & "\Quests.ini", "QUEST3", "FullMP", 0)
End If
End Sub

Private Sub Check9_Click()
If Check9.value = 1 Then
Call PutVar(App.Path & "\Quests.ini", "QUEST3", "FullSP", 1)
Else
Call PutVar(App.Path & "\Quests.ini", "QUEST3", "FullSP", 0)
End If
End Sub

Private Sub Command1_Click()
Call PutVar(App.Path & "\Quests.ini", "QUEST1", "NPCname", Text1.text)
End Sub
Private Sub Command_Click()
frmQuest.Visible = False
End Sub

Private Sub Command10_Click()
Call PutVar(App.Path & "\Quests.ini", "QUEST2", "NPCname", Text14.text)
End Sub

Private Sub Command11_Click()
Call PutVar(App.Path & "\Quests.ini", "QUEST3", "WrongQuestNPCsay", Text15.text)
End Sub

Private Sub Command12_Click()
Call PutVar(App.Path & "\Quests.ini", "QUEST3", "NoItemNPCsay", Text16.text)
End Sub

Private Sub Command13_Click()
Call PutVar(App.Path & "\Quests.ini", "QUEST3", "FinishedNPCSay", Text17.text)
End Sub

Private Sub Command14_Click()
Call PutVar(App.Path & "\Quests.ini", "QUEST3", "NPCsay", Text18.text)
End Sub

Private Sub Command15_Click()
Call PutVar(App.Path & "\Quests.ini", "QUEST3", "NPCname", Text19.text)
End Sub

Private Sub Command16_Click()
Call PutVar(App.Path & "\Quests.ini", "QUEST4", "WrongQuestNPCsay", Text20.text)
End Sub

Private Sub Command17_Click()
Call PutVar(App.Path & "\Quests.ini", "QUEST4", "NoItemNPCsay", Text21.text)
End Sub

Private Sub Command18_Click()
Call PutVar(App.Path & "\Quests.ini", "QUEST4", "FinishedNPCSay", Text22.text)
End Sub

Private Sub Command19_Click()
Call PutVar(App.Path & "\Quests.ini", "QUEST4", "NPCsay", Text23.text)
End Sub

Private Sub Command2_Click()
Call PutVar(App.Path & "\Quests.ini", "QUEST1", "NPCsay", Text2.text)
End Sub

Private Sub Command20_Click()
Call PutVar(App.Path & "\Quests.ini", "QUEST4", "NPCname", Text24.text)
End Sub

Private Sub Command21_Click()
frmQuest.Visible = False
frmQuest1.Visible = True
End Sub

Private Sub Command3_Click()
Call PutVar(App.Path & "\Quests.ini", "QUEST2", "NPCsay", Text13.text)
End Sub

Private Sub Command4_Click()
Call PutVar(App.Path & "\Quests.ini", "QUEST2", "FinishedNPCSay", Text12.text)
End Sub

Private Sub Command5_Click()
Call PutVar(App.Path & "\Quests.ini", "QUEST2", "NoItemNPCsay", Text11.text)
End Sub

Private Sub Command6_Click()
Call PutVar(App.Path & "\Quests.ini", "QUEST2", "WrongQuestNPCsay", Text10.text)
End Sub

Private Sub Command7_Click()
Call PutVar(App.Path & "\Quests.ini", "QUEST1", "FinishedNPCSay", Text7.text)
End Sub

Private Sub Command8_Click()
Call PutVar(App.Path & "\Quests.ini", "QUEST1", "NoItemNPCsay", Text8.text)
End Sub

Private Sub Command9_Click()
Call PutVar(App.Path & "\Quests.ini", "QUEST1", "WrongQuestNPCsay", Text9.text)
End Sub

Private Sub Form_Load()
Text24.text = GetVar(App.Path & "\Quests.ini", "QUEST4", "NPCname")
Text23.text = GetVar(App.Path & "\Quests.ini", "QUEST4", "NPCsay")
HScroll35.value = GetVar(App.Path & "\Quests.ini", "QUEST4", "IfItemNr")
Label90.Caption = GetVar(App.Path & "\Quests.ini", "QUEST4", "IfItemNr")
Text22.text = GetVar(App.Path & "\Quests.ini", "QUEST4", "FinishedNPCSay")
HScroll34.value = GetVar(App.Path & "\Quests.ini", "QUEST4", "ItemNr")
Label89.Caption = GetVar(App.Path & "\Quests.ini", "QUEST4", "ItemNr")
HScroll33.value = GetVar(App.Path & "\Quests.ini", "QUEST4", "StatPoints")
Label88.Caption = GetVar(App.Path & "\Quests.ini", "QUEST4", "StatPoints")
HScroll32.value = GetVar(App.Path & "\Quests.ini", "QUEST4", "Exp")
Label87.Caption = GetVar(App.Path & "\Quests.ini", "QUEST4", "Exp")
Text21.text = GetVar(App.Path & "\Quests.ini", "QUEST4", "NoItemNPCsay")
Text20.text = GetVar(App.Path & "\Quests.ini", "QUEST4", "WrongQuestNPCsay")
If GetVar(App.Path & "\Quests.ini", "QUEST4", "Keep") = 0 Then
Option8.value = True
Else
Option7.value = True
End If
HScroll31.value = GetVar(App.Path & "\Quests.ini", "QUEST4", "Strength")
Label85.Caption = GetVar(App.Path & "\Quests.ini", "QUEST4", "Strength")
HScroll30.value = GetVar(App.Path & "\Quests.ini", "QUEST4", "Defence")
Label83.Caption = GetVar(App.Path & "\Quests.ini", "QUEST4", "Defence")
HScroll29.value = GetVar(App.Path & "\Quests.ini", "QUEST4", "Agility")
Label81.Caption = GetVar(App.Path & "\Quests.ini", "QUEST4", "Agility")
HScroll28.value = GetVar(App.Path & "\Quests.ini", "QUEST4", "Wisdom")
Label79.Caption = GetVar(App.Path & "\Quests.ini", "QUEST4", "Wisdom")
HScroll27.value = GetVar(App.Path & "\Quests.ini", "QUEST4", "LevelReq")
Label87.Caption = GetVar(App.Path & "\Quests.ini", "QUEST4", "LevelReq")
HScroll36.value = GetVar(App.Path & "\Quests.ini", "QUEST4", "Sprite")
Label97.Caption = GetVar(App.Path & "\Quests.ini", "QUEST4", "Sprite")
Check12.value = GetVar(App.Path & "\Quests.ini", "QUEST4", "FullHP")
Check11.value = GetVar(App.Path & "\Quests.ini", "QUEST4", "FullMP")
Check10.value = GetVar(App.Path & "\Quests.ini", "QUEST4", "FullSP")

Text19.text = GetVar(App.Path & "\Quests.ini", "QUEST3", "NPCname")
Text18.text = GetVar(App.Path & "\Quests.ini", "QUEST3", "NPCsay")
HScroll22.value = GetVar(App.Path & "\Quests.ini", "QUEST3", "IfItemNr")
Label72.Caption = GetVar(App.Path & "\Quests.ini", "QUEST3", "IfItemNr")
Text17.text = GetVar(App.Path & "\Quests.ini", "QUEST3", "FinishedNPCSay")
HScroll23.value = GetVar(App.Path & "\Quests.ini", "QUEST3", "ItemNr")
Label73.Caption = GetVar(App.Path & "\Quests.ini", "QUEST3", "ItemNr")
HScroll24.value = GetVar(App.Path & "\Quests.ini", "QUEST3", "StatPoints")
Label74.Caption = GetVar(App.Path & "\Quests.ini", "QUEST3", "StatPoints")
HScroll25.value = GetVar(App.Path & "\Quests.ini", "QUEST3", "Exp")
Label75.Caption = GetVar(App.Path & "\Quests.ini", "QUEST3", "Exp")
Text16.text = GetVar(App.Path & "\Quests.ini", "QUEST3", "NoItemNPCsay")
Text15.text = GetVar(App.Path & "\Quests.ini", "QUEST3", "WrongQuestNPCsay")
If GetVar(App.Path & "\Quests.ini", "QUEST3", "Keep") = 0 Then
Option5.value = True
Else
Option6.value = True
End If
HScroll18.value = GetVar(App.Path & "\Quests.ini", "QUEST3", "Strength")
Label60.Caption = GetVar(App.Path & "\Quests.ini", "QUEST3", "Strength")
HScroll19.value = GetVar(App.Path & "\Quests.ini", "QUEST3", "Defence")
Label62.Caption = GetVar(App.Path & "\Quests.ini", "QUEST3", "Defence")
HScroll20.value = GetVar(App.Path & "\Quests.ini", "QUEST3", "Agility")
Label64.Caption = GetVar(App.Path & "\Quests.ini", "QUEST3", "Agility")
HScroll21.value = GetVar(App.Path & "\Quests.ini", "QUEST3", "Wisdom")
Label66.Caption = GetVar(App.Path & "\Quests.ini", "QUEST3", "Wisdom")
HScroll17.value = GetVar(App.Path & "\Quests.ini", "QUEST3", "LevelReq")
Label58.Caption = GetVar(App.Path & "\Quests.ini", "QUEST3", "LevelReq")
HScroll26.value = GetVar(App.Path & "\Quests.ini", "QUEST3", "Sprite")
Label76.Caption = GetVar(App.Path & "\Quests.ini", "QUEST3", "Sprite")
Check7.value = GetVar(App.Path & "\Quests.ini", "QUEST3", "FullHP")
Check8.value = GetVar(App.Path & "\Quests.ini", "QUEST3", "FullMP")
Check9.value = GetVar(App.Path & "\Quests.ini", "QUEST3", "FullSP")

Text14.text = GetVar(App.Path & "\Quests.ini", "QUEST2", "NPCname")
Text13.text = GetVar(App.Path & "\Quests.ini", "QUEST2", "NPCsay")
HScroll15.value = GetVar(App.Path & "\Quests.ini", "QUEST2", "IfItemNr")
Label50.Caption = GetVar(App.Path & "\Quests.ini", "QUEST2", "IfItemNr")
Text12.text = GetVar(App.Path & "\Quests.ini", "QUEST2", "FinishedNPCSay")
HScroll14.value = GetVar(App.Path & "\Quests.ini", "QUEST2", "ItemNr")
Label49.Caption = GetVar(App.Path & "\Quests.ini", "QUEST2", "ItemNr")
HScroll13.value = GetVar(App.Path & "\Quests.ini", "QUEST2", "StatPoints")
Label49.Caption = GetVar(App.Path & "\Quests.ini", "QUEST2", "StatPoints")
HScroll12.value = GetVar(App.Path & "\Quests.ini", "QUEST2", "Exp")
Label47.Caption = GetVar(App.Path & "\Quests.ini", "QUEST2", "Exp")
Text11.text = GetVar(App.Path & "\Quests.ini", "QUEST2", "NoItemNPCsay")
Text10.text = GetVar(App.Path & "\Quests.ini", "QUEST2", "WrongQuestNPCsay")
If GetVar(App.Path & "\Quests.ini", "QUEST2", "Keep") = 0 Then
Option4.value = True
Else
Option3.value = True
End If
HScroll11.value = GetVar(App.Path & "\Quests.ini", "QUEST2", "Strength")
Label45.Caption = GetVar(App.Path & "\Quests.ini", "QUEST2", "Strength")
HScroll10.value = GetVar(App.Path & "\Quests.ini", "QUEST2", "Defence")
Label43.Caption = GetVar(App.Path & "\Quests.ini", "QUEST2", "Defence")
HScroll9.value = GetVar(App.Path & "\Quests.ini", "QUEST2", "Agility")
Label41.Caption = GetVar(App.Path & "\Quests.ini", "QUEST2", "Agility")
HScroll8.value = GetVar(App.Path & "\Quests.ini", "QUEST2", "Wisdom")
Label39.Caption = GetVar(App.Path & "\Quests.ini", "QUEST2", "Wisdom")
HScroll7.value = GetVar(App.Path & "\Quests.ini", "QUEST2", "LevelReq")
Label38.Caption = GetVar(App.Path & "\Quests.ini", "QUEST2", "LevelReq")
HScroll16.value = GetVar(App.Path & "\Quests.ini", "QUEST2", "Sprite")
Label57.Caption = GetVar(App.Path & "\Quests.ini", "QUEST2", "Sprite")
Check1.value = GetVar(App.Path & "\Quests.ini", "QUEST2", "FullHP")
Check2.value = GetVar(App.Path & "\Quests.ini", "QUEST2", "FullMP")
Check3.value = GetVar(App.Path & "\Quests.ini", "QUEST2", "FullSP")

Text1.text = GetVar(App.Path & "\Quests.ini", "QUEST1", "NPCname")
Text2.text = GetVar(App.Path & "\Quests.ini", "QUEST1", "NPCsay")
Text3.value = GetVar(App.Path & "\Quests.ini", "QUEST1", "IfItemNr")
Label30.Caption = GetVar(App.Path & "\Quests.ini", "QUEST1", "IfItemNr")
Text7.text = GetVar(App.Path & "\Quests.ini", "QUEST1", "FinishedNPCSay")
Text4.value = GetVar(App.Path & "\Quests.ini", "QUEST1", "ItemNr")
Label31.Caption = GetVar(App.Path & "\Quests.ini", "QUEST1", "ItemNr")
Text5.value = GetVar(App.Path & "\Quests.ini", "QUEST1", "StatPoints")
Label32.Caption = GetVar(App.Path & "\Quests.ini", "QUEST1", "StatPoints")
Text6.value = GetVar(App.Path & "\Quests.ini", "QUEST1", "Exp")
Label33.Caption = GetVar(App.Path & "\Quests.ini", "QUEST1", "Exp")
Text8.text = GetVar(App.Path & "\Quests.ini", "QUEST1", "NoItemNPCsay")
Text9.text = GetVar(App.Path & "\Quests.ini", "QUEST1", "WrongQuestNPCsay")
HScroll5.value = GetVar(App.Path & "\Quests.ini", "QUEST1", "LevelReq")
Label34.Caption = GetVar(App.Path & "\Quests.ini", "QUEST1", "LevelReq")
HScroll6.value = GetVar(App.Path & "\Quests.ini", "QUEST1", "Sprite")
Label35.Caption = GetVar(App.Path & "\Quests.ini", "QUEST1", "Sprite")
If GetVar(App.Path & "\Quests.ini", "QUEST1", "Keep") = 0 Then
Option2.value = True
Else
Option1.value = True
End If
HScroll1.value = GetVar(App.Path & "\Quests.ini", "QUEST1", "Strength")
Label23.Caption = GetVar(App.Path & "\Quests.ini", "QUEST1", "Strength")
HScroll2.value = GetVar(App.Path & "\Quests.ini", "QUEST1", "Defence")
Label25.Caption = GetVar(App.Path & "\Quests.ini", "QUEST1", "Defence")
HScroll3.value = GetVar(App.Path & "\Quests.ini", "QUEST1", "Agility")
Label27.Caption = GetVar(App.Path & "\Quests.ini", "QUEST1", "Agility")
HScroll4.value = GetVar(App.Path & "\Quests.ini", "QUEST1", "Wisdom")
Label29.Caption = GetVar(App.Path & "\Quests.ini", "QUEST1", "Wisdom")
Check1.value = GetVar(App.Path & "\Quests.ini", "QUEST1", "FullHP")
Check2.value = GetVar(App.Path & "\Quests.ini", "QUEST1", "FullMP")
Check3.value = GetVar(App.Path & "\Quests.ini", "QUEST1", "FullSP")

End Sub

Private Sub HScroll1_Change()
Label23.Caption = HScroll1.value
Call PutVar(App.Path & "\Quests.ini", "QUEST1", "Strength", HScroll1.value)
End Sub

Private Sub HScroll10_Change()
Label43.Caption = HScroll10.value
Call PutVar(App.Path & "\Quests.ini", "QUEST2", "Defence", HScroll10.value)
End Sub

Private Sub HScroll11_Change()
Label45.Caption = HScroll11.value
Call PutVar(App.Path & "\Quests.ini", "QUEST2", "Strength", HScroll11.value)
End Sub

Private Sub HScroll12_Change()
Label47.Caption = HScroll12.value
Call PutVar(App.Path & "\Quests.ini", "QUEST2", "Exp", HScroll12.value)
End Sub

Private Sub HScroll13_Change()
Label48.Caption = HScroll13.value
Call PutVar(App.Path & "\Quests.ini", "QUEST2", "StatPoints", HScroll13.value)
End Sub

Private Sub HScroll14_Change()
Label49.Caption = HScroll14.value
Call PutVar(App.Path & "\Quests.ini", "QUEST2", "ItemNr", HScroll14.value)
End Sub

Private Sub HScroll15_Change()
Label50.Caption = HScroll15.value
Call PutVar(App.Path & "\Quests.ini", "QUEST2", "IfItemNr", HScroll15.value)
End Sub

Private Sub HScroll16_Change()
Label57.Caption = HScroll16.value
Call PutVar(App.Path & "\Quests.ini", "QUEST2", "Sprite", HScroll16.value)
End Sub

Private Sub HScroll17_Change()
Label58.Caption = HScroll17.value
Call PutVar(App.Path & "\Quests.ini", "QUEST3", "LevelReq", HScroll17.value)
End Sub

Private Sub HScroll18_Change()
Label60.Caption = HScroll18.value
Call PutVar(App.Path & "\Quests.ini", "QUEST3", "Strength", HScroll18.value)
End Sub

Private Sub HScroll19_Change()
Label62.Caption = HScroll19.value
Call PutVar(App.Path & "\Quests.ini", "QUEST3", "Defence", HScroll19.value)
End Sub

Private Sub HScroll2_Change()
Label25.Caption = HScroll2.value
Call PutVar(App.Path & "\Quests.ini", "QUEST1", "Defence", HScroll2.value)
End Sub

Private Sub HScroll20_Change()
Label64.Caption = HScroll20.value
Call PutVar(App.Path & "\Quests.ini", "QUEST3", "Agility", HScroll20.value)
End Sub

Private Sub HScroll21_Change()
Label66.Caption = HScroll21.value
Call PutVar(App.Path & "\Quests.ini", "QUEST3", "Wisdom", HScroll21.value)
End Sub

Private Sub HScroll22_Change()
Label72.Caption = HScroll22.value
Call PutVar(App.Path & "\Quests.ini", "QUEST3", "IfItemNr", HScroll22.value)
End Sub

Private Sub HScroll23_Change()
Label73.Caption = HScroll23.value
Call PutVar(App.Path & "\Quests.ini", "QUEST3", "ItemNr", HScroll23.value)
End Sub

Private Sub HScroll24_Change()
Label74.Caption = HScroll24.value
Call PutVar(App.Path & "\Quests.ini", "QUEST3", "StatPoints", HScroll24.value)
End Sub

Private Sub HScroll25_Change()
Label75.Caption = HScroll25.value
Call PutVar(App.Path & "\Quests.ini", "QUEST3", "Exp", HScroll25.value)
End Sub

Private Sub HScroll26_Change()
Label76.Caption = HScroll26.value
Call PutVar(App.Path & "\Quests.ini", "QUEST3", "Sprite", HScroll26.value)
End Sub

Private Sub HScroll27_Change()
Label78.Caption = HScroll27.value
Call PutVar(App.Path & "\Quests.ini", "QUEST4", "LevelReq", HScroll27.value)
End Sub

Private Sub HScroll28_Change()
Label79.Caption = HScroll28.value
Call PutVar(App.Path & "\Quests.ini", "QUEST4", "Wisdom", HScroll28.value)
End Sub

Private Sub HScroll29_Change()
Label81.Caption = HScroll29.value
Call PutVar(App.Path & "\Quests.ini", "QUEST4", "Agility", HScroll29.value)
End Sub

Private Sub HScroll3_Change()
Label27.Caption = HScroll3.value
Call PutVar(App.Path & "\Quests.ini", "QUEST1", "Agility", HScroll3.value)
End Sub

Private Sub HScroll30_Change()
Label83.Caption = HScroll30.value
Call PutVar(App.Path & "\Quests.ini", "QUEST4", "Defence", HScroll30.value)
End Sub

Private Sub HScroll31_Change()
Label85.Caption = HScroll31.value
Call PutVar(App.Path & "\Quests.ini", "QUEST4", "Strength", HScroll31.value)
End Sub

Private Sub HScroll32_Change()
Label87.Caption = HScroll32.value
Call PutVar(App.Path & "\Quests.ini", "QUEST4", "Exp", HScroll32.value)
End Sub

Private Sub HScroll33_Change()
Label88.Caption = HScroll33.value
Call PutVar(App.Path & "\Quests.ini", "QUEST4", "StatPoints", HScroll33.value)
End Sub

Private Sub HScroll34_Change()
Label89.Caption = HScroll34.value
Call PutVar(App.Path & "\Quests.ini", "QUEST4", "ItemNr", HScroll34.value)
End Sub

Private Sub HScroll35_Change()
Label90.Caption = HScroll35.value
Call PutVar(App.Path & "\Quests.ini", "QUEST4", "IfItemNr", HScroll35.value)
End Sub

Private Sub HScroll36_Change()
Label97.Caption = HScroll36.value
Call PutVar(App.Path & "\Quests.ini", "QUEST4", "Sprite", HScroll36.value)
End Sub

Private Sub HScroll4_Change()
Label29.Caption = HScroll4.value
Call PutVar(App.Path & "\Quests.ini", "QUEST1", "Wisdom", HScroll4.value)
End Sub

Private Sub HScroll5_Change()
Label34.Caption = HScroll5.value
Call PutVar(App.Path & "\Quests.ini", "QUEST1", "LevelReq", HScroll5.value)
End Sub

Private Sub HScroll6_Change()
Label35.Caption = HScroll6.value
Call PutVar(App.Path & "\Quests.ini", "QUEST1", "Sprite", HScroll6.value)
End Sub

Private Sub HScroll7_Change()
Label38.Caption = HScroll7.value
Call PutVar(App.Path & "\Quests.ini", "QUEST2", "LevelReq", HScroll7.value)
End Sub

Private Sub HScroll8_Change()
Label39.Caption = HScroll8.value
Call PutVar(App.Path & "\Quests.ini", "QUEST2", "Wisdom", HScroll8.value)
End Sub

Private Sub HScroll9_Change()
Label41.Caption = HScroll9.value
Call PutVar(App.Path & "\Quests.ini", "QUEST2", "Agility", HScroll9.value)
End Sub

Private Sub Option1_Click()
Call PutVar(App.Path & "\Quests.ini", "QUEST1", "Keep", "1")
End Sub

Private Sub Option2_Click()
Call PutVar(App.Path & "\Quests.ini", "QUEST1", "Keep", "2")
End Sub

Private Sub Option5_Click()
Call PutVar(App.Path & "\Quests.ini", "QUEST3", "Keep", "1")
End Sub

Private Sub Option6_Click()
Call PutVar(App.Path & "\Quests.ini", "QUEST3", "Keep", "0")
End Sub

Private Sub Option7_Click()
Call PutVar(App.Path & "\Quests.ini", "QUEST4", "Keep", "0")
End Sub

Private Sub Option8_Click()
Call PutVar(App.Path & "\Quests.ini", "QUEST4", "Keep", "1")
End Sub

Private Sub Text3_Change()
Label30.Caption = Text3.value
Call PutVar(App.Path & "\Quests.ini", "QUEST1", "IfItemNr", Text3.value)
End Sub

Private Sub Text4_Change()
Label31.Caption = Text4.value
Call PutVar(App.Path & "\Quests.ini", "QUEST1", "ItemNr", Text4.value)
End Sub

Private Sub Text5_Change()
Label32.Caption = Text5.value
Call PutVar(App.Path & "\Quests.ini", "QUEST1", "StatPoints", Text5.value)
End Sub

Private Sub Text6_Change()
Label33.Caption = Text6.value
Call PutVar(App.Path & "\Quests.ini", "QUEST1", "Exp", Text6.value)
End Sub
