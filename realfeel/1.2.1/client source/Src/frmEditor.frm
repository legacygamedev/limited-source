VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmEditor 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Mutli-Data Editor"
   ClientHeight    =   8070
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8070
   ScaleWidth      =   9615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picSprites 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   720
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   118
      Top             =   8160
      Width           =   480
   End
   Begin VB.PictureBox picItems 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   120
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   117
      Top             =   8160
      Width           =   480
   End
   Begin VB.Timer tmrPic 
      Interval        =   50
      Left            =   9225
      Top             =   7650
   End
   Begin TabDlg.SSTab ssEditor 
      Height          =   8055
      Left            =   15
      TabIndex        =   0
      Top             =   0
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   14208
      _Version        =   393216
      Tabs            =   5
      Tab             =   1
      TabsPerRow      =   5
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Item Editor"
      TabPicture(0)   =   "frmEditor.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cmbSound"
      Tab(0).Control(1)=   "cmdItemSave"
      Tab(0).Control(2)=   "lstItemEditor"
      Tab(0).Control(3)=   "cmbItemType"
      Tab(0).Control(4)=   "txtItemName"
      Tab(0).Control(5)=   "scrlItemPic"
      Tab(0).Control(6)=   "picItemPic"
      Tab(0).Control(7)=   "txtItemDescription"
      Tab(0).Control(8)=   "fraItemVitals"
      Tab(0).Control(9)=   "fraItemEquipment"
      Tab(0).Control(10)=   "fraItemSpell"
      Tab(0).Control(11)=   "Label56"
      Tab(0).Control(12)=   "Label1"
      Tab(0).Control(13)=   "Label5"
      Tab(0).Control(14)=   "lblItemPic"
      Tab(0).Control(15)=   "Label8"
      Tab(0).ControlCount=   16
      TabCaption(1)   =   "Spell Editor"
      TabPicture(1)   =   "frmEditor.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "lblSpellLevelReq"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label13"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label14"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "fraSpellVitals"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "fraGiveItem"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "scrlSpellLevelReq"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "cmbSpellClassReq"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "cmbSpellType"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "txtSpellName"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "lstSpellEditor"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "cmdSpellSave"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).ControlCount=   11
      TabCaption(2)   =   "Shop Editor"
      TabPicture(2)   =   "frmEditor.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdShopSave"
      Tab(2).Control(1)=   "lstShopEditor"
      Tab(2).Control(2)=   "txtJoinSay"
      Tab(2).Control(3)=   "txtShopName"
      Tab(2).Control(4)=   "txtLeaveSay"
      Tab(2).Control(5)=   "lstTradeItem"
      Tab(2).Control(6)=   "cmbItemGive"
      Tab(2).Control(7)=   "txtItemGiveValue"
      Tab(2).Control(8)=   "cmbItemGet"
      Tab(2).Control(9)=   "txtItemGetValue"
      Tab(2).Control(10)=   "cmdUpdate"
      Tab(2).Control(11)=   "chkFixesItems"
      Tab(2).Control(12)=   "txtStock"
      Tab(2).Control(13)=   "cmdRestock"
      Tab(2).Control(14)=   "Label22"
      Tab(2).Control(15)=   "Label21"
      Tab(2).Control(16)=   "Label20"
      Tab(2).Control(17)=   "Label19"
      Tab(2).Control(18)=   "Label18"
      Tab(2).Control(19)=   "Label17"
      Tab(2).Control(20)=   "Label16"
      Tab(2).Control(21)=   "Label15"
      Tab(2).Control(22)=   "Label12"
      Tab(2).ControlCount=   23
      TabCaption(3)   =   "Npc Editor"
      TabPicture(3)   =   "frmEditor.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "scrlR"
      Tab(3).Control(1)=   "scrlG"
      Tab(3).Control(2)=   "scrlB"
      Tab(3).Control(3)=   "txtEXP"
      Tab(3).Control(4)=   "txtHP"
      Tab(3).Control(5)=   "cmdNpcSave"
      Tab(3).Control(6)=   "lstNpcEditor"
      Tab(3).Control(7)=   "picSprite"
      Tab(3).Control(8)=   "scrlSprite"
      Tab(3).Control(9)=   "txtNpcName"
      Tab(3).Control(10)=   "cmbBehavior"
      Tab(3).Control(11)=   "scrlRange"
      Tab(3).Control(12)=   "scrlSTR"
      Tab(3).Control(13)=   "scrlDEF"
      Tab(3).Control(14)=   "scrlSPEED"
      Tab(3).Control(15)=   "scrlMAGI"
      Tab(3).Control(16)=   "txtChance"
      Tab(3).Control(17)=   "scrlNum"
      Tab(3).Control(18)=   "scrlValue"
      Tab(3).Control(19)=   "txtAttackSay"
      Tab(3).Control(20)=   "txtSpawnSecs"
      Tab(3).Control(21)=   "chkAfraid"
      Tab(3).Control(22)=   "lblB"
      Tab(3).Control(23)=   "lblG"
      Tab(3).Control(24)=   "lblR"
      Tab(3).Control(25)=   "Label61"
      Tab(3).Control(26)=   "Label58"
      Tab(3).Control(27)=   "Label57"
      Tab(3).Control(28)=   "Label23"
      Tab(3).Control(29)=   "Shape1"
      Tab(3).Control(30)=   "Label55"
      Tab(3).Control(31)=   "Label54"
      Tab(3).Control(32)=   "Label49"
      Tab(3).Control(33)=   "lblSprite"
      Tab(3).Control(34)=   "Label39"
      Tab(3).Control(35)=   "Label38"
      Tab(3).Control(36)=   "Label37"
      Tab(3).Control(37)=   "lblRange"
      Tab(3).Control(38)=   "Label36"
      Tab(3).Control(39)=   "lblSTR"
      Tab(3).Control(40)=   "Label35"
      Tab(3).Control(41)=   "lblDEF"
      Tab(3).Control(42)=   "Label34"
      Tab(3).Control(43)=   "lblSPEED"
      Tab(3).Control(44)=   "Label33"
      Tab(3).Control(45)=   "lblMAGI"
      Tab(3).Control(46)=   "Label32"
      Tab(3).Control(47)=   "Label31"
      Tab(3).Control(48)=   "lblNum"
      Tab(3).Control(49)=   "Label30"
      Tab(3).Control(50)=   "Label29"
      Tab(3).Control(51)=   "lblItemName"
      Tab(3).Control(52)=   "Label28"
      Tab(3).Control(53)=   "lblValue"
      Tab(3).Control(54)=   "Label27"
      Tab(3).Control(55)=   "Label26"
      Tab(3).Control(56)=   "Label25"
      Tab(3).Control(57)=   "Label24"
      Tab(3).ControlCount=   58
      TabCaption(4)   =   "Class Editor"
      TabPicture(4)   =   "frmEditor.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label40"
      Tab(4).Control(1)=   "Label41"
      Tab(4).Control(2)=   "Label42"
      Tab(4).Control(3)=   "Label43"
      Tab(4).Control(4)=   "Label44"
      Tab(4).Control(5)=   "Label45"
      Tab(4).Control(6)=   "Label46"
      Tab(4).Control(7)=   "Label47"
      Tab(4).Control(8)=   "Label48"
      Tab(4).Control(9)=   "lblClassSprite"
      Tab(4).Control(10)=   "Label50"
      Tab(4).Control(11)=   "Label51"
      Tab(4).Control(12)=   "Label52"
      Tab(4).Control(13)=   "Label53"
      Tab(4).Control(14)=   "lstClassEditor"
      Tab(4).Control(15)=   "cmdClassSave"
      Tab(4).Control(16)=   "txtClassName"
      Tab(4).Control(17)=   "scrlClassSprite"
      Tab(4).Control(18)=   "picClassSprite"
      Tab(4).Control(19)=   "txtClassHP"
      Tab(4).Control(20)=   "txtClassMP"
      Tab(4).Control(21)=   "txtClassSP"
      Tab(4).Control(22)=   "txtClassSTR"
      Tab(4).Control(23)=   "txtClassDEF"
      Tab(4).Control(24)=   "txtClassMAGI"
      Tab(4).Control(25)=   "txtClassSPD"
      Tab(4).Control(26)=   "txtClassMap"
      Tab(4).Control(27)=   "txtClassX"
      Tab(4).Control(28)=   "txtClassY"
      Tab(4).Control(29)=   "cmdClassCreate"
      Tab(4).Control(30)=   "cmdClassDelete"
      Tab(4).ControlCount=   31
      Begin VB.HScrollBar scrlR 
         Height          =   255
         Left            =   -73800
         Max             =   255
         TabIndex        =   159
         Top             =   1485
         Width           =   2895
      End
      Begin VB.HScrollBar scrlG 
         Height          =   255
         Left            =   -73800
         Max             =   255
         TabIndex        =   158
         Top             =   1770
         Width           =   2895
      End
      Begin VB.HScrollBar scrlB 
         Height          =   255
         Left            =   -73800
         Max             =   255
         TabIndex        =   157
         Top             =   2055
         Width           =   2895
      End
      Begin VB.CommandButton cmdClassDelete 
         Caption         =   "Delete"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -66720
         TabIndex        =   156
         Top             =   4560
         Width           =   1215
      End
      Begin VB.CommandButton cmdClassCreate 
         Caption         =   "Create"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -67973
         TabIndex        =   155
         Top             =   4560
         Width           =   1215
      End
      Begin VB.TextBox txtEXP 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -68820
         TabIndex        =   154
         Top             =   5625
         Width           =   2895
      End
      Begin VB.TextBox txtHP 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -68820
         TabIndex        =   153
         Top             =   5265
         Width           =   2895
      End
      Begin VB.ComboBox cmbSound 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -72720
         Style           =   2  'Dropdown List
         TabIndex        =   152
         Top             =   4560
         Width           =   2775
      End
      Begin VB.TextBox txtClassY 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -70920
         TabIndex        =   146
         Top             =   2760
         Width           =   735
      End
      Begin VB.TextBox txtClassX 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -72360
         TabIndex        =   144
         Top             =   2760
         Width           =   735
      End
      Begin VB.TextBox txtClassMap 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -73800
         TabIndex        =   142
         Top             =   2760
         Width           =   735
      End
      Begin VB.TextBox txtClassSPD 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -70305
         TabIndex        =   135
         Top             =   1785
         Width           =   735
      End
      Begin VB.TextBox txtClassMAGI 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -70305
         TabIndex        =   134
         Top             =   1425
         Width           =   735
      End
      Begin VB.TextBox txtClassDEF 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -71880
         TabIndex        =   133
         Top             =   1800
         Width           =   735
      End
      Begin VB.TextBox txtClassSTR 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -71880
         TabIndex        =   129
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox txtClassSP 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -73950
         TabIndex        =   128
         Top             =   2055
         Width           =   975
      End
      Begin VB.TextBox txtClassMP 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -73950
         TabIndex        =   127
         Top             =   1695
         Width           =   975
      End
      Begin VB.TextBox txtClassHP 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -73950
         TabIndex        =   126
         Top             =   1335
         Width           =   975
      End
      Begin VB.PictureBox picClassSprite 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   -70320
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   125
         Top             =   720
         Width           =   480
      End
      Begin VB.HScrollBar scrlClassSprite 
         Height          =   255
         Left            =   -73800
         Max             =   255
         TabIndex        =   123
         Top             =   960
         Width           =   2895
      End
      Begin VB.TextBox txtClassName 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -73800
         TabIndex        =   121
         Top             =   600
         Width           =   2895
      End
      Begin VB.CommandButton cmdClassSave 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -69241
         TabIndex        =   120
         Top             =   4560
         Width           =   1215
      End
      Begin VB.ListBox lstClassEditor 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3840
         Left            =   -69240
         TabIndex        =   119
         Top             =   480
         Width           =   3735
      End
      Begin VB.CommandButton cmdNpcSave 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -69240
         TabIndex        =   116
         Top             =   4560
         Width           =   1215
      End
      Begin VB.CommandButton cmdShopSave 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -69240
         TabIndex        =   115
         Top             =   4560
         Width           =   1215
      End
      Begin VB.CommandButton cmdSpellSave 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5760
         TabIndex        =   114
         Top             =   4560
         Width           =   1215
      End
      Begin VB.CommandButton cmdItemSave 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -69240
         TabIndex        =   113
         Top             =   4560
         Width           =   1215
      End
      Begin VB.ListBox lstNpcEditor 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3840
         Left            =   -69240
         TabIndex        =   112
         Top             =   480
         Width           =   3735
      End
      Begin VB.ListBox lstShopEditor 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3840
         Left            =   -69240
         TabIndex        =   111
         Top             =   480
         Width           =   3735
      End
      Begin VB.ListBox lstSpellEditor 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3840
         Left            =   5760
         TabIndex        =   110
         Top             =   480
         Width           =   3735
      End
      Begin VB.ListBox lstItemEditor 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3840
         Left            =   -69240
         TabIndex        =   109
         Top             =   480
         Width           =   3735
      End
      Begin VB.PictureBox picSprite 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   -70110
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   83
         Top             =   1185
         Width           =   480
      End
      Begin VB.HScrollBar scrlSprite 
         Height          =   255
         Left            =   -73800
         Max             =   255
         TabIndex        =   82
         Top             =   1200
         Width           =   2895
      End
      Begin VB.TextBox txtNpcName 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -73920
         TabIndex        =   81
         Top             =   360
         Width           =   3975
      End
      Begin VB.ComboBox cmbBehavior 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmEditor.frx":008C
         Left            =   -73770
         List            =   "frmEditor.frx":00A5
         Style           =   2  'Dropdown List
         TabIndex        =   80
         Top             =   4845
         Width           =   2895
      End
      Begin VB.HScrollBar scrlRange 
         Height          =   255
         Left            =   -73800
         Max             =   255
         TabIndex        =   79
         Top             =   2490
         Value           =   1
         Width           =   2895
      End
      Begin VB.HScrollBar scrlSTR 
         Height          =   255
         Left            =   -73800
         Max             =   9999
         TabIndex        =   78
         Top             =   2850
         Width           =   2895
      End
      Begin VB.HScrollBar scrlDEF 
         Height          =   255
         Left            =   -73800
         Max             =   9999
         TabIndex        =   77
         Top             =   3210
         Width           =   2895
      End
      Begin VB.HScrollBar scrlSPEED 
         Height          =   255
         Left            =   -73800
         Max             =   9999
         TabIndex        =   76
         Top             =   3570
         Width           =   2895
      End
      Begin VB.HScrollBar scrlMAGI 
         Height          =   255
         Left            =   -73800
         Max             =   9999
         TabIndex        =   75
         Top             =   3930
         Width           =   2895
      End
      Begin VB.TextBox txtChance 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -67740
         TabIndex        =   74
         Text            =   "0"
         Top             =   5985
         Width           =   1815
      End
      Begin VB.HScrollBar scrlNum 
         Height          =   255
         Left            =   -74280
         Max             =   255
         TabIndex        =   73
         Top             =   6600
         Value           =   1
         Width           =   3375
      End
      Begin VB.HScrollBar scrlValue 
         Height          =   255
         Left            =   -74280
         Max             =   255
         TabIndex        =   72
         Top             =   6960
         Value           =   1
         Width           =   3375
      End
      Begin VB.TextBox txtAttackSay 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -73920
         TabIndex        =   71
         Top             =   720
         Width           =   3975
      End
      Begin VB.TextBox txtSpawnSecs 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -67740
         TabIndex        =   70
         Text            =   "0"
         Top             =   7065
         Width           =   1815
      End
      Begin VB.CheckBox chkAfraid 
         Caption         =   "Afraid"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   -73740
         TabIndex        =   69
         Top             =   5220
         Width           =   1095
      End
      Begin VB.TextBox txtJoinSay 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -73560
         TabIndex        =   59
         Top             =   840
         Width           =   3975
      End
      Begin VB.TextBox txtShopName 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -73560
         TabIndex        =   58
         Top             =   360
         Width           =   3975
      End
      Begin VB.TextBox txtLeaveSay 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -73560
         TabIndex        =   57
         Top             =   1320
         Width           =   3975
      End
      Begin VB.ListBox lstTradeItem 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1740
         ItemData        =   "frmEditor.frx":00FF
         Left            =   -74880
         List            =   "frmEditor.frx":011B
         TabIndex        =   56
         Top             =   4560
         Width           =   5295
      End
      Begin VB.ComboBox cmbItemGive 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -73560
         Style           =   2  'Dropdown List
         TabIndex        =   55
         Top             =   2280
         Width           =   3975
      End
      Begin VB.TextBox txtItemGiveValue 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -73560
         TabIndex        =   54
         Text            =   "1"
         Top             =   2760
         Width           =   1335
      End
      Begin VB.ComboBox cmbItemGet 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -73560
         Style           =   2  'Dropdown List
         TabIndex        =   53
         Top             =   3240
         Width           =   3975
      End
      Begin VB.TextBox txtItemGetValue 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -73560
         TabIndex        =   52
         Text            =   "1"
         Top             =   3720
         Width           =   1335
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -72000
         TabIndex        =   51
         Top             =   3720
         Width           =   2415
      End
      Begin VB.CheckBox chkFixesItems 
         Caption         =   "Fixes Items"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   50
         Top             =   1800
         Width           =   5295
      End
      Begin VB.TextBox txtStock 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -73560
         TabIndex        =   49
         Text            =   "1"
         Top             =   4080
         Width           =   1335
      End
      Begin VB.ComboBox cmdRestock 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmEditor.frx":013F
         Left            =   -71160
         List            =   "frmEditor.frx":014C
         Style           =   2  'Dropdown List
         TabIndex        =   48
         Top             =   4080
         Width           =   1575
      End
      Begin VB.TextBox txtSpellName 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   960
         TabIndex        =   44
         Top             =   480
         Width           =   3975
      End
      Begin VB.ComboBox cmbSpellType 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmEditor.frx":0163
         Left            =   120
         List            =   "frmEditor.frx":017C
         Style           =   2  'Dropdown List
         TabIndex        =   39
         Top             =   2040
         Width           =   4815
      End
      Begin VB.ComboBox cmbSpellClassReq 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmEditor.frx":01BB
         Left            =   120
         List            =   "frmEditor.frx":01BD
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   960
         Width           =   4815
      End
      Begin VB.HScrollBar scrlSpellLevelReq 
         Height          =   255
         Left            =   960
         Max             =   255
         Min             =   1
         TabIndex        =   30
         Top             =   1560
         Value           =   1
         Width           =   3495
      End
      Begin VB.ComboBox cmbItemType 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmEditor.frx":01BF
         Left            =   -74760
         List            =   "frmEditor.frx":01ED
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   2400
         Width           =   4815
      End
      Begin VB.TextBox txtItemName 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -73920
         TabIndex        =   24
         Top             =   720
         Width           =   3975
      End
      Begin VB.HScrollBar scrlItemPic 
         Height          =   255
         Left            =   -73920
         Max             =   255
         TabIndex        =   19
         Top             =   1080
         Width           =   2895
      End
      Begin VB.PictureBox picItemPic 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   -70350
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   18
         Top             =   1095
         Width           =   480
      End
      Begin VB.TextBox txtItemDescription 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -73755
         TabIndex        =   11
         Top             =   1800
         Width           =   3810
      End
      Begin VB.Frame fraGiveItem 
         Caption         =   "Give Item"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   120
         TabIndex        =   32
         Top             =   2520
         Visible         =   0   'False
         Width           =   4815
         Begin VB.HScrollBar scrlSpellItemNum 
            Height          =   255
            Left            =   1320
            Max             =   255
            Min             =   1
            TabIndex        =   34
            Top             =   360
            Value           =   1
            Width           =   2895
         End
         Begin VB.HScrollBar scrlSpellItemValue 
            Height          =   255
            Left            =   1320
            TabIndex        =   33
            Top             =   840
            Width           =   2895
         End
         Begin VB.Label Label10 
            Caption         =   "Item"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   38
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label lblSpellItemNum 
            Alignment       =   1  'Right Justify
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4200
            TabIndex        =   37
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label9 
            Caption         =   "Value"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   36
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label lblSpellItemValue 
            Alignment       =   1  'Right Justify
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4200
            TabIndex        =   35
            Top             =   840
            Width           =   495
         End
      End
      Begin VB.Frame fraSpellVitals 
         Caption         =   "Vitals Data"
         Height          =   855
         Left            =   120
         TabIndex        =   40
         Top             =   2520
         Visible         =   0   'False
         Width           =   4815
         Begin VB.HScrollBar scrlSpellVitalMod 
            Height          =   255
            Left            =   1320
            Max             =   255
            TabIndex        =   41
            Top             =   360
            Value           =   1
            Width           =   2895
         End
         Begin VB.Label lblSpellVitalMod 
            Alignment       =   1  'Right Justify
            Caption         =   "1"
            Height          =   255
            Left            =   4200
            TabIndex        =   43
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label11 
            Caption         =   "Vital Mod"
            Height          =   255
            Left            =   120
            TabIndex        =   42
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.Frame fraItemVitals 
         Caption         =   "Vitals Data"
         Height          =   855
         Left            =   -74760
         TabIndex        =   20
         Top             =   2880
         Visible         =   0   'False
         Width           =   4815
         Begin VB.HScrollBar scrlItemVitalMod 
            Height          =   255
            Left            =   1320
            Max             =   255
            TabIndex        =   21
            Top             =   360
            Value           =   1
            Width           =   2895
         End
         Begin VB.Label Label4 
            Caption         =   "Vital Mod"
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label lblItemVitalMod 
            Alignment       =   1  'Right Justify
            Caption         =   "1"
            Height          =   375
            Left            =   4200
            TabIndex        =   22
            Top             =   360
            Width           =   495
         End
      End
      Begin VB.Frame fraItemEquipment 
         Caption         =   "Equipment Data"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   -74760
         TabIndex        =   1
         Top             =   2880
         Visible         =   0   'False
         Width           =   4815
         Begin VB.HScrollBar scrlItemDurability 
            Height          =   255
            Left            =   1320
            Max             =   255
            TabIndex        =   6
            Top             =   360
            Value           =   1
            Width           =   2895
         End
         Begin VB.HScrollBar scrlItemStrength 
            Height          =   255
            Left            =   1320
            Max             =   255
            TabIndex        =   5
            Top             =   840
            Value           =   1
            Width           =   2895
         End
         Begin VB.CheckBox chkBow 
            Caption         =   "Bow Type"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   2880
            TabIndex        =   4
            Top             =   1200
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.CheckBox chkArrow 
            Caption         =   "Arrow Type"
            Height          =   270
            Left            =   2880
            TabIndex        =   3
            Top             =   1200
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.CheckBox chkUnBreakable 
            Caption         =   "UnBreakable?"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   960
            TabIndex        =   2
            Top             =   1200
            Width           =   1815
         End
         Begin VB.Label lblItemStrength 
            Alignment       =   1  'Right Justify
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4200
            TabIndex        =   10
            Top             =   840
            Width           =   495
         End
         Begin VB.Label Label2 
            Caption         =   "Durability"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label3 
            Caption         =   "Strength"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   840
            Width           =   975
         End
         Begin VB.Label lblItemDurability 
            Alignment       =   1  'Right Justify
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4200
            TabIndex        =   7
            Top             =   360
            Width           =   495
         End
      End
      Begin VB.Frame fraItemSpell 
         Caption         =   "Spell Data"
         Height          =   1560
         Left            =   -74760
         TabIndex        =   12
         Top             =   2880
         Visible         =   0   'False
         Width           =   4815
         Begin VB.HScrollBar scrlItemSpell 
            Height          =   255
            Left            =   1320
            Max             =   255
            TabIndex        =   13
            Top             =   840
            Value           =   1
            Width           =   2895
         End
         Begin VB.Label lblSpell 
            Alignment       =   1  'Right Justify
            Caption         =   "1"
            Height          =   255
            Left            =   4200
            TabIndex        =   17
            Top             =   840
            Width           =   495
         End
         Begin VB.Label Label7 
            Caption         =   "Num"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label Label6 
            Caption         =   "Name"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label lblItemSpellName 
            Height          =   375
            Left            =   1320
            TabIndex        =   14
            Top             =   360
            Width           =   3375
         End
      End
      Begin VB.Label lblB 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -70920
         TabIndex        =   166
         Top             =   2055
         Width           =   495
      End
      Begin VB.Label lblG 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -70920
         TabIndex        =   165
         Top             =   1770
         Width           =   495
      End
      Begin VB.Label lblR 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -70920
         TabIndex        =   164
         Top             =   1485
         Width           =   495
      End
      Begin VB.Label Label61 
         Caption         =   "B"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -73995
         TabIndex        =   163
         Top             =   2070
         Width           =   210
      End
      Begin VB.Label Label58 
         Caption         =   "G"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -73995
         TabIndex        =   162
         Top             =   1770
         Width           =   210
      End
      Begin VB.Label Label57 
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -73995
         TabIndex        =   161
         Top             =   1485
         Width           =   210
      End
      Begin VB.Label Label23 
         Caption         =   "Tint"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74580
         TabIndex        =   160
         Top             =   1485
         Width           =   405
      End
      Begin VB.Shape Shape1 
         FillStyle       =   0  'Solid
         Height          =   465
         Left            =   -74625
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label Label56 
         Caption         =   "Sound when used (.WAV):"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   151
         Top             =   4560
         Width           =   1935
      End
      Begin VB.Label Label55 
         Caption         =   "EXP:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -69300
         TabIndex        =   150
         Top             =   5625
         Width           =   855
      End
      Begin VB.Label Label54 
         Caption         =   "HP:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -69300
         TabIndex        =   149
         Top             =   5265
         Width           =   855
      End
      Begin VB.Label Label53 
         Caption         =   "Enter the starting location for the class below:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74895
         TabIndex        =   148
         Top             =   2400
         Width           =   4440
      End
      Begin VB.Label Label52 
         Alignment       =   1  'Right Justify
         Caption         =   "Y:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -71520
         TabIndex        =   147
         Top             =   2760
         Width           =   495
      End
      Begin VB.Label Label51 
         Alignment       =   1  'Right Justify
         Caption         =   "X:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -72960
         TabIndex        =   145
         Top             =   2760
         Width           =   495
      End
      Begin VB.Label Label50 
         Alignment       =   1  'Right Justify
         Caption         =   "Map:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74280
         TabIndex        =   143
         Top             =   2760
         Width           =   375
      End
      Begin VB.Label lblClassSprite 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -70920
         TabIndex        =   141
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label49 
         Caption         =   "(This is currently a 1 out of [NUMBER] ratio so enter a second number)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   -69540
         TabIndex        =   140
         Top             =   6345
         Width           =   3150
      End
      Begin VB.Label Label48 
         Alignment       =   1  'Right Justify
         Caption         =   "Speed:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -71040
         TabIndex        =   139
         Top             =   1785
         Width           =   630
      End
      Begin VB.Label Label47 
         Alignment       =   1  'Right Justify
         Caption         =   "Magic:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -70905
         TabIndex        =   138
         Top             =   1425
         Width           =   495
      End
      Begin VB.Label Label46 
         Alignment       =   1  'Right Justify
         Caption         =   "Defense:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -72720
         TabIndex        =   137
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label45 
         Alignment       =   1  'Right Justify
         Caption         =   "Strength:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -72795
         TabIndex        =   136
         Top             =   1440
         Width           =   810
      End
      Begin VB.Label Label44 
         Alignment       =   1  'Right Justify
         Caption         =   "SP:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -75030
         TabIndex        =   132
         Top             =   2055
         Width           =   975
      End
      Begin VB.Label Label43 
         Alignment       =   1  'Right Justify
         Caption         =   "MP:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -75030
         TabIndex        =   131
         Top             =   1695
         Width           =   975
      End
      Begin VB.Label Label42 
         Alignment       =   1  'Right Justify
         Caption         =   "HP:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -75030
         TabIndex        =   130
         Top             =   1335
         Width           =   975
      End
      Begin VB.Label Label41 
         Alignment       =   1  'Right Justify
         Caption         =   "Class Sprite:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   124
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label40 
         Alignment       =   1  'Right Justify
         Caption         =   "Class Name:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   122
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lblSprite 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -70920
         TabIndex        =   108
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label Label39 
         Caption         =   "Sprite"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74640
         TabIndex        =   107
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label38 
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   106
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label37 
         Caption         =   "Behavior"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74610
         TabIndex        =   105
         Top             =   4845
         Width           =   735
      End
      Begin VB.Label lblRange 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -70920
         TabIndex        =   104
         Top             =   2490
         Width           =   495
      End
      Begin VB.Label Label36 
         Caption         =   "Range"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74640
         TabIndex        =   103
         Top             =   2490
         Width           =   855
      End
      Begin VB.Label lblSTR 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -70920
         TabIndex        =   102
         Top             =   2850
         Width           =   495
      End
      Begin VB.Label Label35 
         Caption         =   "STR"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74640
         TabIndex        =   101
         Top             =   2850
         Width           =   855
      End
      Begin VB.Label lblDEF 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -70920
         TabIndex        =   100
         Top             =   3210
         Width           =   495
      End
      Begin VB.Label Label34 
         Caption         =   "DEF"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74640
         TabIndex        =   99
         Top             =   3210
         Width           =   855
      End
      Begin VB.Label lblSPEED 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -70920
         TabIndex        =   98
         Top             =   3570
         Width           =   495
      End
      Begin VB.Label Label33 
         Caption         =   "SPD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74640
         TabIndex        =   97
         Top             =   3570
         Width           =   855
      End
      Begin VB.Label lblMAGI 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -70920
         TabIndex        =   96
         Top             =   3930
         Width           =   495
      End
      Begin VB.Label Label32 
         Caption         =   "MAGI"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74640
         TabIndex        =   95
         Top             =   3930
         Width           =   855
      End
      Begin VB.Label Label31 
         Caption         =   "Drop Item Chance"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -69420
         TabIndex        =   94
         Top             =   5985
         Width           =   1335
      End
      Begin VB.Label lblNum 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -70920
         TabIndex        =   93
         Top             =   6600
         Width           =   495
      End
      Begin VB.Label Label30 
         Caption         =   "Num"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   92
         Top             =   6600
         Width           =   855
      End
      Begin VB.Label Label29 
         Caption         =   "Item"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   91
         Top             =   6120
         Width           =   495
      End
      Begin VB.Label lblItemName 
         Height          =   375
         Left            =   -74280
         TabIndex        =   90
         Top             =   6120
         Width           =   3975
      End
      Begin VB.Label Label28 
         Caption         =   "Value"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   89
         Top             =   6960
         Width           =   735
      End
      Begin VB.Label lblValue 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -70920
         TabIndex        =   88
         Top             =   6960
         Width           =   495
      End
      Begin VB.Label Label27 
         Caption         =   "Start HP"
         Height          =   375
         Left            =   -74640
         TabIndex        =   87
         Top             =   8160
         Width           =   975
      End
      Begin VB.Label Label26 
         Caption         =   "Exp Given"
         Height          =   375
         Left            =   -72240
         TabIndex        =   86
         Top             =   8160
         Width           =   1215
      End
      Begin VB.Label Label25 
         Caption         =   "Say"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   85
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label24 
         Caption         =   " Spawn Rate      (in seconds)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   -69540
         TabIndex        =   84
         Top             =   6990
         Width           =   1335
      End
      Begin VB.Label Label22 
         Caption         =   "Join Say"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   68
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label21 
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   67
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label20 
         Caption         =   "Leave Say"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   66
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label19 
         Caption         =   "Item Give"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   65
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label18 
         Caption         =   "Value"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   64
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Label Label17 
         Caption         =   "Item Get"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   63
         Top             =   3240
         Width           =   1215
      End
      Begin VB.Label Label16 
         Caption         =   "Value"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   62
         Top             =   3720
         Width           =   1215
      End
      Begin VB.Label Label15 
         Caption         =   "Item Stock"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   61
         Top             =   4080
         Width           =   1215
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Caption         =   "Restock Time:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   -72120
         TabIndex        =   60
         Top             =   4065
         Width           =   855
      End
      Begin VB.Label Label14 
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   47
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label13 
         Caption         =   "Level"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label lblSpellLevelReq 
         Alignment       =   1  'Right Justify
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4440
         TabIndex        =   45
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Name:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   29
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Pic:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   28
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lblItemPic 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -71010
         TabIndex        =   27
         Top             =   1065
         Width           =   495
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Description:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   26
         Top             =   1800
         Width           =   1005
      End
   End
End
Attribute VB_Name = "frmEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbItemType_Click()
    If (cmbItemType.ListIndex >= ITEM_TYPE_WEAPON) And (cmbItemType.ListIndex <= ITEM_TYPE_SHIELD) Or cmbItemType.ListIndex = ITEM_TYPE_BOW Then
        fraItemEquipment.Visible = True
        If cmbItemType.ListIndex = ITEM_TYPE_WEAPON Then
            chkBow.Visible = True
        Else
            chkBow.Visible = False
        End If
        
        If cmbItemType.ListIndex = ITEM_TYPE_SHIELD Then
            chkArrow.Visible = True
        Else
            chkArrow.Visible = False
        End If
        
    Else
        fraItemEquipment.Visible = False
    End If
    
    If (cmbItemType.ListIndex >= ITEM_TYPE_POTIONADDHP) And (cmbItemType.ListIndex <= ITEM_TYPE_POTIONSUBSP) Then
        fraItemVitals.Visible = True
    Else
        fraItemVitals.Visible = False
    End If
    
    If (cmbItemType.ListIndex = ITEM_TYPE_SPELL) Then
        fraItemSpell.Visible = True
    Else
        fraItemSpell.Visible = False
    End If
End Sub

Private Sub cmdClassCreate_Click()
    EditorIndex = Max_Classes + 1
    Call SendData("CREATECLASS" & SEP_CHAR & END_CHAR)
    Call SendRequestEditClass
    MsgBox "Class created!"
End Sub

Private Sub cmdClassDelete_Click()
    EditorIndex = lstClassEditor.ListIndex + 1
    Call SendData("DELETECLASS" & SEP_CHAR & (EditorIndex - 1) & SEP_CHAR & END_CHAR)
    If lstClassEditor.ListCount - 1 < EditorIndex Then EditorIndex = lstClassEditor.ListCount - 1
    Call SendRequestEditClass
    MsgBox "Class deleted!"
End Sub

Private Sub cmdClassSave_Click()
    EditorIndex = lstClassEditor.ListIndex + 1
    Call ClassEditorSave
    Call SendRequestEditClass
    MsgBox "Class saved!"
End Sub

Private Sub cmdItemSave_Click()
    EditorIndex = lstItemEditor.ListIndex + 1
    Call ItemEditorSave
    Call SendRequestEditItem
    MsgBox "Item saved!"
End Sub

Private Sub cmdNpcSave_Click()
    EditorIndex = lstNpcEditor.ListIndex + 1
    Call NpcEditorSave
    Call SendRequestEditNpc
    MsgBox "Npc saved!"
End Sub

Private Sub cmdShopSave_Click()
    EditorIndex = lstShopEditor.ListIndex + 1
    Call ShopEditorSave
    Call SendRequestEditShop
    MsgBox "Shop saved!"
End Sub

Private Sub cmdSpellSave_Click()
    EditorIndex = lstSpellEditor.ListIndex + 1
    Call SpellEditorSave
    Call SendRequestEditSpell
    MsgBox "Spell saved!"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    lstItemEditor.Clear
    lstSpellEditor.Clear
    lstShopEditor.Clear
    lstNpcEditor.Clear
    Debug.Print "Cleared lists"
End Sub

Private Sub lstClassEditor_Click()
    EditorIndex = lstClassEditor.ListIndex + 1
    Call SendData("EDITCLASS" & SEP_CHAR & (EditorIndex - 1) & SEP_CHAR & END_CHAR)
End Sub

Private Sub lstItemEditor_Click()
    EditorIndex = lstItemEditor.ListIndex + 1
    Call SendData("EDITITEM" & SEP_CHAR & EditorIndex & SEP_CHAR & END_CHAR)
End Sub

Private Sub lstSpellEditor_Click()
    EditorIndex = lstSpellEditor.ListIndex + 1
    Call SendData("EDITSPELL" & SEP_CHAR & EditorIndex & SEP_CHAR & END_CHAR)
End Sub

Private Sub lstShopEditor_Click()
    EditorIndex = lstShopEditor.ListIndex + 1
    Call SendData("EDITSHOP" & SEP_CHAR & EditorIndex & SEP_CHAR & END_CHAR)
End Sub

Private Sub lstNpcEditor_Click()
    EditorIndex = lstNpcEditor.ListIndex + 1
    Call SendData("EDITNPC" & SEP_CHAR & EditorIndex & SEP_CHAR & END_CHAR)
End Sub

Private Sub scrlClassSprite_Change()
    lblClassSprite.Caption = STR(scrlClassSprite.Value)
End Sub

Private Sub scrlItemPic_Change()
    lblItemPic.Caption = STR(scrlItemPic.Value)
End Sub

Private Sub scrlItemVitalMod_Change()
    lblItemVitalMod.Caption = STR(scrlItemVitalMod.Value)
End Sub

Private Sub scrlItemVitalAdd_Change()
End Sub

Private Sub scrlItemDurability_Change()
    lblItemDurability.Caption = STR(scrlItemDurability.Value)
End Sub

Private Sub scrlItemStrength_Change()
    lblItemStrength.Caption = STR(scrlItemStrength.Value)
End Sub

Private Sub scrlItemSpell_Change()
    lblItemSpellName.Caption = Trim$(Spell(scrlItemSpell.Value).Name)
    lblItemSpellName.Caption = STR(scrlItemSpell.Value)
End Sub

Private Sub cmbSpellType_Click()
    If cmbSpellType.ListIndex <> SPELL_TYPE_GIVEITEM Then
        fraSpellVitals.Visible = True
        fraGiveItem.Visible = False
    Else
        fraSpellVitals.Visible = False
        fraGiveItem.Visible = True
    End If
End Sub

Private Sub scrlR_Change()
    Shape1.FillColor = RGB(scrlR.Value, scrlG.Value, scrlB.Value)
    lblR = scrlR.Value
End Sub

Private Sub scrlB_Change()
    Shape1.FillColor = RGB(scrlR.Value, scrlG.Value, scrlB.Value)
    lblB = scrlB.Value
End Sub

Private Sub scrlG_Change()
    Shape1.FillColor = RGB(scrlR.Value, scrlG.Value, scrlB.Value)
    lblG = scrlG.Value
End Sub

Private Sub scrlSpellItemNum_Change()
    fraGiveItem.Caption = "Give Item " & Trim$(Item(scrlSpellItemNum.Value).Name)
    lblSpellItemNum.Caption = STR(scrlSpellItemNum.Value)
End Sub

Private Sub scrlSpellItemValue_Change()
    lblSpellItemValue.Caption = STR(scrlSpellItemValue.Value)
End Sub

Private Sub scrlSpellLevelReq_Change()
    lblSpellLevelReq.Caption = STR(scrlSpellLevelReq.Value)
End Sub

Private Sub scrlSpellVitalMod_Change()
    lblSpellVitalMod.Caption = STR(scrlSpellVitalMod.Value)
End Sub

Private Sub cmdUpdate_Click()
Dim Index As Long

    Index = lstTradeItem.ListIndex + 1
    Shop(EditorIndex).TradeItem(Index).GiveItem = cmbItemGive.ListIndex
    Shop(EditorIndex).TradeItem(Index).GiveValue = Val(txtItemGiveValue.Text)
    Shop(EditorIndex).TradeItem(Index).GetItem = cmbItemGet.ListIndex
    Shop(EditorIndex).TradeItem(Index).GetValue = Val(txtItemGetValue.Text)
    Shop(EditorIndex).TradeItem(Index).Stock = CInt(txtStock.Text)
    Shop(EditorIndex).TradeItem(Index).MaxStock = CInt(txtStock.Text)
    Call UpdateShopTrade
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
    If scrlNum.Value > 0 Then
        lblItemName.Caption = Trim$(Item(scrlNum.Value).Name)
    End If
End Sub

Private Sub scrlValue_Change()
    lblValue.Caption = STR(scrlValue.Value)
End Sub

Private Sub tmrPic_Timer()
    Call ItemEditorBltItem
    Call NpcEditorBltSprite
    Call ClassEditorBltSprite
End Sub
