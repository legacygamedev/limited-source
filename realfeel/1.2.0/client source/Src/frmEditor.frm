VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
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
      TabIndex        =   119
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
      TabIndex        =   118
      Top             =   8160
      Width           =   480
   End
   Begin VB.Timer tmrPic 
      Interval        =   50
      Left            =   0
      Top             =   0
   End
   Begin TabDlg.SSTab ssEditor 
      Height          =   8055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   14208
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Item Editor"
      TabPicture(0)   =   "frmEditor.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label8"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblItemPic"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label5"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label56"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "fraItemVitals"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtItemDescription"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "picItemPic"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "scrlItemPic"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtItemName"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmbItemType"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lstItemEditor"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cmdItemSave"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cmbSound"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "fraItemSpell"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "fraItemEquipment"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).ControlCount=   16
      TabCaption(1)   =   "Spell Editor"
      TabPicture(1)   =   "frmEditor.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblSpellLevelReq"
      Tab(1).Control(1)=   "Label13"
      Tab(1).Control(2)=   "Label14"
      Tab(1).Control(3)=   "fraSpellVitals"
      Tab(1).Control(4)=   "fraGiveItem"
      Tab(1).Control(5)=   "scrlSpellLevelReq"
      Tab(1).Control(6)=   "cmbSpellClassReq"
      Tab(1).Control(7)=   "cmbSpellType"
      Tab(1).Control(8)=   "txtSpellName"
      Tab(1).Control(9)=   "lstSpellEditor"
      Tab(1).Control(10)=   "cmdSpellSave"
      Tab(1).ControlCount=   11
      TabCaption(2)   =   "Shop Editor"
      TabPicture(2)   =   "frmEditor.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label12"
      Tab(2).Control(1)=   "Label15"
      Tab(2).Control(2)=   "Label16"
      Tab(2).Control(3)=   "Label17"
      Tab(2).Control(4)=   "Label18"
      Tab(2).Control(5)=   "Label19"
      Tab(2).Control(6)=   "Label20"
      Tab(2).Control(7)=   "Label21"
      Tab(2).Control(8)=   "Label22"
      Tab(2).Control(9)=   "cmdRestock"
      Tab(2).Control(10)=   "txtStock"
      Tab(2).Control(11)=   "chkFixesItems"
      Tab(2).Control(12)=   "cmdUpdate"
      Tab(2).Control(13)=   "txtItemGetValue"
      Tab(2).Control(14)=   "cmbItemGet"
      Tab(2).Control(15)=   "txtItemGiveValue"
      Tab(2).Control(16)=   "cmbItemGive"
      Tab(2).Control(17)=   "lstTradeItem"
      Tab(2).Control(18)=   "txtLeaveSay"
      Tab(2).Control(19)=   "txtShopName"
      Tab(2).Control(20)=   "txtJoinSay"
      Tab(2).Control(21)=   "lstShopEditor"
      Tab(2).Control(22)=   "cmdShopSave"
      Tab(2).ControlCount=   23
      TabCaption(3)   =   "Npc Editor"
      TabPicture(3)   =   "frmEditor.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label23"
      Tab(3).Control(1)=   "Label24"
      Tab(3).Control(2)=   "Label25"
      Tab(3).Control(3)=   "Label26"
      Tab(3).Control(4)=   "Label27"
      Tab(3).Control(5)=   "lblValue"
      Tab(3).Control(6)=   "Label28"
      Tab(3).Control(7)=   "lblItemName"
      Tab(3).Control(8)=   "Label29"
      Tab(3).Control(9)=   "Label30"
      Tab(3).Control(10)=   "lblNum"
      Tab(3).Control(11)=   "Label31"
      Tab(3).Control(12)=   "Label32"
      Tab(3).Control(13)=   "lblMAGI"
      Tab(3).Control(14)=   "Label33"
      Tab(3).Control(15)=   "lblSPEED"
      Tab(3).Control(16)=   "Label34"
      Tab(3).Control(17)=   "lblDEF"
      Tab(3).Control(18)=   "Label35"
      Tab(3).Control(19)=   "lblSTR"
      Tab(3).Control(20)=   "Label36"
      Tab(3).Control(21)=   "lblRange"
      Tab(3).Control(22)=   "Label37"
      Tab(3).Control(23)=   "Label38"
      Tab(3).Control(24)=   "Label39"
      Tab(3).Control(25)=   "lblSprite"
      Tab(3).Control(26)=   "Label49"
      Tab(3).Control(27)=   "Label54"
      Tab(3).Control(28)=   "Label55"
      Tab(3).Control(29)=   "chkAfraid"
      Tab(3).Control(30)=   "txtSpawnSecs"
      Tab(3).Control(31)=   "txtAttackSay"
      Tab(3).Control(32)=   "scrlValue"
      Tab(3).Control(33)=   "scrlNum"
      Tab(3).Control(34)=   "txtChance"
      Tab(3).Control(35)=   "scrlMAGI"
      Tab(3).Control(36)=   "scrlSPEED"
      Tab(3).Control(37)=   "scrlDEF"
      Tab(3).Control(38)=   "scrlSTR"
      Tab(3).Control(39)=   "scrlRange"
      Tab(3).Control(40)=   "cmbBehavior"
      Tab(3).Control(41)=   "txtNpcName"
      Tab(3).Control(42)=   "scrlSprite"
      Tab(3).Control(43)=   "picSprite"
      Tab(3).Control(44)=   "lstNpcEditor"
      Tab(3).Control(45)=   "cmdNpcSave"
      Tab(3).Control(46)=   "txtHP"
      Tab(3).Control(47)=   "txtEXP"
      Tab(3).ControlCount=   48
      TabCaption(4)   =   "Class Editor"
      TabPicture(4)   =   "frmEditor.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label40"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Label41"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "Label42"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "Label43"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).Control(4)=   "Label44"
      Tab(4).Control(4).Enabled=   0   'False
      Tab(4).Control(5)=   "Label45"
      Tab(4).Control(5).Enabled=   0   'False
      Tab(4).Control(6)=   "Label46"
      Tab(4).Control(6).Enabled=   0   'False
      Tab(4).Control(7)=   "Label47"
      Tab(4).Control(7).Enabled=   0   'False
      Tab(4).Control(8)=   "Label48"
      Tab(4).Control(8).Enabled=   0   'False
      Tab(4).Control(9)=   "lblClassSprite"
      Tab(4).Control(9).Enabled=   0   'False
      Tab(4).Control(10)=   "Label50"
      Tab(4).Control(10).Enabled=   0   'False
      Tab(4).Control(11)=   "Label51"
      Tab(4).Control(11).Enabled=   0   'False
      Tab(4).Control(12)=   "Label52"
      Tab(4).Control(12).Enabled=   0   'False
      Tab(4).Control(13)=   "Label53"
      Tab(4).Control(13).Enabled=   0   'False
      Tab(4).Control(14)=   "lstClassEditor"
      Tab(4).Control(14).Enabled=   0   'False
      Tab(4).Control(15)=   "cmdClassSave"
      Tab(4).Control(15).Enabled=   0   'False
      Tab(4).Control(16)=   "txtClassName"
      Tab(4).Control(16).Enabled=   0   'False
      Tab(4).Control(17)=   "scrlClassSprite"
      Tab(4).Control(17).Enabled=   0   'False
      Tab(4).Control(18)=   "picClassSprite"
      Tab(4).Control(18).Enabled=   0   'False
      Tab(4).Control(19)=   "txtClassHP"
      Tab(4).Control(19).Enabled=   0   'False
      Tab(4).Control(20)=   "txtClassMP"
      Tab(4).Control(20).Enabled=   0   'False
      Tab(4).Control(21)=   "txtClassSP"
      Tab(4).Control(21).Enabled=   0   'False
      Tab(4).Control(22)=   "txtClassSTR"
      Tab(4).Control(22).Enabled=   0   'False
      Tab(4).Control(23)=   "txtClassDEF"
      Tab(4).Control(23).Enabled=   0   'False
      Tab(4).Control(24)=   "txtClassMAGI"
      Tab(4).Control(24).Enabled=   0   'False
      Tab(4).Control(25)=   "txtClassSPD"
      Tab(4).Control(25).Enabled=   0   'False
      Tab(4).Control(26)=   "txtClassMap"
      Tab(4).Control(26).Enabled=   0   'False
      Tab(4).Control(27)=   "txtClassX"
      Tab(4).Control(27).Enabled=   0   'False
      Tab(4).Control(28)=   "txtClassY"
      Tab(4).Control(28).Enabled=   0   'False
      Tab(4).Control(29)=   "cmdClassCreate"
      Tab(4).Control(29).Enabled=   0   'False
      Tab(4).Control(30)=   "cmdClassDelete"
      Tab(4).Control(30).Enabled=   0   'False
      Tab(4).ControlCount=   31
      Begin VB.CommandButton cmdClassDelete 
         Caption         =   "Delete"
         Height          =   255
         Left            =   -66720
         TabIndex        =   157
         Top             =   4560
         Width           =   1215
      End
      Begin VB.CommandButton cmdClassCreate 
         Caption         =   "Create"
         Height          =   255
         Left            =   -67973
         TabIndex        =   156
         Top             =   4560
         Width           =   1215
      End
      Begin VB.TextBox txtEXP 
         Height          =   285
         Left            =   -74160
         TabIndex        =   155
         Top             =   4200
         Width           =   2895
      End
      Begin VB.TextBox txtHP 
         Height          =   285
         Left            =   -74160
         TabIndex        =   154
         Top             =   3840
         Width           =   2895
      End
      Begin VB.Frame fraItemEquipment 
         Caption         =   "Equipment Data"
         Height          =   1575
         Left            =   240
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
            Height          =   270
            Left            =   960
            TabIndex        =   2
            Top             =   1200
            Width           =   1815
         End
         Begin VB.Label lblItemStrength 
            Alignment       =   1  'Right Justify
            Caption         =   "1"
            Height          =   375
            Left            =   4200
            TabIndex        =   10
            Top             =   840
            Width           =   495
         End
         Begin VB.Label Label2 
            Caption         =   "Durability"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label3 
            Caption         =   "Strength"
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   840
            Width           =   975
         End
         Begin VB.Label lblItemDurability 
            Alignment       =   1  'Right Justify
            Caption         =   "1"
            Height          =   375
            Left            =   4200
            TabIndex        =   7
            Top             =   360
            Width           =   495
         End
      End
      Begin VB.Frame fraItemSpell 
         Caption         =   "Spell Data"
         Height          =   1335
         Left            =   240
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
      Begin VB.ComboBox cmbSound 
         Height          =   315
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   153
         Top             =   4560
         Width           =   2775
      End
      Begin VB.TextBox txtClassY 
         Height          =   285
         Left            =   -70920
         TabIndex        =   147
         Top             =   2760
         Width           =   735
      End
      Begin VB.TextBox txtClassX 
         Height          =   285
         Left            =   -72360
         TabIndex        =   145
         Top             =   2760
         Width           =   735
      End
      Begin VB.TextBox txtClassMap 
         Height          =   285
         Left            =   -73800
         TabIndex        =   143
         Top             =   2760
         Width           =   735
      End
      Begin VB.TextBox txtClassSPD 
         Height          =   285
         Left            =   -70440
         TabIndex        =   136
         Top             =   1800
         Width           =   735
      End
      Begin VB.TextBox txtClassMAGI 
         Height          =   285
         Left            =   -70440
         TabIndex        =   135
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox txtClassDEF 
         Height          =   285
         Left            =   -71880
         TabIndex        =   134
         Top             =   1800
         Width           =   735
      End
      Begin VB.TextBox txtClassSTR 
         Height          =   285
         Left            =   -71880
         TabIndex        =   130
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox txtClassSP 
         Height          =   285
         Left            =   -73800
         TabIndex        =   129
         Top             =   2040
         Width           =   975
      End
      Begin VB.TextBox txtClassMP 
         Height          =   285
         Left            =   -73800
         TabIndex        =   128
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox txtClassHP 
         Height          =   285
         Left            =   -73800
         TabIndex        =   127
         Top             =   1320
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
         TabIndex        =   126
         Top             =   720
         Width           =   480
      End
      Begin VB.HScrollBar scrlClassSprite 
         Height          =   255
         Left            =   -73800
         Max             =   255
         TabIndex        =   124
         Top             =   960
         Width           =   2895
      End
      Begin VB.TextBox txtClassName 
         Height          =   285
         Left            =   -73800
         TabIndex        =   122
         Top             =   600
         Width           =   2895
      End
      Begin VB.CommandButton cmdClassSave 
         Caption         =   "Save"
         Height          =   255
         Left            =   -69241
         TabIndex        =   121
         Top             =   4560
         Width           =   1215
      End
      Begin VB.ListBox lstClassEditor 
         Height          =   3960
         Left            =   -69240
         TabIndex        =   120
         Top             =   480
         Width           =   3735
      End
      Begin VB.CommandButton cmdNpcSave 
         Caption         =   "Save"
         Height          =   255
         Left            =   -69240
         TabIndex        =   117
         Top             =   4560
         Width           =   1215
      End
      Begin VB.CommandButton cmdShopSave 
         Caption         =   "Save"
         Height          =   255
         Left            =   -69240
         TabIndex        =   116
         Top             =   4560
         Width           =   1215
      End
      Begin VB.CommandButton cmdSpellSave 
         Caption         =   "Save"
         Height          =   255
         Left            =   -69240
         TabIndex        =   115
         Top             =   4560
         Width           =   1215
      End
      Begin VB.CommandButton cmdItemSave 
         Caption         =   "Save"
         Height          =   255
         Left            =   5760
         TabIndex        =   114
         Top             =   4560
         Width           =   1215
      End
      Begin VB.ListBox lstNpcEditor 
         Height          =   3960
         Left            =   -69240
         TabIndex        =   113
         Top             =   480
         Width           =   3735
      End
      Begin VB.ListBox lstShopEditor 
         Height          =   3960
         Left            =   -69240
         TabIndex        =   112
         Top             =   480
         Width           =   3735
      End
      Begin VB.ListBox lstSpellEditor 
         Height          =   3960
         Left            =   -69240
         TabIndex        =   111
         Top             =   480
         Width           =   3735
      End
      Begin VB.ListBox lstItemEditor 
         Height          =   3960
         Left            =   5760
         TabIndex        =   110
         Top             =   480
         Width           =   3735
      End
      Begin VB.PictureBox picSprite 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   -70320
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   83
         Top             =   1200
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
         Height          =   285
         Left            =   -73920
         TabIndex        =   81
         Top             =   360
         Width           =   3975
      End
      Begin VB.ComboBox cmbBehavior 
         Height          =   315
         ItemData        =   "frmEditor.frx":008C
         Left            =   -73800
         List            =   "frmEditor.frx":00A5
         Style           =   2  'Dropdown List
         TabIndex        =   80
         Top             =   3360
         Width           =   2895
      End
      Begin VB.HScrollBar scrlRange 
         Height          =   255
         Left            =   -73800
         Max             =   255
         TabIndex        =   79
         Top             =   1560
         Value           =   1
         Width           =   2895
      End
      Begin VB.HScrollBar scrlSTR 
         Height          =   255
         Left            =   -73800
         Max             =   9999
         TabIndex        =   78
         Top             =   1920
         Width           =   2895
      End
      Begin VB.HScrollBar scrlDEF 
         Height          =   255
         Left            =   -73800
         Max             =   9999
         TabIndex        =   77
         Top             =   2280
         Width           =   2895
      End
      Begin VB.HScrollBar scrlSPEED 
         Height          =   255
         Left            =   -73800
         Max             =   9999
         TabIndex        =   76
         Top             =   2640
         Width           =   2895
      End
      Begin VB.HScrollBar scrlMAGI 
         Height          =   255
         Left            =   -73800
         Max             =   9999
         TabIndex        =   75
         Top             =   3000
         Width           =   2895
      End
      Begin VB.TextBox txtChance 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -73080
         TabIndex        =   74
         Text            =   "0"
         Top             =   4560
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
         Height          =   285
         Left            =   -73920
         TabIndex        =   71
         Top             =   720
         Width           =   3975
      End
      Begin VB.TextBox txtSpawnSecs 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -73080
         TabIndex        =   70
         Text            =   "0"
         Top             =   5640
         Width           =   1815
      End
      Begin VB.CheckBox chkAfraid 
         Caption         =   "Afraid"
         Height          =   270
         Left            =   -73920
         TabIndex        =   69
         Top             =   7440
         Width           =   1095
      End
      Begin VB.TextBox txtJoinSay 
         Height          =   285
         Left            =   -73560
         TabIndex        =   59
         Top             =   840
         Width           =   3975
      End
      Begin VB.TextBox txtShopName 
         Height          =   285
         Left            =   -73560
         TabIndex        =   58
         Top             =   360
         Width           =   3975
      End
      Begin VB.TextBox txtLeaveSay 
         Height          =   285
         Left            =   -73560
         TabIndex        =   57
         Top             =   1320
         Width           =   3975
      End
      Begin VB.ListBox lstTradeItem 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
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
         Height          =   315
         Left            =   -73560
         Style           =   2  'Dropdown List
         TabIndex        =   55
         Top             =   2280
         Width           =   3975
      End
      Begin VB.TextBox txtItemGiveValue 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -73560
         TabIndex        =   54
         Text            =   "1"
         Top             =   2760
         Width           =   1335
      End
      Begin VB.ComboBox cmbItemGet 
         Height          =   315
         Left            =   -73560
         Style           =   2  'Dropdown List
         TabIndex        =   53
         Top             =   3240
         Width           =   3975
      End
      Begin VB.TextBox txtItemGetValue 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -73560
         TabIndex        =   52
         Text            =   "1"
         Top             =   3720
         Width           =   1335
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update "
         Height          =   255
         Left            =   -72000
         TabIndex        =   51
         Top             =   3720
         Width           =   2415
      End
      Begin VB.CheckBox chkFixesItems 
         Caption         =   "Fixes Items"
         Height          =   375
         Left            =   -74880
         TabIndex        =   50
         Top             =   1800
         Width           =   5295
      End
      Begin VB.TextBox txtStock 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -73560
         TabIndex        =   49
         Text            =   "1"
         Top             =   4080
         Width           =   1335
      End
      Begin VB.ComboBox cmdRestock 
         Height          =   315
         ItemData        =   "frmEditor.frx":013F
         Left            =   -71160
         List            =   "frmEditor.frx":014C
         Style           =   2  'Dropdown List
         TabIndex        =   48
         Top             =   4080
         Width           =   1575
      End
      Begin VB.TextBox txtSpellName 
         Height          =   285
         Left            =   -74040
         TabIndex        =   44
         Top             =   480
         Width           =   3975
      End
      Begin VB.ComboBox cmbSpellType 
         Height          =   315
         ItemData        =   "frmEditor.frx":0163
         Left            =   -74880
         List            =   "frmEditor.frx":017C
         Style           =   2  'Dropdown List
         TabIndex        =   39
         Top             =   2040
         Width           =   4815
      End
      Begin VB.ComboBox cmbSpellClassReq 
         Height          =   315
         ItemData        =   "frmEditor.frx":01BB
         Left            =   -74880
         List            =   "frmEditor.frx":01BD
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   960
         Width           =   4815
      End
      Begin VB.HScrollBar scrlSpellLevelReq 
         Height          =   255
         Left            =   -74040
         Max             =   255
         Min             =   1
         TabIndex        =   30
         Top             =   1560
         Value           =   1
         Width           =   3495
      End
      Begin VB.ComboBox cmbItemType 
         Height          =   315
         ItemData        =   "frmEditor.frx":01BF
         Left            =   240
         List            =   "frmEditor.frx":01ED
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   2400
         Width           =   4815
      End
      Begin VB.TextBox txtItemName 
         Height          =   285
         Left            =   1080
         TabIndex        =   24
         Top             =   720
         Width           =   3975
      End
      Begin VB.HScrollBar scrlItemPic 
         Height          =   255
         Left            =   1080
         Max             =   255
         TabIndex        =   19
         Top             =   1080
         Width           =   2895
      End
      Begin VB.PictureBox picItemPic 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   4560
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   18
         Top             =   1200
         Width           =   480
      End
      Begin VB.TextBox txtItemDescription 
         Height          =   285
         Left            =   1080
         TabIndex        =   11
         Top             =   1800
         Width           =   3975
      End
      Begin VB.Frame fraItemVitals 
         Caption         =   "Vitals Data"
         Height          =   855
         Left            =   240
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
      Begin VB.Frame fraGiveItem 
         Caption         =   "Give Item"
         Height          =   1335
         Left            =   -74880
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
            Height          =   255
            Left            =   120
            TabIndex        =   38
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label lblSpellItemNum 
            Alignment       =   1  'Right Justify
            Caption         =   "1"
            Height          =   255
            Left            =   4200
            TabIndex        =   37
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label9 
            Caption         =   "Value"
            Height          =   255
            Left            =   120
            TabIndex        =   36
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label lblSpellItemValue 
            Alignment       =   1  'Right Justify
            Caption         =   "0"
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
         Left            =   -74880
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
      Begin VB.Label Label56 
         Caption         =   "Sound when used (.WAV):"
         Height          =   255
         Left            =   240
         TabIndex        =   152
         Top             =   4560
         Width           =   1935
      End
      Begin VB.Label Label55 
         Caption         =   "EXP:"
         Height          =   255
         Left            =   -74640
         TabIndex        =   151
         Top             =   4200
         Width           =   855
      End
      Begin VB.Label Label54 
         Caption         =   "HP:"
         Height          =   255
         Left            =   -74640
         TabIndex        =   150
         Top             =   3840
         Width           =   855
      End
      Begin VB.Label Label53 
         Caption         =   "Enter the starting location for the class below:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   149
         Top             =   2400
         Width           =   3255
      End
      Begin VB.Label Label52 
         Alignment       =   1  'Right Justify
         Caption         =   "Y:"
         Height          =   255
         Left            =   -71520
         TabIndex        =   148
         Top             =   2760
         Width           =   495
      End
      Begin VB.Label Label51 
         Alignment       =   1  'Right Justify
         Caption         =   "X:"
         Height          =   255
         Left            =   -72960
         TabIndex        =   146
         Top             =   2760
         Width           =   495
      End
      Begin VB.Label Label50 
         Alignment       =   1  'Right Justify
         Caption         =   "Map:"
         Height          =   255
         Left            =   -74280
         TabIndex        =   144
         Top             =   2760
         Width           =   375
      End
      Begin VB.Label lblClassSprite 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   255
         Left            =   -70920
         TabIndex        =   142
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label49 
         Caption         =   "(This is currently a 1 out of [NUMBER] ratio so enter a second number)"
         Height          =   615
         Left            =   -74880
         TabIndex        =   141
         Top             =   4920
         Width           =   1695
      End
      Begin VB.Label Label48 
         Alignment       =   1  'Right Justify
         Caption         =   "Speed:"
         Height          =   255
         Left            =   -71040
         TabIndex        =   140
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label Label47 
         Alignment       =   1  'Right Justify
         Caption         =   "Magic:"
         Height          =   255
         Left            =   -71040
         TabIndex        =   139
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label Label46 
         Alignment       =   1  'Right Justify
         Caption         =   "Defense:"
         Height          =   255
         Left            =   -72720
         TabIndex        =   138
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label45 
         Alignment       =   1  'Right Justify
         Caption         =   "Strength:"
         Height          =   255
         Left            =   -72720
         TabIndex        =   137
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label44 
         Alignment       =   1  'Right Justify
         Caption         =   "SP:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   133
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label43 
         Alignment       =   1  'Right Justify
         Caption         =   "MP:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   132
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label42 
         Alignment       =   1  'Right Justify
         Caption         =   "HP:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   131
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label41 
         Alignment       =   1  'Right Justify
         Caption         =   "Class Sprite:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   125
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label40 
         Alignment       =   1  'Right Justify
         Caption         =   "Class Name:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   123
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lblSprite 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   255
         Left            =   -70920
         TabIndex        =   109
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label Label39 
         Caption         =   "Sprite"
         Height          =   255
         Left            =   -74640
         TabIndex        =   108
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label38 
         Caption         =   "Name"
         Height          =   255
         Left            =   -74760
         TabIndex        =   107
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label37 
         Caption         =   "Behavior"
         Height          =   375
         Left            =   -74640
         TabIndex        =   106
         Top             =   3360
         Width           =   735
      End
      Begin VB.Label lblRange 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   255
         Left            =   -70920
         TabIndex        =   105
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label Label36 
         Caption         =   "Range"
         Height          =   255
         Left            =   -74640
         TabIndex        =   104
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label lblSTR 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   255
         Left            =   -70920
         TabIndex        =   103
         Top             =   1920
         Width           =   495
      End
      Begin VB.Label Label35 
         Caption         =   "STR"
         Height          =   255
         Left            =   -74640
         TabIndex        =   102
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label lblDEF 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   255
         Left            =   -70920
         TabIndex        =   101
         Top             =   2280
         Width           =   495
      End
      Begin VB.Label Label34 
         Caption         =   "DEF"
         Height          =   255
         Left            =   -74640
         TabIndex        =   100
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label lblSPEED 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   255
         Left            =   -70920
         TabIndex        =   99
         Top             =   2640
         Width           =   495
      End
      Begin VB.Label Label33 
         Caption         =   "SPD"
         Height          =   255
         Left            =   -74640
         TabIndex        =   98
         Top             =   2640
         Width           =   855
      End
      Begin VB.Label lblMAGI 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   255
         Left            =   -70920
         TabIndex        =   97
         Top             =   3000
         Width           =   495
      End
      Begin VB.Label Label32 
         Caption         =   "MAGI"
         Height          =   255
         Left            =   -74640
         TabIndex        =   96
         Top             =   3000
         Width           =   855
      End
      Begin VB.Label Label31 
         Caption         =   "Drop Item Chance"
         Height          =   255
         Left            =   -74760
         TabIndex        =   95
         Top             =   4560
         Width           =   1335
      End
      Begin VB.Label lblNum 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   255
         Left            =   -70920
         TabIndex        =   94
         Top             =   6600
         Width           =   495
      End
      Begin VB.Label Label30 
         Caption         =   "Num"
         Height          =   255
         Left            =   -74880
         TabIndex        =   93
         Top             =   6600
         Width           =   855
      End
      Begin VB.Label Label29 
         Caption         =   "Item"
         Height          =   255
         Left            =   -74880
         TabIndex        =   92
         Top             =   6120
         Width           =   495
      End
      Begin VB.Label lblItemName 
         Height          =   375
         Left            =   -74280
         TabIndex        =   91
         Top             =   6120
         Width           =   3975
      End
      Begin VB.Label Label28 
         Caption         =   "Value"
         Height          =   255
         Left            =   -74880
         TabIndex        =   90
         Top             =   6960
         Width           =   735
      End
      Begin VB.Label lblValue 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   255
         Left            =   -70920
         TabIndex        =   89
         Top             =   6960
         Width           =   495
      End
      Begin VB.Label Label27 
         Caption         =   "Start HP"
         Height          =   375
         Left            =   -74640
         TabIndex        =   88
         Top             =   8160
         Width           =   975
      End
      Begin VB.Label Label26 
         Caption         =   "Exp Given"
         Height          =   375
         Left            =   -72240
         TabIndex        =   87
         Top             =   8160
         Width           =   1215
      End
      Begin VB.Label Label25 
         Caption         =   "Say"
         Height          =   255
         Left            =   -74760
         TabIndex        =   86
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label24 
         Caption         =   "Spawn Rate (in seconds)"
         Height          =   255
         Left            =   -74880
         TabIndex        =   85
         Top             =   5640
         Width           =   1815
      End
      Begin VB.Label Label23 
         Caption         =   "NPC Extras:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   84
         Top             =   7440
         Width           =   1335
      End
      Begin VB.Label Label22 
         Caption         =   "Join Say"
         Height          =   255
         Left            =   -74880
         TabIndex        =   68
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label21 
         Caption         =   "Name"
         Height          =   255
         Left            =   -74880
         TabIndex        =   67
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label20 
         Caption         =   "Leave Say"
         Height          =   255
         Left            =   -74880
         TabIndex        =   66
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label19 
         Caption         =   "Item Give"
         Height          =   255
         Left            =   -74880
         TabIndex        =   65
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label18 
         Caption         =   "Value"
         Height          =   255
         Left            =   -74880
         TabIndex        =   64
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Label Label17 
         Caption         =   "Item Get"
         Height          =   255
         Left            =   -74880
         TabIndex        =   63
         Top             =   3240
         Width           =   1215
      End
      Begin VB.Label Label16 
         Caption         =   "Value"
         Height          =   255
         Left            =   -74880
         TabIndex        =   62
         Top             =   3720
         Width           =   1215
      End
      Begin VB.Label Label15 
         Caption         =   "Item Stock"
         Height          =   255
         Left            =   -74880
         TabIndex        =   61
         Top             =   4080
         Width           =   1215
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Caption         =   "Restock Time:"
         Height          =   615
         Left            =   -72120
         TabIndex        =   60
         Top             =   4080
         Width           =   855
      End
      Begin VB.Label Label14 
         Caption         =   "Name"
         Height          =   375
         Left            =   -74880
         TabIndex        =   47
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label13 
         Caption         =   "Level"
         Height          =   255
         Left            =   -74880
         TabIndex        =   46
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label lblSpellLevelReq 
         Alignment       =   1  'Right Justify
         Caption         =   "1"
         Height          =   255
         Left            =   -70560
         TabIndex        =   45
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Name:"
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Pic:"
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lblItemPic 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   375
         Left            =   3960
         TabIndex        =   27
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Description:"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   1800
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
    lblItemSpellName.Caption = Trim$(Spell(scrlSpell.Value).Name)
    lblItemSpell.Caption = STR(scrlSpell.Value)
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

Private Sub scrlSpellItemNum_Change()
    fraSpellGiveItem.Caption = "Give Item " & Trim$(Item(scrlSpellItemNum.Value).Name)
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
