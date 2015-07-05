VERSION 5.00
Begin VB.Form frmEditor_Item 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item Editor"
   ClientHeight    =   8280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9720
   BeginProperty Font 
      Name            =   "Microsoft Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEditor_Item.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8280
   ScaleWidth      =   9720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton cmdChangeDataSize 
      Caption         =   "Change Data Size"
      Height          =   375
      Left            =   120
      TabIndex        =   136
      Top             =   7800
      Width           =   3135
   End
   Begin VB.Frame Frame2 
      Caption         =   "Properties"
      Height          =   1695
      Left            =   3360
      TabIndex        =   57
      Top             =   0
      Width           =   6255
      Begin VB.PictureBox picSpriteBorder 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   540
         Left            =   2280
         ScaleHeight     =   34
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   34
         TabIndex        =   137
         TabStop         =   0   'False
         Top             =   240
         Width           =   540
         Begin VB.PictureBox Picture 
            BackColor       =   &H00404040&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   15
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   138
            TabStop         =   0   'False
            Top             =   15
            Width           =   480
            Begin VB.PictureBox picItem 
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
               Height          =   480
               Left            =   0
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   139
               TabStop         =   0   'False
               Top             =   0
               Width           =   480
            End
         End
      End
      Begin VB.CheckBox chkStackable 
         Caption         =   "Stackable"
         Height          =   195
         Left            =   5137
         TabIndex        =   112
         Top             =   240
         Width           =   1095
      End
      Begin VB.ComboBox cmbSound 
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
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   900
         Width           =   1455
      End
      Begin VB.ComboBox cmbType 
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
         ItemData        =   "frmEditor_Item.frx":038A
         Left            =   720
         List            =   "frmEditor_Item.frx":03AC
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   540
         Width           =   1455
      End
      Begin VB.TextBox txtPrice 
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
         Left            =   4080
         TabIndex        =   5
         Top             =   240
         Width           =   1035
      End
      Begin VB.HScrollBar scrlRarity 
         Height          =   255
         Left            =   960
         Max             =   6
         TabIndex        =   7
         Top             =   1320
         Width           =   1275
      End
      Begin VB.ComboBox cmbBind 
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
         ItemData        =   "frmEditor_Item.frx":0414
         Left            =   4080
         List            =   "frmEditor_Item.frx":0421
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   540
         Width           =   2055
      End
      Begin VB.HScrollBar scrlAnim 
         Height          =   255
         Left            =   4080
         Max             =   5
         TabIndex        =   8
         Top             =   1320
         Width           =   2055
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
         ScrollBars      =   1  'Horizontal
         TabIndex        =   1
         Top             =   240
         Width           =   1455
      End
      Begin VB.HScrollBar scrlPic 
         Height          =   255
         Left            =   2400
         Max             =   255
         TabIndex        =   2
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "Type:"
         Height          =   255
         Left            =   120
         TabIndex        =   98
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "Sound:"
         Height          =   255
         Left            =   120
         TabIndex        =   97
         Top             =   960
         Width           =   615
      End
      Begin VB.Label lblRarity 
         AutoSize        =   -1  'True
         Caption         =   "Rarity: 1"
         Height          =   195
         Left            =   120
         TabIndex        =   63
         Top             =   1320
         Width           =   585
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Bind Type:"
         Height          =   195
         Left            =   2940
         TabIndex        =   62
         Top             =   600
         Width           =   765
      End
      Begin VB.Label lblPrice 
         AutoSize        =   -1  'True
         Caption         =   "Price:"
         Height          =   195
         Left            =   2940
         TabIndex        =   61
         Top             =   240
         Width           =   405
      End
      Begin VB.Label lblAnim 
         AutoSize        =   -1  'True
         Caption         =   "Animation: None"
         Height          =   255
         Left            =   4140
         TabIndex        =   60
         Top             =   1080
         Width           =   1950
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Name:"
         Height          =   195
         Left            =   120
         TabIndex        =   59
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   465
      End
      Begin VB.Label lblPic 
         AutoSize        =   -1  'True
         Caption         =   "Icon: 0"
         Height          =   195
         Left            =   2400
         TabIndex        =   58
         Top             =   1080
         UseMnemonic     =   0   'False
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   4080
      TabIndex        =   45
      Top             =   7800
      Width           =   1455
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   7440
      TabIndex        =   47
      Top             =   7800
      Width           =   1455
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   5760
      TabIndex        =   46
      Top             =   7800
      Width           =   1455
   End
   Begin VB.Frame Frame3 
      Caption         =   "Item List"
      Height          =   7695
      Left            =   120
      TabIndex        =   48
      Top             =   0
      Width           =   3135
      Begin VB.CommandButton cmdPaste 
         Caption         =   "Paste"
         Height          =   315
         Left            =   2400
         TabIndex        =   140
         Top             =   240
         Width           =   615
      End
      Begin VB.ListBox lstIndex 
         Height          =   6885
         Left            =   120
         TabIndex        =   111
         Top             =   660
         Width           =   2895
      End
      Begin VB.CommandButton cmdCopy 
         Caption         =   "Copy"
         Height          =   315
         Left            =   1680
         TabIndex        =   110
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox txtSearch 
         CausesValidation=   0   'False
         Height          =   270
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Requirements"
      Height          =   1935
      Left            =   3360
      TabIndex        =   49
      Top             =   1680
      Width           =   6255
      Begin VB.ComboBox cmbSkillReq 
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
         ItemData        =   "frmEditor_Item.frx":044A
         Left            =   480
         List            =   "frmEditor_Item.frx":0457
         Style           =   2  'Dropdown List
         TabIndex        =   115
         Top             =   360
         Width           =   2415
      End
      Begin VB.HScrollBar scrlLevelReq 
         Height          =   255
         LargeChange     =   10
         Left            =   1080
         Max             =   255
         TabIndex        =   108
         Top             =   1080
         Width           =   975
      End
      Begin VB.ComboBox cmbProficiencyReq 
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
         ItemData        =   "frmEditor_Item.frx":046F
         Left            =   3840
         List            =   "frmEditor_Item.frx":0471
         Style           =   2  'Dropdown List
         TabIndex        =   106
         Top             =   360
         Width           =   2295
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
         ItemData        =   "frmEditor_Item.frx":0473
         Left            =   2760
         List            =   "frmEditor_Item.frx":0480
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   720
         Width           =   1095
      End
      Begin VB.HScrollBar scrlAccessReq 
         Height          =   255
         Left            =   1080
         Max             =   5
         TabIndex        =   10
         Top             =   720
         Width           =   975
      End
      Begin VB.ComboBox cmbClassReq 
         Height          =   315
         Left            =   4440
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   720
         Width           =   1695
      End
      Begin VB.HScrollBar scrlStatReq 
         Height          =   255
         Index           =   1
         LargeChange     =   10
         Left            =   3120
         TabIndex        =   12
         Top             =   1080
         Width           =   925
      End
      Begin VB.HScrollBar scrlStatReq 
         Height          =   255
         Index           =   2
         LargeChange     =   10
         Left            =   5160
         TabIndex        =   13
         Top             =   1080
         Width           =   925
      End
      Begin VB.HScrollBar scrlStatReq 
         Height          =   255
         Index           =   3
         LargeChange     =   10
         Left            =   1080
         TabIndex        =   14
         Top             =   1440
         Width           =   925
      End
      Begin VB.HScrollBar scrlStatReq 
         Height          =   255
         Index           =   4
         LargeChange     =   10
         Left            =   3120
         TabIndex        =   15
         Top             =   1440
         Width           =   925
      End
      Begin VB.HScrollBar scrlStatReq 
         Height          =   255
         Index           =   5
         LargeChange     =   10
         Left            =   5160
         TabIndex        =   16
         Top             =   1440
         Width           =   925
      End
      Begin VB.Label lblSkillReq 
         AutoSize        =   -1  'True
         Caption         =   "Skill:"
         Height          =   195
         Left            =   120
         TabIndex        =   116
         Top             =   420
         Width           =   330
      End
      Begin VB.Label lblLevelReq 
         AutoSize        =   -1  'True
         Caption         =   "Level: 0"
         Height          =   195
         Left            =   120
         TabIndex        =   109
         Top             =   1080
         Width           =   570
      End
      Begin VB.Label lblProficiencyReq 
         AutoSize        =   -1  'True
         Caption         =   "Proficiency:"
         Height          =   195
         Left            =   3000
         TabIndex        =   107
         Top             =   420
         Width           =   825
      End
      Begin VB.Label lblGenderReq 
         AutoSize        =   -1  'True
         Caption         =   "Gender:"
         Height          =   195
         Left            =   2160
         TabIndex        =   100
         Top             =   780
         Width           =   570
      End
      Begin VB.Label lblAccessReq 
         AutoSize        =   -1  'True
         Caption         =   "Access: 0"
         Height          =   195
         Left            =   120
         TabIndex        =   56
         Top             =   720
         Width           =   705
      End
      Begin VB.Label lblClassReq 
         AutoSize        =   -1  'True
         Caption         =   "Class:"
         Height          =   195
         Left            =   3960
         TabIndex        =   55
         Top             =   780
         Width           =   420
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         Caption         =   "Str: 0"
         Height          =   195
         Index           =   1
         Left            =   2160
         TabIndex        =   54
         Top             =   1080
         UseMnemonic     =   0   'False
         Width           =   375
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         Caption         =   "End: 0"
         Height          =   195
         Index           =   2
         Left            =   4200
         TabIndex        =   53
         Top             =   1080
         UseMnemonic     =   0   'False
         Width           =   465
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         Caption         =   "Int: 0"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   52
         Top             =   1440
         UseMnemonic     =   0   'False
         Width           =   360
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         Caption         =   "Agi: 0"
         Height          =   195
         Index           =   4
         Left            =   2160
         TabIndex        =   51
         Top             =   1440
         UseMnemonic     =   0   'False
         Width           =   405
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         Caption         =   "Spi: 0"
         Height          =   195
         Index           =   5
         Left            =   4200
         TabIndex        =   50
         Top             =   1440
         UseMnemonic     =   0   'False
         Width           =   405
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Description"
      Height          =   1095
      Left            =   3360
      TabIndex        =   99
      Top             =   3600
      Width           =   6255
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
         Height          =   735
         Left            =   120
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   17
         Top             =   240
         Width           =   6015
      End
   End
   Begin VB.Frame fraTeleport 
      Caption         =   "Teleport"
      Height          =   2415
      Left            =   3360
      TabIndex        =   82
      Top             =   4680
      Width           =   2895
      Begin VB.CheckBox chkReusable 
         Caption         =   "Reusable"
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
         Index           =   0
         Left            =   1680
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   2040
         Width           =   1140
      End
      Begin VB.HScrollBar scrlY 
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   1680
         Width           =   2535
      End
      Begin VB.HScrollBar scrlX 
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   1080
         Width           =   2535
      End
      Begin VB.HScrollBar scrlMap 
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label lblY 
         Caption         =   "Y: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   85
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label lblX 
         Caption         =   "X: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   84
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label lblMap 
         Caption         =   "Map: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   83
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame fraAutoLife 
      Caption         =   "Auto Life Data"
      Height          =   1815
      Left            =   3360
      TabIndex        =   89
      Top             =   4680
      Visible         =   0   'False
      Width           =   3375
      Begin VB.CheckBox chkWarpAway 
         Caption         =   "Warp Away"
         Height          =   255
         Left            =   240
         TabIndex        =   101
         TabStop         =   0   'False
         Top             =   1440
         Width           =   1140
      End
      Begin VB.CheckBox chkReusable 
         Caption         =   "Reusable"
         Height          =   255
         Index           =   2
         Left            =   2160
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   1440
         Width           =   1025
      End
      Begin VB.HScrollBar scrlHP 
         Height          =   255
         Left            =   240
         Min             =   1
         TabIndex        =   18
         Top             =   480
         Value           =   1
         Width           =   2895
      End
      Begin VB.HScrollBar scrlMP 
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   1080
         Width           =   2895
      End
      Begin VB.Label lblHP 
         AutoSize        =   -1  'True
         Caption         =   "HP: 1"
         Height          =   195
         Left            =   240
         TabIndex        =   91
         Top             =   240
         Width           =   405
      End
      Begin VB.Label lblMP 
         AutoSize        =   -1  'True
         Caption         =   "MP: 0"
         Height          =   195
         Left            =   240
         TabIndex        =   90
         Top             =   840
         Width           =   420
      End
   End
   Begin VB.Frame fraTitle 
      Caption         =   "Title Data"
      Height          =   1095
      Left            =   3360
      TabIndex        =   117
      Top             =   4680
      Visible         =   0   'False
      Width           =   3735
      Begin VB.HScrollBar scrlTitle 
         Height          =   255
         Left            =   1680
         Max             =   255
         Min             =   1
         TabIndex        =   119
         Top             =   360
         Value           =   1
         Width           =   1815
      End
      Begin VB.CheckBox chkReusable 
         Caption         =   "Reusable"
         Height          =   255
         Index           =   4
         Left            =   2520
         TabIndex        =   118
         TabStop         =   0   'False
         Top             =   720
         Width           =   1025
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "Title: 1"
         Height          =   195
         Left            =   240
         TabIndex        =   120
         Top             =   360
         Width           =   480
      End
   End
   Begin VB.Frame fraRecipe 
      Caption         =   "Recipe Data"
      Height          =   3015
      Left            =   3360
      TabIndex        =   121
      Top             =   4680
      Visible         =   0   'False
      Width           =   6255
      Begin VB.HScrollBar scrlItem1 
         Height          =   255
         Left            =   120
         Max             =   5
         TabIndex        =   128
         Top             =   1080
         Width           =   2535
      End
      Begin VB.HScrollBar scrlItem2 
         Height          =   255
         Left            =   2880
         Max             =   5
         TabIndex        =   127
         Top             =   1080
         Width           =   2535
      End
      Begin VB.HScrollBar scrlResult 
         Height          =   255
         Left            =   120
         Max             =   5
         TabIndex        =   126
         Top             =   1800
         Width           =   2535
      End
      Begin VB.HScrollBar scrlSkill 
         Height          =   255
         Left            =   2880
         Max             =   5
         TabIndex        =   125
         Top             =   1800
         Width           =   2535
      End
      Begin VB.HScrollBar ScrlSkillExp 
         Height          =   255
         Left            =   120
         Max             =   5
         TabIndex        =   124
         Top             =   2640
         Width           =   2535
      End
      Begin VB.HScrollBar ScrlToolRequired 
         Height          =   255
         Left            =   3000
         Max             =   5
         TabIndex        =   123
         Top             =   360
         Width           =   2535
      End
      Begin VB.HScrollBar ScrlSkillLevelReq 
         Height          =   255
         Left            =   2880
         Max             =   5
         TabIndex        =   122
         Top             =   2640
         Width           =   2535
      End
      Begin VB.Label lblToolRequired 
         AutoSize        =   -1  'True
         Caption         =   "Tool Required: None"
         Height          =   180
         Left            =   120
         TabIndex        =   135
         Top             =   360
         Width           =   1530
      End
      Begin VB.Label lblItem1 
         Caption         =   "Item 1: None"
         Height          =   255
         Left            =   120
         TabIndex        =   134
         Top             =   720
         Width           =   2535
      End
      Begin VB.Label lblItem2 
         Caption         =   "Item 2: None"
         Height          =   255
         Left            =   2880
         TabIndex        =   133
         Top             =   720
         Width           =   3135
      End
      Begin VB.Label lblResult 
         Caption         =   "Result: None"
         Height          =   255
         Left            =   120
         TabIndex        =   132
         Top             =   1440
         Width           =   2535
      End
      Begin VB.Label lblSkill 
         Caption         =   "Skill: None"
         Height          =   255
         Left            =   2880
         TabIndex        =   131
         Top             =   1440
         Width           =   2535
      End
      Begin VB.Label lblSkillExp 
         Caption         =   "Skill Exp: None"
         Height          =   255
         Left            =   120
         TabIndex        =   130
         Top             =   2280
         Width           =   2535
      End
      Begin VB.Label lblSkillLevelReq 
         Caption         =   "Skill Level Req: None"
         Height          =   255
         Left            =   2880
         TabIndex        =   129
         Top             =   2280
         Width           =   2535
      End
   End
   Begin VB.Frame fraEquipment 
      Caption         =   "Equipment Data"
      Height          =   3015
      Left            =   3360
      TabIndex        =   64
      Top             =   4680
      Visible         =   0   'False
      Width           =   6255
      Begin VB.CheckBox chkIndestructable 
         Caption         =   "Indestructable"
         Height          =   255
         Left            =   4800
         TabIndex        =   114
         Top             =   2400
         Width           =   1335
      End
      Begin VB.CheckBox chkTwoHanded 
         Caption         =   "Two Handed"
         Height          =   255
         Left            =   4800
         TabIndex        =   113
         Top             =   2640
         Width           =   1335
      End
      Begin VB.ComboBox cmbEquipSlot 
         Height          =   315
         ItemData        =   "frmEditor_Item.frx":0498
         Left            =   1680
         List            =   "frmEditor_Item.frx":04B7
         Style           =   2  'Dropdown List
         TabIndex        =   103
         Top             =   300
         Width           =   1935
      End
      Begin VB.HScrollBar scrlDurability 
         Height          =   255
         Left            =   4560
         TabIndex        =   38
         Top             =   1080
         Width           =   1575
      End
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   540
         Left            =   1680
         ScaleHeight     =   34
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   34
         TabIndex        =   86
         TabStop         =   0   'False
         Top             =   2280
         Width           =   540
         Begin VB.PictureBox Picture6 
            BackColor       =   &H00404040&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   15
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   87
            TabStop         =   0   'False
            Top             =   15
            Width           =   480
            Begin VB.PictureBox picPaperdoll 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   480
               Left            =   0
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   128
               TabIndex        =   88
               TabStop         =   0   'False
               Top             =   0
               Width           =   1920
            End
         End
      End
      Begin VB.HScrollBar scrlPaperdoll 
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   2520
         Width           =   1455
      End
      Begin VB.HScrollBar scrlSpeed 
         Height          =   255
         LargeChange     =   100
         Left            =   1680
         Max             =   10000
         Min             =   100
         SmallChange     =   100
         TabIndex        =   37
         Top             =   1080
         Value           =   1000
         Width           =   1335
      End
      Begin VB.HScrollBar scrlStatBonus 
         Height          =   255
         Index           =   5
         LargeChange     =   10
         Left            =   3240
         TabIndex        =   43
         Top             =   1800
         Width           =   900
      End
      Begin VB.HScrollBar scrlStatBonus 
         Height          =   255
         Index           =   4
         LargeChange     =   10
         Left            =   1200
         TabIndex        =   42
         Top             =   1800
         Width           =   900
      End
      Begin VB.HScrollBar scrlStatBonus 
         Height          =   255
         Index           =   3
         LargeChange     =   10
         Left            =   5225
         TabIndex        =   41
         Top             =   1440
         Width           =   900
      End
      Begin VB.HScrollBar scrlStatBonus 
         Height          =   255
         Index           =   2
         LargeChange     =   10
         Left            =   3240
         TabIndex        =   40
         Top             =   1440
         Width           =   900
      End
      Begin VB.HScrollBar scrlStatBonus 
         Height          =   255
         Index           =   1
         LargeChange     =   10
         Left            =   1200
         TabIndex        =   39
         Top             =   1440
         Width           =   900
      End
      Begin VB.ComboBox cmbTool 
         Height          =   315
         ItemData        =   "frmEditor_Item.frx":04F6
         Left            =   4200
         List            =   "frmEditor_Item.frx":0506
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   300
         Width           =   1935
      End
      Begin VB.HScrollBar scrlChanceModifier 
         Height          =   255
         LargeChange     =   10
         Left            =   1680
         Max             =   255
         Min             =   1
         TabIndex        =   92
         Top             =   720
         Value           =   1
         Width           =   4455
      End
      Begin VB.HScrollBar scrlDamage 
         Height          =   255
         LargeChange     =   10
         Left            =   1680
         TabIndex        =   36
         Top             =   720
         Width           =   4455
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Slot:"
         Height          =   195
         Left            =   120
         TabIndex        =   102
         Top             =   360
         Width           =   315
      End
      Begin VB.Label lblDurability 
         AutoSize        =   -1  'True
         Caption         =   "Durability: 0"
         Height          =   195
         Left            =   3240
         TabIndex        =   96
         ToolTipText     =   "In seconds."
         Top             =   1080
         UseMnemonic     =   0   'False
         Width           =   825
      End
      Begin VB.Label lblPaperdoll 
         AutoSize        =   -1  'True
         Caption         =   "Paperdoll: 0"
         Height          =   195
         Left            =   120
         TabIndex        =   77
         Top             =   2280
         Width           =   840
      End
      Begin VB.Label lblSpeed 
         AutoSize        =   -1  'True
         Caption         =   "Speed: 1 s"
         Height          =   195
         Left            =   120
         TabIndex        =   72
         ToolTipText     =   "In seconds."
         Top             =   1080
         UseMnemonic     =   0   'False
         Width           =   765
      End
      Begin VB.Label lblStatBonus 
         AutoSize        =   -1  'True
         Caption         =   "+ Spi: 0"
         Height          =   195
         Index           =   5
         Left            =   2160
         TabIndex        =   71
         Top             =   1800
         UseMnemonic     =   0   'False
         Width           =   540
      End
      Begin VB.Label lblStatBonus 
         AutoSize        =   -1  'True
         Caption         =   "+ Agi: 0"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   70
         Top             =   1800
         UseMnemonic     =   0   'False
         Width           =   540
      End
      Begin VB.Label lblStatBonus 
         AutoSize        =   -1  'True
         Caption         =   "+ Int: 0"
         Height          =   195
         Index           =   3
         Left            =   4200
         TabIndex        =   69
         Top             =   1440
         UseMnemonic     =   0   'False
         Width           =   495
      End
      Begin VB.Label lblStatBonus 
         AutoSize        =   -1  'True
         Caption         =   "+ End: 0"
         Height          =   195
         Index           =   2
         Left            =   2160
         TabIndex        =   68
         Top             =   1440
         UseMnemonic     =   0   'False
         Width           =   600
      End
      Begin VB.Label lblStatBonus 
         AutoSize        =   -1  'True
         Caption         =   "+ Str: 0"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   65
         Top             =   1440
         UseMnemonic     =   0   'False
         Width           =   510
      End
      Begin VB.Label lblChance 
         AutoSize        =   -1  'True
         Caption         =   "Chance: 1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   120
         TabIndex        =   93
         Top             =   720
         UseMnemonic     =   0   'False
         Width           =   780
      End
      Begin VB.Label lblTool 
         AutoSize        =   -1  'True
         Caption         =   "Tool:"
         Height          =   195
         Left            =   3720
         TabIndex        =   66
         Top             =   360
         Width           =   360
      End
      Begin VB.Label lblDamage 
         AutoSize        =   -1  'True
         Caption         =   "Damage: 0"
         Height          =   195
         Left            =   120
         TabIndex        =   67
         Top             =   720
         UseMnemonic     =   0   'False
         Width           =   780
      End
   End
   Begin VB.Frame fraSprite 
      Caption         =   "Sprite Data"
      Height          =   1095
      Left            =   3360
      TabIndex        =   94
      Top             =   4680
      Visible         =   0   'False
      Width           =   3735
      Begin VB.CheckBox chkReusable 
         Caption         =   "Reusable"
         Height          =   255
         Index           =   3
         Left            =   2520
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   720
         Width           =   1025
      End
      Begin VB.HScrollBar scrlSprite 
         Height          =   255
         Left            =   1440
         Max             =   255
         Min             =   1
         TabIndex        =   21
         Top             =   360
         Value           =   1
         Width           =   2055
      End
      Begin VB.Label lblSprite 
         AutoSize        =   -1  'True
         Caption         =   "Sprite: 1"
         Height          =   195
         Left            =   240
         TabIndex        =   95
         Top             =   360
         Width           =   585
      End
   End
   Begin VB.Frame fraSpell 
      Caption         =   "Spell Data"
      Height          =   1215
      Left            =   3360
      TabIndex        =   74
      Top             =   4680
      Visible         =   0   'False
      Width           =   3735
      Begin VB.HScrollBar scrlSpell 
         Height          =   255
         Left            =   1440
         Max             =   255
         TabIndex        =   23
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label lblSpellName 
         AutoSize        =   -1  'True
         Caption         =   "Name: None"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   240
         TabIndex        =   76
         Top             =   360
         Width           =   930
      End
      Begin VB.Label lblSpell 
         AutoSize        =   -1  'True
         Caption         =   "Spell: 0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   240
         TabIndex        =   75
         Top             =   720
         Width           =   585
      End
   End
   Begin VB.Frame fraConsume 
      Caption         =   "Consumable Data"
      Height          =   2775
      Left            =   3360
      TabIndex        =   73
      Top             =   4680
      Visible         =   0   'False
      Width           =   4455
      Begin VB.HScrollBar scrlDuration 
         Height          =   255
         Left            =   3000
         Max             =   60
         TabIndex        =   28
         Top             =   1680
         Width           =   1335
      End
      Begin VB.CheckBox chkHoT 
         Caption         =   "Heal Over Time"
         Height          =   375
         Left            =   3000
         TabIndex        =   104
         TabStop         =   0   'False
         Top             =   960
         Width           =   1380
      End
      Begin VB.CheckBox chkReusable 
         Caption         =   "Reusable"
         Height          =   255
         Index           =   1
         Left            =   3000
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   720
         Width           =   1025
      End
      Begin VB.CheckBox chkInstaCast 
         Caption         =   "Instant Cast"
         Height          =   255
         Left            =   3000
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   480
         Width           =   1335
      End
      Begin VB.HScrollBar scrlCastSpell 
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   2280
         Width           =   2775
      End
      Begin VB.HScrollBar scrlAddEXP 
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   1680
         Width           =   2775
      End
      Begin VB.HScrollBar scrlAddMP 
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   1080
         Width           =   2775
      End
      Begin VB.HScrollBar scrlAddHP 
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   480
         Width           =   2775
      End
      Begin VB.Label lblDuration 
         Caption         =   "Duration: 0 s"
         Height          =   255
         Left            =   3000
         TabIndex        =   105
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label lblCastSpell 
         Caption         =   "Cast Spell: None"
         Height          =   255
         Left            =   120
         TabIndex        =   81
         Top             =   2040
         Width           =   2775
      End
      Begin VB.Label lblAddEXP 
         Caption         =   "Add Exp: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   80
         Top             =   1440
         Width           =   2775
      End
      Begin VB.Label lblAddMP 
         Caption         =   "Add MP: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   79
         Top             =   840
         Width           =   2775
      End
      Begin VB.Label lblAddHP 
         Caption         =   "Add HP: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   78
         Top             =   240
         Width           =   2775
      End
   End
End
Attribute VB_Name = "frmEditor_Item"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private lastIndex As Long
Private TmpIndex As Long

Private Sub chkHoT_Click()
    If EditorIndex < 1 Or EditorIndex > MAX_ITEMS Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Item(EditorIndex).HoT = chkHoT.Value
    
    If chkHoT.Value = 1 Then
        lblDuration.Enabled = True
        scrlDuration.Enabled = True
    Else
        lblDuration.Enabled = False
        scrlDuration.Enabled = False
    End If
    Exit Sub

' Error handler
ErrorHandler:
    HandleError "chkHoT_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub chkIndestructable_Click()
    If EditorIndex < 1 Or EditorIndex > MAX_ITEMS Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Item(EditorIndex).Indestructable = chkIndestructable.Value
    
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "chkIndestructable_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub chkInstaCast_Click()
    If EditorIndex < 1 Or EditorIndex > MAX_ITEMS Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Item(EditorIndex).InstaCast = chkInstaCast.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "chkInstaCast_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub chkReusable_Click(Index As Integer)
    If EditorIndex < 1 Or EditorIndex > MAX_ITEMS Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If frmEditor_Item.chkReusable(Index) = 1 Then
        Item(EditorIndex).IsReusable = True
    Else
        Item(EditorIndex).IsReusable = False
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "chkReusable_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub chkStackable_Click()
    If EditorIndex < 1 Or EditorIndex > MAX_ITEMS Then Exit Sub
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Item(EditorIndex).stackable = chkStackable.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "chkStackable_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub chkTwoHanded_Click()
    If EditorIndex < 1 Or EditorIndex > MAX_ITEMS Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Item(EditorIndex).TwoHanded = chkTwoHanded.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "chkTwoHanded_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub chkWarpAway_Click()
    If EditorIndex < 1 Or EditorIndex > MAX_ITEMS Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If fraAutoLife.Visible = False Then Exit Sub
    
    Item(EditorIndex).Data1 = chkWarpAway.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "chkWarpAway_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmbBind_Click()
    If EditorIndex < 1 Or EditorIndex > MAX_ITEMS Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Item(EditorIndex).BindType = cmbBind.ListIndex
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmbBind_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmbEquipSlot_Click()
    If EditorIndex < 1 Or EditorIndex > MAX_ITEMS Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Item(EditorIndex).EquipSlot = cmbEquipSlot.ListIndex
    
    With frmEditor_Item
        ' Specific options when selecting the weapon type
        lblDamage.Caption = "Damage: " & scrlDamage.Value
        
        If .cmbEquipSlot.ListIndex = Equipment.Weapon Then
            .cmbTool.Enabled = True
            .lblTool.Enabled = True
            .lblDamage.Enabled = True
            .scrlSpeed.Enabled = True
            .lblSpeed.Enabled = True
            .lblDamage.Caption = "Damage: " & scrlDamage.Value
        Else
            cmbTool.ListIndex = 0
            .scrlSpeed.Enabled = False
            .lblSpeed.Enabled = False
            .cmbTool.Enabled = False
            .lblTool.Enabled = False
            .lblDamage.Caption = "Defense: " & scrlDamage.Value
        End If
    End With
    
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmbEquipSlot_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmbProficiencyReq_Click()
    If EditorIndex < 1 Or EditorIndex > MAX_ITEMS Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Item(EditorIndex).ProficiencyReq = cmbProficiencyReq.ListIndex
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmbProficiencyReq_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmbClassReq_Click()
    If EditorIndex < 1 Or EditorIndex > MAX_ITEMS Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Item(EditorIndex).ClassReq = cmbClassReq.ListIndex
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmbClassReq_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmbGenderReq_Click()
    If EditorIndex < 1 Or EditorIndex > MAX_ITEMS Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Item(EditorIndex).GenderReq = cmbGenderReq.ListIndex
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmbGenderReq_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmbSound_Click()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If cmbSound.ListIndex >= 0 Then
        Audio.StopSounds
        Item(EditorIndex).Sound = cmbSound.List(cmbSound.ListIndex)
        Audio.PlaySound Item(EditorIndex).Sound, -1, -1, True
    Else
        Item(EditorIndex).Sound = vbNullString
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdSound_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmbTool_Click()
    If EditorIndex < 1 Or EditorIndex > MAX_ITEMS Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Item(EditorIndex).Tool = cmbTool.ListIndex
    
    ' Resources
    If frmEditor_Item.cmbTool.ListIndex = 0 Then
        scrlChanceModifier.Visible = False
        lblChance.Visible = False
        lblDamage.Visible = True
        scrlDamage.Visible = True
    Else
        scrlChanceModifier.Visible = True
        lblChance.Visible = True
        lblDamage.Visible = False
        scrlDamage.Visible = False
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmbTool_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdChangeDataSize_Click()
    Dim Res As VbMsgBoxResult, val As String
    Dim dataModified As Boolean, I As Long
    
    If EditorIndex < 1 Or EditorIndex > MAX_ITEMS Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    For I = 1 To MAX_ITEMS
        If Item_Changed(I) And I <> EditorIndex Then
        
            dataModified = True
            Exit For
        End If
    Next
    
    If dataModified Then
        Res = MsgBox("Do you want to continue and discard the changes you made to your data?", vbYesNo)
        
        If Res = vbNo Then Exit Sub
    End If
    
    val = InputBox("Enter the amount you want the new data size to be.", "Change Data Size", MAX_ITEMS)
    
    If Not IsNumeric(val) Then
        Exit Sub
    End If
    
    Res = Abs(val)
    
    If Res = MAX_ITEMS Then Exit Sub
    
    Call SendChangeDataSize(Res, EDITOR_ITEM)
    
    Unload frmEditor_Item
    MAX_ITEMS = Res
    ReDim Item(MAX_ITEMS)
    
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdChangeDataSize_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdDelete_Click()
    Dim TmpIndex As Long
    
    If EditorIndex < 1 Or EditorIndex > MAX_ITEMS Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ClearItem EditorIndex
    
    TmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Item(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = TmpIndex

    ItemEditorInit
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdDelete_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub Form_Load()
    Dim I As Long
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    frmMain.SubDaFocus Me.hWnd
    scrlPic.max = NumItems
    scrlAnim.max = MAX_ANIMATIONS
    scrlPaperdoll.max = NumPaperdolls
    frmEditor_Item.scrlLevelReq.max = MAX_LEVEL
    frmEditor_Item.scrlSprite.max = NumCharacters
    txtName.MaxLength = NAME_LENGTH
    txtSearch.MaxLength = NAME_LENGTH
    txtDesc.MaxLength = 256
    
    cmbProficiencyReq.Clear
    cmbProficiencyReq.AddItem "None"
    
    For I = 1 To Proficiency_Count - 1
        cmbProficiencyReq.AddItem GetProficiencyName(I)
    Next
    
    cmbSkillReq.Clear
    cmbSkillReq.AddItem "None"
    
    For I = 1 To Skill_Count - 1
        cmbSkillReq.AddItem GetSkillName(I)
    Next
    
    scrlItem1.max = MAX_ITEMS
    scrlItem2.max = MAX_ITEMS
    scrlResult.max = MAX_ITEMS
    ScrlSkillExp.max = 32767
    ScrlToolRequired.max = MAX_ITEMS
    ScrlSkillLevelReq.max = MAX_LEVEL
    scrlSkill.max = Skill_Count - 1
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "Form_Load", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdSave_Click()
    If EditorIndex < 1 Or EditorIndex > MAX_ITEMS Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    EditorSave = True
    Call ItemEditorSave
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdSave_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdClose_Click()
    If EditorIndex < 1 Or EditorIndex > MAX_ITEMS Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    frmAdmin.chkEditor(EDITOR_ITEM).FontBold = False
    frmAdmin.picEye(EDITOR_ITEM).Visible = False
    BringWindowToTop (frmAdmin.hWnd)
    Unload frmEditor_Item
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdClose_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmbType_Click()
    If EditorIndex < 1 Or EditorIndex > MAX_ITEMS Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If (cmbType.ListIndex = ITEM_TYPE_EQUIPMENT) Then
        fraEquipment.Visible = True
        Me.cmbEquipSlot.ListIndex = Item(EditorIndex).EquipSlot
        Me.chkStackable.Value = 0
        Me.chkStackable.Enabled = False
    Else
        chkStackable.Enabled = True
        fraEquipment.Visible = False
    End If

    If cmbType.ListIndex = ITEM_TYPE_CONSUME Then
        fraConsume.Visible = True
    Else
        fraConsume.Visible = False
    End If

    If (cmbType.ListIndex = ITEM_TYPE_SPELL) Then
        fraSpell.Visible = True
    Else
        fraSpell.Visible = False
    End If
    
    If (cmbType.ListIndex = ITEM_TYPE_TELEPORT) Then
        fraTeleport.Visible = True
    Else
        fraTeleport.Visible = False
    End If
    
    If (cmbType.ListIndex = ITEM_TYPE_AUTOLIFE) Then
        fraAutoLife.Visible = True
    Else
        fraAutoLife.Visible = False
    End If
    
    If (cmbType.ListIndex = ITEM_TYPE_SPRITE) Then
        fraSprite.Visible = True
    Else
        fraSprite.Visible = False
    End If
    
    If (cmbType.ListIndex = ITEM_TYPE_TITLE) Then
        fraTitle.Visible = True
    Else
        fraTitle.Visible = False
    End If
    
    If (frmEditor_Item.cmbType.ListIndex = ITEM_TYPE_RECIPE) Then
        fraRecipe.Visible = True
    Else
        fraRecipe.Visible = False
    End If
    
    With frmEditor_Item
        ' Specific options when selecting the weapon type
        lblDamage.Caption = "Damage: " & scrlDamage.Value
        
        If .cmbEquipSlot.ListIndex = Equipment.Weapon Then
            .cmbTool.Enabled = True
            .lblTool.Enabled = True
            .lblDamage.Enabled = True
            .scrlSpeed.Enabled = True
            .lblSpeed.Enabled = True
            .lblDamage.Caption = "Damage: " & scrlDamage.Value
        Else
            cmbTool.ListIndex = 0
            .scrlSpeed.Enabled = False
            .lblSpeed.Enabled = False
            .cmbTool.Enabled = False
            .lblTool.Enabled = False
            .lblDamage.Caption = "Defense: " & scrlDamage.Value
        End If
    End With
    
    Item(EditorIndex).Type = cmbType.ListIndex
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmbType_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    frmMain.UnsubDaFocus Me.hWnd
    If EditorSave = False Then
        ItemEditorCancel
    Else
        EditorSave = False
    End If
    frmAdmin.chkEditor(EDITOR_ITEM).Value = False
    BringWindowToTop (frmAdmin.hWnd)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "Form_Unload", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub lstIndex_Click()
    If EditorIndex < 1 Or EditorIndex > MAX_ITEMS Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ItemEditorInit
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "lstIndex_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlAccessReq_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_ITEMS Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lblAccessReq.Caption = "Access: " & scrlAccessReq.Value
    Item(EditorIndex).AccessReq = scrlAccessReq.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlAccessReq_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlAddEXP_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_ITEMS Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lblAddEXP.Caption = "Add Exp: " & scrlAddEXP.Value
    Item(EditorIndex).AddEXP = scrlAddEXP.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlAddEXP_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlAddHP_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_ITEMS Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lblAddHP.Caption = "Add HP: " & scrlAddHP.Value
    Item(EditorIndex).AddHP = scrlAddHP.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlAddHP_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlAddMP_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_ITEMS Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lblAddMP.Caption = "Add MP: " & scrlAddMP.Value
    Item(EditorIndex).AddMP = scrlAddMP.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlAddMP_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlAnim_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_ITEMS Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If scrlAnim.Value = 0 Then
        lblAnim.Caption = "Animation: None"
    Else
        lblAnim.Caption = "Animation: " & Trim$(Animation(scrlAnim.Value).Name)
    End If
    Item(EditorIndex).Animation = scrlAnim.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlAnim_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlCastSpell_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_ITEMS Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If scrlCastSpell.Value > 0 Then
        lblCastSpell.Caption = "Cast Spell: " & Trim$(Spell(scrlCastSpell.Value).Name)
    Else
        lblCastSpell.Caption = "Cast Spell: None"
    End If
    Item(EditorIndex).CastSpell = scrlCastSpell.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlCastSpell_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlDamage_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_ITEMS Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If fraEquipment.Visible = False Then Exit Sub
    
    If (cmbEquipSlot.ListIndex = Equipment.Weapon) Then
        lblDamage.Caption = "Damage: " & scrlDamage.Value
    Else
        lblDamage.Caption = "Defense: " & scrlDamage.Value
    End If
    Item(EditorIndex).Data2 = scrlDamage.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlDamage_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlDurability_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_ITEMS Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If fraEquipment.Visible = False Then Exit Sub
    
    Item(EditorIndex).Data1 = frmEditor_Item.scrlDurability.Value
    lblDurability.Caption = "Durability: " & Item(EditorIndex).Data1
    
    If Item(EditorIndex).Data1 > 0 Then
        chkStackable.Value = 0
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlDurability_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlDuration_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_ITEMS Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If fraConsume.Visible = False Then Exit Sub
    
    lblDuration.Caption = "Duration: " & scrlDuration.Value & " s"
    Item(EditorIndex).Data1 = scrlDuration.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlDuration_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlHP_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_ITEMS Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If fraAutoLife.Visible = False Then Exit Sub
    
    lblHP.Caption = "HP: " & scrlHP.Value
    Item(EditorIndex).AddHP = scrlHP.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlHP_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlMP_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_ITEMS Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lblMP.Caption = "MP: " & scrlMP.Value
    Item(EditorIndex).AddMP = scrlMP.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlMP_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlLevelReq_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_ITEMS Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lblLevelReq.Caption = "Level: " & scrlLevelReq
    Item(EditorIndex).LevelReq = scrlLevelReq.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlLevelReq_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlMap_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_ITEMS Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lblMap.Caption = "Map: " & scrlMap.Value
    Item(EditorIndex).Data1 = scrlMap.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlMap_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlChanceModifier_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_ITEMS Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lblChance.Caption = "Chance: " & scrlChanceModifier.Value
    Item(EditorIndex).ChanceModifier = scrlChanceModifier.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlMinChance_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlPaperdoll_Change()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If EditorIndex < 1 Or EditorIndex > MAX_ITEMS Then Exit Sub
    
    lblPaperdoll.Caption = "Paperdoll: " & scrlPaperdoll.Value
    Item(EditorIndex).Paperdoll = scrlPaperdoll.Value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlPaperdoll_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlPic_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_ITEMS Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lblPic.Caption = "Icon: " & scrlPic.Value
    Item(EditorIndex).Pic = scrlPic.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlPic_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlRarity_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_ITEMS Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lblRarity.Caption = "Rarity: " & scrlRarity.Value
    Item(EditorIndex).Rarity = scrlRarity.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlRarity_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlSpeed_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_ITEMS Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lblSpeed.Caption = "Speed: " & scrlSpeed.Value / 1000 & " s"
    Item(EditorIndex).WeaponSpeed = scrlSpeed.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlSpeed_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlSprite_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_ITEMS Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If fraSprite.Visible = False Then Exit Sub
    
    lblSprite.Caption = "Sprite: " & scrlSprite.Value
    Item(EditorIndex).Data1 = scrlSprite.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlSprite_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlStatBonus_Change(Index As Integer)
    Dim text As String
    
    If EditorIndex < 1 Or EditorIndex > MAX_ITEMS Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Select Case Index
        Case 1
            text = "+ Str: "
        Case 2
            text = "+ End: "
        Case 3
            text = "+ Int: "
        Case 4
            text = "+ Agi: "
        Case 5
            text = "+ Spi: "
    End Select
            
    lblStatBonus(Index).Caption = text & scrlStatBonus(Index).Value
    Item(EditorIndex).Add_Stat(Index) = scrlStatBonus(Index).Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlStatBonus_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlStatReq_Change(Index As Integer)
    Dim text As String
    
    If EditorIndex < 1 Or EditorIndex > MAX_ITEMS Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Select Case Index
        Case 1
            text = "Str: "
        Case 2
            text = "End: "
        Case 3
            text = "Int: "
        Case 4
            text = "Agi: "
        Case 5
            text = "Spi: "
    End Select
    
    lblStatReq(Index).Caption = text & scrlStatReq(Index).Value
    Item(EditorIndex).Stat_Req(Index) = scrlStatReq(Index).Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlStatReq_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlSpell_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_ITEMS Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If fraSpell.Visible = False Then Exit Sub
    
    Call UpdateSpellScrollBars
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlSpell_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlTitle_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_ITEMS Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If fraTitle.Visible = False Then Exit Sub
    
    If Not Trim$(title(scrlTitle.Value).Name) = vbNullString Then
        lblTitle.Caption = "Title: " & Trim$(title(scrlTitle.Value).Name)
    Else
        lblTitle.Caption = "Title: None"
    End If
    
    Item(EditorIndex).Data1 = scrlTitle.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlTitle_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub ScrlToolRequired_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_ITEMS Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If fraRecipe.Visible = False Then Exit Sub
    
    If ScrlToolRequired.Value > 0 Then
        lblToolRequired.Caption = "Tool Required: " & Trim$(Item(ScrlToolRequired.Value).Name)
    Else
        lblToolRequired.Caption = "Tool Required: None"
    End If
    
    Item(EditorIndex).ToolRequired = ScrlToolRequired.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "ScrlToolRequired_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlResult_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_ITEMS Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If fraRecipe.Visible = False Then Exit Sub
    
    If scrlResult.Value > 0 Then
        lblResult.Caption = "Result: " & Trim$(Item(scrlResult.Value).Name)
    Else
        lblResult.Caption = "Result: None"
    End If
    
    Item(EditorIndex).Data3 = scrlResult.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlResult_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlSkill_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_ITEMS Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If fraRecipe.Visible = False Then Exit Sub
    
    If scrlSkill.Value > 0 Then
        lblSkill.Caption = "Skill: " & GetSkillName(scrlSkill.Value)
    Else
        lblSkill.Caption = "Skill: None"
    End If
    
    Item(EditorIndex).Skill = scrlSkill.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlSkill_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub ScrlSkillExp_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_ITEMS Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If fraRecipe.Visible = False Then Exit Sub
    
    If ScrlSkillExp.Value > 0 Then
        lblSkillExp.Caption = "Skill Exp: " & ScrlSkillExp.Value
    Else
        lblSkillExp.Caption = "Skill Exp: None"
    End If
    
    Item(EditorIndex).SkillExp = ScrlSkillExp.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "ScrlSkillExp_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub ScrlSkillLevelReq_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_ITEMS Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If fraRecipe.Visible = False Then Exit Sub
    
    If ScrlSkillLevelReq.Value > 0 Then
        lblSkillLevelReq.Caption = "Skill Level Required: " & ScrlSkillLevelReq.Value
    Else
        lblSkillLevelReq.Caption = "Skill Level Required: None"
    End If
    
    Item(EditorIndex).SkillLevelReq = ScrlSkillLevelReq.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "ScrlSkillLevelReq_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlItem1_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_ITEMS Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If fraRecipe.Visible = False Then Exit Sub
    
    If scrlItem1.Value > 0 Then
        lblItem1.Caption = "Item 1: " & Trim$(Item(scrlItem1.Value).Name)
    Else
        lblItem1.Caption = "Item 1: None"
    End If
    
    Item(EditorIndex).Data1 = scrlItem1.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlItem1_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlItem2_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_ITEMS Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If fraRecipe.Visible = False Then Exit Sub
    
    If scrlItem2.Value > 0 Then
        lblItem2.Caption = "Item 2: " & Trim$(Item(scrlItem2.Value).Name)
    Else
        lblItem2.Caption = "Item 2: None"
    End If
    
    Item(EditorIndex).Data2 = scrlItem2.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlItem2_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlX_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_ITEMS Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If fraTeleport.Visible = False Then Exit Sub
    
    lblX.Caption = "X: " & scrlX.Value
    Item(EditorIndex).Data2 = scrlX.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlX_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlY_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_ITEMS Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lblY.Caption = "Y: " & scrlY.Value
    Item(EditorIndex).Data3 = scrlY.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlY_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtDesc_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_ITEMS Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If EditorIndex < 1 Or EditorIndex > MAX_ITEMS Then Exit Sub

    Item(EditorIndex).Desc = Trim$(txtDesc.text)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "txtDesc_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
    Dim TmpIndex As Long
    
    If EditorIndex < 1 Or EditorIndex > MAX_ITEMS Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    TmpIndex = lstIndex.ListIndex
    Item(EditorIndex).Name = Trim$(txtName.text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Item(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = TmpIndex
    Exit Sub
    
' Error handlerin
ErrorHandler:
    HandleError "txtName_Validate", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtPrice_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_ITEMS Then Exit Sub

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If Not IsNumeric(txtPrice.text) Then txtPrice.text = 0
    If txtPrice.text > MAX_LONG Then txtPrice.text = MAX_LONG
    If txtPrice.text < 0 Then txtPrice.text = 0
    Item(EditorIndex).Price = val(frmEditor_Item.txtPrice.text)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "txtPrice_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtName_GotFocus()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    txtName.SelStart = Len(txtName)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "txtName_GotFocus", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtPrice_GotFocus()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    txtPrice.SelStart = Len(txtPrice)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "txtPrice_GotFocus", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtDesc_GotFocus()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    txtDesc.SelStart = Len(txtDesc)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "txtDesc_GotFocus", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtSearch_Change()
    Dim Find As String, I As Long
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    For I = 0 To lstIndex.ListCount - 1
        Find = Trim$(I + 1 & ": " & txtSearch.text)
        
        ' Make sure we dont try to check a name that's too small
        If Len(lstIndex.List(I)) >= Len(Find) Then
            If UCase$(Mid$(Trim$(lstIndex.List(I)), 1, Len(Find))) = UCase$(Find) Then
                lstIndex.ListIndex = I
                Exit For
            End If
        End If
    Next
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "txtSearch_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtSearch_GotFocus()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    txtSearch.SelStart = Len(txtSearch)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "txtSearch_GotFocus", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Dim buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If KeyAscii = vbKeyReturn Then
        cmdSave_Click
        KeyAscii = 0
    ElseIf KeyAscii = vbKeyEscape Then
        cmdClose_Click
        KeyAscii = 0
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "Form_KeyPress", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdCopy_Click()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    TmpIndex = lstIndex.ListIndex
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdCopy_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdPaste_Click()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
         
    lstIndex.RemoveItem EditorIndex - 1
    Call CopyMemory(ByVal VarPtr(Item(EditorIndex)), ByVal VarPtr(Item(TmpIndex + 1)), LenB(Item(TmpIndex + 1)))
    lstIndex.AddItem EditorIndex & ": " & Trim$(Item(EditorIndex).Name), EditorIndex - 1
    lstIndex.ListIndex = EditorIndex - 1
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdPaste_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub
