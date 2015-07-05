VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmMainGame 
   BackColor       =   &H00004080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mirage"
   ClientHeight    =   8550
   ClientLeft      =   3960
   ClientTop       =   2790
   ClientWidth     =   11445
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "frmMirage.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmMirage.frx":000C
   ScaleHeight     =   570
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   763
   Visible         =   0   'False
   Begin VB.PictureBox picItemDesc 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4725
      Left            =   0
      Picture         =   "frmMirage.frx":13E6A2
      ScaleHeight     =   313
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   120
      TabIndex        =   14
      Top             =   0
      Visible         =   0   'False
      Width           =   1830
      Begin VB.Label lblItemDescName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ItemDescName"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H003E8CA6&
         Height          =   240
         Left            =   60
         TabIndex        =   18
         Top             =   75
         Width           =   1695
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblRequirement 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Requirement"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H001D2B34&
         Height          =   180
         Left            =   435
         TabIndex        =   17
         Top             =   1155
         Width           =   960
      End
      Begin VB.Label lblItemDescReq 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "DescReq"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H001D2B34&
         Height          =   180
         Left            =   600
         TabIndex        =   16
         Top             =   1425
         Width           =   630
      End
      Begin VB.Label lblItemName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Type"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H001D2B34&
         Height          =   255
         Left            =   75
         TabIndex        =   15
         Top             =   600
         Width           =   1680
      End
   End
   Begin VB.PictureBox picNpcQuests 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3480
      Left            =   2400
      Picture         =   "frmMirage.frx":15A1DC
      ScaleHeight     =   232
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   215
      TabIndex        =   65
      Top             =   1320
      Visible         =   0   'False
      Width           =   3225
      Begin VB.PictureBox picNpcAcceptQuest 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   260
         Left            =   1320
         Picture         =   "frmMirage.frx":17ED5E
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   55
         TabIndex        =   76
         ToolTipText     =   "Show your stats."
         Top             =   3045
         Visible         =   0   'False
         Width           =   820
      End
      Begin VB.PictureBox picNpcQuestInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   260
         Left            =   1320
         Picture         =   "frmMirage.frx":17FAC0
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   55
         TabIndex        =   75
         ToolTipText     =   "Show your stats."
         Top             =   3045
         Visible         =   0   'False
         Width           =   820
      End
      Begin VB.PictureBox picTurnInQuest 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   260
         Left            =   1320
         Picture         =   "frmMirage.frx":180822
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   55
         TabIndex        =   74
         ToolTipText     =   "Show your stats."
         Top             =   3045
         Visible         =   0   'False
         Width           =   820
      End
      Begin VB.PictureBox picQuestInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2580
         Left            =   180
         Picture         =   "frmMirage.frx":181584
         ScaleHeight     =   172
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   193
         TabIndex        =   68
         Top             =   450
         Visible         =   0   'False
         Width           =   2895
         Begin VB.PictureBox picQuestReward 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   480
            Index           =   2
            Left            =   1575
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   73
            Top             =   1950
            Width           =   480
         End
         Begin VB.PictureBox picQuestReward 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   480
            Index           =   1
            Left            =   885
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   72
            Top             =   1950
            Width           =   480
         End
         Begin VB.PictureBox picQuestReward 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   480
            Index           =   0
            Left            =   195
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   71
            Top             =   1950
            Width           =   480
         End
         Begin VB.Shape shpSelected 
            BorderColor     =   &H000000FF&
            BorderWidth     =   4
            Height          =   510
            Left            =   195
            Top             =   1950
            Visible         =   0   'False
            Width           =   510
         End
         Begin VB.Label lblQuestName 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Quest Name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   0
            TabIndex        =   70
            Top             =   0
            Width           =   1050
         End
         Begin VB.Label lblQuestDescription 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Quest Description"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   1515
            Left            =   0
            TabIndex        =   69
            Top             =   240
            Width           =   2820
         End
      End
      Begin VB.ListBox lstNpcQuests 
         Appearance      =   0  'Flat
         BackColor       =   &H00577E8F&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H002B3339&
         Height          =   2370
         ItemData        =   "frmMirage.frx":199B76
         Left            =   250
         List            =   "frmMirage.frx":199B7D
         TabIndex        =   67
         ToolTipText     =   "Current Abilities."
         Top             =   480
         Width           =   2715
      End
      Begin VB.PictureBox picNpcCloseQuest 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   260
         Left            =   2160
         Picture         =   "frmMirage.frx":199B8F
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   55
         TabIndex        =   66
         ToolTipText     =   "Show your stats."
         Top             =   3045
         Width           =   820
      End
   End
   Begin VB.PictureBox picQuest 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   9720
      Picture         =   "frmMirage.frx":19A8F1
      ScaleHeight     =   285
      ScaleWidth      =   1350
      TabIndex        =   61
      ToolTipText     =   "Show your stats."
      Top             =   7560
      Width           =   1350
   End
   Begin VB.PictureBox picQuests 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3300
      Left            =   8190
      Picture         =   "frmMirage.frx":19BF63
      ScaleHeight     =   220
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   203
      TabIndex        =   59
      Top             =   210
      Visible         =   0   'False
      Width           =   3045
      Begin VB.PictureBox picDropQuest 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   260
         Left            =   2125
         Picture         =   "frmMirage.frx":1BCD95
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   55
         TabIndex        =   63
         ToolTipText     =   "Show your stats."
         Top             =   2640
         Width           =   820
      End
      Begin VB.PictureBox picInfoQuest 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   260
         Left            =   2125
         Picture         =   "frmMirage.frx":1BDAF7
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   55
         TabIndex        =   62
         ToolTipText     =   "Show your stats."
         Top             =   2350
         Width           =   820
      End
      Begin VB.ListBox lstQuests 
         Appearance      =   0  'Flat
         BackColor       =   &H00577E8F&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H002B3339&
         Height          =   1830
         ItemData        =   "frmMirage.frx":1BE859
         Left            =   165
         List            =   "frmMirage.frx":1BE860
         TabIndex        =   60
         ToolTipText     =   "Current Abilities."
         Top             =   435
         Width           =   2715
      End
      Begin VB.Label lblQuestProgress 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Quest Progress:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   165
         TabIndex        =   64
         Top             =   2295
         Width           =   1365
      End
   End
   Begin VB.PictureBox picItemInfo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5955
      Left            =   3120
      Picture         =   "frmMirage.frx":1BE86F
      ScaleHeight     =   395
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   118
      TabIndex        =   8
      Top             =   240
      Visible         =   0   'False
      Width           =   1800
      Begin VB.Label lblItemInfoOk 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   120
         TabIndex        =   13
         Top             =   1230
         Width           =   1575
      End
      Begin VB.Label lblItemInfoName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ItemDescName"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H003E8CA6&
         Height          =   240
         Left            =   45
         TabIndex        =   12
         Top             =   75
         Width           =   1695
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblItemInfoType 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Type"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H001D2B34&
         Height          =   180
         Left            =   120
         TabIndex        =   11
         Top             =   1875
         Width           =   1560
      End
      Begin VB.Label lblItemInfoDesc 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Desc"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H001D2B34&
         Height          =   180
         Left            =   705
         TabIndex        =   10
         Top             =   2670
         Width           =   360
      End
      Begin VB.Label lblItemInfoReq 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Requirement"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H001D2B34&
         Height          =   180
         Left            =   405
         TabIndex        =   9
         Top             =   2400
         Width           =   960
      End
   End
   Begin VB.PictureBox picStats 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   3270
      Left            =   8160
      Picture         =   "frmMirage.frx":1E16F9
      ScaleHeight     =   218
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   201
      TabIndex        =   26
      Top             =   240
      Visible         =   0   'False
      Width           =   3015
      Begin VB.Label lblName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H001D2B34&
         Height          =   180
         Left            =   120
         TabIndex        =   45
         Top             =   315
         Width           =   2640
      End
      Begin VB.Label lblPoints 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Points"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H001D2B34&
         Height          =   180
         Left            =   120
         TabIndex        =   44
         Top             =   600
         Width           =   2640
      End
      Begin VB.Label lblStat 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Stat"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H001D2B34&
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   43
         Top             =   810
         Width           =   330
      End
      Begin VB.Label lblStatUp 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H001D2B34&
         Height          =   180
         Index           =   0
         Left            =   2775
         TabIndex        =   42
         Top             =   780
         Width           =   120
      End
      Begin VB.Label lblStat 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Stat"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H001D2B34&
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   41
         Top             =   975
         Width           =   330
      End
      Begin VB.Label lblStatUp 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H001D2B34&
         Height          =   180
         Index           =   1
         Left            =   2775
         TabIndex        =   40
         Top             =   945
         Width           =   120
      End
      Begin VB.Label lblStat 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Stat"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H001D2B34&
         Height          =   180
         Index           =   2
         Left            =   120
         TabIndex        =   39
         Top             =   1140
         Width           =   330
      End
      Begin VB.Label lblStatUp 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H001D2B34&
         Height          =   180
         Index           =   2
         Left            =   2775
         TabIndex        =   38
         Top             =   1110
         Width           =   120
      End
      Begin VB.Label lblStat 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Stat"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H001D2B34&
         Height          =   180
         Index           =   3
         Left            =   120
         TabIndex        =   37
         Top             =   1305
         Width           =   330
      End
      Begin VB.Label lblStatUp 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H001D2B34&
         Height          =   180
         Index           =   3
         Left            =   2775
         TabIndex        =   36
         Top             =   1275
         Width           =   120
      End
      Begin VB.Label lblStat 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Stat"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H001D2B34&
         Height          =   180
         Index           =   4
         Left            =   120
         TabIndex        =   35
         Top             =   1470
         Width           =   330
      End
      Begin VB.Label lblStatUp 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H001D2B34&
         Height          =   180
         Index           =   4
         Left            =   2775
         TabIndex        =   34
         Top             =   1440
         Width           =   120
      End
      Begin VB.Label lblDamage 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Damage:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H001D2B34&
         Height          =   180
         Left            =   120
         TabIndex        =   33
         Top             =   2400
         Width           =   660
      End
      Begin VB.Label lblProtection 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Protection:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H001D2B34&
         Height          =   180
         Left            =   120
         TabIndex        =   32
         Top             =   2580
         Width           =   855
      End
      Begin VB.Label lblMagicProtection 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Magic Protection:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H001D2B34&
         Height          =   180
         Left            =   120
         TabIndex        =   31
         Top             =   2940
         Width           =   1320
      End
      Begin VB.Label lblMagicDamage 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Magic Bonus:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H001D2B34&
         Height          =   180
         Left            =   120
         TabIndex        =   30
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label lblVital 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Vital"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H001D2B34&
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   29
         Top             =   1800
         Width           =   360
      End
      Begin VB.Label lblVital 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Vital"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H001D2B34&
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   28
         Top             =   1965
         Width           =   360
      End
      Begin VB.Label lblVital 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Vital"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H001D2B34&
         Height          =   180
         Index           =   2
         Left            =   120
         TabIndex        =   27
         Top             =   2130
         Width           =   360
      End
   End
   Begin VB.TextBox txtMyTextBox 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   600
      TabIndex        =   58
      Top             =   6285
      Width           =   7290
   End
   Begin VB.PictureBox picStat 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   8280
      Picture         =   "frmMirage.frx":201993
      ScaleHeight     =   285
      ScaleWidth      =   1350
      TabIndex        =   54
      ToolTipText     =   "Show your stats."
      Top             =   7560
      Width           =   1350
   End
   Begin VB.PictureBox PicSpells 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3300
      Left            =   8190
      Picture         =   "frmMirage.frx":203005
      ScaleHeight     =   3300
      ScaleWidth      =   3045
      TabIndex        =   48
      Top             =   240
      Visible         =   0   'False
      Width           =   3045
      Begin VB.PictureBox HotBar 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00577E8F&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   0
         Left            =   240
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   53
         Top             =   2535
         Width           =   480
      End
      Begin VB.ListBox lstSpells 
         Appearance      =   0  'Flat
         BackColor       =   &H00577E8F&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H002B3339&
         Height          =   1830
         ItemData        =   "frmMirage.frx":223E37
         Left            =   165
         List            =   "frmMirage.frx":223E3E
         TabIndex        =   52
         ToolTipText     =   "Current Abilities."
         Top             =   435
         Width           =   2715
      End
      Begin VB.PictureBox HotBar 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00577E8F&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   1
         Left            =   930
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   51
         Top             =   2535
         Width           =   480
      End
      Begin VB.PictureBox HotBar 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00577E8F&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   2
         Left            =   1620
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   50
         Top             =   2535
         Width           =   480
      End
      Begin VB.PictureBox HotBar 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00577E8F&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   3
         Left            =   2310
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   49
         Top             =   2535
         Width           =   480
      End
   End
   Begin VB.PictureBox picPouch 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   3270
      Left            =   8190
      Picture         =   "frmMirage.frx":223E4D
      ScaleHeight     =   218
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   201
      TabIndex        =   46
      Top             =   210
      Width           =   3015
      Begin VB.PictureBox picTempInv 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   600
         Left            =   120
         Picture         =   "frmMirage.frx":2440E7
         ScaleHeight     =   50
         ScaleMode       =   0  'User
         ScaleWidth      =   40
         TabIndex        =   47
         Top             =   375
         Visible         =   0   'False
         Width           =   600
      End
   End
   Begin VB.PictureBox PicMP 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'None
      FillColor       =   &H00008080&
      ForeColor       =   &H0000C0C0&
      Height          =   150
      Left            =   8400
      Picture         =   "frmMirage.frx":2453E9
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   178
      TabIndex        =   25
      Top             =   4350
      Width           =   2670
   End
   Begin VB.PictureBox picHP 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   150
      Left            =   8400
      Picture         =   "frmMirage.frx":24691D
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   178
      TabIndex        =   24
      Top             =   4050
      Width           =   2670
   End
   Begin VB.PictureBox PicExp 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'None
      FillColor       =   &H00008080&
      ForeColor       =   &H0000C0C0&
      Height          =   150
      Left            =   8400
      Picture         =   "frmMirage.frx":247E4F
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   178
      TabIndex        =   23
      Top             =   4650
      Width           =   2670
   End
   Begin VB.PictureBox picEquipment 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   0
      Left            =   8430
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   22
      Top             =   5595
      Width           =   480
   End
   Begin VB.PictureBox picEquipment 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   1
      Left            =   9120
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   21
      Top             =   5595
      Width           =   480
   End
   Begin VB.PictureBox picEquipment 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   2
      Left            =   9810
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   20
      Top             =   5595
      Width           =   480
   End
   Begin VB.PictureBox picEquipment 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   3
      Left            =   10500
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   19
      Top             =   5595
      Width           =   480
   End
   Begin VB.PictureBox picScreen 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
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
      Height          =   5760
      Left            =   210
      MousePointer    =   4  'Icon
      ScaleHeight     =   382
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   510
      TabIndex        =   7
      Top             =   210
      Width           =   7680
   End
   Begin MSWinsockLib.Winsock Socket 
      Left            =   420
      Top             =   810
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.PictureBox picChatControl 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2535
      Left            =   7920
      ScaleHeight     =   2535
      ScaleWidth      =   3165
      TabIndex        =   1
      Top             =   8640
      Visible         =   0   'False
      Width           =   3165
      Begin VB.OptionButton optParty 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Option1"
         Height          =   195
         Left            =   1080
         TabIndex        =   6
         Top             =   1320
         Width           =   195
      End
      Begin VB.OptionButton optGuild 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   720
         TabIndex        =   5
         Top             =   1320
         Width           =   195
      End
      Begin VB.OptionButton optglobal 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Option1"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   360
         TabIndex        =   3
         Top             =   1170
         Width           =   195
      End
      Begin VB.OptionButton optMap 
         Caption         =   "Option1"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   360
         TabIndex        =   2
         Top             =   765
         Value           =   -1  'True
         Width           =   165
      End
   End
   Begin RichTextLib.RichTextBox txtChat 
      Height          =   1680
      Left            =   240
      TabIndex        =   57
      Top             =   6660
      Width           =   7650
      _ExtentX        =   13494
      _ExtentY        =   2963
      _Version        =   393217
      BackColor       =   0
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmMirage.frx":2493A9
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblPouch 
      BackStyle       =   0  'Transparent
      Height          =   240
      Left            =   9735
      TabIndex        =   56
      ToolTipText     =   "Item pouch."
      Top             =   6840
      Width           =   1305
   End
   Begin VB.Label lblAbilities 
      BackStyle       =   0  'Transparent
      Height          =   210
      Left            =   8325
      TabIndex        =   55
      Top             =   6840
      Width           =   1275
   End
   Begin VB.Label lblLeave 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   9720
      TabIndex        =   4
      ToolTipText     =   "Quit the game."
      Top             =   7950
      Width           =   1350
   End
   Begin VB.Label lblEXP 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Left            =   5430
      TabIndex        =   0
      Top             =   1470
      Width           =   1230
   End
End
Attribute VB_Name = "frmMainGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public WithEvents Hypertext As clsHyperText
Attribute Hypertext.VB_VarHelpID = -1

Private Sub Form_Load()
        
    Me.Width = 11535
    Me.Height = 9015
    
    txtChat.SelHangingIndent = 8 ' Set hanging indent For chat box
    
    If ShowItemLinks Then
        Set Hypertext = New clsHyperText
        Hypertext.Initialize txtChat
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    'GameDestroy
    InGame = False
End Sub

Private Sub lblLeave_Click()
    'Call GameDestroy
    InGame = False
End Sub

Private Sub Hypertext_Clicked(Button As Integer, ItemNum As String, ItemName As String)
    If Button = vbLeftButton Then
        ShowItemInfo ItemNum
    End If
End Sub

Private Sub ShowItemInfo(ByVal ItemNum As Long)
Dim rec As RECT
Dim rec_pos As RECT
    
    lblItemInfoName.Caption = Trim$(Item(ItemNum).Name)
    
    Select Case Item(ItemNum).Type
        Case ITEM_TYPE_NONE
            lblItemInfoType.Caption = "Item"
            lblItemInfoReq.Caption = ItemReq(ItemNum)
            lblItemInfoDesc.Caption = vbNullString

        Case ITEM_TYPE_EQUIPMENT
            lblItemInfoType.Caption = EquipmentName(Item(ItemNum).Data1)
            lblItemInfoReq.Caption = ItemReq(ItemNum)
            lblItemInfoDesc.Caption = ItemDesc(ItemNum)
            
        Case ITEM_TYPE_POTION
            lblItemInfoType.Caption = "Potion"
            lblItemInfoReq.Caption = ItemReq(ItemNum)
            lblItemInfoDesc.Caption = ItemDesc(ItemNum)

        Case ITEM_TYPE_KEY
            lblItemInfoType.Caption = "Key"
            lblItemInfoReq.Caption = ItemReq(ItemNum)
            lblItemInfoDesc.Caption = vbNullString
        
        Case ITEM_TYPE_SPELL
            lblItemInfoType.Caption = "Spell"
            lblItemInfoReq.Caption = ItemReq(ItemNum)
            lblItemInfoDesc.Caption = Trim$(Spell(Item(ItemNum).Data1).Name) & vbNewLine
            
    End Select
    
    lblItemInfoDesc.Top = lblItemInfoReq.Top + lblItemInfoReq.Height
    picItemInfo.Height = lblItemInfoDesc.Top + lblItemInfoDesc.Height
    
    picItemInfo.Visible = True
    
    With rec
        .Top = Item(ItemNum).Pic * PIC_Y
        .Bottom = .Top + PIC_Y
        .Left = 0
        .Right = .Left + PIC_X
    End With
    
    With rec_pos
        .Top = 37
        .Bottom = .Top + PIC_Y
        .Left = 43
        .Right = .Left + PIC_X
    End With
    
    DD_ItemSurf.BltToDC frmMainGame.picItemInfo.hdc, rec, rec_pos
    picItemInfo.Refresh
    
End Sub

Private Sub Form_Click()
    picScreen.SetFocus
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If picItemDesc.Visible Then
        picItemDesc.Visible = False
    End If
End Sub


Private Sub lblPouch_Click()
    picPouch.Visible = True
    UpdateInventory
    
    PicSpells.Visible = False
    picStats.Visible = False
    picQuests.Visible = False
    'picChatControl.Visible = False
End Sub

Private Sub lblAbilities_Click()
    PicSpells.Visible = True
    picPouch.Visible = False
    picStats.Visible = False
    picQuests.Visible = False
    'picChatControl.Visible = False
End Sub

Private Sub picQuest_Click()
    PicSpells.Visible = False
    picPouch.Visible = False
    picStats.Visible = False
    picQuests.Visible = True
    UpdateQuestList
End Sub

Private Sub picStat_Click()
    PicSpells.Visible = False
    picPouch.Visible = False
    picStats.Visible = True
    picQuests.Visible = False
End Sub

Private Sub lblItemInfoOk_Click()
    picItemInfo.Visible = Not picItemInfo.Visible
End Sub

Private Sub lblStatUp_Click(Index As Integer)
    SendUseStatPoint Index + 1
End Sub

Private Sub lblStatUp_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long
    For i = 1 To Stats.Stat_Count
        lblStatUp(i - 1).ForeColor = QBColor(Black)
    Next
    lblStatUp(Index).ForeColor = QBColor(White)
End Sub

Private Sub lstSpells_GotFocus()
    frmMainGame.picScreen.SetFocus
End Sub

Private Sub optMap_Click()
    picScreen.SetFocus
    Call AddText("Now communicating in immediate chat (current map).", AlertColor)
End Sub

Private Sub optglobal_Click()
    picScreen.SetFocus
    Call AddText("Now communicating in realm-wide chat (everyone).", AlertColor)
End Sub

Private Sub optGuild_Click()
    picScreen.SetFocus
    Call AddText("Now communicating in guild chat (Guild).", AlertColor)
End Sub

Private Sub optparty_Click()
    picScreen.SetFocus
    Call AddText("Now communicating in party chat (Party).", AlertColor)
End Sub

Private Sub picScreen_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    MapEditorMouseDown Button, Shift, CLng(X), CLng(Y)

    If Not InEditor Then
        X = TileView.Left + ((X + Camera.Left) \ PIC_X)
        Y = TileView.Top + ((Y + Camera.Top) \ PIC_Y)
        
        If Not IsValidMapPoint(CLng(X), CLng(Y)) Then Exit Sub
    
        If Button = 1 Then SendSearch X, Y
        If Button = 2 Then
            If Player(MyIndex).Access >= 1 Then
                GettingMap = True
                SendClickWarp X, Y
            End If
        End If
    End If

End Sub

Private Sub picScreen_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CurX = TileView.Left + ((X + Camera.Left) \ PIC_X)
    CurY = TileView.Top + ((Y + Camera.Top) \ PIC_Y)
    MouseX = X
    MouseY = Y
    MapEditorMouseDown Button, Shift, CLng(X), CLng(Y)
End Sub

Private Sub Socket_Close()
     '  Have to clear out key otherwise we wouldn't be able to reconnect to the server
    PacketInIndex = 0
    PacketOutIndex = 0
End Sub

Private Sub Socket_DataArrival(ByVal bytesTotal As Long)
    'If IsConnected Then
        IncomingData bytesTotal
    'End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Call HandleKeypresses(KeyAscii)
    frmMainGame.txtMyTextBox.Text = MyText
    If Len(MyText) > 4 Then frmMainGame.txtMyTextBox.SelStart = Len(MyText) + 1
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Call CheckInput(1, KeyCode, Shift)
End Sub

Private Sub txtMyTextBox_GotFocus()
    picScreen.SetFocus
End Sub

Private Sub txtchat_GotFocus()
    picScreen.SetFocus
End Sub

Private Sub HotBar_Click(Index As Integer)
Dim SpellSlot As Long
Dim SpellNum As Long
Dim rec As RECT
Dim rec_pos As RECT

    SpellSlot = lstSpells.ListIndex + 1
    SpellNum = Player(MyIndex).Spell(SpellSlot).SpellNum
    
    If SpellNum > 0 Then
        HotBarSpell(Index) = SpellSlot
        HotBar(Index).ToolTipText = Trim$(Spell(SpellNum).Name)
        
        HotBar(Index).Cls
        
        ' Draw the shit,yo
        If Spell(SpellNum).Animation > 0 Then
            With rec_pos
                .Top = 0
                .Bottom = PIC_Y
                .Left = 0
                .Right = PIC_X
            End With
            
            Select Case Animation(Spell(SpellNum).Animation).AnimationSize
                Case 1
                    With rec
                        .Top = Animation(Spell(SpellNum).Animation).Animation * PIC_Y
                        .Bottom = .Top + PIC_Y
                        .Left = (Animation(Spell(SpellNum).Animation).AnimationFrames \ 2) * PIC_X
                        .Right = .Left + PIC_X
                    End With
                                       
                    DD_AnimationSurf.BltToDC HotBar(Index).hdc, rec, rec_pos
                Case 2
                    With rec
                        .Top = Animation(Spell(SpellNum).Animation).Animation * (PIC_Y * 2)
                        .Bottom = .Top + (PIC_Y * 2)
                        .Left = (Animation(Spell(SpellNum).Animation).AnimationFrames \ 2) * (PIC_X * 2)
                        .Right = .Left + (PIC_X * 2)
                    End With
                                       
                    DD_AnimationSurf2.BltToDC HotBar(Index).hdc, rec, rec_pos
            End Select
        End If
        
        HotBar(Index).Refresh
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    Call CheckInput(0, KeyCode, Shift)
        
    If KeyCode = vbKeyF1 Then
        If frmMainGame.optMap.Value = False Then
            frmMainGame.optMap.Value = True
        End If
    End If

    If KeyCode = vbKeyF2 Then
        If frmMainGame.optglobal.Value = False Then
            frmMainGame.optglobal.Value = True
        End If
    End If
    
    If KeyCode = vbKeyF3 Then
        If frmMainGame.optGuild.Value = False Then
            frmMainGame.optGuild.Value = True
        End If
    End If

    If KeyCode = vbKeyF4 Then
        If frmMainGame.optParty.Value = False Then
            frmMainGame.optParty.Value = True
        End If
    End If
    
    If KeyCode = vbKeyF5 Then
        CastSpell HotBarSpell(0)
    End If
    
    If KeyCode = vbKeyF6 Then
        CastSpell HotBarSpell(1)
    End If
    
    If KeyCode = vbKeyF7 Then
        CastSpell HotBarSpell(2)
    End If
    
    If KeyCode = vbKeyF8 Then
        CastSpell HotBarSpell(3)
    End If
    
    If KeyCode = vbKeyEnd Then
        CastSpell frmMainGame.lstSpells.ListIndex + 1
    End If
End Sub

Private Sub picPouch_DblClick()
Dim InvNum As Long
    
    DragInvSlotNum = 0
    If ShiftDown Then Exit Sub

    InvNum = IsItem(InvX, InvY)
    If InvNum <> 0 Then
         If Current_InvItemNum(MyIndex, InvNum) = ITEM_TYPE_NONE Then Exit Sub
         
         Call SendUseItem(InvNum)
         Exit Sub
    End If
End Sub

Private Sub picPouch_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim InvNum As Long

    InvNum = IsItem(X, Y)
    
    If InvNum <> 0 Then
        If ShiftDown Then
            If Button = 1 Then
                MyText = MyText & " [" & Trim$(Item(Current_InvItemNum(MyIndex, InvNum)).Name) & "]"
                frmMainGame.txtMyTextBox.Text = MyText
                If Len(MyText) > 4 Then frmMainGame.txtMyTextBox.SelStart = Len(MyText) + 1
                Exit Sub
            End If
        End If
        
        If Button = 1 Then
            DragInvSlotNum = InvNum
        ElseIf Button = 2 Then
            
            ' Check if it's bound to you
            If Current_InvItemBound(MyIndex, InvNum) Then
                If MsgBox("If you drop this item, it will be destroyed. Are you sure you want to do this?", vbOKCancel) = vbCancel Then Exit Sub
            End If
                        
            ' Check to see if item stacks so we can drop certain amount
            If Item(Current_InvItemNum(MyIndex, InvNum)).Stack Then
                DropNum = InvNum
                frmMainGame.picItemDesc.Visible = False
                If Not frmDrop.Visible Then
                    frmDrop.Top = Y - 20
                    frmDrop.Left = X - 20
                    frmDrop.Show vbModal
                End If
            Else
                Call SendDropItem(InvNum, 1)
            End If
        End If
    End If
End Sub

Private Sub picPouch_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long
Dim rec_pos As RECT

    If DragInvSlotNum > 0 Then
        For i = 1 To MAX_INV
            With rec_pos
                .Top = InvTop + ((InvOffsetY + 32) * ((i - 1) \ InvColumns))
                .Bottom = .Top + PIC_Y
                .Left = InvLeft + ((InvOffsetX + 32) * (((i - 1) Mod InvColumns)))
                .Right = .Left + PIC_X
            End With

            If X >= rec_pos.Left And X <= rec_pos.Right Then
                If Y >= rec_pos.Top And Y <= rec_pos.Bottom Then '
                    If DragInvSlotNum <> i Then
                        SwitchInvSlots DragInvSlotNum, i
                        Exit For
                    End If
                End If
            End If
        Next
    End If

    DragInvSlotNum = 0
    picTempInv.Visible = False
End Sub

Private Sub picPouch_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim InvNum As Long, ItemNum As Long, ItemType As Long
Dim X2 As Long, Y2 As Long

    InvX = X
    InvY = Y

    If DragInvSlotNum > 0 Then
        Call BltInventoryItem(CLng(X), CLng(Y))
    Else
        InvNum = IsItem(X, Y)

        If InvNum <> 0 Then
            ItemNum = Current_InvItemNum(MyIndex, InvNum)
            ItemType = Item(ItemNum).Type

            lblItemDescName.Caption = Trim$(Item(ItemNum).Name)
            Select Case ItemType
                Case ITEM_TYPE_NONE
                    lblItemName.Caption = "Item"
                    lblRequirement.Caption = ItemReq(ItemNum)
                    lblItemDescReq.Caption = "Value: " & Current_InvItemValue(MyIndex, InvNum) & vbNewLine

                Case ITEM_TYPE_EQUIPMENT
                    lblItemName.Caption = EquipmentName(Item(ItemNum).Data1)
                    lblRequirement.Caption = ItemReq(ItemNum)
                    lblItemDescReq.Caption = ItemDesc(ItemNum)

                Case ITEM_TYPE_POTION
                    lblItemName.Caption = "Potion"
                    lblRequirement.Caption = ItemReq(ItemNum)
                    lblItemDescReq.Caption = ItemDesc(ItemNum)

                Case ITEM_TYPE_KEY
                    lblItemName.Caption = "Key"
                    lblRequirement.Caption = ItemReq(ItemNum)
                    lblItemDescReq.Caption = "Carrying: " & Current_InvItemValue(MyIndex, InvNum) & vbNewLine

                Case ITEM_TYPE_SPELL
                    lblItemName.Caption = "Spell"
                    lblRequirement.Caption = ItemReq(ItemNum)
                    lblItemDescReq.Caption = Trim$(Spell(Item(ItemNum).Data1).Name) & vbNewLine

            End Select

            lblItemDescReq.Top = lblRequirement.Top + lblRequirement.Height
            picItemDesc.Height = lblItemDescReq.Top + lblItemDescReq.Height

            X2 = (X - picItemDesc.Width) + picPouch.Left
            Y2 = (Y + picPouch.Top) + 20

            If X2 < (picScreen.Left + picScreen.Width) Then
                X2 = picScreen.Left + picScreen.Width
            End If

            picItemDesc.Top = Y2
            picItemDesc.Left = X2

            picItemDesc.Visible = True
            Exit Sub
        End If
    End If

    picItemDesc.Visible = False
End Sub

Private Sub picEquipment_Click(Index As Integer)
    If ShiftDown Then
        MyText = MyText & " [" & Trim$(Item(Current_EquipmentSlot(MyIndex, Index + 1)).Name) & "]"
        frmMainGame.txtMyTextBox.Text = MyText
        If Len(MyText) > 4 Then frmMainGame.txtMyTextBox.SelStart = Len(MyText) + 1
    End If
End Sub

Private Sub picEquipment_DblClick(Index As Integer)
Dim i As Long
Dim EquipmentSlot As Long

    If ShiftDown Then Exit Sub
    
    ' Doesnt' matter if they are dead
    If Player(MyIndex).IsDead Then Exit Sub
    
    EquipmentSlot = Index + 1
    If EquipmentSlot <= 0 Then Exit Sub
    If EquipmentSlot > Slots.Slot_Count Then Exit Sub

    i = Current_EquipmentSlot(MyIndex, EquipmentSlot)
    If i > 0 Then
        SendUnequipSlot EquipmentSlot
    End If
End Sub

Private Sub picEquipment_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long, X2 As Long, Y2 As Long
Dim EquipmentSlot As Long
    
    If Button <= 0 Then
        EquipmentSlot = Index + 1
        If EquipmentSlot <= 0 Then Exit Sub
        If EquipmentSlot > Slots.Slot_Count Then Exit Sub
        
        i = Current_EquipmentSlot(MyIndex, EquipmentSlot)
        If i > 0 Then
            lblItemDescName.Caption = Trim$(Item(i).Name)
            
            lblItemName.Caption = EquipmentName(Item(i).Data1)
            lblRequirement.Caption = ItemReq(i)
            lblItemDescReq.Caption = ItemDesc(i)
            
            lblItemDescReq.Top = lblRequirement.Top + lblRequirement.Height
            picItemDesc.Height = lblItemDescReq.Top + lblItemDescReq.Height
        
            Y2 = (Y + picEquipment(Index).Top) - picItemDesc.Height - 12
            
            X2 = (X - picItemDesc.Width) + picEquipment(Index).Left
            If X2 < (picScreen.Left + picScreen.Width) Then X2 = picScreen.Left + picScreen.Width
            
            picItemDesc.Left = X2
            picItemDesc.Top = Y2
            
            picItemDesc.Visible = True
        End If
    Else
        picItemDesc.Visible = False
    End If
End Sub

Private Function IsItem(ByVal X As Single, ByVal Y As Single) As Long
Dim tempRec As RECT
Dim i As Long
    
    For i = 1 To MAX_INV
        If Current_InvItemNum(MyIndex, i) > 0 And Current_InvItemNum(MyIndex, i) <= MAX_ITEMS Then
            With tempRec
                .Top = InvTop + ((InvOffsetY + 32) * ((i - 1) \ InvColumns))
                .Bottom = .Top + PIC_Y
                .Left = InvLeft + ((InvOffsetX + 32) * (((i - 1) Mod InvColumns)))
                .Right = .Left + PIC_X
            End With
            
            If X >= tempRec.Left And X <= tempRec.Right Then
                If Y >= tempRec.Top And Y <= tempRec.Bottom Then
                    
                    IsItem = i
                    Exit Function
                End If
            End If
        End If
    Next
End Function

'
' Quest Things
'

Private Sub picInfoQuest_Click()
Dim i As Long
Dim RewardCount As Long
Dim SelectRewardCount As Long
Dim QuestNum As Long
Dim QuestProgressNum As Long
Dim rec As RECT
Dim rec_pos As RECT

    If lstQuests.ListCount <= 0 Then Exit Sub
    
    ' Load in the data
    QuestNum = lstQuests.ItemData(lstQuests.ListIndex)
    CurrentSelectedQuest = QuestNum
    
    lblQuestName.Caption = Trim$(Quest(QuestNum).Name)
    lblQuestDescription.Caption = QuestDescription(QuestNum)
    
    picTurnInQuest.Visible = False
    picNpcAcceptQuest.Visible = False
    
    ' Draw the rewards
    With rec_pos
        .Top = 0
        .Bottom = PIC_Y
        .Left = 0
        .Right = PIC_X
    End With
    For i = 1 To MAX_QUEST_REWARDS
        picQuestReward(i - 1).Cls
        If Quest(QuestNum).Rewards(i).ItemNum Then
            With rec
                .Top = Item(Quest(QuestNum).Rewards(i).ItemNum).Pic * PIC_Y
                .Bottom = .Top + PIC_Y
                .Left = 0
                .Right = .Left + PIC_X
            End With
            
            DD_ItemSurf.BltToDC picQuestReward(i - 1).hdc, rec, rec_pos
            DrawText picQuestReward(i - 1).hdc, rec_pos.Left, rec_pos.Top, Quest(QuestNum).Rewards(i).ItemValue, QBColor(White)
            ' If it's a selection draw a S
            If Quest(QuestNum).Rewards(i).SelectionOnly Then
                DrawText picQuestReward(i - 1).hdc, 20, 16, "S", QBColor(White)
            End If
        End If
        picQuestReward(i - 1).Refresh
    Next
        
    picNpcQuests.Visible = True
    
    ' Show the info area
    picQuestInfo.Visible = True
    
    picScreen.SetFocus
End Sub

Private Sub picDropQuest_Click()
Dim QuestNum As Long
Dim QuestProgressNum As Long
    
    If lstQuests.ListCount <= 0 Then Exit Sub
    
    QuestNum = lstQuests.ItemData(lstQuests.ListIndex)
    QuestProgressNum = OnQuest(QuestNum)
    If QuestProgressNum Then
        SendDropQuest QuestProgressNum
    End If
    
    picScreen.SetFocus
End Sub

Private Sub lstQuests_Click()
Dim QuestNum As Long
Dim QuestProgressNum As Long
Dim i As Long

    If Not InGame Then Exit Sub
    
    LastQuestClicked = lstQuests.ListIndex
    
    If LastQuestClicked < 0 Then Exit Sub
    If LastQuestClicked > lstQuests.ListCount - 1 Then Exit Sub
    
    QuestNum = lstQuests.ItemData(LastQuestClicked)
    
    lblQuestProgress.Caption = vbNullString
    
    ' Check if you're on quest
    QuestProgressNum = OnQuest(QuestNum)
    If QuestProgressNum > 0 Then
        ' Quest Needs
        lblQuestProgress.Caption = "Quest Progress: " & vbNewLine
        For i = 1 To MAX_QUEST_NEEDS
            Select Case Quest(QuestNum).QuestNeeds(i).QuestType
                Case QuestTypes.KillNpc
                    lblQuestProgress.Caption = lblQuestProgress.Caption & Player(MyIndex).QuestProgress(QuestProgressNum).Progress(i) & " / " & Quest(QuestNum).QuestNeeds(i).Data2 & " " & Trim$(Npc(Quest(QuestNum).QuestNeeds(i).Data1).Name) & vbNewLine
                Case QuestTypes.ItemCollection
                    lblQuestProgress.Caption = lblQuestProgress.Caption & Player(MyIndex).QuestProgress(QuestProgressNum).Progress(i) & " / " & Quest(QuestNum).QuestNeeds(i).Data2 & " " & Trim$(Item(Quest(QuestNum).QuestNeeds(i).Data1).Name) & vbNewLine
                Case QuestTypes.ExploreMap
                    If Player(MyIndex).QuestProgress(QuestProgressNum).Progress(i) Then
                        lblQuestProgress.Caption = lblQuestProgress.Caption & "Explore an area."
                    Else
                        lblQuestProgress.Caption = lblQuestProgress.Caption & "Explored!"
                    End If
            End Select
        Next
    End If
    
    picScreen.SetFocus
End Sub

' Npc Quest Stuff

Private Sub picNpcQuestInfo_Click()
Dim i As Long
Dim RewardCount As Long
Dim SelectRewardCount As Long
Dim QuestNum As Long
Dim QuestProgressNum As Long
Dim rec As RECT
Dim rec_pos As RECT

    If lstNpcQuests.ListCount <= 0 Then Exit Sub
    
    ' Load in the data
    QuestNum = lstNpcQuests.ItemData(lstNpcQuests.ListIndex)
    CurrentSelectedQuest = QuestNum
    
    lblQuestName.Caption = Trim$(Quest(QuestNum).Name)
    lblQuestDescription.Caption = QuestDescription(QuestNum)
        
    ' Check if you're on quest
    QuestProgressNum = OnQuest(QuestNum)
    If QuestProgressNum > 0 Then
        ' If this is the turn in NPC, show the "TurnIn" button
        If Quest(QuestNum).EndNPC = MapNpc(QuestMapNpcNum).Num Then
            picTurnInQuest.Visible = True
            picNpcAcceptQuest.Visible = False
        End If
    Else
        picTurnInQuest.Visible = False
        picNpcAcceptQuest.Visible = True
    End If
    
    ' Draw the rewards
    With rec_pos
        .Top = 0
        .Bottom = PIC_Y
        .Left = 0
        .Right = PIC_X
    End With
    For i = 1 To MAX_QUEST_REWARDS
        picQuestReward(i - 1).Cls
        If Quest(QuestNum).Rewards(i).ItemNum Then
            With rec
                .Top = Item(Quest(QuestNum).Rewards(i).ItemNum).Pic * PIC_Y
                .Bottom = .Top + PIC_Y
                .Left = 0
                .Right = .Left + PIC_X
            End With
            
            DD_ItemSurf.BltToDC picQuestReward(i - 1).hdc, rec, rec_pos
            DrawText picQuestReward(i - 1).hdc, rec_pos.Left, rec_pos.Top, Quest(QuestNum).Rewards(i).ItemValue, QBColor(White)
            ' If it's a selection draw a S
            If Quest(QuestNum).Rewards(i).SelectionOnly Then
                DrawText picQuestReward(i - 1).hdc, 20, 16, "S", QBColor(White)
            End If
        End If
        picQuestReward(i - 1).Refresh
    Next
    
    ' Hide the info button now
    picNpcQuestInfo.Visible = False
    
    ' Show the info area
    picQuestInfo.Visible = True
End Sub

Private Sub picQuestReward_Click(Index As Integer)
    
    ' Make sure it's a valid reward
    If Quest(CurrentSelectedQuest).Rewards(Index + 1).ItemNum <= 0 Then Exit Sub
    
    ' Make sure it's a selection
    If Not Quest(CurrentSelectedQuest).Rewards(Index + 1).SelectionOnly Then Exit Sub
    
    SelectReward = Index + 1
    shpSelected.Left = picQuestReward(Index).Left - 1
    shpSelected.Visible = True
End Sub

Private Sub picQuestReward_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ItemNum As Long
Dim X2 As Long
Dim Y2 As Long

    ' Make sure it's a valid reward
    If Quest(CurrentSelectedQuest).Rewards(Index + 1).ItemNum <= 0 Then Exit Sub
        
    ItemNum = Quest(CurrentSelectedQuest).Rewards(Index + 1).ItemNum

    lblItemDescName.Caption = Trim$(Item(ItemNum).Name)
    Select Case Item(ItemNum).Type
        Case ITEM_TYPE_NONE
            lblItemName.Caption = "Item"
            lblRequirement.Caption = ItemReq(ItemNum)
            lblItemDescReq.Caption = "Value: " & Quest(CurrentSelectedQuest).Rewards(Index + 1).ItemValue & vbNewLine

        Case ITEM_TYPE_EQUIPMENT
            lblItemName.Caption = EquipmentName(Item(ItemNum).Data1)
            lblRequirement.Caption = ItemReq(ItemNum)
            lblItemDescReq.Caption = ItemDesc(ItemNum)

        Case ITEM_TYPE_POTION
            lblItemName.Caption = "Potion"
            lblRequirement.Caption = ItemReq(ItemNum)
            lblItemDescReq.Caption = ItemDesc(ItemNum)

        Case ITEM_TYPE_KEY
            lblItemName.Caption = "Key"
            lblRequirement.Caption = ItemReq(ItemNum)
            lblItemDescReq.Caption = "Carrying: " & Quest(CurrentSelectedQuest).Rewards(Index + 1).ItemValue & vbNewLine

        Case ITEM_TYPE_SPELL
            lblItemName.Caption = "Spell"
            lblRequirement.Caption = ItemReq(ItemNum)
            lblItemDescReq.Caption = Trim$(Spell(Item(ItemNum).Data1).Name) & vbNewLine

    End Select

    lblItemDescReq.Top = lblRequirement.Top + lblRequirement.Height
    picItemDesc.Height = lblItemDescReq.Top + lblItemDescReq.Height

    picItemDesc.Top = picNpcQuests.Top '(Y + picQuestReward(Index).Top) + 20
    picItemDesc.Left = picNpcQuests.Left + picNpcQuests.Width 'X + picQuestReward(Index).Left + picItemDesc.Width '- picItemDesc.Width) '+ picQuestReward(Index).Left

    picItemDesc.Visible = True
End Sub

Private Sub picTurnInQuest_Click()
Dim QuestNum As Long
Dim QuestProgressNum As Long
    
    QuestNum = lstNpcQuests.ItemData(lstNpcQuests.ListIndex)
    QuestProgressNum = OnQuest(QuestNum)
    SendCompleteQuest QuestProgressNum, SelectReward
    
    ' Close it
    picNpcCloseQuest_Click
End Sub

Private Sub picNpcAcceptQuest_Click()
Dim QuestNum As Long

    'Get the quest num
    QuestNum = lstNpcQuests.ItemData(lstNpcQuests.ListIndex)
    SendAcceptQuest QuestNum
    
    ' Close it
    picNpcCloseQuest_Click
End Sub

Private Sub picNpcCloseQuest_Click()
    QuestMapNpcNum = 0
    CurrentSelectedQuest = 0
    SelectReward = 0
    
    shpSelected.Visible = False
    
    ' close the info window
    picQuestInfo.Visible = False
    picNpcQuests.Visible = False
    
    ' Hide buttons
    picNpcAcceptQuest.Visible = False
    picNpcQuestInfo.Visible = False
    
    picScreen.SetFocus
End Sub

Private Sub picQuestInfo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picItemDesc.Visible = False
End Sub

Private Sub picNpcQuests_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picItemDesc.Visible = False
End Sub
