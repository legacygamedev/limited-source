VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMirage 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Chaos Engine"
   ClientHeight    =   8955
   ClientLeft      =   300
   ClientTop       =   330
   ClientWidth     =   12000
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   FontTransparent =   0   'False
   ForeColor       =   &H00000000&
   Icon            =   "frmMirage.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmMirage.frx":17D2A
   ScaleHeight     =   597
   ScaleMode       =   0  'User
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.PictureBox picChatBox 
      Appearance      =   0  'Flat
      BackColor       =   &H00000080&
      ForeColor       =   &H80000008&
      Height          =   2055
      Left            =   7080
      ScaleHeight     =   135
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   327
      TabIndex        =   223
      Top             =   6960
      Width           =   4935
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00808000&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         ItemData        =   "frmMirage.frx":19D5A
         Left            =   240
         List            =   "frmMirage.frx":19D70
         Style           =   2  'Dropdown List
         TabIndex        =   225
         Top             =   360
         Width           =   885
      End
      Begin VB.TextBox txtMyTextBox 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   80
         MaxLength       =   255
         TabIndex        =   224
         Top             =   720
         Width           =   4755
      End
      Begin RichTextLib.RichTextBox txtChat 
         Height          =   900
         Left            =   75
         TabIndex        =   226
         Top             =   1080
         Width           =   4755
         _ExtentX        =   8387
         _ExtentY        =   1588
         _Version        =   393217
         BackColor       =   16744576
         Enabled         =   -1  'True
         ScrollBars      =   2
         Appearance      =   0
         TextRTF         =   $"frmMirage.frx":19DA1
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
      Begin VB.Label lblGold 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Gold:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   1920
         TabIndex        =   230
         Top             =   120
         Width           =   1905
      End
      Begin VB.Label Label66 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "It is now:"
         ForeColor       =   &H00C0FFC0&
         Height          =   255
         Left            =   1440
         TabIndex        =   229
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label GameClock 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "(time)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   2880
         TabIndex        =   228
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label40 
         BackStyle       =   0  'Transparent
         Caption         =   " X"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   4680
         TabIndex        =   227
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picItems 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   3600
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   222
      Top             =   8280
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Timer tmrDisease 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4680
      Top             =   8400
   End
   Begin VB.Timer tmrPoison 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4200
      Top             =   8400
   End
   Begin VB.PictureBox picStat 
      Appearance      =   0  'Flat
      BackColor       =   &H00000080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5415
      Left            =   4320
      Picture         =   "frmMirage.frx":19E1C
      ScaleHeight     =   361
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   229
      TabIndex        =   117
      Top             =   1080
      Visible         =   0   'False
      Width           =   3435
      Begin VB.Label lblArrows 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Arrows"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080C0FF&
         Height          =   225
         Left            =   720
         TabIndex        =   220
         Top             =   3600
         Width           =   1905
      End
      Begin VB.Label lblPoints 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Points"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   180
         Left            =   1800
         TabIndex        =   141
         Top             =   600
         Width           =   1185
      End
      Begin VB.Label lblLevel 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Level"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   180
         Left            =   360
         TabIndex        =   140
         Top             =   600
         Width           =   1305
      End
      Begin VB.Label lblSPEED 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Speed"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   225
         Left            =   720
         TabIndex        =   139
         Top             =   1800
         Width           =   1905
      End
      Begin VB.Label lblMAGI 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Magic"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   225
         Left            =   720
         TabIndex        =   138
         Top             =   2040
         Width           =   1905
      End
      Begin VB.Label lblDEF 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Defense"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   225
         Left            =   600
         TabIndex        =   137
         Top             =   1560
         Width           =   2265
      End
      Begin VB.Label lblSTR 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Strength"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Left            =   570
         TabIndex        =   136
         Top             =   1320
         Width           =   2265
      End
      Begin VB.Label AddStr 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   165
         Left            =   2880
         TabIndex        =   135
         Top             =   1320
         Width           =   105
      End
      Begin VB.Label AddMagi 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   165
         Left            =   2880
         TabIndex        =   134
         Top             =   2040
         Width           =   105
      End
      Begin VB.Label AddSpeed 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   165
         Left            =   2880
         TabIndex        =   133
         Top             =   1800
         Width           =   105
      End
      Begin VB.Label AddDef 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   165
         Left            =   2880
         TabIndex        =   132
         Top             =   1560
         Width           =   105
      End
      Begin VB.Label Label38 
         BackStyle       =   0  'Transparent
         Caption         =   " X"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3120
         TabIndex        =   131
         Top             =   120
         Width           =   255
      End
      Begin VB.Label lblName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   225
         Left            =   720
         TabIndex        =   130
         Top             =   960
         Width           =   1905
      End
      Begin VB.Label lblKills 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Kills"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080C0FF&
         Height          =   225
         Left            =   720
         TabIndex        =   129
         Top             =   2400
         Width           =   1905
      End
      Begin VB.Label lblAccess 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Access"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Left            =   720
         TabIndex        =   128
         Top             =   2640
         Width           =   1905
      End
      Begin VB.Label lblClass 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Class"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C000&
         Height          =   225
         Left            =   720
         TabIndex        =   127
         Top             =   2880
         Width           =   1905
      End
      Begin VB.Label lblSex 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Sex"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0FF&
         Height          =   225
         Left            =   720
         TabIndex        =   126
         Top             =   3120
         Width           =   1905
      End
      Begin VB.Label Label37 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Guild Name :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Left            =   480
         TabIndex        =   125
         Top             =   4440
         Width           =   855
      End
      Begin VB.Label lblGuild 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Guild"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   240
         Left            =   1410
         TabIndex        =   124
         Top             =   4440
         Width           =   465
      End
      Begin VB.Label Label35 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Your Rank :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   195
         Left            =   600
         TabIndex        =   123
         Top             =   4800
         Width           =   765
      End
      Begin VB.Label lblRank 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Rank"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   1440
         TabIndex        =   122
         Top             =   4800
         Width           =   510
      End
      Begin VB.Label lblLeaveGuild 
         BackStyle       =   0  'Transparent
         Caption         =   "Leave Guild"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   1200
         TabIndex        =   121
         Top             =   5160
         Width           =   1005
      End
      Begin VB.Label Label63 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Guild Information"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   270
         Left            =   720
         TabIndex        =   120
         Top             =   4080
         Width           =   1980
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         Caption         =   "  Character Information"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   270
         Left            =   360
         TabIndex        =   119
         Top             =   240
         Width           =   2745
      End
      Begin VB.Label lblAlign 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Align"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   720
         TabIndex        =   118
         Top             =   3360
         Width           =   1905
      End
   End
   Begin VB.PictureBox Picture15 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   12000
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   114
      Top             =   8880
      Width           =   495
   End
   Begin VB.PictureBox Picture13 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   12000
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   113
      Top             =   8880
      Width           =   615
   End
   Begin VB.PictureBox picOptions 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
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
      Height          =   3405
      Left            =   5880
      Picture         =   "frmMirage.frx":57B60
      ScaleHeight     =   227
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   366
      TabIndex        =   3
      Top             =   2160
      Visible         =   0   'False
      Width           =   5490
      Begin VB.CheckBox chkFPS 
         BackColor       =   &H00FF8080&
         Caption         =   "Toggle FPS On Screen"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1560
         TabIndex        =   221
         Top             =   2400
         Value           =   1  'Checked
         Width           =   1725
      End
      Begin VB.CheckBox chkclicktomove 
         BackColor       =   &H00FF8080&
         Caption         =   "Mouse Movement"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   3120
         TabIndex        =   146
         Top             =   2040
         Value           =   1  'Checked
         Width           =   1485
      End
      Begin VB.CheckBox chkEmoticons 
         BackColor       =   &H00FF8080&
         Caption         =   "Emoticons"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1560
         TabIndex        =   144
         Top             =   2040
         Value           =   1  'Checked
         Width           =   1005
      End
      Begin VB.CheckBox chkWeather 
         BackColor       =   &H00FF8080&
         Caption         =   "Weather Graphics"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   3120
         TabIndex        =   143
         Top             =   1680
         Value           =   1  'Checked
         Width           =   1365
      End
      Begin VB.CheckBox chkNight 
         BackColor       =   &H00FF8080&
         Caption         =   "Night Graphics"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1560
         TabIndex        =   142
         Top             =   1680
         Value           =   1  'Checked
         Width           =   1245
      End
      Begin VB.CheckBox chkEmoSound 
         BackColor       =   &H00FF8080&
         Caption         =   "Emoticon Sounds"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   3480
         TabIndex        =   49
         Top             =   960
         Value           =   1  'Checked
         Width           =   1365
      End
      Begin VB.CheckBox chkAutoScroll 
         BackColor       =   &H00FF8080&
         Caption         =   "Auto Scroll"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   3120
         TabIndex        =   44
         Top             =   1320
         Value           =   1  'Checked
         Width           =   1005
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Ok"
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
         Left            =   480
         TabIndex        =   43
         Top             =   2880
         Width           =   495
      End
      Begin VB.HScrollBar scrlBltText 
         Height          =   255
         Left            =   3120
         Max             =   20
         Min             =   4
         TabIndex        =   17
         Top             =   2880
         Value           =   6
         Width           =   1935
      End
      Begin VB.CheckBox chksound 
         BackColor       =   &H00FF8080&
         Caption         =   "Sound"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   2520
         TabIndex        =   15
         Top             =   960
         Value           =   1  'Checked
         Width           =   765
      End
      Begin VB.CheckBox chkmusic 
         BackColor       =   &H00FF8080&
         Caption         =   "Music"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   1560
         TabIndex        =   14
         Top             =   960
         Value           =   1  'Checked
         Width           =   735
      End
      Begin VB.CheckBox chknpcdamage 
         BackColor       =   &H00FF8080&
         Caption         =   "Damage Above Heads"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   2520
         TabIndex        =   11
         Top             =   600
         Value           =   1  'Checked
         Width           =   1725
      End
      Begin VB.CheckBox chkplayerdamage 
         BackColor       =   &H00FF8080&
         Caption         =   "Damage Above Head"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   2520
         TabIndex        =   10
         Top             =   240
         Value           =   1  'Checked
         Width           =   1725
      End
      Begin VB.CheckBox chkbubblebar 
         BackColor       =   &H00FF8080&
         Caption         =   "Speech Bubbles"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1560
         TabIndex        =   6
         Top             =   1320
         Value           =   1  'Checked
         Width           =   1365
      End
      Begin VB.CheckBox chknpcname 
         BackColor       =   &H00FF8080&
         Caption         =   "Names"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   1560
         TabIndex        =   5
         Top             =   600
         Value           =   1  'Checked
         Width           =   765
      End
      Begin VB.CheckBox chkplayername 
         BackColor       =   &H00FF8080&
         Caption         =   "Names"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1560
         TabIndex        =   4
         Top             =   240
         Value           =   1  'Checked
         Width           =   765
      End
      Begin VB.Label Label22 
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Misc:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   360
         TabIndex        =   145
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   " X"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   5040
         TabIndex        =   46
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label18 
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Chat Data:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   360
         TabIndex        =   42
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label14 
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Sound/Music:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   360
         TabIndex        =   41
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label9 
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "NPC Data:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   360
         TabIndex        =   40
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label6 
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Player Data: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   360
         TabIndex        =   39
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblLines 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "On Screen Text Line Amount: 6"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   165
         Left            =   3120
         TabIndex        =   18
         Top             =   2640
         Width           =   1965
      End
   End
   Begin VB.PictureBox picSpellIcons 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   585
      TabIndex        =   82
      Top             =   9000
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox picPlayerSpells 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3735
      Left            =   120
      Picture         =   "frmMirage.frx":95734
      ScaleHeight     =   249
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   229
      TabIndex        =   52
      Top             =   5160
      Visible         =   0   'False
      Width           =   3435
      Begin VB.PictureBox picSpellsl 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2430
         Left            =   405
         ScaleHeight     =   162
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   177
         TabIndex        =   56
         Top             =   945
         Width           =   2655
         Begin VB.PictureBox picSpellssss 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF8080&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   3090
            Left            =   75
            ScaleHeight     =   206
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   167
            TabIndex        =   57
            Top             =   60
            Width           =   2505
            Begin VB.PictureBox picSpell 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   17
               Left            =   720
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   77
               Top             =   2520
               Width           =   480
            End
            Begin VB.PictureBox picSpell 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   16
               Left            =   120
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   76
               Top             =   2520
               Width           =   480
            End
            Begin VB.PictureBox picSpell 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   15
               Left            =   1920
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   75
               Top             =   1920
               Width           =   480
            End
            Begin VB.PictureBox picSpell 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   5
               Left            =   720
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   74
               Top             =   720
               Width           =   480
            End
            Begin VB.PictureBox picSpell 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   4
               Left            =   120
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   73
               Top             =   720
               Width           =   480
            End
            Begin VB.PictureBox picSpell 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   3
               Left            =   1920
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   72
               Top             =   120
               Width           =   480
            End
            Begin VB.PictureBox picSpell 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   2
               Left            =   1320
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   71
               Top             =   120
               Width           =   480
            End
            Begin VB.PictureBox picSpell 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   1
               Left            =   720
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   70
               Top             =   120
               Width           =   480
            End
            Begin VB.PictureBox picSpell 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   0
               Left            =   120
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   69
               Top             =   120
               Width           =   480
            End
            Begin VB.PictureBox picSpell 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   18
               Left            =   1320
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   68
               Top             =   2520
               Width           =   480
            End
            Begin VB.PictureBox picSpell 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   19
               Left            =   1920
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   67
               Top             =   2520
               Width           =   480
            End
            Begin VB.PictureBox picSpell 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   6
               Left            =   1320
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   66
               Top             =   720
               Width           =   480
            End
            Begin VB.PictureBox picSpell 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   7
               Left            =   1920
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   65
               Top             =   720
               Width           =   480
            End
            Begin VB.PictureBox picSpell 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   8
               Left            =   120
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   64
               Top             =   1320
               Width           =   480
            End
            Begin VB.PictureBox picSpell 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   9
               Left            =   720
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   63
               Top             =   1320
               Width           =   480
            End
            Begin VB.PictureBox picSpell 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   10
               Left            =   1320
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   62
               Top             =   1335
               Width           =   480
            End
            Begin VB.PictureBox picSpell 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   11
               Left            =   1905
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   61
               Top             =   1320
               Width           =   480
            End
            Begin VB.PictureBox picSpell 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   12
               Left            =   120
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   60
               Top             =   1920
               Width           =   480
            End
            Begin VB.PictureBox picSpell 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   13
               Left            =   720
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   59
               Top             =   1920
               Width           =   480
            End
            Begin VB.PictureBox picSpell 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   14
               Left            =   1320
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   58
               Top             =   1920
               Width           =   480
            End
            Begin VB.Shape shpMem 
               BorderColor     =   &H0000FF00&
               BorderWidth     =   3
               Height          =   540
               Left            =   15
               Top             =   45
               Visible         =   0   'False
               Width           =   540
            End
            Begin VB.Shape shpSel 
               BorderColor     =   &H000000FF&
               BorderWidth     =   2
               Height          =   525
               Left            =   105
               Top             =   105
               Visible         =   0   'False
               Width           =   525
            End
         End
      End
      Begin VB.PictureBox picUp 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   2400
         Picture         =   "frmMirage.frx":CD0C6
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   55
         Top             =   555
         Width           =   270
      End
      Begin VB.PictureBox picDown 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   2760
         Picture         =   "frmMirage.frx":CD35E
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   54
         Top             =   555
         Width           =   270
      End
      Begin VB.VScrollBar scrlUpDown 
         Height          =   330
         Left            =   3525
         Max             =   2
         TabIndex        =   53
         Top             =   840
         Width           =   240
      End
      Begin VB.Label Label28 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Forget"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   1560
         TabIndex        =   83
         Top             =   600
         Width           =   555
      End
      Begin VB.Label lblCastSpell 
         BackStyle       =   0  'Transparent
         Height          =   390
         Left            =   435
         TabIndex        =   81
         Top             =   495
         Width           =   690
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   3060
         TabIndex        =   80
         Top             =   135
         Width           =   240
      End
      Begin VB.Label LabelCastSpell 
         BackStyle       =   0  'Transparent
         Caption         =   "Cast"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   180
         Left            =   630
         TabIndex        =   79
         Top             =   600
         Width           =   405
      End
      Begin VB.Label Label30 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  Spells"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   345
         Left            =   1320
         TabIndex        =   78
         Top             =   240
         Width           =   945
      End
   End
   Begin VB.PictureBox picInv3 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
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
      Height          =   3735
      Left            =   120
      Picture         =   "frmMirage.frx":CD5E9
      ScaleHeight     =   249
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   229
      TabIndex        =   0
      Top             =   5160
      Visible         =   0   'False
      Width           =   3435
      Begin VB.VScrollBar VScroll1 
         Height          =   330
         Left            =   3480
         Max             =   3
         TabIndex        =   45
         Top             =   1680
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox Up 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1440
         Picture         =   "frmMirage.frx":104F7B
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   13
         Top             =   480
         Width           =   270
      End
      Begin VB.PictureBox Down 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1800
         Picture         =   "frmMirage.frx":105213
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   12
         Top             =   480
         Width           =   270
      End
      Begin VB.PictureBox Picture8 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2415
         Left            =   480
         ScaleHeight     =   161
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   169
         TabIndex        =   9
         Top             =   960
         Width           =   2535
         Begin VB.PictureBox Picture9 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF8080&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   3855
            Left            =   0
            ScaleHeight     =   257
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   169
            TabIndex        =   84
            Top             =   0
            Width           =   2535
            Begin VB.PictureBox picInv 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   8
               Left            =   120
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   108
               Top             =   1320
               Width           =   480
            End
            Begin VB.PictureBox picInv 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   7
               Left            =   1920
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   107
               Top             =   720
               Width           =   480
            End
            Begin VB.PictureBox picInv 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   6
               Left            =   1320
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   106
               Top             =   720
               Width           =   480
            End
            Begin VB.PictureBox picInv 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   5
               Left            =   720
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   105
               Top             =   720
               Width           =   480
            End
            Begin VB.PictureBox picInv 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   4
               Left            =   120
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   104
               Top             =   720
               Width           =   480
            End
            Begin VB.PictureBox picInv 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   3
               Left            =   1920
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   103
               Top             =   120
               Width           =   480
            End
            Begin VB.PictureBox picInv 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   2
               Left            =   1320
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   102
               Top             =   120
               Width           =   480
            End
            Begin VB.PictureBox picInv 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   1
               Left            =   720
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   101
               Top             =   120
               Width           =   480
            End
            Begin VB.PictureBox picInv 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   0
               Left            =   120
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   100
               Top             =   120
               Width           =   480
            End
            Begin VB.PictureBox picInv 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   9
               Left            =   720
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   99
               Top             =   1320
               Width           =   480
            End
            Begin VB.PictureBox picInv 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   10
               Left            =   1320
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   98
               Top             =   1320
               Width           =   480
            End
            Begin VB.PictureBox picInv 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   11
               Left            =   1920
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   97
               Top             =   1320
               Width           =   480
            End
            Begin VB.PictureBox picInv 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   12
               Left            =   120
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   96
               Top             =   1920
               Width           =   480
            End
            Begin VB.PictureBox picInv 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   13
               Left            =   720
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   95
               Top             =   1920
               Width           =   480
            End
            Begin VB.PictureBox picInv 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   14
               Left            =   1320
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   94
               Top             =   1920
               Width           =   480
            End
            Begin VB.PictureBox picInv 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   15
               Left            =   1920
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   93
               Top             =   1920
               Width           =   480
            End
            Begin VB.PictureBox picInv 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   16
               Left            =   120
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   92
               Top             =   2520
               Width           =   480
            End
            Begin VB.PictureBox picInv 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   17
               Left            =   720
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   91
               Top             =   2520
               Width           =   480
            End
            Begin VB.PictureBox picInv 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   18
               Left            =   1320
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   90
               Top             =   2520
               Width           =   480
            End
            Begin VB.PictureBox picInv 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   19
               Left            =   1920
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   89
               Top             =   2520
               Width           =   480
            End
            Begin VB.PictureBox picInv 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   20
               Left            =   120
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   88
               Top             =   3120
               Width           =   480
            End
            Begin VB.PictureBox picInv 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   21
               Left            =   720
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   87
               Top             =   3120
               Width           =   480
            End
            Begin VB.PictureBox picInv 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   22
               Left            =   1320
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   86
               Top             =   3120
               Width           =   480
            End
            Begin VB.PictureBox picInv 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   23
               Left            =   1920
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   85
               Top             =   3120
               Width           =   480
            End
            Begin VB.Shape EquipS 
               BorderColor     =   &H0000FFFF&
               BorderWidth     =   3
               Height          =   540
               Index           =   9
               Left            =   0
               Top             =   0
               Width           =   540
            End
            Begin VB.Shape EquipS 
               BorderColor     =   &H0000FFFF&
               BorderWidth     =   3
               Height          =   540
               Index           =   8
               Left            =   0
               Top             =   0
               Width           =   540
            End
            Begin VB.Shape EquipS 
               BorderColor     =   &H0000FFFF&
               BorderWidth     =   3
               Height          =   540
               Index           =   7
               Left            =   0
               Top             =   0
               Width           =   540
            End
            Begin VB.Shape EquipS 
               BorderColor     =   &H0000FFFF&
               BorderWidth     =   3
               Height          =   540
               Index           =   6
               Left            =   0
               Top             =   0
               Width           =   540
            End
            Begin VB.Shape EquipS 
               BorderColor     =   &H0000FFFF&
               BorderWidth     =   3
               Height          =   540
               Index           =   5
               Left            =   0
               Top             =   0
               Width           =   540
            End
            Begin VB.Shape EquipS 
               BorderColor     =   &H0000FFFF&
               BorderWidth     =   3
               Height          =   540
               Index           =   4
               Left            =   0
               Top             =   0
               Width           =   540
            End
            Begin VB.Shape SelectedItem 
               BorderColor     =   &H000000FF&
               BorderWidth     =   2
               Height          =   525
               Left            =   105
               Top             =   105
               Width           =   525
            End
            Begin VB.Shape EquipS 
               BorderColor     =   &H0000FFFF&
               BorderWidth     =   3
               Height          =   540
               Index           =   3
               Left            =   0
               Top             =   0
               Width           =   540
            End
            Begin VB.Shape EquipS 
               BorderColor     =   &H0000FFFF&
               BorderWidth     =   3
               Height          =   540
               Index           =   2
               Left            =   0
               Top             =   120
               Width           =   540
            End
            Begin VB.Shape EquipS 
               BorderColor     =   &H0000FFFF&
               BorderWidth     =   3
               Height          =   540
               Index           =   1
               Left            =   -360
               Top             =   120
               Width           =   540
            End
            Begin VB.Shape EquipS 
               BorderColor     =   &H0000FFFF&
               BorderWidth     =   3
               Height          =   540
               Index           =   0
               Left            =   0
               Top             =   0
               Width           =   540
            End
         End
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   " X"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   3120
         TabIndex        =   48
         Top             =   120
         Width           =   375
      End
      Begin VB.Label lblDropItem 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Drop Item"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   2280
         TabIndex        =   2
         Top             =   480
         Width           =   795
      End
      Begin VB.Label lblUseItem 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Use Item"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   480
         TabIndex        =   1
         Top             =   480
         Width           =   690
      End
   End
   Begin VB.PictureBox picEquip 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3780
      Left            =   120
      Picture         =   "frmMirage.frx":10549E
      ScaleHeight     =   252
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   229
      TabIndex        =   8
      Top             =   1080
      Visible         =   0   'False
      Width           =   3435
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   555
         Left            =   240
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   35
         TabIndex        =   50
         Top             =   3000
         Width           =   555
         Begin VB.PictureBox HandsImage 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   15
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   51
            Top             =   15
            Width           =   495
         End
      End
      Begin VB.PictureBox AmuletImage2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   555
         Left            =   2640
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   35
         TabIndex        =   37
         Top             =   840
         Width           =   555
         Begin VB.PictureBox AmuletImage 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   15
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   38
            Top             =   15
            Width           =   495
         End
      End
      Begin VB.PictureBox Picture14 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   555
         Left            =   840
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   35
         TabIndex        =   35
         Top             =   1440
         Width           =   555
         Begin VB.PictureBox GlovesImage 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   15
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   36
            Top             =   15
            Width           =   495
         End
      End
      Begin VB.PictureBox Picture12 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   555
         Left            =   840
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   35
         TabIndex        =   33
         Top             =   2040
         Width           =   555
         Begin VB.PictureBox Ring1Image 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   15
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   34
            Top             =   15
            Width           =   495
         End
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   555
         Left            =   2040
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   35
         TabIndex        =   31
         Top             =   2040
         Width           =   555
         Begin VB.PictureBox Ring2Image 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   15
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   32
            Top             =   15
            Width           =   495
         End
      End
      Begin VB.PictureBox Picture10 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   555
         Left            =   1440
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   35
         TabIndex        =   29
         Top             =   2760
         Width           =   555
         Begin VB.PictureBox BootsImage 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   15
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   30
            Top             =   15
            Width           =   495
         End
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   555
         Left            =   1440
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   35
         TabIndex        =   27
         Top             =   2160
         Width           =   555
         Begin VB.PictureBox LegsImage 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   15
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   28
            Top             =   15
            Width           =   495
         End
      End
      Begin VB.PictureBox Picture6 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   555
         Left            =   240
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   35
         TabIndex        =   25
         Top             =   840
         Width           =   555
         Begin VB.PictureBox WeaponImage 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   15
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   26
            Top             =   15
            Width           =   495
         End
      End
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   555
         Left            =   1440
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   35
         TabIndex        =   23
         Top             =   1560
         Width           =   555
         Begin VB.PictureBox ArmorImage 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   15
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   24
            Top             =   15
            Width           =   495
         End
      End
      Begin VB.PictureBox Picture7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   555
         Left            =   2040
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   35
         TabIndex        =   21
         Top             =   1440
         Width           =   555
         Begin VB.PictureBox ShieldImage 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   15
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   22
            Top             =   15
            Width           =   495
         End
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   555
         Left            =   1440
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   35
         TabIndex        =   19
         Top             =   960
         Width           =   555
         Begin VB.PictureBox HelmetImage 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   15
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   20
            Top             =   15
            Width           =   495
         End
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   " X"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   3120
         TabIndex        =   47
         Top             =   120
         Width           =   375
      End
   End
   Begin VB.Timer tmrSnowDrop 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   6120
      Top             =   8400
   End
   Begin VB.Timer tmrRainDrop 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5160
      Top             =   8400
   End
   Begin VB.PictureBox ScreenShot 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   11880
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   16
      Top             =   9120
      Visible         =   0   'False
      Width           =   615
   End
   Begin MSWinsockLib.Winsock Socket 
      Left            =   5640
      Top             =   8400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.PictureBox picScreen 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
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
      Height          =   6675
      Left            =   1320
      ScaleHeight     =   443
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   637
      TabIndex        =   7
      Top             =   0
      Width           =   9585
      Begin VB.Timer tmrHunger 
         Enabled         =   0   'False
         Interval        =   65000
         Left            =   0
         Top             =   0
      End
      Begin VB.PictureBox picIntroGreeting 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
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
         Height          =   3435
         Left            =   3600
         Picture         =   "frmMirage.frx":13CE30
         ScaleHeight     =   229
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   365
         TabIndex        =   213
         Top             =   2760
         Width           =   5475
         Begin VB.CommandButton Command4 
            Caption         =   "Begin Playing The Chaos Engine"
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
            Left            =   1440
            TabIndex        =   214
            Top             =   2760
            Width           =   2895
         End
         Begin VB.Label Label42 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Welcome To The Chao Engine"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   300
            Left            =   1200
            TabIndex        =   218
            Top             =   240
            Width           =   3270
         End
         Begin VB.Label Label43 
            BackStyle       =   0  'Transparent
            Height          =   375
            Left            =   5040
            TabIndex        =   217
            Top             =   0
            Width           =   375
         End
         Begin VB.Label Label47 
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   $"frmMirage.frx":1660D4
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   810
            Left            =   1320
            TabIndex        =   216
            Top             =   720
            Width           =   3135
         End
         Begin VB.Label Label41 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00404040&
            Caption         =   "3rd Party Programs Are Prohibited"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   270
            Left            =   840
            TabIndex        =   215
            Top             =   1680
            Width           =   3885
         End
      End
      Begin VB.PictureBox picStatus 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000080&
         ForeColor       =   &H80000008&
         Height          =   1335
         Left            =   6120
         ScaleHeight     =   87
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   207
         TabIndex        =   200
         Top             =   120
         Visible         =   0   'False
         Width           =   3135
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "FOOD:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080C0FF&
            Height          =   225
            Left            =   -120
            TabIndex        =   211
            Top             =   1080
            Width           =   780
         End
         Begin VB.Label Label33 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "SP:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080C0FF&
            Height          =   195
            Left            =   120
            TabIndex        =   204
            Top             =   600
            Width           =   255
         End
         Begin VB.Label Label31 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "EXP:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   195
            Left            =   120
            TabIndex        =   203
            Top             =   840
            Width           =   345
         End
         Begin VB.Label Label29 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "MP:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFF80&
            Height          =   195
            Left            =   120
            TabIndex        =   202
            Top             =   360
            Width           =   300
         End
         Begin VB.Label Label32 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "HP:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   120
            TabIndex        =   201
            Top             =   120
            Width           =   270
         End
         Begin VB.Label lblEXP 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0000C000&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   180
            Left            =   480
            TabIndex        =   207
            Top             =   840
            Width           =   2385
         End
         Begin VB.Shape shpTNL 
            BackColor       =   &H000000FF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            FillColor       =   &H000000FF&
            Height          =   165
            Left            =   495
            Top             =   855
            Width           =   2370
         End
         Begin VB.Label lblMP 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00CB884B&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   180
            Left            =   480
            TabIndex        =   206
            Top             =   360
            Width           =   2385
         End
         Begin VB.Shape shpMP 
            BackColor       =   &H00CB884B&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            Height          =   165
            Left            =   495
            Top             =   375
            Width           =   2370
         End
         Begin VB.Label lblHP 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0000C000&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   180
            Left            =   495
            TabIndex        =   205
            Top             =   120
            Width           =   2385
         End
         Begin VB.Shape shpHP 
            BackColor       =   &H0000C000&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            Height          =   165
            Left            =   510
            Top             =   135
            Width           =   2370
         End
         Begin VB.Label lblSP 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0000C000&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   180
            Left            =   480
            TabIndex        =   208
            Top             =   600
            Width           =   2385
         End
         Begin VB.Shape shpSP 
            BackColor       =   &H00FFFF80&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00FFFF80&
            Height          =   165
            Left            =   495
            Top             =   615
            Width           =   2370
         End
         Begin VB.Label lblHunger 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0000C000&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   180
            Left            =   600
            TabIndex        =   212
            Top             =   1080
            Width           =   2385
         End
         Begin VB.Shape shpHunger 
            BackColor       =   &H0000FFFF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            Height          =   165
            Left            =   600
            Top             =   1095
            Width           =   2385
         End
      End
      Begin VB.PictureBox picGuildAdmin 
         Appearance      =   0  'Flat
         BackColor       =   &H00000080&
         ForeColor       =   &H80000008&
         Height          =   2775
         Left            =   120
         ScaleHeight     =   2745
         ScaleWidth      =   2985
         TabIndex        =   188
         Top             =   1080
         Visible         =   0   'False
         Width           =   3015
         Begin VB.CommandButton cmdDisown 
            Appearance      =   0  'Flat
            BackColor       =   &H00789298&
            Caption         =   "Disown"
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
            Left            =   1560
            TabIndex        =   199
            Top             =   2400
            Width           =   1095
         End
         Begin VB.CommandButton cmdAccess 
            Appearance      =   0  'Flat
            BackColor       =   &H00789298&
            Caption         =   "Change Access"
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
            Left            =   360
            TabIndex        =   198
            Top             =   2400
            Width           =   1095
         End
         Begin VB.CommandButton cmdTrainee 
            Appearance      =   0  'Flat
            BackColor       =   &H00789298&
            Caption         =   "Make Trainee"
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
            Left            =   1560
            TabIndex        =   197
            Top             =   2040
            Width           =   1095
         End
         Begin VB.CommandButton cmdMember 
            Appearance      =   0  'Flat
            BackColor       =   &H00789298&
            Caption         =   "Make Member"
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
            Left            =   360
            TabIndex        =   196
            Top             =   2040
            Width           =   1095
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Guild Members Online"
            Height          =   255
            Left            =   360
            TabIndex        =   195
            Top             =   1680
            Width           =   2295
         End
         Begin VB.TextBox txtAccess 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   600
            TabIndex        =   194
            Top             =   1320
            Width           =   1815
         End
         Begin VB.TextBox txtName 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   600
            TabIndex        =   192
            Top             =   720
            Width           =   1815
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Access:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   165
            Left            =   600
            TabIndex        =   193
            Top             =   1080
            Width           =   465
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Name:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   165
            Left            =   600
            TabIndex        =   191
            Top             =   480
            Width           =   420
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   " X"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   2760
            TabIndex        =   190
            Top             =   0
            Width           =   255
         End
         Begin VB.Label lblGuildName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Guild Name"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   165
            Left            =   1080
            TabIndex        =   189
            Top             =   120
            Width           =   735
         End
      End
      Begin VB.PictureBox picFriend 
         Appearance      =   0  'Flat
         BackColor       =   &H00000080&
         ForeColor       =   &H80000008&
         Height          =   3495
         Left            =   6360
         ScaleHeight     =   231
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   199
         TabIndex        =   182
         Top             =   240
         Visible         =   0   'False
         Width           =   3015
         Begin VB.TextBox txtPlayerName 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   120
            TabIndex        =   186
            Top             =   360
            Width           =   2655
         End
         Begin VB.CommandButton cmdRemove 
            Caption         =   "Remove"
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
            Left            =   1560
            TabIndex        =   185
            Top             =   600
            Width           =   1215
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "Add"
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
            TabIndex        =   184
            Top             =   600
            Width           =   1215
         End
         Begin VB.ListBox lstFriend 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF8080&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2340
            ItemData        =   "frmMirage.frx":16616F
            Left            =   120
            List            =   "frmMirage.frx":166171
            TabIndex        =   183
            Top             =   960
            Width           =   2700
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   " X"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   2760
            TabIndex        =   187
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.PictureBox picWhosOnline 
         Appearance      =   0  'Flat
         BackColor       =   &H00000080&
         ForeColor       =   &H80000008&
         Height          =   2775
         Left            =   9000
         ScaleHeight     =   2745
         ScaleWidth      =   2865
         TabIndex        =   179
         Top             =   120
         Visible         =   0   'False
         Width           =   2895
         Begin VB.ListBox lstOnline 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF8080&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2340
            ItemData        =   "frmMirage.frx":166173
            Left            =   120
            List            =   "frmMirage.frx":166175
            TabIndex        =   180
            Top             =   240
            Width           =   2630
         End
         Begin VB.Label Label23 
            BackStyle       =   0  'Transparent
            Caption         =   " X"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   2640
            TabIndex        =   181
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.PictureBox itmDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H00000080&
         ForeColor       =   &H80000008&
         Height          =   4095
         Left            =   8520
         ScaleHeight     =   271
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   223
         TabIndex        =   165
         Top             =   120
         Visible         =   0   'False
         Width           =   3375
         Begin VB.Label descAS 
            Alignment       =   2  'Center
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "Attack Speed:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   720
            TabIndex        =   219
            Top             =   2760
            Width           =   1815
         End
         Begin VB.Label descPR 
            Alignment       =   2  'Center
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "Price:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   720
            TabIndex        =   210
            Top             =   2520
            Width           =   1815
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   " X"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   3120
            TabIndex        =   178
            Top             =   0
            Width           =   255
         End
         Begin VB.Label desc 
            Alignment       =   2  'Center
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   735
            Left            =   240
            TabIndex        =   177
            Top             =   3240
            Width           =   3015
         End
         Begin VB.Label Label10 
            Alignment       =   2  'Center
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "-Description-"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFF00&
            Height          =   210
            Left            =   720
            TabIndex        =   176
            Top             =   3000
            Width           =   1815
         End
         Begin VB.Label descMS 
            Alignment       =   2  'Center
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "Magi: XXXXX Speed: XXXX"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   600
            TabIndex        =   175
            Top             =   2280
            Width           =   2055
         End
         Begin VB.Label descSD 
            Alignment       =   2  'Center
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "Str: XXXX Def: XXXXX"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   600
            TabIndex        =   174
            Top             =   2040
            Width           =   2055
         End
         Begin VB.Label descHpMp 
            Alignment       =   2  'Center
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "HP: XXXX MP: XXXX SP: XXXX"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   600
            TabIndex        =   173
            Top             =   1800
            Width           =   2175
         End
         Begin VB.Label Label15 
            Alignment       =   2  'Center
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "-Add-"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   210
            Left            =   720
            TabIndex        =   172
            Top             =   1560
            Width           =   1815
         End
         Begin VB.Label descMagic 
            Alignment       =   2  'Center
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "Magic"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   720
            TabIndex        =   171
            Top             =   1320
            Width           =   1815
         End
         Begin VB.Label descSpeed 
            Alignment       =   2  'Center
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "Speed"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   720
            TabIndex        =   170
            Top             =   1080
            Width           =   1815
         End
         Begin VB.Label descDef 
            Alignment       =   2  'Center
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "Defence"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   720
            TabIndex        =   169
            Top             =   840
            Width           =   1815
         End
         Begin VB.Label descStr 
            Alignment       =   2  'Center
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "Strength"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   720
            TabIndex        =   168
            Top             =   600
            Width           =   1815
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "-Requirements-"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   210
            Left            =   720
            TabIndex        =   167
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label descName 
            Alignment       =   2  'Center
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "Name"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   360
            TabIndex        =   166
            Top             =   120
            Width           =   2535
         End
      End
      Begin VB.PictureBox picSpellInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H00000080&
         ForeColor       =   &H80000008&
         Height          =   1695
         Left            =   8520
         ScaleHeight     =   1665
         ScaleWidth      =   3345
         TabIndex        =   157
         Top             =   120
         Visible         =   0   'False
         Width           =   3375
         Begin VB.Label Label26 
            BackStyle       =   0  'Transparent
            Caption         =   " X"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   3120
            TabIndex        =   164
            Top             =   0
            Width           =   255
         End
         Begin VB.Label descSELE 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Element: "
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   240
            TabIndex        =   163
            Top             =   1440
            Width           =   3000
         End
         Begin VB.Label descSMana 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Mana Cost: "
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   240
            TabIndex        =   162
            Top             =   1200
            Width           =   3000
         End
         Begin VB.Label descSLevel 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Level: "
            ForeColor       =   &H00FFFFFF&
            Height          =   270
            Left            =   240
            TabIndex        =   161
            Top             =   960
            Width           =   2910
         End
         Begin VB.Label descSClass 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Class: "
            ForeColor       =   &H00FFFFFF&
            Height          =   270
            Left            =   240
            TabIndex        =   160
            Top             =   720
            Width           =   2880
         End
         Begin VB.Label Label27 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "(Requirements)"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   1200
            TabIndex        =   159
            Top             =   480
            Width           =   1110
         End
         Begin VB.Label descSName 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Spell Name"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   270
            Left            =   960
            TabIndex        =   158
            Top             =   120
            Width           =   1500
         End
      End
      Begin VB.PictureBox picHotBar1 
         Appearance      =   0  'Flat
         BackColor       =   &H00000080&
         ForeColor       =   &H80000008&
         Height          =   525
         Left            =   120
         ScaleHeight     =   495
         ScaleWidth      =   5895
         TabIndex        =   147
         Top             =   360
         Width           =   5925
         Begin VB.PictureBox Picture21 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   4800
            Picture         =   "frmMirage.frx":166177
            ScaleHeight     =   465
            ScaleWidth      =   465
            TabIndex        =   209
            Top             =   0
            Width           =   495
         End
         Begin VB.PictureBox Picture11 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   2400
            Picture         =   "frmMirage.frx":166DB9
            ScaleHeight     =   465
            ScaleWidth      =   465
            TabIndex        =   156
            Top             =   0
            Width           =   495
         End
         Begin VB.PictureBox Picture28 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   600
            Picture         =   "frmMirage.frx":1679FB
            ScaleHeight     =   465
            ScaleWidth      =   465
            TabIndex        =   155
            Top             =   0
            Width           =   495
         End
         Begin VB.PictureBox Picture22 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   5400
            Picture         =   "frmMirage.frx":16863D
            ScaleHeight     =   465
            ScaleWidth      =   465
            TabIndex        =   154
            Top             =   0
            Width           =   495
         End
         Begin VB.PictureBox Picture20 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   4200
            Picture         =   "frmMirage.frx":16927F
            ScaleHeight     =   465
            ScaleWidth      =   480
            TabIndex        =   153
            Top             =   0
            Width           =   510
         End
         Begin VB.PictureBox Picture19 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   1200
            Picture         =   "frmMirage.frx":169EC1
            ScaleHeight     =   465
            ScaleWidth      =   465
            TabIndex        =   152
            Top             =   0
            Width           =   495
         End
         Begin VB.PictureBox Picture16 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   3000
            Picture         =   "frmMirage.frx":16AB03
            ScaleHeight     =   465
            ScaleWidth      =   480
            TabIndex        =   151
            Top             =   0
            Width           =   510
         End
         Begin VB.PictureBox Picture17 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   3600
            Picture         =   "frmMirage.frx":16B745
            ScaleHeight     =   465
            ScaleWidth      =   480
            TabIndex        =   150
            Top             =   0
            Width           =   510
         End
         Begin VB.PictureBox Picture18 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   1800
            Picture         =   "frmMirage.frx":16C387
            ScaleHeight     =   465
            ScaleWidth      =   465
            TabIndex        =   149
            Top             =   0
            Width           =   495
         End
         Begin VB.PictureBox Picture27 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   0
            Picture         =   "frmMirage.frx":16CFC9
            ScaleHeight     =   465
            ScaleWidth      =   465
            TabIndex        =   148
            Top             =   0
            Width           =   495
         End
      End
      Begin VB.PictureBox picExitOptions 
         Appearance      =   0  'Flat
         BackColor       =   &H00000080&
         ForeColor       =   &H80000008&
         Height          =   3495
         Left            =   4080
         ScaleHeight     =   3465
         ScaleWidth      =   4665
         TabIndex        =   109
         Top             =   2760
         Visible         =   0   'False
         Width           =   4695
         Begin VB.CommandButton cmdOptions 
            Caption         =   "Options Menu"
            Height          =   495
            Left            =   960
            TabIndex        =   116
            Top             =   2040
            Width           =   2895
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Switch Server"
            Enabled         =   0   'False
            Height          =   495
            Left            =   960
            TabIndex        =   115
            Top             =   1440
            Width           =   2895
         End
         Begin VB.CommandButton Command6 
            Caption         =   "Quit Game"
            Height          =   495
            Left            =   960
            TabIndex        =   112
            Top             =   2880
            Width           =   2895
         End
         Begin VB.CommandButton cmdReturnMain 
            Caption         =   "Switch Character"
            Enabled         =   0   'False
            Height          =   495
            Left            =   960
            TabIndex        =   111
            Top             =   840
            Width           =   2895
         End
         Begin VB.CommandButton cmdResume 
            Caption         =   "Resume Playing"
            Height          =   495
            Left            =   960
            TabIndex        =   110
            Top             =   240
            Width           =   2895
         End
      End
   End
End
Attribute VB_Name = "frmMirage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Copyright (c) 2006 Chaos Engine Source. All rights reserved.
' This code is licensed under the Chaos Engine General License.

Option Explicit

Dim SpellMemorized As Long

Private Sub Close_Click()
    Unload Me
End Sub

Private Sub cmdLeaveGuild_Click()
Dim Packet As String
    Packet = "GUILDLEAVE" & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Private Sub chkclicktomove_Click()
Call PutVar(App.Path & "\Main\Config\config.ini", "CONFIG", "MouseMovement", chkclicktomove.Value)
End Sub

Private Sub chkEmoticons_Click()
Call PutVar(App.Path & "\Main\Config\config.ini", "CONFIG", "Emoticons", chkEmoticons.Value)
End Sub

Private Sub chkFPS_Click()
Call PutVar(App.Path & "\Main\Config\config.ini", "CONFIG", "FPS", chkFPS.Value)
End Sub

Private Sub chkNight_Click()
Call PutVar(App.Path & "\Main\Config\config.ini", "CONFIG", "Night", chkNight.Value)
End Sub

Private Sub chkWeather_Click()
Call PutVar(App.Path & "\Main\Config\config.ini", "CONFIG", "Weather", chkWeather.Value)
End Sub

Private Sub Combo1_Change()
    txtMyTextBox.SetFocus
End Sub

Private Sub Combo1_Click()
On Error Resume Next
    txtMyTextBox.SetFocus
End Sub

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    txtMyTextBox.SetFocus
End Sub

Private Sub AddDef_Click()
    Call SendData("usestatpoint" & SEP_CHAR & 1 & SEP_CHAR & END_CHAR)
End Sub

Private Sub AddMagi_Click()
    Call SendData("usestatpoint" & SEP_CHAR & 2 & SEP_CHAR & END_CHAR)
End Sub

Private Sub AddSpeed_Click()
    Call SendData("usestatpoint" & SEP_CHAR & 3 & SEP_CHAR & END_CHAR)
End Sub

Private Sub AddStr_Click()
    Call SendData("usestatpoint" & SEP_CHAR & 0 & SEP_CHAR & END_CHAR)
End Sub

Private Sub chksound_Click()
    Call PutVar(App.Path & "\Main\Config\config.ini", "CONFIG", "Sound", chkSound.Value)
End Sub

Private Sub chkbubblebar_Click()
    Call PutVar(App.Path & "\Main\Config\config.ini", "CONFIG", "SpeechBubbles", chkbubblebar.Value)
End Sub

Private Sub chkEmoSound_Click()
    Call PutVar(App.Path & "\Main\Config\config.ini", "CONFIG", "EmoticonSound", chkEmoSound.Value)
End Sub

Private Sub chknpcdamage_Click()
    Call PutVar(App.Path & "\Main\Config\config.ini", "CONFIG", "NPCDamage", chknpcdamage.Value)
End Sub

Private Sub chknpcname_Click()
    Call PutVar(App.Path & "\Main\Config\config.ini", "CONFIG", "NPCName", chknpcname.Value)
End Sub

Private Sub chkplayerdamage_Click()
    Call PutVar(App.Path & "\Main\Config\config.ini", "CONFIG", "PlayerDamage", chkplayerdamage.Value)
End Sub

Private Sub chkAutoScroll_Click()
    Call PutVar(App.Path & "\Main\Config\config.ini", "CONFIG", "AutoScroll", chkAutoScroll.Value)
End Sub

Private Sub chkplayername_Click()
    Call PutVar(App.Path & "\Main\Config\config.ini", "CONFIG", "PlayerName", chkplayername.Value)
End Sub

Private Sub chkmusic_Click()
    Call PutVar(App.Path & "\Main\Config\config.ini", "CONFIG", "Music", chkmusic.Value)
    If MyIndex <= 0 Then Exit Sub
    Call PlayMidi(Trim(Map(GetPlayerMap(MyIndex)).Music))
End Sub

Private Sub cmdAdd_Click()
Dim Packet As String
    If txtPlayerName.Text <> "" Then
        Packet = "ADDFRIEND" & SEP_CHAR & txtPlayerName.Text & SEP_CHAR & END_CHAR
        Call SendData(Packet)
    End If
End Sub

Private Sub cmdLeave_Click()
Dim Packet As String
    Packet = "GUILDLEAVE" & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Private Sub cmdMember_Click()
Dim Packet As String
    Packet = "GUILDMEMBER" & SEP_CHAR & txtName.Text & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Private Sub cmdOptions_Click()
picOptions.Visible = True
End Sub

Private Sub cmdRemove_Click()
Dim Packet As String
    If txtPlayerName.Text <> "" Then
        Packet = "REMOVEFRIEND" & SEP_CHAR & txtPlayerName.Text & SEP_CHAR & END_CHAR
        Call SendData(Packet)
    End If
End Sub

Private Sub cmdResume_Click()
picExitOptions.Visible = False
End Sub

Private Sub cmdReturnMain_Click()
'Call TcpDestroy
'frmMainMenu.Visible = True
'frmMirage.Visible = False
End Sub

Private Sub Command1_Click()
    picOptions.Visible = False
End Sub

Private Sub Command2_Click()
 If GetPlayerGuildAccess(MyIndex) > 0 Then
        If frmGuildMembers.Visible = False Then
            frmGuildMembers.Visible = True
        Else
            Unload frmGuildMembers
        End If
    End If
End Sub

Private Sub Command4_Click()
frmMirage.tmrHunger.Enabled = True
picIntroGreeting.Visible = False
End Sub

Private Sub Command5_Click()
tmrPoison.Enabled = True
End Sub

Private Sub Command6_Click()
Call GameDestroy
End Sub

Private Sub form_load()
Dim I As Long
Dim Ending As String
Set mclsStyle = New clsWindowed
Set mclsStyle.Client = Me
    For I = 1 To 3
        If I = 1 Then Ending = ".gif"
        If I = 2 Then Ending = ".jpg"
        If I = 3 Then Ending = ".png"
 
    Next I
    Combo1.ListIndex = 0
    
    txtChat.SelHangingIndent = 8 ' Set hanging indent For chat box
End Sub

Private Sub form_mouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    itmDesc.Visible = False
End Sub

Private Sub Form_Terminate()
    Call GameDestroy
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call GameDestroy
End Sub

Private Sub Label13_Click()
    picGuildAdmin.Visible = False
End Sub

Private Sub Label19_Click()
    picPlayerSpells.Visible = False
End Sub

Private Sub Label21_Click()
    picEquip.Visible = False
End Sub

Private Sub Label20_Click()
    picOptions.Visible = False
End Sub

Private Sub Label22_Click()
    picStat.Visible = False
End Sub

Private Sub Label23_Click()
    picWhosOnline.Visible = False
End Sub

Private Sub Label24_Click()
    picInv3.Visible = False
End Sub

Private Sub Label25_Click()
    
End Sub

Private Sub Label26_Click()
picSpellInfo.Visible = False
End Sub

Private Sub Label28_Click()
If Player(MyIndex).Spell(SpellIndex) > 0 Then
If GetTickCount > Player(MyIndex).AttackTimer + 1000 Then
If MsgBox("Are you sure you want to forget this spell?", vbYesNo, "Forget Spell") = vbNo Then Exit Sub
Call SendData("forgetspell" & SEP_CHAR & SpellIndex & SEP_CHAR & END_CHAR)
'picPlayerSpells.Visible = False
End If
Else
Call AddText("No spell here.", BrightRed)
End If
End Sub

Private Sub Label3_Click()
    picStatus.Visible = False
End Sub

Private Sub Label38_Click()
picStat.Visible = False
End Sub

Private Sub Label39_Click()

End Sub

Private Sub Label4_Click()
    picFriend.Visible = False
End Sub

Private Sub Label40_Click()
frmMirage.picChatBox.Visible = False
End Sub

Private Sub Label6_Click()
    Call SendData("getstats" & SEP_CHAR & END_CHAR)
End Sub

Private Sub Label7_Click()
    itmDesc.Visible = False
End Sub

Private Sub lblCloseOnline_Click()
Call SendOnlineList
picWhosOnline.Visible = False
End Sub

Private Sub lblClosePicGuildAdmin_Click()
picGuildAdmin.Visible = False
End Sub

Private Sub lblCastSpell_Click()
If Player(MyIndex).Spell(SpellIndex) > 0 Then
        If GetTickCount > Player(MyIndex).AttackTimer + 1000 Then
            If Player(MyIndex).MovingH = 0 Then
                Call SendData("cast" & SEP_CHAR & SpellIndex & SEP_CHAR & END_CHAR)
                Player(MyIndex).Attacking = 1
                Player(MyIndex).AttackTimer = GetTickCount
                Player(MyIndex).CastedSpell = YES
            Else
                Call AddText("Cannot cast while walking!", BrightRed)
            End If
        End If
    Else
        Call AddText("No spell here.", BrightRed)
    End If
End Sub

Private Sub lblCht_Click()

End Sub

Private Sub lblEqp_Click()
    picEquip.Visible = True
    Call UpdateVisInv
End Sub

Private Sub lblInv_Click()
    
End Sub

Private Sub lblOpt_Click()
    picOptions.Visible = True
End Sub

Private Sub lblQit_Click()
    Call GameDestroy
End Sub

Private Sub lblLeaveGuild_Click()
Dim Packet As String
    Packet = "GUILDLEAVE" & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Private Sub lblSpl_Click()
    
End Sub

Private Sub lblWho_Click()
    
End Sub

Private Sub lstFriend_DblClick()
    Call SendData("playerchat" & SEP_CHAR & Trim(lstFriend.Text) & SEP_CHAR & END_CHAR)
End Sub

Private Sub lstOnline_DblClick()
    Call SendData("playerchat" & SEP_CHAR & Trim(lstOnline.Text) & SEP_CHAR & END_CHAR)
End Sub

Private Sub picDown_Click()
If scrlUpDown.Value = 2 Then Exit Sub
    scrlUpDown.Value = scrlUpDown.Value + 1
    picSpellssss.Top = scrlUpDown.Value * -PIC_Y
End Sub

Private Sub picInv_DblClick(Index As Integer)
Dim d As Long

If Player(MyIndex).Inv(Inventory).Num <= 0 Then Exit Sub

Call SendUseItem(Inventory)

For d = 1 To MAX_INV
    If Player(MyIndex).Inv(d).Num > 0 Then
        If Item(GetPlayerInvItemNum(MyIndex, d)).Type = ITEM_TYPE_POTIONADDMP Or ITEM_TYPE_POTIONADDHP Or ITEM_TYPE_POTIONADDSP Or ITEM_TYPE_POTIONSUBHP Or ITEM_TYPE_POTIONSUBMP Or ITEM_TYPE_POTIONSUBSP Then
            picInv(d - 1).Picture = LoadPicture()
        End If
    End If
Next d
Call UpdateVisInv
End Sub

Private Sub picInv_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Inventory = Index + 1
    frmMirage.SelectedItem.Top = frmMirage.picInv(Inventory - 1).Top - 1
    frmMirage.SelectedItem.Left = frmMirage.picInv(Inventory - 1).Left - 1
    
    If Button = 1 Then
        Call UpdateVisInv
    ElseIf Button = 2 Then
        Call DropItems
    End If
End Sub

Private Sub picInv_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim d As Long
d = Index

    If Player(MyIndex).Inv(d + 1).Num > 0 Then
        If Item(GetPlayerInvItemNum(MyIndex, d + 1)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(MyIndex, d + 1)).Stackable = 1 Then
            descName.Caption = Trim(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Name) & " (" & GetPlayerInvItemValue(MyIndex, d + 1) & ")"
            descName.ForeColor = Trim(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Rarity)
        Else
            If GetPlayerWeaponSlot(MyIndex) = d + 1 Then
                descName.Caption = Trim(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Name) & " (worn)"
                descName.ForeColor = Trim(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Rarity)
            ElseIf GetPlayerArmorSlot(MyIndex) = d + 1 Then
                descName.Caption = Trim(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Name) & " (worn)"
                descName.ForeColor = Trim(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Rarity)
            ElseIf GetPlayerHelmetSlot(MyIndex) = d + 1 Then
                descName.Caption = Trim(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Name) & " (worn)"
                descName.ForeColor = Trim(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Rarity)
            ElseIf GetPlayerShieldSlot(MyIndex) = d + 1 Then
                descName.Caption = Trim(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Name) & " (worn)"
                descName.ForeColor = Trim(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Rarity)
            ElseIf GetPlayerLegsSlot(MyIndex) = d + 1 Then
                descName.Caption = Trim(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Name) & " (worn)"
                descName.ForeColor = Trim(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Rarity)
            ElseIf GetPlayerBootsSlot(MyIndex) = d + 1 Then
                descName.Caption = Trim(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Name) & " (worn)"
                descName.ForeColor = Trim(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Rarity)
            ElseIf GetPlayerGlovesSlot(MyIndex) = d + 1 Then
                descName.Caption = Trim(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Name) & " (worn)"
                descName.ForeColor = Trim(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Rarity)
            ElseIf GetPlayerRing1Slot(MyIndex) = d + 1 Then
                descName.Caption = Trim(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Name) & " (worn)"
                descName.ForeColor = Trim(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Rarity)
            ElseIf GetPlayerRing2Slot(MyIndex) = d + 1 Then
                descName.Caption = Trim(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Name) & " (worn)"
                descName.ForeColor = Trim(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Rarity)
            ElseIf GetPlayerAmuletSlot(MyIndex) = d + 1 Then
                descName.Caption = Trim(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Name) & " (worn)"
                descName.ForeColor = Trim(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Rarity)
            Else
                descName.Caption = Trim(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Name)
                descName.ForeColor = Trim(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Rarity)
            End If
        End If
        descStr.Caption = Item(GetPlayerInvItemNum(MyIndex, d + 1)).StrReq & " Strength"
        descDef.Caption = Item(GetPlayerInvItemNum(MyIndex, d + 1)).DefReq & " Defence"
        descSpeed.Caption = Item(GetPlayerInvItemNum(MyIndex, d + 1)).SpeedReq & " Speed"
        descMagic.Caption = Item(GetPlayerInvItemNum(MyIndex, d + 1)).MagicReq & " Magic"
        descHpMp.Caption = "HP: " & Item(GetPlayerInvItemNum(MyIndex, d + 1)).AddHP & " MP: " & Item(GetPlayerInvItemNum(MyIndex, d + 1)).AddMP & " SP: " & Item(GetPlayerInvItemNum(MyIndex, d + 1)).AddSP
        descSD.Caption = "Str: " & Item(GetPlayerInvItemNum(MyIndex, d + 1)).AddStr & " Def: " & Item(GetPlayerInvItemNum(MyIndex, d + 1)).AddDef
        descMS.Caption = "Magi: " & Item(GetPlayerInvItemNum(MyIndex, d + 1)).AddMagi & " Speed: " & Item(GetPlayerInvItemNum(MyIndex, d + 1)).AddSpeed
        descAS.Caption = "Attack Speed: " & Item(GetPlayerInvItemNum(MyIndex, d + 1)).AttackSpeed
        desc.Caption = Trim(Item(GetPlayerInvItemNum(MyIndex, d + 1)).desc)
        descPR.Caption = "Item Price: " & Item(GetPlayerInvItemNum(MyIndex, d + 1)).Price & " Gold"
        
        itmDesc.Visible = True
        Call itmDesc.ZOrder(0)
    Else
        itmDesc.Visible = False
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Dim d As Long, I As Long
Dim ii As Long

    Call CheckInput(0, KeyCode, Shift)
If KeyCode = vbKeyF1 Then
    If frmMirage.picHotBar1.Visible = False Then
        frmMirage.picHotBar1.Visible = True
    Else
        frmMirage.picHotBar1.Visible = False
    End If
  End If
  
   If KeyCode = vbKeyF2 Then
   If frmUserPanel.Visible = False Then
        frmUserPanel.Visible = True
        Else
        frmUserPanel.Visible = False
        End If
        Exit Sub
    End If
    
    If KeyCode = vbKeyF3 Then
        If Player(MyIndex).Access > 0 Then
        If frmadmin.Visible = False Then
            frmadmin.Visible = True
            Else
            frmadmin.Visible = False
            End If
        End If
    End If
    
    ' The Guild Creator
    If KeyCode = vbKeyF4 Then
        If Player(MyIndex).Access > 0 Then
            frmGuild.Show vbModeless, frmMirage
        End If
    End If
    
    ' Exit Options
    If KeyCode = vbKeyEscape Then
        frmMirage.picExitOptions.Visible = True
    End If

    ' The Guild Maker
    If KeyCode = vbKeyF5 Then
      If Player(MyIndex).Guildaccess = 5 Then
        frmMirage.picGuildAdmin.Visible = True
        frmMirage.picInv3.Visible = False
        frmMirage.picEquip.Visible = False
        frmMirage.picPlayerSpells.Visible = False
        frmMirage.picWhosOnline.Visible = False
       End If
     End If
      
    If KeyCode = vbKeyInsert Then
        If SpellMemorized > 0 Then
            If GetTickCount > Player(MyIndex).AttackTimer + 1000 Then
                If Player(MyIndex).MovingH = 0 And Player(MyIndex).MovingV = 0 Then
                    Call SendData("cast" & SEP_CHAR & SpellMemorized & SEP_CHAR & END_CHAR)
                    Player(MyIndex).Attacking = 1
                    Player(MyIndex).AttackTimer = GetTickCount
                    Player(MyIndex).CastedSpell = YES
                Else
                    Call AddText("Cannot cast while walking!", BrightRed)
                End If
            End If
        Else
            Call AddText("No spell here memorized.", BrightRed)
        End If
    Else
        Call CheckInput(0, KeyCode, Shift)
    End If
    
     If KeyCode = vbKeyF6 Then
    If MiniMap = False Then
                MiniMap = True
            Else
                MiniMap = False
            End If
            Exit Sub
        End If
    
    If KeyCode = vbKeyF11 Then
        ScreenShot.Picture = CaptureForm(frmMirage)
        I = 0
        ii = 0
        If LCase(Dir(App.Path & "\Main\Screenshots", vbDirectory)) <> "screenshots" Then
            Call MkDir(App.Path & "\Main\Screenshots")
        End If
        Do
            If ScreenFileExist("Screenshot" & I & ".bmp") = True Then
                I = I + 1
            Else
                Call SavePicture(ScreenShot.Picture, App.Path & "\Main\Screenshots\Screenshot" & I & ".bmp")
                ii = 1
            End If
           
            DoEvents
        Loop Until ii = 1
    ElseIf KeyCode = vbKeyF12 Then
        ScreenShot.Picture = CaptureArea(frmMirage, picScreen.Left, picScreen.Top, picScreen.Width, picScreen.Height)
        I = 0
        ii = 0
        If LCase(Dir(App.Path & "\Main\Screenshots", vbDirectory)) <> "screenshots" Then
            Call MkDir(App.Path & "\Main\Screenshots")
        End If
        Do
            If ScreenFileExist("Screenshot" & I & ".bmp") = True Then
                I = I + 1
            Else
                Call SavePicture(ScreenShot.Picture, App.Path & "\Main\Screenshots\Screenshot" & I & ".bmp")
                ii = 1
            End If
           
            DoEvents
        Loop Until ii = 1
    End If
    
    If KeyCode = vbKeyEnd Then
    d = GetPlayerDir(MyIndex)
    
        If Player(MyIndex).MovingH = NO And Player(MyIndex).MovingV = NO Then
            If Player(MyIndex).Dir = DIR_DOWN Then
                Call SetPlayerDir(MyIndex, DIR_LEFT)
                If d <> DIR_LEFT Then
                    Call SendPlayerDir
                End If
            ElseIf Player(MyIndex).Dir = DIR_LEFT Then
                Call SetPlayerDir(MyIndex, DIR_UP)
                If d <> DIR_UP Then
                    Call SendPlayerDir
                End If
            ElseIf Player(MyIndex).Dir = DIR_UP Then
                Call SetPlayerDir(MyIndex, DIR_RIGHT)
                If d <> DIR_RIGHT Then
                    Call SendPlayerDir
                End If
            ElseIf Player(MyIndex).Dir = DIR_RIGHT Then
                Call SetPlayerDir(MyIndex, DIR_DOWN)
                If d <> DIR_DOWN Then
                    Call SendPlayerDir
                End If
            End If
        End If
    End If
End Sub

Private Sub picOptions_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    SOffsetX = x
    SOffsetY = y
    Call picOptions.ZOrder(0)
End Sub

Private Sub picOptions_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call MovePicture(frmMirage.picOptions, Button, Shift, x, y)
End Sub

Private Sub picPlayerSpells_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    SOffsetX = x
    SOffsetY = y
    Call picPlayerSpells.ZOrder(0)
End Sub

Private Sub picPlayerSpells_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call MovePicture(frmMirage.picPlayerSpells, Button, Shift, x, y)
End Sub

Private Sub picSpell_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim s As Long
s = Index + 1

    If Player(MyIndex).Spell(s) > 0 Then
        descSName.Caption = Spell(Player(MyIndex).Spell(s)).Name
        descSLevel.Caption = "Level: " & Spell(Player(MyIndex).Spell(s)).LevelReq
        If Spell(Player(MyIndex).Spell(s)).ClassReq > 0 Then
            descSClass.Caption = "Class: " & Class(Spell(Player(MyIndex).Spell(s)).ClassReq).Name
        Else
            descSClass.Caption = "Class: All"
        End If
        descSMana.Caption = "Mana Cost: " & Spell(Player(MyIndex).Spell(s)).MPCost
        descSELE.Caption = "Element: None"
        
        picSpellInfo.Visible = True
        Call picSpellInfo.ZOrder(0)
    Else
        picSpellInfo.Visible = False
    End If
End Sub

Private Sub picSpellInfo_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    SOffsetX = x
    SOffsetY = y
End Sub

Private Sub picSpellInfo_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call MovePicture(frmMirage.picSpellInfo, Button, Shift, x, y)
End Sub

Private Sub picSpell_Click(Index As Integer)
SpellIndex = Index + 1
shpSel.Visible = True
frmMirage.shpSel.Top = frmMirage.picSpell(SpellIndex - 1).Top - 1
frmMirage.shpSel.Left = frmMirage.picSpell(SpellIndex - 1).Left - 1
End Sub

Private Sub picSpell_DblClick(Index As Integer)
If Player(MyIndex).Spell(SpellIndex) > 0 Then
    If SpellMemorized <> SpellIndex Then
        SpellMemorized = SpellIndex
        Call AddText("Successfully memorized spell!", White)
        frmMirage.shpMem.Visible = True
        frmMirage.shpMem.Top = frmMirage.picSpell(SpellIndex - 1).Top - 2
        frmMirage.shpMem.Left = frmMirage.picSpell(SpellIndex - 1).Left - 2
    End If
Else
    Call AddText("No spell to memorize.", BrightRed)
End If
End Sub

Private Sub picSpellssss_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
picSpellInfo.Visible = False
End Sub

Private Sub picStatus_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    'SOffsetX = x
    'SOffsetY = y
    'Call picStatus.ZOrder(0)
End Sub

Private Sub picStatus_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call MovePicture(frmMirage.picStatus, Button, Shift, x, y)
End Sub

Private Sub picStat_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    SOffsetX = x
    SOffsetY = y
    Call picStat.ZOrder(0)
End Sub

Private Sub Picture11_Click()
If picInv3.Visible = False Then
Call UpdateVisInv
    picInv3.Visible = True
    Else
    picInv3.Visible = False
    End If
End Sub

Private Sub Picture16_Click()
picEquip.Visible = True
    Call UpdateVisInv
End Sub

Private Sub Picture17_Click()
Call SendData("spells" & SEP_CHAR & END_CHAR)
    picPlayerSpells.Visible = True
End Sub

Private Sub Picture18_Click()
If picStatus.Visible = False Then
picStatus.Visible = True
Else
picStatus.Visible = False
End If
End Sub

Private Sub Picture19_Click()
frmMirage.lblName.Caption = "Name: " & GetPlayerName(MyIndex)
    frmMirage.lblKills.Caption = "Kills: " & GetPlayerPK(MyIndex)
    frmMirage.lblAlign.Caption = "Align: " & GetPlayerAlignment(MyIndex)
    If GetPlayerAccess(MyIndex) = 0 Then
    frmMirage.lblAccess.Caption = "Access: " & "None"
    End If
    If GetPlayerAccess(MyIndex) = 1 Then
    frmMirage.lblAccess.Caption = "Access: " & "Moderator"
    End If
    If GetPlayerAccess(MyIndex) = 2 Then
    frmMirage.lblAccess.Caption = "Access: " & "Mapper"
    End If
    If GetPlayerAccess(MyIndex) = 3 Then
    frmMirage.lblAccess.Caption = "Access: " & "Developer"
    End If
    If GetPlayerAccess(MyIndex) = 4 Then
    frmMirage.lblAccess.Caption = "Access: " & "Server Owner"
    End If
    If GetPlayerClass(MyIndex) = 1 Then
    frmMirage.lblClass.Caption = "Class: " & "Warrior"
    End If
    If GetPlayerClass(MyIndex) = 2 Then
    frmMirage.lblClass.Caption = "Class: " & "Mage"
    End If
    If GetPlayerClass(MyIndex) = 3 Then
    frmMirage.lblClass.Caption = "Class: " & "Archer"
    End If
    If GetPlayerSprite(MyIndex) = 126 Then
    frmMirage.lblSex.Caption = "Sex: " & "Male"
    End If
    If GetPlayerSprite(MyIndex) = 127 Then
    frmMirage.lblSex.Caption = "Sex: " & "Female"
    End If
    frmMirage.lblGuild.Caption = GetPlayerGuild(MyIndex)
    frmMirage.lblRank.Caption = GetPlayerGuildAccess(MyIndex)
    picStat.Visible = True
End Sub

Private Sub Picture20_Click()
frmTradeSkills.Visible = True
frmTradeSkills.Timer1.Enabled = True
End Sub

Private Sub Picture21_Click()
Dim Msg, Style, Response

If GetPlayerGuildAccess(MyIndex) >= 1 Then
'Label3.Visible = True
txtAccess.Visible = False
cmdTrainee.Visible = False
cmdMember.Visible = False
txtName.Visible = False
cmdDisown.Visible = False
cmdAccess.Visible = False
If GetPlayerGuildAccess(MyIndex) >= 2 Then
txtAccess.Visible = False
cmdTrainee.Visible = False
cmdMember.Visible = False
txtName.Visible = False
cmdDisown.Visible = False
cmdAccess.Visible = False
If Player(MyIndex).Guildaccess >= 3 Then
cmdTrainee.Visible = True
cmdMember.Visible = True
txtName.Visible = True
cmdAccess.Visible = False
'Label11.Visible = True
'Label38.Visible = False
If Player(MyIndex).Guildaccess >= 4 Then
cmdAccess.Visible = True
cmdTrainee.Visible = True
cmdMember.Visible = True
txtName.Visible = True
txtAccess.Visible = True
'Label11.Visible = True
cmdDisown.Visible = True
'Label38.Visible = False
End If
End If
End If
End If

If GetPlayerGuildAccess(MyIndex) > 0 Then
picGuildAdmin.Visible = True
Else
Msg = "Your not in a Guild! Do you wish to make one?"
Style = vbYesNo + vbDefaultButton2
Response = MsgBox(Msg, Style)
If Response = vbYes Then
frmGuildDeed.Visible = True
End If
End If
End Sub

Private Sub Picture22_Click()
Dim I As Long

    For I = 1 To MAX_PLAYERS
        If IsPlaying(I) = True Then
            If MouseDownX = GetPlayerX(I) And MouseDownY = GetPlayerY(I) Then
                Call SendData("playerchat" & SEP_CHAR & GetPlayerName(I) & SEP_CHAR & END_CHAR)
                Exit Sub
            End If
        End If
    Next I
    
    Call AddText("Target a player to chat with first!", Red)
End Sub

Private Sub Picture27_Click()
frmUserPanel.lblUserName.Caption = "Name: " & "" & GetPlayerName(MyIndex)
frmUserPanel.Visible = True
End Sub

Private Sub Picture28_Click()
If picChatBox.Visible = False Then
picChatBox.Visible = True
Else
picChatBox.Visible = False
End If
End Sub

Private Sub picUp_Click()
If scrlUpDown.Value = 0 Then Exit Sub
    scrlUpDown.Value = scrlUpDown.Value - 1
    picSpellssss.Top = scrlUpDown.Value * -PIC_Y
End Sub

Private Sub picWhosOnline_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'    SOffsetX = x
   ' SOffsetY = y
   ' Call picWhosOnline.ZOrder(0)
End Sub

Private Sub picWhosOnline_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call MovePicture(frmMirage.picWhosOnline, Button, Shift, x, y)
End Sub

Private Sub picInv3_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    SOffsetX = x
    SOffsetY = y
    Call picInv3.ZOrder(0)
End Sub

Private Sub picInv3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call MovePicture(frmMirage.picInv3, Button, Shift, x, y)
End Sub

Private Sub picFriend_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'    SOffsetX = x
    'SOffsetY = y
    'Call picFriend.ZOrder(0)
End Sub

Private Sub picFriend_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call MovePicture(frmMirage.picFriend, Button, Shift, x, y)
End Sub

Private Sub itmDesc_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   ' SOffsetX = x
  '  SOffsetY = y
End Sub

Private Sub itmDesc_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call MovePicture(frmMirage.itmDesc, Button, Shift, x, y)
End Sub

Private Sub picGuildAdmin_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    SOffsetX = x
    SOffsetY = y
    Call picGuildAdmin.ZOrder(0)
End Sub

Private Sub picGuildAdmin_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call MovePicture(frmMirage.picGuildAdmin, Button, Shift, x, y)
End Sub

Private Sub picEquip_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    SOffsetX = x
    SOffsetY = y
    Call picEquip.ZOrder(0)
End Sub

Private Sub picEquip_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call MovePicture(frmMirage.picEquip, Button, Shift, x, y)
End Sub

Private Sub picStat_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call MovePicture(frmMirage.picStat, Button, Shift, x, y)
End Sub

Private Sub picChatbox_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call MovePicture(frmMirage.picChatBox, Button, Shift, x, y)
End Sub

Private Sub picScreen_GotFocus()
On Error Resume Next
    frmMirage.txtMyTextBox.SetFocus
End Sub

Private Sub picScreen_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim I As Long

    If InSpawnEditor Then
        If SpawnLocator > 0 Then
            TempNpcSpawn(SpawnLocator).Used = 1
            TempNpcSpawn(SpawnLocator).x = Int((x + (NewPlayerX * PIC_X)) / PIC_X)
            TempNpcSpawn(SpawnLocator).y = Int((y + (NewPlayerY * PIC_Y)) / PIC_Y)
            frmMapProperties.Spawn(SpawnLocator - 1).Caption = "(" & TempNpcSpawn(SpawnLocator).x & ", " & TempNpcSpawn(SpawnLocator).y & ")"
            SpawnLocator = 0
        End If
        
        Exit Sub
    End If

    If (Button = 1 Or Button = 2) And InEditor = True Then
        Call EditorMouseDown(Button, Shift, (x + (NewPlayerX * PIC_X)), (y + (NewPlayerY * PIC_Y)))
    End If
    
    If MouseCheck = True Then
    If Button = 1 And InEditor = False Then
                ControlDown = True
                Call CheckAttack
    End If
    End If
    
    If MouseCheck = False Then
    If Button = 1 And InEditor = False Then
                ControlDown = False
        Call PlayerSearch(Button, Shift, (x + (NewPlayerX * PIC_X)), (y + (NewPlayerY * PIC_Y)))
    End If
    End If
    
    If MouseCheck = False Then
    If Button = 2 Then
    Call PlayerSearch(Button, Shift, (x + (NewPlayerX * PIC_X)), (y + (NewPlayerY * PIC_Y)))
    End If
    End If
    
    If (Button = 1 Or Button = 2) And InEditor = False Then
        If Button = 1 And Player(MyIndex).Pet.Alive = YES Then
            Call PetMove(Button, Shift, (x + (NewPlayerX * PIC_X)), (y + (NewPlayerY * PIC_Y)))
        End If
      End If
      
        If MouseCheck = True Then
            If Button = 2 Then
                XToGo = (x + (NewPlayerX * PIC_X)) / PIC_X
                YToGo = (y + (NewPlayerY * PIC_Y)) / PIC_Y
                Call CheckMapGetItem
            End If
           End If
           
            If MouseCheck = True Then
            If Button = 1 Then
                Call PlayerSearch(Button, Shift, (x + (NewPlayerX * PIC_X)), (y + (NewPlayerY * PIC_Y)))
            End If
            End If
End Sub

Private Sub picScreen_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If (Button = 1 Or Button = 2) And InEditor = True Then
        Call EditorMouseDown(Button, Shift, (x + (NewPlayerX * PIC_X)), (y + (NewPlayerY * PIC_Y)))
    End If
    
    If InEditor Then
        MouseX = Int(x / PIC_X) * PIC_X
        MouseY = Int(y / PIC_Y) * PIC_Y
    End If
    
    frmMapEditor.Caption = "Map Editor - " & "X: " & Int((x + (NewPlayerX * PIC_X)) / PIC_X) & " Y: " & Int((y + (NewPlayerY * PIC_Y)) / PIC_Y)
End Sub

Private Sub Picture9_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    itmDesc.Visible = False
End Sub

Private Sub tmrDisease_Timer()
Dim Index As Long
Dim I As Long
Static Secs As Long

    If Secs <= 0 Then Secs = 30

    If Secs = 30 Then Call SendDisease
    If Secs = 28 Then Call SendDisease
    If Secs = 25 Then Call SendDisease
    If Secs = 22 Then Call SendDisease
    If Secs = 20 Then Call SendDisease
    If Secs = 18 Then Call SendDisease
    If Secs = 15 Then Call SendDisease
    If Secs = 12 Then Call SendDisease
    If Secs = 10 Then Call SendDisease
    If Secs = 8 Then Call SendDisease
    If Secs < 6 Then
        Call SendDisease
    End If
    Secs = Secs - 1

    If Secs <= 0 Then
        tmrDisease.Enabled = False
        Call AddText("The Effects of Disease Have Worn Off !", White)
    End If
End Sub

Private Sub tmrHunger_Timer()
Dim Index As Long
Dim I As Long
Static Secs As Long

    Call SendHunger
End Sub

Private Sub tmrPoison_Timer()
Dim Index As Long
Dim I As Long
Static Secs As Long

    If Secs <= 0 Then Secs = 30

    If Secs = 30 Then Call SendPoison
    If Secs = 25 Then Call SendPoison
    If Secs = 20 Then Call SendPoison
    If Secs = 15 Then Call SendPoison
    If Secs = 10 Then Call SendPoison
    If Secs < 6 Then
        Call SendPoison
    End If
    Secs = Secs - 1

    If Secs <= 0 Then
        tmrPoison.Enabled = False
        Call AddText("The Effects of Poison Have Worn Off !", White)
    End If

End Sub

Private Sub txtChat_GotFocus()
    frmMirage.txtMyTextBox.SetFocus
End Sub

Private Sub scrlBltText_Change()
Dim I As Long
    For I = 1 To MAX_BLT_LINE
        BattlePMsg(I).Index = 1
        BattlePMsg(I).Time = I
        BattleMMsg(I).Index = 1
        BattleMMsg(I).Time = I
    Next I
    
    MAX_BLT_LINE = scrlBltText.Value
    ReDim BattlePMsg(1 To MAX_BLT_LINE) As BattleMsgRec
    ReDim BattleMMsg(1 To MAX_BLT_LINE) As BattleMsgRec
    lblLines.Caption = "On Screen Text Line Amount: " & scrlBltText.Value
End Sub

Private Sub Socket_DataArrival(ByVal bytesTotal As Long)
    If IsConnected Then
        Call IncomingData(bytesTotal)
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Call HandleKeypresses(KeyAscii)
    If (KeyAscii = vbKeyReturn) Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Call CheckInput(1, KeyCode, Shift)
End Sub

Private Sub tmrRainDrop_Timer()
    If BLT_RAIN_DROPS > RainIntensity Then
        tmrRainDrop.Enabled = False
        Exit Sub
    End If
    If BLT_RAIN_DROPS > 0 Then
        If DropRain(BLT_RAIN_DROPS).Randomized = False Then
            Call RNDRainDrop(BLT_RAIN_DROPS)
        End If
    End If
    BLT_RAIN_DROPS = BLT_RAIN_DROPS + 1
    If tmrRainDrop.Interval > 30 Then
        tmrRainDrop.Interval = tmrRainDrop.Interval - 10
    End If
End Sub

Private Sub tmrSnowDrop_Timer()
    If BLT_SNOW_DROPS > RainIntensity Then
        tmrSnowDrop.Enabled = False
        Exit Sub
    End If
    If BLT_SNOW_DROPS > 0 Then
        If DropSnow(BLT_SNOW_DROPS).Randomized = False Then
            Call RNDSnowDrop(BLT_SNOW_DROPS)
        End If
    End If
    BLT_SNOW_DROPS = BLT_SNOW_DROPS + 1
    If tmrSnowDrop.Interval > 30 Then
        tmrSnowDrop.Interval = tmrSnowDrop.Interval - 10
    End If
End Sub

Private Sub picInv3entory_Click()
    picInv3.Visible = True
End Sub

Private Sub lblUseItem_Click()
Dim d As Long

If Player(MyIndex).Inv(Inventory).Num <= 0 Then Exit Sub

Call SendUseItem(Inventory)

For d = 1 To MAX_INV
    If Player(MyIndex).Inv(d).Num > 0 Then
        If Item(GetPlayerInvItemNum(MyIndex, d)).Type = ITEM_TYPE_POTIONADDMP Or ITEM_TYPE_POTIONADDHP Or ITEM_TYPE_POTIONADDSP Or ITEM_TYPE_POTIONSUBHP Or ITEM_TYPE_POTIONSUBMP Or ITEM_TYPE_POTIONSUBSP Then
            picInv(d - 1).Picture = LoadPicture()
        End If
    End If
Next d
Call UpdateVisInv
End Sub

Private Sub lblDropItem_Click()
    Call DropItems
End Sub

Sub DropItems()
Dim InvNum As Long
Dim GoldAmount As String
On Error GoTo Done
If Inventory <= 0 Then Exit Sub

    InvNum = Inventory
    If GetPlayerInvItemNum(MyIndex, InvNum) > 0 And GetPlayerInvItemNum(MyIndex, InvNum) <= MAX_ITEMS Then
      If Item(GetPlayerInvItemNum(MyIndex, InvNum)).Bound = 0 Then
        If Item(GetPlayerInvItemNum(MyIndex, InvNum)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(MyIndex, InvNum)).Stackable = 1 Then
            GoldAmount = InputBox("How much " & Trim(Item(GetPlayerInvItemNum(MyIndex, InvNum)).Name) & "(" & GetPlayerInvItemValue(MyIndex, InvNum) & ") would you like to drop?", "Drop " & Trim(Item(GetPlayerInvItemNum(MyIndex, InvNum)).Name), 0, frmMirage.Left, frmMirage.Top)
            If IsNumeric(GoldAmount) Then
                Call SendDropItem(InvNum, GoldAmount)
            End If
        Else
            Call SendDropItem(InvNum, 0)
        End If
    End If
End If
   
    picInv(InvNum - 1).Picture = LoadPicture()
    Call UpdateVisInv
    Exit Sub
Done:
    If Item(GetPlayerInvItemNum(MyIndex, InvNum)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(MyIndex, InvNum)).Stackable = 1 Then
        MsgBox "The variable cant handle that amount!"
    End If
End Sub

Private Sub lblCast_Click()
        If GetTickCount > Player(MyIndex).AttackTimer + 1000 Then
            If Player(MyIndex).MovingH = 0 And Player(MyIndex).MovingV = 0 Then
                Call SendData("cast" & SEP_CHAR & SpellIndex & SEP_CHAR & END_CHAR)
                Player(MyIndex).Attacking = 1
                Player(MyIndex).AttackTimer = GetTickCount
                Player(MyIndex).CastedSpell = YES
            Else
                Call AddText("Cannot cast while walking!", BrightRed)
            End If
    Else
        Call AddText("No spell here.", BrightRed)
    End If
End Sub

Private Sub lblCancel_Click()
    picInv3.Visible = False
End Sub

Private Sub lblSpellsCancel_Click()
    picPlayerSpells.Visible = False
End Sub

Private Sub picSpells_Click()
    Call SendData("spells" & SEP_CHAR & END_CHAR)
End Sub

Private Sub picStats_Click()
    Call SendData("getstats" & SEP_CHAR & END_CHAR)
End Sub

Private Sub picTrade_Click()
    Call SendData("trade" & SEP_CHAR & END_CHAR)
End Sub

Private Sub picQuit_Click()
    Call GameDestroy
End Sub

Private Sub cmdAccess_Click()
Dim Packet As String

    Packet = "GUILDCHANGEACCESS" & SEP_CHAR & txtName.Text & SEP_CHAR & txtAccess.Text & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Private Sub cmdDisown_Click()
Dim Packet As String

    Packet = "GUILDDISOWN" & SEP_CHAR & txtName.Text & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Private Sub cmdTrainee_Click()
Dim Packet As String
    
    Packet = "GUILDTRAINEE" & SEP_CHAR & txtName.Text & SEP_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Private Sub picOffline_Click()
    Call SendOnlineList
    lstOnline.Visible = False
    'Label9.Visible = False
End Sub

Private Sub picOnline_Click()
    Call SendOnlineList
    lstOnline.Visible = True
    'Label9.Visible = True
End Sub

Private Sub Up_Click()
If VScroll1.Value = 0 Then Exit Sub
    VScroll1.Value = VScroll1.Value - 1
    Picture9.Top = VScroll1.Value * -PIC_Y
End Sub

Private Sub Down_Click()
If VScroll1.Value = 3 Then Exit Sub
    VScroll1.Value = VScroll1.Value + 1
    Picture9.Top = VScroll1.Value * -PIC_Y
End Sub
