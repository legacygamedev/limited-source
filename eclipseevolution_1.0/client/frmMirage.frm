VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form frmMirage 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Main Game"
   ClientHeight    =   9030
   ClientLeft      =   555
   ClientTop       =   780
   ClientWidth     =   12045
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
   Icon            =   "frmMirage.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMirage.frx":08CA
   ScaleHeight     =   602
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   803
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.PictureBox piccharstats 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   3375
      Left            =   0
      ScaleHeight     =   225
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   160
      TabIndex        =   117
      Top             =   3480
      Width           =   2400
      Begin VB.Label Label27 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "POINTS :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   140
         Top             =   2880
         Width           =   1125
      End
      Begin VB.Label Label26 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "DEFENCE :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   139
         Top             =   2520
         Width           =   1125
      End
      Begin VB.Label Label25 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "STRENGTH :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   138
         Top             =   2160
         Width           =   1125
      End
      Begin VB.Label Label24 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "SPEED :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   137
         Top             =   1800
         Width           =   1125
      End
      Begin VB.Label Label23 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "MAGIC :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   136
         Top             =   1440
         Width           =   1125
      End
      Begin VB.Label Label22 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "LEVEL :  "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   135
         Top             =   1080
         Width           =   1125
      End
      Begin VB.Label lblPoints 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "POINTS"
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
         Height          =   300
         Left            =   840
         TabIndex        =   134
         Top             =   2880
         Width           =   1050
      End
      Begin VB.Label lblSTATWINDOW 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "CHARACTER"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   132
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label lblLevel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "LEVEL"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   300
         Left            =   840
         TabIndex        =   126
         Top             =   1080
         Width           =   1050
      End
      Begin VB.Label lblMAGI 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "MAGIC"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   300
         Left            =   840
         TabIndex        =   125
         Top             =   1440
         Width           =   1050
      End
      Begin VB.Label lblSPEED 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "SPEED"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   300
         Left            =   840
         TabIndex        =   124
         Top             =   1800
         Width           =   1050
      End
      Begin VB.Label AddSpeed 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   2040
         TabIndex        =   123
         Top             =   1800
         Width           =   165
      End
      Begin VB.Label AddMagi 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   2040
         TabIndex        =   122
         Top             =   1440
         Width           =   165
      End
      Begin VB.Label lblSTR 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "STRENGTH"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   300
         Left            =   840
         TabIndex        =   121
         Top             =   2160
         Width           =   1050
      End
      Begin VB.Label lblDEF 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "DEFENCE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   300
         Left            =   840
         TabIndex        =   120
         Top             =   2520
         Width           =   1050
      End
      Begin VB.Label AddStr 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   2040
         TabIndex        =   119
         Top             =   2160
         Width           =   165
      End
      Begin VB.Label AddDef 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   2040
         TabIndex        =   118
         Top             =   2520
         Width           =   165
      End
   End
   Begin VB.PictureBox picOptions 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Height          =   4665
      Left            =   2520
      ScaleHeight     =   309
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   173
      TabIndex        =   98
      Top             =   120
      Visible         =   0   'False
      Width           =   2625
      Begin VB.CheckBox chkplayerbar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Mini HP Bar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   120
         TabIndex        =   110
         Top             =   840
         Value           =   1  'Checked
         Width           =   2325
      End
      Begin VB.CheckBox chkplayername 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   120
         TabIndex        =   109
         Top             =   360
         Value           =   1  'Checked
         Width           =   2325
      End
      Begin VB.CheckBox chknpcname 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   120
         TabIndex        =   108
         Top             =   1320
         Value           =   1  'Checked
         Width           =   2325
      End
      Begin VB.CheckBox chkbubblebar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   120
         TabIndex        =   107
         Top             =   3000
         Value           =   1  'Checked
         Width           =   2325
      End
      Begin VB.CheckBox chknpcbar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Show NPC HP Bars"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   120
         TabIndex        =   106
         Top             =   1800
         Value           =   1  'Checked
         Width           =   2325
      End
      Begin VB.CheckBox chkplayerdamage 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   120
         TabIndex        =   105
         Top             =   600
         Value           =   1  'Checked
         Width           =   2325
      End
      Begin VB.CheckBox chknpcdamage 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   120
         TabIndex        =   104
         Top             =   1560
         Value           =   1  'Checked
         Width           =   2325
      End
      Begin VB.CheckBox chkmusic 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   120
         TabIndex        =   103
         Top             =   2280
         Value           =   1  'Checked
         Width           =   2325
      End
      Begin VB.CheckBox chksound 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   120
         TabIndex        =   102
         Top             =   2520
         Value           =   1  'Checked
         Width           =   2325
      End
      Begin VB.HScrollBar scrlBltText 
         Height          =   255
         Left            =   120
         Max             =   20
         Min             =   4
         TabIndex        =   101
         Top             =   3840
         Value           =   6
         Width           =   2295
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Save Settings"
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
         TabIndex        =   100
         Top             =   4320
         Width           =   2325
      End
      Begin VB.CheckBox chkAutoScroll 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   120
         TabIndex        =   99
         Top             =   3240
         Value           =   1  'Checked
         Width           =   2325
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
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
         Height          =   255
         Left            =   120
         TabIndex        =   133
         Top             =   3480
         Width           =   2295
      End
      Begin VB.Label lblLines 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00000000&
         Height          =   165
         Left            =   120
         TabIndex        =   115
         Top             =   3600
         Width           =   2325
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "-Player Data-"
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
         Height          =   210
         Left            =   120
         TabIndex        =   114
         Top             =   120
         Width           =   2295
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "-NPC Data-"
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
         Height          =   210
         Left            =   120
         TabIndex        =   113
         Top             =   1080
         Width           =   2295
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "-Sound/Music Data-"
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
         Height          =   210
         Left            =   120
         TabIndex        =   112
         Top             =   2040
         Width           =   2295
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "-Chat Data-"
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
         Height          =   210
         Left            =   120
         TabIndex        =   111
         Top             =   2760
         Width           =   2295
      End
   End
   Begin VB.PictureBox itmDesc 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      FillColor       =   &H00828B82&
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
      Height          =   4335
      Left            =   2520
      ScaleHeight     =   287
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   175
      TabIndex        =   86
      Top             =   120
      Visible         =   0   'False
      Width           =   2655
      Begin VB.Label descName 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   390
         TabIndex        =   97
         Top             =   0
         Width           =   1815
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "-Requirements-"
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
         Height          =   210
         Left            =   360
         TabIndex        =   96
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label descStr 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Strength"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   360
         TabIndex        =   95
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label descDef 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Defence"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   360
         TabIndex        =   94
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label descSpeed 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Speed"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   360
         TabIndex        =   93
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "-Add-"
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
         Height          =   210
         Left            =   360
         TabIndex        =   92
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label descHpMp 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "HP: XXXX MP: XXXX SP: XXXX"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   0
         TabIndex        =   91
         Top             =   1440
         Width           =   2655
      End
      Begin VB.Label descSD 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Str: XXXX Def: XXXXX"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   0
         TabIndex        =   90
         Top             =   1680
         Width           =   2655
      End
      Begin VB.Label desc 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   1815
         Left            =   120
         TabIndex        =   89
         Top             =   2400
         Width           =   2415
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "-Description-"
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
         Height          =   210
         Left            =   360
         TabIndex        =   88
         Top             =   2160
         Width           =   1815
      End
      Begin VB.Label descMS 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Magi: XXXXX Speed: XXXX"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   0
         TabIndex        =   87
         Top             =   1920
         Width           =   2655
      End
   End
   Begin VB.PictureBox picUber 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7200
      Left            =   2400
      ScaleHeight     =   480
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   640
      TabIndex        =   85
      Top             =   0
      Width           =   9600
   End
   Begin VB.PictureBox picItems 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
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
      Height          =   2.25000e5
      Left            =   14040
      Picture         =   "frmMirage.frx":1CF1
      ScaleHeight     =   2.23636e5
      ScaleMode       =   0  'User
      ScaleWidth      =   477.091
      TabIndex        =   84
      Top             =   240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Timer Mp3timer 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   12840
      Top             =   5640
   End
   Begin VB.Timer NightTimer 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   12840
      Top             =   4200
   End
   Begin VB.Timer tmrSnowDrop 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   12840
      Top             =   4680
   End
   Begin VB.Timer tmrRainDrop 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   12840
      Top             =   5160
   End
   Begin VB.PictureBox ScreenShot 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   12840
      ScaleHeight     =   495
      ScaleWidth      =   525
      TabIndex        =   63
      Top             =   3480
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.PictureBox picInv3 
      Appearance      =   0  'Flat
      BackColor       =   &H00828B82&
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
      Height          =   3375
      Left            =   0
      ScaleHeight     =   225
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   160
      TabIndex        =   1
      Top             =   3480
      Visible         =   0   'False
      Width           =   2400
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
         Left            =   870
         Picture         =   "frmMirage.frx":161633
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   62
         Top             =   3075
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
         Left            =   1230
         Picture         =   "frmMirage.frx":1618CB
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   61
         Top             =   3075
         Width           =   270
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   330
         Left            =   2640
         Max             =   3
         TabIndex        =   60
         Top             =   2400
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox Picture8 
         Appearance      =   0  'Flat
         BackColor       =   &H00789298&
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
         Height          =   3015
         Left            =   0
         ScaleHeight     =   201
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   177
         TabIndex        =   34
         Top             =   0
         Width           =   2655
         Begin VB.PictureBox Picture9 
            Appearance      =   0  'Flat
            BackColor       =   &H00828B82&
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
            Height          =   3255
            Left            =   0
            ScaleHeight     =   217
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   160
            TabIndex        =   35
            Top             =   0
            Width           =   2400
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
               Left            =   105
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   59
               Top             =   1095
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
               Left            =   1815
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   58
               Top             =   570
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
               Left            =   1245
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   57
               Top             =   570
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
               Left            =   675
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   56
               Top             =   570
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
               Left            =   105
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   55
               Top             =   570
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
               Left            =   1815
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   54
               Top             =   45
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
               Left            =   1245
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   53
               Top             =   45
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
               Left            =   675
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   52
               Top             =   45
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
               Left            =   105
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   51
               Top             =   45
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
               Left            =   675
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   50
               Top             =   1095
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
               Left            =   1245
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   49
               Top             =   1095
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
               Left            =   1815
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   48
               Top             =   1095
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
               Left            =   105
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   47
               Top             =   1620
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
               Left            =   675
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   46
               Top             =   1620
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
               Left            =   1245
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   45
               Top             =   1620
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
               Left            =   1815
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   44
               Top             =   1620
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
               Left            =   105
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   43
               Top             =   2145
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
               Left            =   675
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   42
               Top             =   2145
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
               Left            =   1245
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   41
               Top             =   2145
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
               Left            =   1815
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   40
               Top             =   2145
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
               Left            =   105
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   39
               Top             =   2670
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
               Left            =   675
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   38
               Top             =   2670
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
               Left            =   1245
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   37
               Top             =   2670
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
               Left            =   1815
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   36
               Top             =   2670
               Width           =   480
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
               Index           =   6
               Left            =   0
               Top             =   0
               Width           =   540
            End
            Begin VB.Shape SelectedItem 
               BorderColor     =   &H000000FF&
               BorderWidth     =   2
               Height          =   525
               Left            =   90
               Top             =   30
               Width           =   525
            End
            Begin VB.Shape EquipS 
               BorderColor     =   &H0000FFFF&
               BorderWidth     =   3
               Height          =   540
               Index           =   3
               Left            =   600
               Top             =   600
               Width           =   540
            End
            Begin VB.Shape EquipS 
               BorderColor     =   &H0000FFFF&
               BorderWidth     =   3
               Height          =   540
               Index           =   2
               Left            =   600
               Top             =   720
               Width           =   540
            End
            Begin VB.Shape EquipS 
               BorderColor     =   &H0000FFFF&
               BorderWidth     =   3
               Height          =   540
               Index           =   1
               Left            =   0
               Top             =   720
               Width           =   540
            End
            Begin VB.Shape EquipS 
               BorderColor     =   &H0000FFFF&
               BorderWidth     =   3
               Height          =   540
               Index           =   0
               Left            =   600
               Top             =   720
               Width           =   540
            End
         End
      End
      Begin VB.Line Line2 
         X1              =   24
         X2              =   191
         Y1              =   202
         Y2              =   202
      End
      Begin VB.Line Line1 
         X1              =   8
         X2              =   168
         Y1              =   200
         Y2              =   200
      End
      Begin VB.Label lblDropItem 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   1560
         TabIndex        =   3
         Top             =   3105
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
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   15
         TabIndex        =   2
         Top             =   3105
         Width           =   690
      End
   End
   Begin VB.PictureBox picPlayerSpells 
      Appearance      =   0  'Flat
      BackColor       =   &H00789298&
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
      Height          =   3345
      Left            =   0
      ScaleHeight     =   223
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   160
      TabIndex        =   4
      Top             =   3480
      Visible         =   0   'False
      Width           =   2400
      Begin VB.ListBox lstSpells 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   2835
         ItemData        =   "frmMirage.frx":161B56
         Left            =   45
         List            =   "frmMirage.frx":161B58
         TabIndex        =   5
         Top             =   60
         Width           =   2310
      End
      Begin VB.Label lblForgetSpell 
         BackStyle       =   0  'Transparent
         Caption         =   "Forget"
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
         Height          =   195
         Left            =   1440
         TabIndex        =   77
         Top             =   3120
         Width           =   495
      End
      Begin VB.Label lblCast 
         BackStyle       =   0  'Transparent
         Caption         =   "Cast"
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
         Height          =   195
         Left            =   480
         TabIndex        =   6
         Top             =   3120
         Width           =   375
      End
   End
   Begin VB.PictureBox picWhosOnline 
      Appearance      =   0  'Flat
      BackColor       =   &H00789298&
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
      Height          =   3345
      Left            =   0
      ScaleHeight     =   223
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   160
      TabIndex        =   14
      Top             =   3480
      Visible         =   0   'False
      Width           =   2400
      Begin VB.ListBox lstOnline 
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
         Height          =   2835
         ItemData        =   "frmMirage.frx":161B5A
         Left            =   45
         List            =   "frmMirage.frx":161B5C
         TabIndex        =   15
         Top             =   60
         Width           =   2310
      End
   End
   Begin VB.PictureBox picGuildAdmin 
      Appearance      =   0  'Flat
      BackColor       =   &H00789298&
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
      Height          =   3345
      Left            =   0
      ScaleHeight     =   3345
      ScaleWidth      =   2400
      TabIndex        =   17
      Top             =   3480
      Visible         =   0   'False
      Width           =   2400
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
         Left            =   750
         TabIndex        =   23
         Top             =   585
         Width           =   1575
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
         Left            =   750
         TabIndex        =   22
         Top             =   345
         Width           =   1575
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
         Left            =   300
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   975
         Width           =   1815
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
         Left            =   300
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   1305
         Width           =   1815
      End
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
         Left            =   300
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   1650
         Width           =   1815
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
         Left            =   300
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   1980
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
         Height          =   165
         Left            =   180
         TabIndex        =   25
         Top             =   615
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
         Height          =   165
         Left            =   210
         TabIndex        =   24
         Top             =   360
         Width           =   420
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00789298&
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
      Height          =   3345
      Left            =   0
      ScaleHeight     =   3345
      ScaleWidth      =   2400
      TabIndex        =   27
      Top             =   3480
      Visible         =   0   'False
      Width           =   2400
      Begin VB.Label cmdLeave 
         BackStyle       =   0  'Transparent
         Caption         =   "Leave Guild"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   840
         TabIndex        =   32
         Top             =   2280
         Width           =   765
      End
      Begin VB.Label lblRank 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rank"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   1425
         TabIndex        =   31
         Top             =   975
         Width           =   1080
      End
      Begin VB.Label lblGuild 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Guild"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   1425
         TabIndex        =   30
         Top             =   660
         Width           =   1065
      End
      Begin VB.Label Label17 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H80000008&
         Height          =   165
         Left            =   570
         TabIndex        =   29
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label16 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H80000008&
         Height          =   165
         Left            =   480
         TabIndex        =   28
         Top             =   645
         Width           =   825
      End
   End
   Begin VB.PictureBox picEquip 
      Appearance      =   0  'Flat
      BackColor       =   &H00789298&
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
      Height          =   3345
      Left            =   0
      ScaleHeight     =   3345
      ScaleWidth      =   2400
      TabIndex        =   33
      Top             =   3480
      Width           =   2400
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
         Left            =   960
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   35
         TabIndex        =   78
         Top             =   2280
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
            TabIndex        =   79
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
         Left            =   960
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   35
         TabIndex        =   74
         Top             =   840
         Width           =   555
         Begin VB.PictureBox NecklaceImage 
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
            TabIndex        =   75
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
         Left            =   1680
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   35
         TabIndex        =   72
         Top             =   1560
         Width           =   555
         Begin VB.PictureBox RingImage 
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
            TabIndex        =   73
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
         TabIndex        =   70
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
            TabIndex        =   71
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
         Left            =   960
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   35
         TabIndex        =   68
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
            TabIndex        =   69
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
         Left            =   1680
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   35
         TabIndex        =   66
         Top             =   840
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
            TabIndex        =   67
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
         Left            =   960
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   35
         TabIndex        =   64
         Top             =   120
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
            TabIndex        =   65
            Top             =   15
            Width           =   495
         End
      End
   End
   Begin VB.PictureBox picScreen 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
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
      Height          =   7200
      Left            =   13560
      ScaleHeight     =   480
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   640
      TabIndex        =   13
      Top             =   120
      Visible         =   0   'False
      Width           =   9600
   End
   Begin VB.TextBox txtMyTextBox 
      Appearance      =   0  'Flat
      BackColor       =   &H00828B82&
      BorderStyle     =   0  'None
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
      Height          =   300
      Left            =   2400
      MaxLength       =   255
      TabIndex        =   11
      Top             =   7320
      Width           =   9600
   End
   Begin MSWinsockLib.Winsock Socket 
      Left            =   12960
      Top             =   720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox txtChat 
      Height          =   1320
      Left            =   2400
      TabIndex        =   0
      Top             =   7680
      Width           =   9600
      _ExtentX        =   16933
      _ExtentY        =   2328
      _Version        =   393217
      BackColor       =   8555394
      BorderStyle     =   0
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmMirage.frx":161B5E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Close 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Something?"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   6360
      TabIndex        =   131
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "CHat"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   6240
      TabIndex        =   130
      Top             =   960
      Width           =   2115
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "  "
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
      Height          =   300
      Left            =   0
      TabIndex        =   129
      Top             =   2760
      Width           =   2400
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "AM / PM ?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   7080
      TabIndex        =   128
      Top             =   960
      Width           =   825
   End
   Begin VB.Label lblcharstats 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "  "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   300
      Left            =   0
      TabIndex        =   127
      Top             =   960
      Width           =   2400
   End
   Begin VB.Label lblLabel20 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "  "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   300
      Left            =   0
      TabIndex        =   116
      Top             =   3120
      Width           =   2400
   End
   Begin WMPLibCtl.WindowsMediaPlayer mp3player 
      Height          =   30
      Left            =   -30
      TabIndex        =   83
      Top             =   -30
      Visible         =   0   'False
      Width           =   30
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   53
      _cy             =   53
   End
   Begin WMPLibCtl.WindowsMediaPlayer Mp3musicplayer 
      Height          =   30
      Left            =   -30
      TabIndex        =   82
      Top             =   -30
      Visible         =   0   'False
      Width           =   30
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   53
      _cy             =   53
   End
   Begin VB.Label GameClock 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   81
      Top             =   240
      Width           =   2205
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "It is now:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   80
      Top             =   75
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label lblEXP 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      BackStyle       =   0  'Transparent
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
      Height          =   225
      Left            =   75
      TabIndex        =   76
      Top             =   6930
      Width           =   2250
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "  "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   300
      Left            =   0
      TabIndex        =   26
      Top             =   2400
      Width           =   2400
   End
   Begin VB.Label lblWhosOnline 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "  "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   300
      Left            =   0
      TabIndex        =   16
      Top             =   1680
      Width           =   2400
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "  "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   300
      Left            =   0
      TabIndex        =   12
      Top             =   600
      Width           =   2400
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "  "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   300
      Left            =   0
      TabIndex        =   10
      Top             =   2040
      Width           =   2400
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "  "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   300
      Left            =   0
      TabIndex        =   9
      Top             =   1320
      Width           =   2400
   End
   Begin VB.Label lblMP 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00CB884B&
      BackStyle       =   0  'Transparent
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
      Height          =   225
      Left            =   75
      TabIndex        =   8
      Top             =   7230
      Width           =   2250
   End
   Begin VB.Label lblHP 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      BackStyle       =   0  'Transparent
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
      Height          =   225
      Left            =   75
      TabIndex        =   7
      Top             =   7530
      Width           =   2250
   End
   Begin VB.Shape shpHP 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   225
      Left            =   75
      Top             =   7530
      Width           =   2250
   End
   Begin VB.Shape shpMP 
      BackColor       =   &H00CB884B&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   225
      Left            =   75
      Top             =   7230
      Width           =   2250
   End
   Begin VB.Shape shpTNL 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000C000&
      Height          =   225
      Left            =   75
      Top             =   6930
      Width           =   2250
   End
End
Attribute VB_Name = "frmMirage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SpellMemorized As Long

Private Sub Close_Click()
    Call StopMidi
    Unload Me
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
    WriteINI "CONFIG", "Sound", chksound.Value, App.Path & "\config.ini"
End Sub

Private Sub chkbubblebar_Click()
    WriteINI "CONFIG", "SpeechBubbles", chkbubblebar.Value, App.Path & "\config.ini"
End Sub

Private Sub chknpcbar_Click()
    WriteINI "CONFIG", "NpcBar", chknpcbar.Value, App.Path & "\config.ini"
End Sub

Private Sub chknpcdamage_Click()
    WriteINI "CONFIG", "NPCDamage", chknpcdamage.Value, App.Path & "\config.ini"
End Sub

Private Sub chknpcname_Click()
    WriteINI "CONFIG", "NPCName", chknpcname.Value, App.Path & "\config.ini"
End Sub

Private Sub chkplayerbar_Click()
    WriteINI "CONFIG", "PlayerBar", chkplayerbar.Value, App.Path & "\config.ini"
End Sub

Private Sub chkplayerdamage_Click()
    WriteINI "CONFIG", "PlayerDamage", chkplayerdamage.Value, App.Path & "\config.ini"
End Sub

Private Sub chkAutoScroll_Click()
    WriteINI "CONFIG", "AutoScroll", chkAutoScroll.Value, App.Path & "\config.ini"
End Sub

Private Sub chkplayername_Click()
    WriteINI "CONFIG", "PlayerName", chkplayername.Value, App.Path & "\config.ini"
End Sub

Private Sub chkmusic_Click()
    WriteINI "CONFIG", "Music", chkmusic.Value, App.Path & "\config.ini"
    If MyIndex <= 0 Then Exit Sub
    Call PlayMidi(Trim$(Map(GetPlayerMap(MyIndex)).Music))
End Sub

Private Sub cmdLeave_Click()
Dim packet As String
    packet = "GUILDLEAVE" & SEP_CHAR & END_CHAR
    Call SendData(packet)
End Sub

Private Sub cmdMember_Click()
Dim packet As String
    packet = "GUILDMEMBER" & SEP_CHAR & txtName.Text & SEP_CHAR & END_CHAR
    Call SendData(packet)
End Sub

Private Sub Command1_Click()
    picOptions.Visible = False
End Sub



Private Sub Form_Load()
Dim i As Long
Dim Ending As String
    For i = 1 To 3
        If i = 1 Then Ending = ".gif"
        If i = 2 Then Ending = ".jpg"
        If i = 3 Then Ending = ".png"
 
        If FileExist("GUI\Game" & Ending) Then frmMirage.Picture = LoadPicture(App.Path & "\GUI\Game" & Ending)
    Next i
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    itmDesc.Visible = False
End Sub

Private Sub Form_Terminate()
    Call GameDestroy
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call GameDestroy
End Sub

Private Sub HScroll1_Change()



End Sub

Private Sub Label1_Click()
Dim i As Long

For i = 1 To MAX_PLAYERS
    If IsPlaying(i) = True Then
        If MouseDownX = GetPlayerX(i) And MouseDownY = GetPlayerY(i) Then
            Call SendData("playerchat" & SEP_CHAR & GetPlayerName(i) & SEP_CHAR & END_CHAR)
            Exit Sub
        End If
    End If
Next i
End Sub

Private Sub Label13_Click()
' Set Their Guild Name and Their Rank
frmMirage.lblGuild.Caption = GetPlayerGuild(MyIndex)
frmMirage.lblRank.Caption = GetPlayerGuildAccess(MyIndex)
Picture1.Visible = True
picInv3.Visible = False
picPlayerSpells.Visible = False
'picEquip.Visible = False
picWhosOnline.Visible = False
frmMirage.piccharstats.Visible = False
End Sub

Private Sub Label19_Click()
    picEquip.Visible = True
    picInv3.Visible = False
    picPlayerSpells.Visible = False
    picWhosOnline.Visible = False
    Picture1.Visible = False
    frmMirage.picGuildAdmin.Visible = False
    frmMirage.piccharstats.Visible = False
    Call UpdateVisInv
End Sub

Private Sub Label2_Click()
    picOptions.Visible = True
    chkplayername.Value = Trim$(ReadINI("CONFIG", "playername", App.Path & "\config.ini"))
chkplayerdamage.Value = Trim$(ReadINI("CONFIG", "playerdamage", App.Path & "\config.ini"))
chkplayerbar.Value = Trim$(ReadINI("CONFIG", "playerbar", App.Path & "\config.ini"))
chknpcname.Value = Trim$(ReadINI("CONFIG", "npcname", App.Path & "\config.ini"))
chknpcdamage.Value = Trim$(ReadINI("CONFIG", "npcdamage", App.Path & "\config.ini"))
chknpcbar.Value = Trim$(ReadINI("CONFIG", "npcbar", App.Path & "\config.ini"))
chkmusic.Value = Trim$(ReadINI("CONFIG", "music", App.Path & "\config.ini"))
chksound.Value = Trim$(ReadINI("CONFIG", "sound", App.Path & "\config.ini"))
chkbubblebar.Value = Trim$(ReadINI("CONFIG", "speechbubbles", App.Path & "\config.ini"))
chkAutoScroll.Value = Trim$(ReadINI("CONFIG", "AutoScroll", App.Path & "\config.ini"))
End Sub

Private Sub Label21_Click()
'picEquip.Visible = False
End Sub



Private Sub Label3_Click()
Call GameDestroy
End Sub


Private Sub Label7_Click()
    Call UpdateVisInv
    picInv3.Visible = True
    Picture1.Visible = False
    frmMirage.picGuildAdmin.Visible = False
    ''picEquip.Visible = False
    picPlayerSpells.Visible = False
    picWhosOnline.Visible = False
    frmMirage.piccharstats.Visible = False
End Sub

Private Sub Label8_Click()
    Call SendData("spells" & SEP_CHAR & END_CHAR)
    picInv3.Visible = False
    frmMirage.picGuildAdmin.Visible = False
    'picEquip.Visible = False
    Picture1.Visible = False
    picWhosOnline.Visible = False
    frmMirage.piccharstats.Visible = False
End Sub

Private Sub lblCloseOnline_Click()
Call SendOnlineList
picWhosOnline.Visible = False
End Sub

Private Sub lblClosePicGuildAdmin_Click()
picGuildAdmin.Visible = False
End Sub

Private Sub lblcharstats_Click()
picWhosOnline.Visible = True
picInv3.Visible = False
'picEquip.Visible = False
Picture1.Visible = False
frmMirage.picGuildAdmin.Visible = False
picPlayerSpells.Visible = False
frmMirage.piccharstats.Visible = True
End Sub




Private Sub lblForgetSpell_Click()
If Player(MyIndex).Spell(lstSpells.ListIndex + 1) > 0 Then
If GetTickCount > Player(MyIndex).AttackTimer + 1000 Then
If MsgBox("Are you sure you want to forget this spell?", vbYesNo, "Forget Spell") = vbNo Then Exit Sub
Call SendData("forgetspell" & SEP_CHAR & lstSpells.ListIndex + 1 & SEP_CHAR & END_CHAR)
picPlayerSpells.Visible = False
End If
Else
Call AddText("No spell here.", BrightRed)
End If
End Sub




Private Sub lblLabel20_Click()
Call GameDestroy
End Sub

Private Sub lblSTATWINDOW_Click()
    Call SendData("getstats" & SEP_CHAR & END_CHAR)
End Sub

Private Sub lblWhosOnline_Click()
Call SendOnlineList
picWhosOnline.Visible = True
picInv3.Visible = False
'picEquip.Visible = False
Picture1.Visible = False
frmMirage.picGuildAdmin.Visible = False
picPlayerSpells.Visible = False
frmMirage.piccharstats.Visible = False
End Sub

Private Sub lstOnline_DblClick()
    Call SendData("playerchat" & SEP_CHAR & Trim$(lstOnline.Text) & SEP_CHAR & END_CHAR)
End Sub

Private Sub lstSpells_DblClick()
    If Player(MyIndex).Spell(lstSpells.ListIndex + 1) > 0 Then
        SpellMemorized = lstSpells.ListIndex + 1
        Call AddText("Successfully memorized spell!", White)
    Else
        Call AddText("No spell here to memorize.", BrightRed)
    End If
End Sub




Private Sub MENUreloader_Timer()

frmCustom1.Visible = True

End Sub

Private Sub Mp3timer_Timer()

frmMirage.Mp3musicplayer.settings.playCount = 500

End Sub




Private Sub picInv_DblClick(Index As Integer)
Dim d As Long

If Player(MyIndex).Inv(Inventory).num <= 0 Then Exit Sub

Call SendUseItem(Inventory)

For d = 1 To MAX_INV
    If Player(MyIndex).Inv(d).num > 0 Then
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

    If Player(MyIndex).Inv(d + 1).num > 0 Then
        If Item(GetPlayerInvItemNum(MyIndex, d + 1)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(MyIndex, d + 1)).Stackable = 1 Then
            If Trim$(Item(GetPlayerInvItemNum(MyIndex, d + 1)).desc) = "" Then
                itmDesc.Height = 17
                itmDesc.Top = 224
            Else
                itmDesc.Height = 233
                itmDesc.Top = 8
            End If
        Else
            If Trim$(Item(GetPlayerInvItemNum(MyIndex, d + 1)).desc) = "" Then
                itmDesc.Height = 145
                itmDesc.Top = 96
            Else
                itmDesc.Height = 233
                itmDesc.Top = 8
            End If
        End If
        If Item(GetPlayerInvItemNum(MyIndex, d + 1)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(MyIndex, d + 1)).Stackable = 1 Then
            descName.Caption = Trim$(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Name) & " (" & GetPlayerInvItemValue(MyIndex, d + 1) & ")"
        Else
            If GetPlayerWeaponSlot(MyIndex) = d + 1 Then
                descName.Caption = Trim$(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Name) & " (worn)"
            ElseIf GetPlayerArmorSlot(MyIndex) = d + 1 Then
                descName.Caption = Trim$(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Name) & " (worn)"
            ElseIf GetPlayerHelmetSlot(MyIndex) = d + 1 Then
                descName.Caption = Trim$(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Name) & " (worn)"
            ElseIf GetPlayerShieldSlot(MyIndex) = d + 1 Then
                descName.Caption = Trim$(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Name) & " (worn)"
            ElseIf GetPlayerLegsSlot(MyIndex) = d + 1 Then
                descName.Caption = Trim$(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Name) & " (worn)"
            ElseIf GetPlayerRingSlot(MyIndex) = d + 1 Then
                descName.Caption = Trim$(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Name) & " (worn)"
            ElseIf GetPlayerNecklaceSlot(MyIndex) = d + 1 Then
                descName.Caption = Trim$(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Name) & " (worn)"
            Else
                descName.Caption = Trim$(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Name)
            End If
        End If
        descStr.Caption = Item(GetPlayerInvItemNum(MyIndex, d + 1)).StrReq & " Strength"
        descDef.Caption = Item(GetPlayerInvItemNum(MyIndex, d + 1)).DefReq & " Defence"
        descSpeed.Caption = Item(GetPlayerInvItemNum(MyIndex, d + 1)).SpeedReq & " Speed"
        descHpMp.Caption = "HP: " & Item(GetPlayerInvItemNum(MyIndex, d + 1)).AddHP & " MP: " & Item(GetPlayerInvItemNum(MyIndex, d + 1)).AddMP & " SP: " & Item(GetPlayerInvItemNum(MyIndex, d + 1)).AddSP
        descSD.Caption = "Str: " & Item(GetPlayerInvItemNum(MyIndex, d + 1)).AddStr & " Def: " & Item(GetPlayerInvItemNum(MyIndex, d + 1)).AddDef
        descMS.Caption = "Magi: " & Item(GetPlayerInvItemNum(MyIndex, d + 1)).AddMagi & " Speed: " & Item(GetPlayerInvItemNum(MyIndex, d + 1)).AddSpeed
        desc.Caption = Trim$(Item(GetPlayerInvItemNum(MyIndex, d + 1)).desc)
        
        itmDesc.Visible = True
    Else
        itmDesc.Visible = False
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Dim d As Long
Dim i As Long
Dim ii As Long

    Call CheckInput(0, KeyCode, Shift)
    If KeyCode = vbKeyF1 Then
        If Player(MyIndex).Access > 0 Then
            frmadmin.Visible = False
            frmadmin.Visible = True
        End If
    End If
    
    
    
    
    
    
    
If KeyCode = vbKeyF2 Then
    For i = 1 To MAX_INV
    
        If GetPlayerInvItemNum(MyIndex, i) > 0 And GetPlayerInvItemNum(MyIndex, i) <= MAX_ITEMS Then
            If Item(GetPlayerInvItemNum(MyIndex, i)).Type = ITEM_TYPE_POTIONADDHP Then
                Call SendUseItem(i)
                Call AddText("You Drink a healing potion", Yellow)
                Exit Sub
            End If
        Else
            If i = MAX_INV Then
                Call AddText("You dont have any healing potions!", Red)
            End If
        End If
        
    Next i
    
End If

If KeyCode = vbKeyF3 Then
For i = 1 To MAX_INV

    If GetPlayerInvItemNum(MyIndex, i) > 0 And GetPlayerInvItemNum(MyIndex, i) <= MAX_ITEMS Then
        If Item(GetPlayerInvItemNum(MyIndex, i)).Type = ITEM_TYPE_POTIONADDMP Then
        Call SendUseItem(i)
        Call AddText("You Drink a Mana Restoring potion", Yellow)
        
        Exit Sub
        End If
    Else
    If i = MAX_INV Then Call AddText("You dont have any Mana potions!", Red)
    End If
    
Next i
End If
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    ' The Guild Creator
    If KeyCode = vbKeyF4 Then
        If Player(MyIndex).Access > 0 Then
            frmGuild.Show vbModeless, frmMirage
        End If
    End If

    If KeyCode = vbKeyPageUp Then
    Call SendHotScript1
    End If
    
    If KeyCode = vbKeyDelete Then
    Call SendHotScript2
    End If
    
    If KeyCode = vbKeyEnd Then
    Call SendHotScript3
    End If
    
    If KeyCode = vbKeyPageDown Then
    Call SendHotScript4
    End If

    ' The Guild Maker
    If KeyCode = vbKeyF5 Then
        frmMirage.picGuildAdmin.Visible = True
        frmMirage.picInv3.Visible = False
        frmMirage.Picture1.Visible = False
        'frmMirage.picEquip.Visible = False
        frmMirage.picPlayerSpells.Visible = False
        frmMirage.picWhosOnline.Visible = False
      End If
      
    If KeyCode = vbKeyInsert Then
        If SpellMemorized > 0 Then
            If GetTickCount > Player(MyIndex).AttackTimer + 1000 Then
                If Player(MyIndex).Moving = 0 Then
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
    
    If KeyCode = vbKeyF11 Then
        ScreenShot.Picture = CaptureForm(frmMirage)
        i = 0
        ii = 0
        Do
            If FileExist("Screenshot" & i & ".bmp") = True Then
                i = i + 1
            Else
                Call SavePicture(ScreenShot.Picture, App.Path & "\Screenshot" & i & ".bmp")
                ii = 1
            End If
            
            Call sleep(1)
            DoEvents
        Loop Until ii = 1
    ElseIf KeyCode = vbKeyF12 Then
        ScreenShot.Picture = CaptureArea(frmMirage, 8, 6, 634, 479)
        i = 0
        ii = 0
        Do
            If FileExist("Screenshot" & i & ".bmp") = True Then
                i = i + 1
            Else
                Call SavePicture(ScreenShot.Picture, App.Path & "\Screenshot" & i & ".bmp")
                ii = 1
            End If
            
            Call sleep(1)
            DoEvents
        Loop Until ii = 1
    End If
    
    If KeyCode = vbKeyHome Then
    d = GetPlayerDir(MyIndex)
    
        If Player(MyIndex).Moving = NO Then
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

Private Sub PicOptions_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    SOffsetX = x
    SOffsetY = y
End Sub

Private Sub PicOptions_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call MovePicture(frmMirage.picOptions, Button, Shift, x, y)
End Sub



Private Sub picScreen_GotFocus()
On Error Resume Next
    txtMyTextBox.SetFocus
End Sub

Private Sub picUber_GotFocus()
On Error Resume Next
    txtMyTextBox.SetFocus
End Sub
Private Sub picUber_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim i As Long
Dim xcalc As Single
Dim ycalc As Single
Dim xconv As Long
Dim yconv As Long

x = x / (frmMirage.picUber.Width / frmMirage.picScreen.Width)
y = y / (frmMirage.picUber.Height / frmMirage.picScreen.Height)

xcalc = (NewPlayerX * PIC_X) / (frmMirage.picUber.Width / frmMirage.picScreen.Width)
ycalc = (NewPlayerY * PIC_Y) / (frmMirage.picUber.Height / frmMirage.picScreen.Height)

x = x + xcalc
y = y + ycalc

    If (Button = 1 Or Button = 2) And InEditor = True Then
        'Call AddText("Clicked x:" & (x) & " y:" & (y), BrightRed)
        Call EditorMouseDown(Button, Shift, x, y)
    End If
    
    If Button = 1 And InEditor = False Then
        'Call AddText("Clicked xcalc" & (x + xcalc) & " ycalc" & (y + ycalc), BrightRed)
        Call PlayerSearch(Button, Shift, x, y)
        'Call PlayerSearch(Button, Shift, (x + (NewPlayerX * PIC_X)), ycalc)
    End If
    
    If (Button = 1 Or Button = 2) And InHouseEditor = True Then
        'Call AddText("Clicked xcalc" & (x + xcalc) & " ycalc" & (y + ycalc), BrightRed)
        Call HouseEditorMouseDown(Button, Shift, x, y)
        'Call HouseEditorMouseDown(Button, Shift, (x + (NewPlayerX * PIC_X)), ycalc)
    End If
    
End Sub

Private Sub picUber_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim xcalc As Single
Dim ycalc As Single
Dim xconv As Long
Dim yconv As Long

x = x / (frmMirage.picUber.Width / frmMirage.picScreen.Width)
y = y / (frmMirage.picUber.Height / frmMirage.picScreen.Height)

xcalc = (NewPlayerX * PIC_X) / (frmMirage.picUber.Width / frmMirage.picScreen.Width)
ycalc = (NewPlayerY * PIC_Y) / (frmMirage.picUber.Height / frmMirage.picScreen.Height)

    If (Button = 1 Or Button = 2) And InEditor = True Then
        Call EditorMouseDown(Button, Shift, Int((x + (xcalc))), Int((y + (ycalc))))
    End If
    
    If (Button = 1 Or Button = 2) And InHouseEditor = True Then
        Call HouseEditorMouseDown(Button, Shift, Int((x + (xcalc))), Int((y + (ycalc))))
    End If

    frmMapEditor.Caption = "Map Editor - " & "X: " & Int(Int((x + (xcalc))) / 32) & " Y: " & Int(Int((y + (ycalc))) / 32)
    frmHouseEditor.Caption = "House Editor - " & "X: " & Int(Int((x + (xcalc))) / 32) & " Y: " & Int(Int((y + (ycalc))) / 32)

' (x + (NewPlayerX * PIC_X)) BECAME Int((x + (xcalc)))
' Int((y + (ycalc)))

End Sub

Private Sub picScreen_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    If (Button = 1 Or Button = 2) And InEditor = True Then
        Call EditorMouseDown(Button, Shift, (x + (NewPlayerX * PIC_X)), (y + (NewPlayerY * PIC_Y)))
    End If
    
    If (Button = 1 Or Button = 2) And InHouseEditor = True Then
        Call HouseEditorMouseDown(Button, Shift, (x + (NewPlayerX * PIC_X)), (y + (NewPlayerY * PIC_Y)))
    End If

    frmMapEditor.Caption = "Map Editor - " & "X: " & Int((x + (NewPlayerX * PIC_X)) / 32) & " Y: " & Int((y + (NewPlayerY * PIC_Y)) / 32)
    frmHouseEditor.Caption = "House Editor - " & "X: " & Int((x + (NewPlayerX * PIC_X)) / 32) & " Y: " & Int((y + (NewPlayerY * PIC_Y)) / 32)

End Sub

Private Sub picScreen_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim i As Long

    If (Button = 1 Or Button = 2) And InEditor = True Then
        'Call AddText("Clicked real xcalc" & (x + (NewPlayerX * PIC_X)) & " ycalc" & (y + (NewPlayerY * PIC_Y)), BrightRed)
        Call EditorMouseDown(Button, Shift, (x + (NewPlayerX * PIC_X)), (y + (NewPlayerY * PIC_Y)))
    End If
    
    If Button = 1 And InEditor = False Then
        'Call AddText("Clicked real xcalc" & (x + (NewPlayerX * PIC_X)) & " ycalc" & (y + (NewPlayerY * PIC_Y)), BrightRed)
        Call PlayerSearch(Button, Shift, (x + (NewPlayerX * PIC_X)), (y + (NewPlayerY * PIC_Y)))
    End If
    
        If (Button = 1 Or Button = 2) And InHouseEditor = True Then
        'Call AddText("Clicked real xcalc" & (x + (NewPlayerX * PIC_X)) & " ycalc" & (y + (NewPlayerY * PIC_Y)), BrightRed)
 
        Call HouseEditorMouseDown(Button, Shift, (x + (NewPlayerX * PIC_X)), (y + (NewPlayerY * PIC_Y)))
    End If
End Sub



Private Sub Picture9_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    itmDesc.Visible = False
End Sub

Private Sub scrlBltText_Change()
Dim i As Long
    For i = 1 To MAX_BLT_LINE
        BattlePMsg(i).Index = 1
        BattlePMsg(i).Time = i
        BattleMMsg(i).Index = 1
        BattleMMsg(i).Time = i
    Next i
    
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

Private Sub txtChat_GotFocus()
    frmMirage.txtMyTextBox.SetFocus
End Sub

Private Sub picInv3entory_Click()
    picInv3.Visible = True
End Sub

Private Sub lblUseItem_Click()
Dim d As Long

If Player(MyIndex).Inv(Inventory).num <= 0 Then Exit Sub

Call SendUseItem(Inventory)

For d = 1 To MAX_INV
    If Player(MyIndex).Inv(d).num > 0 Then
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
            GoldAmount = InputBox("How much " & Trim$(Item(GetPlayerInvItemNum(MyIndex, InvNum)).Name) & "(" & GetPlayerInvItemValue(MyIndex, InvNum) & ") would you like to drop?", "Drop " & Trim$(Item(GetPlayerInvItemNum(MyIndex, InvNum)).Name), 0, frmMirage.Left, frmMirage.Top)
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
    If Player(MyIndex).Spell(lstSpells.ListIndex + 1) > 0 Then
        If GetTickCount > Player(MyIndex).AttackTimer + 1000 Then
            If Player(MyIndex).Moving = 0 Then
                Call SendData("cast" & SEP_CHAR & lstSpells.ListIndex + 1 & SEP_CHAR & END_CHAR)
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
Dim packet As String

    packet = "GUILDCHANGEACCESS" & SEP_CHAR & txtName.Text & SEP_CHAR & txtAccess.Text & SEP_CHAR & END_CHAR
    Call SendData(packet)
End Sub

Private Sub cmdDisown_Click()
Dim packet As String

    packet = "GUILDDISOWN" & SEP_CHAR & txtName.Text & SEP_CHAR & END_CHAR
    Call SendData(packet)
End Sub

Private Sub cmdTrainee_Click()
Dim packet As String
    
    packet = "GUILDTRAINEE" & SEP_CHAR & txtName.Text & SEP_CHAR & END_CHAR
    Call SendData(packet)
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
Private Sub lstSpells_GotFocus()
'picScreen.SetFocus
picUber.SetFocus
End Sub


