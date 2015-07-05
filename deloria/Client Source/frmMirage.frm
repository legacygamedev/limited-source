VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "Mswinsck.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmMirage 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Deloria"
   ClientHeight    =   9000
   ClientLeft      =   255
   ClientTop       =   -105
   ClientWidth     =   12000
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMirage.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMirage.frx":0442
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Timer tmrRainDrop 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   120
      Top             =   600
   End
   Begin VB.Timer tmrSnowDrop 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   600
      Top             =   600
   End
   Begin VB.PictureBox picDayNight 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
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
      Height          =   6255
      Left            =   12000
      ScaleHeight     =   417
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   545
      TabIndex        =   123
      Top             =   8880
      Width           =   8175
   End
   Begin VB.PictureBox picOptions 
      Appearance      =   0  'Flat
      BackColor       =   &H00492A04&
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
      Height          =   2445
      Left            =   9240
      ScaleHeight     =   161
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   173
      TabIndex        =   104
      Top             =   330
      Visible         =   0   'False
      Width           =   2625
      Begin VB.CheckBox Check1 
         BackColor       =   &H00492A04&
         Caption         =   "Enable Joypad"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFC0&
         Height          =   195
         Left            =   120
         TabIndex        =   127
         Top             =   1920
         Width           =   1485
      End
      Begin VB.CheckBox chkTime 
         BackColor       =   &H00492A04&
         Caption         =   "Enable Day/Night"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFC0&
         Height          =   195
         Left            =   120
         TabIndex        =   125
         Top             =   1680
         Value           =   1  'Checked
         Width           =   1485
      End
      Begin VB.CheckBox chkWeather 
         BackColor       =   &H00492A04&
         Caption         =   "Enable Weather"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFC0&
         Height          =   195
         Left            =   120
         TabIndex        =   124
         Top             =   1440
         Value           =   1  'Checked
         Width           =   1365
      End
      Begin VB.CheckBox chkBroadcast 
         BackColor       =   &H00492A04&
         Caption         =   "Enable Broadcast"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFC0&
         Height          =   195
         Left            =   120
         TabIndex        =   121
         Top             =   1200
         Value           =   1  'Checked
         Width           =   1365
      End
      Begin VB.CheckBox chkplayername 
         BackColor       =   &H00492A04&
         Caption         =   "Player Names"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFC0&
         Height          =   195
         Left            =   120
         TabIndex        =   109
         Top             =   240
         Value           =   1  'Checked
         Width           =   1245
      End
      Begin VB.CheckBox chkmusic 
         BackColor       =   &H00492A04&
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
         ForeColor       =   &H00FFFFC0&
         Height          =   225
         Left            =   120
         TabIndex        =   108
         Top             =   480
         Value           =   1  'Checked
         Width           =   735
      End
      Begin VB.CheckBox chksound 
         BackColor       =   &H00492A04&
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
         ForeColor       =   &H00FFFFC0&
         Height          =   225
         Left            =   120
         TabIndex        =   107
         Top             =   720
         Value           =   1  'Checked
         Width           =   765
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
         Left            =   2040
         TabIndex        =   106
         Top             =   2040
         Width           =   495
      End
      Begin VB.CheckBox chkFPS 
         BackColor       =   &H00492A04&
         Caption         =   "Display FPS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFC0&
         Height          =   195
         Left            =   120
         TabIndex        =   105
         Top             =   960
         Value           =   1  'Checked
         Width           =   1125
      End
      Begin VB.Label Label26 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Controls"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   120
         TabIndex        =   126
         Top             =   2160
         Width           =   600
      End
      Begin VB.Label Label35 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Game Options"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   210
         Left            =   0
         TabIndex        =   110
         Top             =   0
         Width           =   2655
      End
   End
   Begin VB.PictureBox picScreen 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Height          =   6120
      Left            =   3600
      ScaleHeight     =   408
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   552
      TabIndex        =   15
      Top             =   360
      Width           =   8280
   End
   Begin VB.PictureBox picStats 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2535
      Left            =   240
      ScaleHeight     =   169
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   209
      TabIndex        =   82
      Top             =   3480
      Width           =   3135
      Begin VB.Label lblGender 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "XXX"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   1560
         TabIndex        =   102
         Top             =   480
         Width           =   1500
      End
      Begin VB.Label lblClass 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "XXX"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   1560
         TabIndex        =   101
         Top             =   240
         Width           =   1500
      End
      Begin VB.Label lblLevel 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "XXX"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   1560
         TabIndex        =   100
         Top             =   0
         Width           =   1500
      End
      Begin VB.Label lblPoints 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "XXX"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   1560
         TabIndex        =   99
         Top             =   2280
         Width           =   1380
      End
      Begin VB.Label Label25 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   1560
         TabIndex        =   98
         Top             =   2040
         Width           =   1380
      End
      Begin VB.Label lblVit 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "XXX"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   1560
         TabIndex        =   97
         Top             =   1320
         Width           =   1380
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Gender:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   195
         Left            =   240
         TabIndex        =   96
         Top             =   480
         Width           =   1200
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Class:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   195
         Left            =   240
         TabIndex        =   95
         Top             =   240
         Width           =   1200
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Level:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   195
         Left            =   240
         TabIndex        =   94
         Top             =   0
         Width           =   1200
      End
      Begin VB.Label Label34 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Train Skills:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   195
         Left            =   240
         TabIndex        =   93
         Top             =   2280
         Width           =   1200
      End
      Begin VB.Label Label33 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Wisdom:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   195
         Left            =   240
         TabIndex        =   92
         Top             =   2040
         Width           =   1200
      End
      Begin VB.Label Label32 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Intelligence:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   195
         Left            =   240
         TabIndex        =   91
         Top             =   1800
         Width           =   1200
      End
      Begin VB.Label Label31 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Stamina:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   195
         Left            =   240
         TabIndex        =   90
         Top             =   1560
         Width           =   1200
      End
      Begin VB.Label Label30 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Vitality:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   195
         Left            =   240
         TabIndex        =   89
         Top             =   1320
         Width           =   1200
      End
      Begin VB.Label Label29 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Defense:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   195
         Left            =   240
         TabIndex        =   88
         Top             =   1080
         Width           =   1200
      End
      Begin VB.Label Label28 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Strength:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   195
         Left            =   240
         TabIndex        =   87
         Top             =   840
         Width           =   1200
      End
      Begin VB.Label lblSTR 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "XXX"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   1560
         TabIndex        =   86
         Top             =   840
         Width           =   1380
      End
      Begin VB.Label lblDEF 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "XXX"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   1560
         TabIndex        =   85
         Top             =   1080
         Width           =   1380
      End
      Begin VB.Label lblMAGI 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "XXX"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   1560
         TabIndex        =   84
         Top             =   1800
         Width           =   1380
      End
      Begin VB.Label lblSPEED 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "XXX"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   1560
         TabIndex        =   83
         Top             =   1560
         Width           =   1380
      End
   End
   Begin VB.PictureBox ScreenShot 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   73
      Top             =   9120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox picInv3 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Height          =   2505
      Left            =   240
      ScaleHeight     =   167
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   209
      TabIndex        =   1
      Top             =   3480
      Width           =   3135
      Begin VB.PictureBox Up 
         BorderStyle     =   0  'None
         Height          =   270
         Left            =   1200
         Picture         =   "frmMirage.frx":29764
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   71
         Top             =   2235
         Width           =   270
      End
      Begin VB.PictureBox Down 
         BorderStyle     =   0  'None
         Height          =   270
         Left            =   1560
         Picture         =   "frmMirage.frx":299FC
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   70
         Top             =   2235
         Width           =   270
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   330
         Left            =   3120
         Max             =   3
         TabIndex        =   69
         Top             =   2520
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox Picture8 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2175
         Left            =   0
         ScaleHeight     =   145
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   209
         TabIndex        =   43
         Top             =   0
         Width           =   3135
         Begin VB.PictureBox Picture9 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   3135
            Left            =   0
            ScaleHeight     =   209
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   209
            TabIndex        =   44
            Top             =   0
            Width           =   3135
            Begin VB.PictureBox picInv 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   8
               Left            =   1920
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   68
               Top             =   720
               Width           =   480
            End
            Begin VB.PictureBox picInv 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   7
               Left            =   1320
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   67
               Top             =   720
               Width           =   480
            End
            Begin VB.PictureBox picInv 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   6
               Left            =   720
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   66
               Top             =   720
               Width           =   480
            End
            Begin VB.PictureBox picInv 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   5
               Left            =   120
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   65
               Top             =   720
               Width           =   480
            End
            Begin VB.PictureBox picInv 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   4
               Left            =   2520
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   64
               Top             =   120
               Width           =   480
            End
            Begin VB.PictureBox picInv 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   3
               Left            =   1920
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   63
               Top             =   120
               Width           =   480
            End
            Begin VB.PictureBox picInv 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   2
               Left            =   1320
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   62
               Top             =   120
               Width           =   480
            End
            Begin VB.PictureBox picInv 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   1
               Left            =   720
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   61
               Top             =   120
               Width           =   480
            End
            Begin VB.PictureBox picInv 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   0
               Left            =   120
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   60
               Top             =   120
               Width           =   480
            End
            Begin VB.PictureBox picInv 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   9
               Left            =   2520
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   59
               Top             =   720
               Width           =   480
            End
            Begin VB.PictureBox picInv 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   10
               Left            =   120
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   58
               Top             =   1320
               Width           =   480
            End
            Begin VB.PictureBox picInv 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   11
               Left            =   720
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   57
               Top             =   1320
               Width           =   480
            End
            Begin VB.PictureBox picInv 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   12
               Left            =   1320
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   56
               Top             =   1320
               Width           =   480
            End
            Begin VB.PictureBox picInv 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   13
               Left            =   1920
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   55
               Top             =   1320
               Width           =   480
            End
            Begin VB.PictureBox picInv 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   14
               Left            =   2520
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   54
               Top             =   1320
               Width           =   480
            End
            Begin VB.PictureBox picInv 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   15
               Left            =   120
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   53
               Top             =   1920
               Width           =   480
            End
            Begin VB.PictureBox picInv 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   16
               Left            =   720
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   52
               Top             =   1920
               Width           =   480
            End
            Begin VB.PictureBox picInv 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   17
               Left            =   1320
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   51
               Top             =   1920
               Width           =   480
            End
            Begin VB.PictureBox picInv 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   18
               Left            =   1920
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   50
               Top             =   1920
               Width           =   480
            End
            Begin VB.PictureBox picInv 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   19
               Left            =   2520
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   49
               Top             =   1920
               Width           =   480
            End
            Begin VB.PictureBox picInv 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   20
               Left            =   120
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   48
               Top             =   2520
               Width           =   480
            End
            Begin VB.PictureBox picInv 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   21
               Left            =   720
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   47
               Top             =   2520
               Width           =   480
            End
            Begin VB.PictureBox picInv 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   22
               Left            =   1320
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   46
               Top             =   2520
               Width           =   480
            End
            Begin VB.PictureBox picInv 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   480
               Index           =   23
               Left            =   1920
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   45
               Top             =   2520
               Width           =   480
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
      Begin VB.ListBox lstInv 
         Appearance      =   0  'Flat
         BackColor       =   &H0084ADB3&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   2190
         ItemData        =   "frmMirage.frx":29C87
         Left            =   3000
         List            =   "frmMirage.frx":29C89
         TabIndex        =   2
         Top             =   2520
         Visible         =   0   'False
         Width           =   2625
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         X1              =   4
         X2              =   208
         Y1              =   146
         Y2              =   146
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
         ForeColor       =   &H8000000F&
         Height          =   210
         Left            =   2280
         TabIndex        =   4
         Top             =   2280
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
         ForeColor       =   &H8000000F&
         Height          =   210
         Left            =   15
         TabIndex        =   3
         Top             =   2280
         Width           =   690
      End
   End
   Begin VB.PictureBox picPlayerSpells 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Height          =   2505
      Left            =   240
      ScaleHeight     =   167
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   209
      TabIndex        =   5
      Top             =   3480
      Width           =   3135
      Begin VB.ListBox lstSpells 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFC0&
         Height          =   2190
         ItemData        =   "frmMirage.frx":29C8B
         Left            =   90
         List            =   "frmMirage.frx":29C8D
         TabIndex        =   6
         Top             =   75
         Width           =   2940
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
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1440
         TabIndex        =   7
         Top             =   2325
         Width           =   375
      End
   End
   Begin VB.PictureBox picWhosOnline 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2505
      Left            =   240
      ScaleHeight     =   167
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   209
      TabIndex        =   16
      Top             =   3480
      Width           =   3135
      Begin VB.ListBox lstOnline 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFC0&
         Height          =   2340
         Left            =   0
         TabIndex        =   17
         Top             =   0
         Width           =   2940
      End
   End
   Begin VB.PictureBox picGuildAdmin 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2505
      Left            =   240
      ScaleHeight     =   2505
      ScaleWidth      =   3135
      TabIndex        =   18
      Top             =   3480
      Width           =   3135
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
         Left            =   825
         TabIndex        =   21
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
         Left            =   825
         TabIndex        =   20
         Top             =   345
         Width           =   1575
      End
      Begin VB.CommandButton cmdAccess 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
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
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   840
         Width           =   1215
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
         Left            =   255
         TabIndex        =   23
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
         ForeColor       =   &H00FFFFFF&
         Height          =   165
         Left            =   285
         TabIndex        =   22
         Top             =   360
         Width           =   420
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2505
      Left            =   240
      ScaleHeight     =   2505
      ScaleWidth      =   3135
      TabIndex        =   25
      Top             =   3480
      Width           =   3135
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
         ForeColor       =   &H80000005&
         Height          =   180
         Left            =   1200
         TabIndex        =   30
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
         ForeColor       =   &H00FFFFC0&
         Height          =   165
         Left            =   1425
         TabIndex        =   29
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
         ForeColor       =   &H00FFFFC0&
         Height          =   165
         Left            =   1425
         TabIndex        =   28
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
         ForeColor       =   &H80000005&
         Height          =   165
         Left            =   570
         TabIndex        =   27
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
         ForeColor       =   &H80000005&
         Height          =   165
         Left            =   480
         TabIndex        =   26
         Top             =   645
         Width           =   825
      End
   End
   Begin VB.PictureBox picEquip 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2505
      Left            =   240
      ScaleHeight     =   2505
      ScaleWidth      =   3135
      TabIndex        =   31
      Top             =   3480
      Width           =   3135
      Begin VB.PictureBox AmuletImage 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   2040
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   131
         Top             =   1080
         Width           =   495
      End
      Begin VB.PictureBox RingImage 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   1560
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   130
         Top             =   600
         Width           =   495
      End
      Begin VB.PictureBox GlovesImage 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   600
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   129
         Top             =   960
         Width           =   495
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1650
         Left            =   480
         Picture         =   "frmMirage.frx":29C8F
         ScaleHeight     =   1650
         ScaleWidth      =   2085
         TabIndex        =   34
         Top             =   390
         Width           =   2085
         Begin VB.PictureBox BootsImage 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   0
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   128
            Top             =   0
            Width           =   495
         End
         Begin VB.PictureBox HelmetImage 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   585
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   39
            Top             =   90
            Width           =   495
         End
         Begin VB.PictureBox ArmorImage 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   1545
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   37
            Top             =   225
            Width           =   495
         End
         Begin VB.PictureBox ShieldImage 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   105
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   36
            Top             =   1035
            Width           =   495
         End
         Begin VB.PictureBox WeaponImage 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   1560
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   35
            Top             =   1035
            Width           =   495
         End
         Begin VB.PictureBox Picture4 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   555
            Left            =   555
            ScaleHeight     =   525
            ScaleWidth      =   525
            TabIndex        =   38
            Top             =   60
            Width           =   555
         End
         Begin VB.PictureBox Picture5 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H80000008&
            Height          =   555
            Left            =   1515
            ScaleHeight     =   525
            ScaleWidth      =   525
            TabIndex        =   40
            Top             =   195
            Width           =   555
         End
         Begin VB.PictureBox Picture6 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H80000008&
            Height          =   555
            Left            =   1530
            ScaleHeight     =   525
            ScaleWidth      =   525
            TabIndex        =   41
            Top             =   1005
            Width           =   555
         End
         Begin VB.PictureBox Picture7 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   555
            Left            =   75
            ScaleHeight     =   525
            ScaleWidth      =   525
            TabIndex        =   42
            Top             =   1005
            Width           =   555
         End
      End
      Begin VB.PictureBox picItems 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2.25000e5
         Left            =   2400
         Picture         =   "frmMirage.frx":35149
         ScaleHeight     =   2.23636e5
         ScaleMode       =   0  'User
         ScaleWidth      =   477.091
         TabIndex        =   32
         Top             =   2760
         Visible         =   0   'False
         Width           =   480
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   240
      TabIndex        =   14
      Top             =   8520
      Width           =   11490
   End
   Begin MSWinsockLib.Winsock Socket 
      Left            =   120
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox txtChat 
      Height          =   1455
      Left            =   240
      TabIndex        =   0
      Top             =   7080
      Width           =   11490
      _ExtentX        =   20267
      _ExtentY        =   2566
      _Version        =   393217
      BackColor       =   0
      BorderStyle     =   0
      Enabled         =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmMirage.frx":194A8B
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSWinsockLib.Winsock MapSocket 
      Left            =   600
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Guild"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   225
      Left            =   4410
      TabIndex        =   122
      Top             =   6645
      Width           =   1140
   End
   Begin VB.Label Label24 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "_"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   11505
      TabIndex        =   120
      Top             =   15
      Width           =   180
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Refresh"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   225
      Left            =   8190
      TabIndex        =   119
      Top             =   6645
      Width           =   1140
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   225
      Left            =   10710
      TabIndex        =   118
      Top             =   6645
      Width           =   1140
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Options"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   225
      Left            =   9450
      TabIndex        =   117
      Top             =   6645
      Width           =   1140
   End
   Begin VB.Label lblWhosOnline 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Whos Online"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   225
      Left            =   6930
      TabIndex        =   116
      Top             =   6645
      Width           =   1140
   End
   Begin VB.Label lblItemName 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   180
      Left            =   720
      TabIndex        =   115
      Top             =   6120
      Width           =   2700
   End
   Begin VB.Label lblItemNameS 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
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
      Height          =   195
      Left            =   240
      TabIndex        =   114
      Top             =   6120
      Width           =   465
   End
   Begin VB.Label lblGold 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H0000C000&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   195
      Left            =   2550
      TabIndex        =   113
      Top             =   960
      Width           =   105
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Gold:"
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
      Height          =   195
      Left            =   2160
      TabIndex        =   112
      Top             =   960
      Width           =   375
   End
   Begin VB.Label Label36 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   480
      TabIndex        =   111
      Top             =   2940
      Width           =   495
   End
   Begin VB.Label lblMapName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "MapName"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   195
      Left            =   4290
      TabIndex        =   103
      Top             =   75
      Width           =   7080
   End
   Begin VB.Label lblName 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H0000C000&
      BackStyle       =   0  'Transparent
      Caption         =   "XXX"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   195
      Left            =   840
      TabIndex        =   81
      Top             =   960
      Width           =   360
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Stamina:"
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
      Height          =   195
      Left            =   360
      TabIndex        =   80
      Top             =   1680
      Width           =   630
   End
   Begin VB.Shape shpSP 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   135
      Left            =   360
      Top             =   1920
      Width           =   2865
   End
   Begin VB.Label lblSP 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      BackStyle       =   0  'Transparent
      Caption         =   "XX/XX"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   180
      Left            =   1080
      TabIndex        =   79
      Top             =   1680
      Width           =   2130
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Experience:"
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
      Height          =   195
      Left            =   360
      TabIndex        =   78
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mana:"
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
      Height          =   195
      Left            =   360
      TabIndex        =   77
      Top             =   2040
      Width           =   450
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Health:"
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
      Height          =   195
      Left            =   360
      TabIndex        =   76
      Top             =   1320
      Width           =   525
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
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
      Height          =   195
      Left            =   360
      TabIndex        =   75
      Top             =   960
      Width           =   465
   End
   Begin VB.Label desc 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
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
      Height          =   420
      Left            =   240
      TabIndex        =   74
      Top             =   6330
      Width           =   3120
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Help"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   225
      Left            =   5670
      TabIndex        =   72
      Top             =   6645
      Width           =   1140
   End
   Begin VB.Label Label19 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   1875
      TabIndex        =   33
      Top             =   2940
      Width           =   495
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   10755
      TabIndex        =   24
      Top             =   6180
      Width           =   45
   End
   Begin VB.Shape shpTNL 
      BorderColor     =   &H00C0C000&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00C0C000&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   360
      Top             =   2640
      Width           =   2865
   End
   Begin VB.Label lblEXP 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "XX/XX"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   180
      Left            =   1320
      TabIndex        =   13
      Top             =   2400
      Width           =   1905
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   2580
      TabIndex        =   12
      Top             =   2940
      Width           =   495
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   1155
      TabIndex        =   11
      Top             =   2940
      Width           =   495
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
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
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   11760
      TabIndex        =   10
      Top             =   15
      Width           =   180
   End
   Begin VB.Label lblMP 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00CB884B&
      BackStyle       =   0  'Transparent
      Caption         =   "XX/XX"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   180
      Left            =   1080
      TabIndex        =   9
      Top             =   2040
      Width           =   2130
   End
   Begin VB.Label lblHP 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      BackStyle       =   0  'Transparent
      Caption         =   "XX/XX"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   180
      Left            =   1080
      TabIndex        =   8
      Top             =   1320
      Width           =   2130
   End
   Begin VB.Shape shpHP 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   135
      Left            =   360
      Top             =   1560
      Width           =   2865
   End
   Begin VB.Shape shpMP 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   135
      Left            =   360
      Top             =   2280
      Width           =   2865
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   135
      Left            =   360
      Top             =   1560
      Width           =   2865
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   135
      Left            =   360
      Top             =   1920
      Width           =   2865
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   135
      Left            =   360
      Top             =   2280
      Width           =   2865
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   135
      Left            =   360
      Top             =   2640
      Width           =   2865
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
    Unload Me
End Sub

Private Sub Check1_Click()
Dim FileName As String
FileName = App.Path & "\Input.ini"

    If Check1.Value = Checked Then
        WriteINI "CONFIG", "Joypad", 1, FileName
        ID = True
    Else
        WriteINI "CONFIG", "Joypad", 0, FileName
        ID = False
    End If
End Sub

Private Sub chksound_Click()
    WriteINI "CONFIG", "Sound", chksound.Value, App.Path & "\config.ini"
End Sub

Private Sub chkplayername_Click()
    WriteINI "CONFIG", "PlayerName", chkplayername.Value, App.Path & "\config.ini"
End Sub

Private Sub chkmusic_Click()
    WriteINI "CONFIG", "Music", chkmusic.Value, App.Path & "\config.ini"
    If ReadINI("CONFIG", "Music", App.Path & "\config.ini") = 1 Then
        Call PlayMidi(Trim(CheckMap(GetPlayerMap(MyIndex)).Music))
    Else
        Call StopMidi
    End If
End Sub

Private Sub chkTime_Click()
    If chkTime.Value = Checked Then
        WriteINI "UC", "Day/Night", 1, App.Path & "\UC.ini"
    Else
        WriteINI "UC", "Day/Night", 0, App.Path & "\UC.ini"
    End If
End Sub

Private Sub chkWeather_Click()
    If chkWeather.Value = Checked Then
        WriteINI "UC", "Weather", 1, App.Path & "\UC.ini"
    Else
        WriteINI "UC", "Weather", 0, App.Path & "\UC.ini"
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

Private Sub Command1_Click()
    picOptions.Visible = False
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    SOffsetX = x
    SOffsetY = y
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call MoveForm(frmMirage, Button, Shift, x, y)
End Sub

Private Sub Form_Terminate()
    Call GameDestroy
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call GameDestroy
End Sub

Private Sub KeepNotes_Click()
frmKeepNotes.Visible = True
End Sub

Private Sub Label1_Click()
DoEvents
    frmHelp.lstBasics.Text = ""
    frmHelp.lstChat.Text = ""
    frmHelp.lstGuildParty.Text = ""
    frmHelp.lstAdmin.Text = ""

    AddHelp frmHelp.lstBasics, ":: Basic Controls ::", BrightGreen
    AddHelp frmHelp.lstBasics, "-Use the arrows keys to move (Up, Down, Left, Right).", &HFFFFC0
    AddHelp frmHelp.lstBasics, "-Use CTRL to attack enemies.", &HFFFFC0
    AddHelp frmHelp.lstBasics, "-Hit enter over items to pick them up.", &HFFFFC0
    AddHelp frmHelp.lstBasics, "-Click on players/monsters/items to view their information.", &HFFFFC0
    AddHelp frmHelp.lstBasics, "", &HFFFFC0
    AddHelp frmHelp.lstBasics, ":: Basic Help ::", BrightGreen
    AddHelp frmHelp.lstBasics, "-Equip your newbie items first before going out into the wilderness.", &HFFFFC0
    AddHelp frmHelp.lstBasics, "-To open up a shop, step on the flashing stars.", &HFFFFC0
    
    AddHelp frmHelp.lstChat, ":: Chat Commands ::", BrightGreen
    AddHelp frmHelp.lstChat, "-Type ' before your text to broadcast.", &HFFFFC0
    AddHelp frmHelp.lstChat, "-Type - before your text to emote chat.", &HFFFFC0
    AddHelp frmHelp.lstChat, "-Type @ or /gchat before your text to talk with your guild.", &HFFFFC0
    AddHelp frmHelp.lstChat, "-Type # or /pchat before your text to talk with your party.", &HFFFFC0
    AddHelp frmHelp.lstChat, "-Type !<playername> before your text to message a player.", &HFFFFC0
    AddHelp frmHelp.lstChat, "-Type /emotes to display all emoticon commands.", &HFFFFC0
    AddHelp frmHelp.lstChat, "-Type /train in training house to train your stats.", &HFFFFC0
    AddHelp frmHelp.lstChat, "-Type /calladmins if your in need of an admin.", &HFFFFC0
    AddHelp frmHelp.lstChat, "-Type /who to find out whos online.", &HFFFFC0
    AddHelp frmHelp.lstChat, "-Type /inv to open your inventory.", &HFFFFC0
    AddHelp frmHelp.lstChat, "-Type /stats to display your stats.", &HFFFFC0
    AddHelp frmHelp.lstChat, "-Type /refresh if your stuck.", &HFFFFC0
    AddHelp frmHelp.lstChat, "-Type /chat to accept a chat.", &HFFFFC0
    AddHelp frmHelp.lstChat, "-Type /chatdecline to decline a chat.", &HFFFFC0
    AddHelp frmHelp.lstChat, "-Type /trade <playername> to trade with a player.", &HFFFFC0
    AddHelp frmHelp.lstChat, "-Type /accept to accept a trade.", &HFFFFC0
    AddHelp frmHelp.lstChat, "-Type /decline to decline a trade.", &HFFFFC0
    AddHelp frmHelp.lstChat, "-Type /pass <newpassword> to change current password.", &HFFFFC0
    AddHelp frmHelp.lstChat, "-Type /reply <messagehere> or /r <messagehere> to reply to a player.", &HFFFFC0
    
    AddHelp frmHelp.lstGuildParty, ":: Guild Commands ::", BrightGreen
    AddHelp frmHelp.lstGuildParty, "-Type /createguild <guildname> to create a guild. Cost 5000 Gold.", &HFFFFC0
    AddHelp frmHelp.lstGuildParty, "-Type /guildinvite <playername> to invite someone into your guild.", &HFFFFC0
    AddHelp frmHelp.lstGuildParty, "-Type /guildaccept to accept guild invitation.", &HFFFFC0
    AddHelp frmHelp.lstGuildParty, "-Type /guilddecline to decline guild invitation.", &HFFFFC0
    AddHelp frmHelp.lstGuildParty, "-Type /guildkick <playername> to kick someone from guild.", &HFFFFC0
    AddHelp frmHelp.lstGuildParty, "-Type /guildleave to leave current guild.", &HFFFFC0
    AddHelp frmHelp.lstGuildParty, "-Type /guildwho to see who's online in the guild.", &HFFFFC0
    AddHelp frmHelp.lstGuildParty, "", &HFFFFC0
    AddHelp frmHelp.lstGuildParty, ":: Party Commands ::", BrightGreen
    AddHelp frmHelp.lstGuildParty, "-Type /party <playername> to party with a player.", &HFFFFC0
    AddHelp frmHelp.lstGuildParty, "-Type /join to accept a party invitation.", &HFFFFC0
    AddHelp frmHelp.lstGuildParty, "-Type /leave to leave a party.", &HFFFFC0
    
    AddHelp frmHelp.lstAdmin, ":: Admin Commands ::", BrightGreen
    If GetPlayerAccess(MyIndex) <= 0 Then AddHelp frmHelp.lstAdmin, "You arent enough access to view this!", &HFFFFC0
    If GetPlayerAccess(MyIndex) >= 1 Then
        AddHelp frmHelp.lstAdmin, "-Type /jail <playername> to jail a player!", &HFFFFC0
        AddHelp frmHelp.lstAdmin, "-Type /unjail <playername> to unjail a player!", &HFFFFC0
        AddHelp frmHelp.lstAdmin, "-Type /ban <playername> to ban certain player.", &HFFFFC0
        AddHelp frmHelp.lstAdmin, "-Type /banlist to view banlist.", &HFFFFC0
        AddHelp frmHelp.lstAdmin, "-Type /mute <playername> to mute/unmute a player.", &HFFFFC0
        AddHelp frmHelp.lstAdmin, "-Type /massmute mute broadcast for all players.", &HFFFFC0
        AddHelp frmHelp.lstAdmin, "-Type /kick <playername> kick player from game.", &HFFFFC0
        AddHelp frmHelp.lstAdmin, "-Type /daynight to switch the game time.", &HFFFFC0
        AddHelp frmHelp.lstAdmin, "-Type /weather <none(0)/rain(1)/snow(2)/thunder(3)> to switch the game weather.", &HFFFFC0
        AddHelp frmHelp.lstAdmin, "-Type /intensity <1 - 50> to change the intensity of the weather.", &HFFFFC0
        AddHelp frmHelp.lstAdmin, "-Type "" to use global chat.", &HFFFFC0
        AddHelp frmHelp.lstAdmin, "-Type = to use admin chat.", &HFFFFC0
        AddHelp frmHelp.lstAdmin, "-Type /warpmeto <playername> to warp to a player.", &HFFFFC0
        AddHelp frmHelp.lstAdmin, "-Type /warptome <playername> to warp a player to you.", &HFFFFC0
    End If
    If GetPlayerAccess(MyIndex) >= 2 Then
        AddHelp frmHelp.lstAdmin, "-Type /loc to find your location.", &HFFFFC0
        AddHelp frmHelp.lstAdmin, "-Type /mapeditor to open up the mapeditor.", &HFFFFC0
        AddHelp frmHelp.lstAdmin, "-Type /mapreport to display a visual report of all the maps.", &HFFFFC0
        AddHelp frmHelp.lstAdmin, "-Type /warpto <mapnumber> to warp to a map.", &HFFFFC0
        AddHelp frmHelp.lstAdmin, "-Type /setsprite <spritenumber> to change your sprite.", &HFFFFC0
        AddHelp frmHelp.lstAdmin, "-Type /respawn to respawn the current map.", &HFFFFC0
        AddHelp frmHelp.lstAdmin, "-Type /motd <motdhere> to change the current MOTD.", &HFFFFC0
    End If
    If GetPlayerAccess(MyIndex) >= 3 Then
        AddHelp frmHelp.lstAdmin, "-Type /edititem to edit items.", &HFFFFC0
        AddHelp frmHelp.lstAdmin, "-Type /editemoticon to edit emoticons.", &HFFFFC0
        AddHelp frmHelp.lstAdmin, "-Type /editnpc to edit NPCs.", &HFFFFC0
        AddHelp frmHelp.lstAdmin, "-Type /editshop to edit shops.", &HFFFFC0
        AddHelp frmHelp.lstAdmin, "-Type /editspell to edit spells.", &HFFFFC0
    End If
    If GetPlayerAccess(MyIndex) >= 4 Then
        AddHelp frmHelp.lstAdmin, "-Type /masswarp to warp everyone to you.", &HFFFFC0
        AddHelp frmHelp.lstAdmin, "-Type /setaccess <playername> <access> to set player's access.", &HFFFFC0
        AddHelp frmHelp.lstAdmin, "-Type /destroybanlist to destroy the banlist.", &HFFFFC0
    End If

    frmHelp.Visible = True
End Sub

Private Sub Label13_Click()
    frmMirage.lblGuild.Caption = GetPlayerGuild(MyIndex)
    frmMirage.lblRank.Caption = GetPlayerGuildAccess(MyIndex)
    Picture1.ZOrder (0)
End Sub

Private Sub Label19_Click()
    picEquip.ZOrder (0)
    Call UpdateVisInv
End Sub

Private Sub label2_Click()
    picOptions.Visible = True
End Sub

Private Sub Label20_Click()
    Call SendData("refresh" & SEP_CHAR & END_CHAR)
End Sub

Private Sub Label21_Click()
picEquip.Visible = False
End Sub

Private Sub Label24_Click()
    frmMirage.WindowState = 1
End Sub

Private Sub Label26_Click()
    frmInput.Show vbModeless, frmMirage
End Sub

Private Sub Label3_Click()
Call GameDestroy
End Sub

Private Sub Label36_Click()
    picStats.ZOrder (0)
End Sub

Private Sub Label4_Click()
    frmMirage.picGuildAdmin.ZOrder (0)
End Sub

Private Sub Label6_Click()
    Call SendData("getstats" & SEP_CHAR & END_CHAR)
End Sub

Private Sub Label7_Click()
    Call UpdateVisInv
    picInv3.ZOrder (0)
    
    If Player(MyIndex).Inv(Inventory).Num > 0 Then
        If Item(GetPlayerInvItemNum(MyIndex, Inventory)).Type = ITEM_TYPE_CURRENCY Then
            lblItemName.Caption = Trim(Item(GetPlayerInvItemNum(MyIndex, Inventory)).name) & " (" & GetPlayerInvItemValue(MyIndex, Inventory) & ")"
        Else
            lblItemName.Caption = Trim(Item(GetPlayerInvItemNum(MyIndex, Inventory)).name)
        End If
        
        desc.Caption = Trim(Item(GetPlayerInvItemNum(MyIndex, Inventory)).desc)
    Else
        lblItemName.Caption = ""
        desc.Caption = ""
    End If
End Sub

Private Sub Label8_Click()
    Call SendData("spells" & SEP_CHAR & END_CHAR)
End Sub

Private Sub lblCloseOnline_Click()
Call SendOnlineList
picWhosOnline.Visible = False
End Sub

Private Sub lblClosePicGuildAdmin_Click()
    picGuildAdmin.Visible = False
End Sub

Private Sub lblOptionsCancel_Click()
    picOptions.Visible = False
End Sub

Private Sub lblGold_Click()
Dim d As Long
    For d = 1 To MAX_INV
        If GetPlayerInvItemNum(MyIndex, d) > 0 Then
            If Item(GetPlayerInvItemNum(MyIndex, d)).Type = ITEM_TYPE_CURRENCY Then
                If GetPlayerInvItemNum(MyIndex, d) = 1 Then
                    Inventory = d
                    Call DropItems
                    Exit Sub
                End If
            End If
        End If
    Next d
End Sub

Private Sub lblMapName_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    SOffsetX = x
    SOffsetY = y
End Sub

Private Sub lblMapName_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call MoveForm(frmMirage, Button, Shift, x, y)
End Sub

Private Sub lblWhosOnline_Click()
    Call SendOnlineList
    picWhosOnline.ZOrder (0)
End Sub

Private Sub lstOnline_DblClick()
    Call SendData("playerchat" & SEP_CHAR & Trim(lstOnline.Text) & SEP_CHAR & END_CHAR)
End Sub

Private Sub lstSpells_DblClick()
    If Player(MyIndex).Spell(lstSpells.ListIndex + 1) > 0 Then
        SpellMemorized = lstSpells.ListIndex + 1
        Call AddText("Successfully memorized spell!", White)
    Else
        Call AddText("No spell here to memorize.", BrightRed)
    End If
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
Dim d As Long
d = Index

    Inventory = Index + 1
    frmMirage.SelectedItem.Top = frmMirage.picInv(Inventory - 1).Top - 1
    frmMirage.SelectedItem.Left = frmMirage.picInv(Inventory - 1).Left - 1
    
    If Button = 1 Then
        Call UpdateVisInv
    ElseIf Button = 2 Then
        Call DropItems
    End If
    
    If Player(MyIndex).Inv(d + 1).Num > 0 Then
        If Item(GetPlayerInvItemNum(MyIndex, d + 1)).Type = ITEM_TYPE_CURRENCY Then
            lblItemName.Caption = Trim(Item(GetPlayerInvItemNum(MyIndex, d + 1)).name) & " (" & GetPlayerInvItemValue(MyIndex, d + 1) & ")"
        Else
            lblItemName.Caption = Trim(Item(GetPlayerInvItemNum(MyIndex, d + 1)).name)
        End If
        
        desc.Caption = Trim(Item(GetPlayerInvItemNum(MyIndex, d + 1)).desc)
    Else
        lblItemName.Caption = ""
        desc.Caption = ""
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Dim d As Long, i As Long
Dim ii As Long

    Call CheckInput(0, KeyCode, Shift)
    If KeyCode = vbKeyF1 Then
        If Player(MyIndex).Access > 0 Then
            frmadmin.Visible = False
            frmadmin.Visible = True
        End If
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
            
            DoEvents
        Loop Until ii = 1
    ElseIf KeyCode = vbKeyF12 Then
        ScreenShot.Picture = CaptureArea(frmMirage, 8, 8, 634, 478)
        i = 0
        ii = 0
        Do
            If FileExist("Screenshot" & i & ".bmp") = True Then
                i = i + 1
            Else
                Call SavePicture(ScreenShot.Picture, App.Path & "\Screenshot" & i & ".bmp")
                ii = 1
            End If
            
            DoEvents
        Loop Until ii = 1
    End If
    
    If KeyCode = vbKeyEnd Then
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

Private Sub picScreen_GotFocus()
    txtMyTextBox.SetFocus
End Sub

Private Sub picScreen_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim i As Long

    If (Button = 1 Or Button = 2) And InEditor = True Then
        Call EditorMouseDown(Button, Shift, (x + (NewPlayerX * PIC_X)), (y + (NewPlayerY * PIC_Y)))
    End If
    
    If Button = 1 And InEditor = False Then
        Call PlayerSearch(Button, Shift, (x + (NewPlayerX * PIC_X)), (y + (NewPlayerY * PIC_Y)))
    End If
End Sub

Private Sub picScreen_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If (Button = 1 Or Button = 2) And InEditor = True Then
        Call EditorMouseDown(Button, Shift, (x + (NewPlayerX * PIC_X)), (y + (NewPlayerY * PIC_Y)))
    End If
    frmMapEditor.lblXY.Caption = "Cursor - X:" & Int((x + (NewPlayerX * PIC_X)) / PIC_X) & " - Y:" & Int((y + (NewPlayerY * PIC_Y)) / PIC_Y)
End Sub

Private Sub Socket_DataArrival(ByVal bytesTotal As Long)
    If IsConnected Then
        Call IncomingData(bytesTotal)
    End If
End Sub

Private Sub MapSocket_DataArrival(ByVal bytesTotal As Long)
    If IsMapConnected Then
        Call IncomingMapData(bytesTotal)
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Call HandleKeypresses(KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
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
    txtMyTextBox.SetFocus
    'frmMirage.picScreen.SetFocus
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
        If Item(GetPlayerInvItemNum(MyIndex, InvNum)).Type = ITEM_TYPE_CURRENCY Then
            GoldAmount = InputBox("How much " & Trim(Item(GetPlayerInvItemNum(MyIndex, InvNum)).name) & "(" & GetPlayerInvItemValue(MyIndex, InvNum) & ") would you like to drop?", "Drop " & Trim(Item(GetPlayerInvItemNum(MyIndex, InvNum)).name), 0, frmMirage.Left, frmMirage.Top)
            If IsNumeric(GoldAmount) Then
                Call SendDropItem(InvNum, GoldAmount)
            End If
        Else
            Call SendDropItem(InvNum, 0)
        End If
    End If
   
    picInv(InvNum - 1).Picture = LoadPicture()
    Call UpdateVisInv
    Exit Sub
Done:
    If Item(GetPlayerInvItemNum(MyIndex, InvNum)).Type = ITEM_TYPE_CURRENCY Then
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
If VScroll1.Value = 2 Then Exit Sub
    VScroll1.Value = VScroll1.Value + 1
    Picture9.Top = VScroll1.Value * -PIC_Y
End Sub
