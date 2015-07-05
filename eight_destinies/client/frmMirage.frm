VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCN.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMirage 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Main Game"
   ClientHeight    =   8985
   ClientLeft      =   315
   ClientTop       =   345
   ClientWidth     =   12600
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
   Picture         =   "frmMirage.frx":08CA
   ScaleHeight     =   599
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   840
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.PictureBox picOptions 
      Appearance      =   0  'Flat
      BackColor       =   &H00789298&
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
      Height          =   2745
      Left            =   7470
      ScaleHeight     =   181
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   141
      TabIndex        =   19
      Top             =   4515
      Width           =   2145
      Visible         =   0   'False
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
         Left            =   1560
         TabIndex        =   110
         Top             =   2400
         Width           =   495
      End
      Begin VB.CheckBox chksound 
         BackColor       =   &H00789298&
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
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Left            =   120
         TabIndex        =   80
         Top             =   2400
         Value           =   1  'Checked
         Width           =   765
      End
      Begin VB.CheckBox chkmusic 
         BackColor       =   &H00789298&
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
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Left            =   120
         TabIndex        =   78
         Top             =   2160
         Value           =   1  'Checked
         Width           =   735
      End
      Begin VB.CheckBox chknpcdamage 
         BackColor       =   &H00789298&
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
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Left            =   120
         TabIndex        =   75
         Top             =   1440
         Value           =   1  'Checked
         Width           =   1725
      End
      Begin VB.CheckBox chkplayerdamage 
         BackColor       =   &H00789298&
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
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   74
         Top             =   480
         Value           =   1  'Checked
         Width           =   1725
      End
      Begin VB.CheckBox chknpcbar 
         BackColor       =   &H00789298&
         Caption         =   "Show HP Bar"
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
         Height          =   225
         Left            =   120
         TabIndex        =   46
         Top             =   1680
         Value           =   1  'Checked
         Width           =   1560
      End
      Begin VB.CheckBox chknpcname 
         BackColor       =   &H00789298&
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
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Left            =   120
         TabIndex        =   22
         Top             =   1200
         Value           =   1  'Checked
         Width           =   765
      End
      Begin VB.CheckBox chkplayername 
         BackColor       =   &H00789298&
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
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Value           =   1  'Checked
         Width           =   765
      End
      Begin VB.CheckBox chkplayerbar 
         BackColor       =   &H00789298&
         Caption         =   "Show HP Bar"
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
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   690
         Value           =   1  'Checked
         Width           =   1200
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
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   -240
         TabIndex        =   109
         Top             =   1920
         Width           =   2655
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
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   -240
         TabIndex        =   108
         Top             =   960
         Width           =   2655
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
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   -240
         TabIndex        =   107
         Top             =   0
         Width           =   2655
      End
   End
   Begin VB.Timer tmrSnowDrop 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   12000
      Top             =   6960
   End
   Begin VB.Timer tmrRainDrop 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   11520
      Top             =   6960
   End
   Begin VB.PictureBox itmDesc 
      Appearance      =   0  'Flat
      BackColor       =   &H00789298&
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
      Height          =   3615
      Left            =   7560
      ScaleHeight     =   239
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   135
      TabIndex        =   82
      Top             =   105
      Width           =   2055
      Visible         =   0   'False
      Begin VB.Label descLvl 
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "Level"
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
         Height          =   135
         Left            =   120
         TabIndex        =   115
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label descDUR 
         BackStyle       =   0  'Transparent
         Caption         =   "Durability:"
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
         Height          =   255
         Left            =   120
         TabIndex        =   114
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label descMagi 
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "INT"
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
         Height          =   135
         Left            =   120
         TabIndex        =   112
         Top             =   1170
         Width           =   1215
      End
      Begin VB.Label descMS 
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "DEX: XXXX INT: XXXX"
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
         Height          =   165
         Left            =   120
         TabIndex        =   93
         Top             =   2100
         Width           =   1695
      End
      Begin VB.Label Label10 
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "DESCRIPTION:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   165
         Left            =   120
         TabIndex        =   92
         Top             =   2340
         Width           =   1095
      End
      Begin VB.Label desc 
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
         ForeColor       =   &H0000FFFF&
         Height          =   855
         Left            =   120
         TabIndex        =   91
         Top             =   2610
         Width           =   1815
      End
      Begin VB.Label descSD 
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "STR: XXXX CON: XXXX"
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
         Height          =   165
         Left            =   120
         TabIndex        =   90
         Top             =   1950
         Width           =   1695
      End
      Begin VB.Label descHpMp 
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
         ForeColor       =   &H0000FFFF&
         Height          =   165
         Left            =   120
         TabIndex        =   89
         Top             =   1800
         Width           =   1935
      End
      Begin VB.Label Label15 
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "BONUS:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   165
         Left            =   120
         TabIndex        =   88
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label descSpeed 
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "DEX"
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
         Height          =   165
         Left            =   120
         TabIndex        =   87
         Top             =   1020
         Width           =   1215
      End
      Begin VB.Label descDef 
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "CON"
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
         Height          =   165
         Left            =   120
         TabIndex        =   86
         Top             =   870
         Width           =   1215
      End
      Begin VB.Label descStr 
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "STR"
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
         Height          =   165
         Left            =   120
         TabIndex        =   85
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "REQUIREMENTS:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   165
         Left            =   120
         TabIndex        =   84
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label descName 
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
         ForeColor       =   &H0000FFFF&
         Height          =   165
         Left            =   120
         TabIndex        =   83
         Top             =   30
         Width           =   1815
      End
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
      Left            =   12000
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   81
      Top             =   8880
      Width           =   615
      Visible         =   0   'False
   End
   Begin VB.PictureBox picInv3 
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
      Height          =   2505
      Left            =   9840
      ScaleHeight     =   167
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   175
      TabIndex        =   1
      Top             =   3720
      Width           =   2625
      Visible         =   0   'False
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
         Left            =   1980
         Picture         =   "frmMirage.frx":25D0
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   77
         Top             =   2160
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
         Left            =   2280
         Picture         =   "frmMirage.frx":2868
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   76
         Top             =   2160
         Width           =   270
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   330
         Left            =   2640
         Max             =   3
         TabIndex        =   73
         Top             =   2400
         Width           =   255
         Visible         =   0   'False
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
         Height          =   1815
         Left            =   0
         ScaleHeight     =   121
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   177
         TabIndex        =   47
         Top             =   0
         Width           =   2655
         Begin VB.PictureBox Picture9 
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
            Height          =   5535
            Left            =   45
            ScaleHeight     =   369
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   177
            TabIndex        =   48
            Top             =   0
            Width           =   2655
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
               Index           =   35
               Left            =   1920
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   135
               Top             =   4920
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
               Index           =   34
               Left            =   1320
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   134
               Top             =   4920
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
               Index           =   33
               Left            =   720
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   133
               Top             =   4920
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
               Index           =   32
               Left            =   120
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   132
               Top             =   4920
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
               Index           =   31
               Left            =   1920
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   131
               Top             =   4320
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
               Index           =   30
               Left            =   1320
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   130
               Top             =   4320
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
               Index           =   29
               Left            =   720
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   129
               Top             =   4320
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
               Index           =   28
               Left            =   120
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   128
               Top             =   4320
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
               Index           =   27
               Left            =   1920
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   127
               Top             =   3720
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
               Index           =   26
               Left            =   1320
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   126
               Top             =   3720
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
               Index           =   25
               Left            =   720
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   125
               Top             =   3720
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
               Index           =   24
               Left            =   120
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   124
               Top             =   3720
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
               Index           =   8
               Left            =   120
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   72
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
               TabIndex        =   71
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
               TabIndex        =   70
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
               TabIndex        =   69
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
               TabIndex        =   68
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
               TabIndex        =   67
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
               TabIndex        =   66
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
               TabIndex        =   65
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
               TabIndex        =   64
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
               TabIndex        =   63
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
               TabIndex        =   62
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
               TabIndex        =   61
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
               TabIndex        =   60
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
               TabIndex        =   59
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
               TabIndex        =   58
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
               TabIndex        =   57
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
               TabIndex        =   56
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
               TabIndex        =   55
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
               TabIndex        =   54
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
               TabIndex        =   53
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
               TabIndex        =   52
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
               TabIndex        =   51
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
               TabIndex        =   50
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
               TabIndex        =   49
               Top             =   3120
               Width           =   480
            End
            Begin VB.Shape EquipS 
               BorderColor     =   &H0000FFFF&
               BorderWidth     =   3
               Height          =   540
               Index           =   9
               Left            =   240
               Top             =   240
               Width           =   540
            End
            Begin VB.Shape EquipS 
               BorderColor     =   &H0000FFFF&
               BorderWidth     =   3
               Height          =   540
               Index           =   8
               Left            =   240
               Top             =   0
               Width           =   540
            End
            Begin VB.Shape EquipS 
               BorderColor     =   &H0000FFFF&
               BorderWidth     =   3
               Height          =   540
               Index           =   7
               Left            =   120
               Top             =   240
               Width           =   540
            End
            Begin VB.Shape EquipS 
               BorderColor     =   &H0000FFFF&
               BorderWidth     =   3
               Height          =   540
               Index           =   6
               Left            =   120
               Top             =   0
               Width           =   540
            End
            Begin VB.Shape EquipS 
               BorderColor     =   &H0000FFFF&
               BorderWidth     =   3
               Height          =   540
               Index           =   5
               Left            =   0
               Top             =   240
               Width           =   540
            End
            Begin VB.Shape EquipS 
               BorderColor     =   &H0000FFFF&
               BorderWidth     =   3
               Height          =   540
               Index           =   4
               Left            =   120
               Top             =   120
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
      Begin VB.Line Line2 
         BorderColor     =   &H000000C0&
         X1              =   4
         X2              =   171
         Y1              =   130
         Y2              =   130
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000C0&
         X1              =   8
         X2              =   168
         Y1              =   138
         Y2              =   138
      End
      Begin VB.Label lblDropItem 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "DROP ITEM"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   165
         Left            =   990
         TabIndex        =   3
         Top             =   2280
         Width           =   915
      End
      Begin VB.Label lblUseItem 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "USE ITEM"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   165
         Left            =   60
         TabIndex        =   2
         Top             =   2280
         Width           =   825
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
      Height          =   2505
      Left            =   9840
      ScaleHeight     =   167
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   175
      TabIndex        =   4
      Top             =   3720
      Width           =   2625
      Visible         =   0   'False
      Begin VB.ListBox lstSpells 
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
         ItemData        =   "frmMirage.frx":2AF3
         Left            =   45
         List            =   "frmMirage.frx":2AF5
         TabIndex        =   5
         Top             =   60
         Width           =   2535
      End
      Begin VB.Label lblCast 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "CAST"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   1080
         TabIndex        =   6
         Top             =   2280
         Width           =   495
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
      Height          =   2505
      Left            =   9840
      ScaleHeight     =   2505
      ScaleWidth      =   2625
      TabIndex        =   24
      Top             =   3720
      Width           =   2625
      Visible         =   0   'False
      Begin VB.ListBox lstOnline 
         Appearance      =   0  'Flat
         BackColor       =   &H0084ADB3&
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
         Height          =   2340
         ItemData        =   "frmMirage.frx":2AF7
         Left            =   80
         List            =   "frmMirage.frx":2AF9
         TabIndex        =   25
         Top             =   75
         Width           =   2460
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
      Height          =   2505
      Left            =   9840
      ScaleHeight     =   2505
      ScaleWidth      =   2625
      TabIndex        =   27
      Top             =   3720
      Width           =   2625
      Visible         =   0   'False
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
         TabIndex        =   33
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
         TabIndex        =   32
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
         Left            =   420
         Style           =   1  'Graphical
         TabIndex        =   31
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
         Left            =   420
         Style           =   1  'Graphical
         TabIndex        =   30
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
         Left            =   420
         Style           =   1  'Graphical
         TabIndex        =   29
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
         Left            =   420
         Style           =   1  'Graphical
         TabIndex        =   28
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
         ForeColor       =   &H000000C0&
         Height          =   165
         Left            =   255
         TabIndex        =   35
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
         ForeColor       =   &H000000C0&
         Height          =   165
         Left            =   285
         TabIndex        =   34
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
      Height          =   2505
      Left            =   9840
      ScaleHeight     =   2505
      ScaleWidth      =   2625
      TabIndex        =   37
      Top             =   3720
      Width           =   2625
      Visible         =   0   'False
      Begin VB.Label cmdLeave 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "LEAVE GUILD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   180
         Left            =   730
         TabIndex        =   42
         Top             =   2280
         Width           =   1095
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
         ForeColor       =   &H000000C0&
         Height          =   165
         Left            =   1425
         TabIndex        =   41
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
         ForeColor       =   &H000000C0&
         Height          =   165
         Left            =   1425
         TabIndex        =   40
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
         ForeColor       =   &H000000C0&
         Height          =   165
         Left            =   570
         TabIndex        =   39
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
         ForeColor       =   &H000000C0&
         Height          =   165
         Left            =   480
         TabIndex        =   38
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
      Height          =   2505
      Left            =   9840
      ScaleHeight     =   2505
      ScaleWidth      =   2625
      TabIndex        =   43
      Top             =   3720
      Width           =   2625
      Visible         =   0   'False
      Begin VB.PictureBox Picture13 
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
         Left            =   480
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   35
         TabIndex        =   138
         Top             =   80
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
            TabIndex        =   139
            Top             =   15
            Width           =   495
         End
      End
      Begin VB.PictureBox Picture12 
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
         Left            =   1680
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   35
         TabIndex        =   136
         Top             =   1280
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
            TabIndex        =   137
            Top             =   15
            Width           =   495
         End
      End
      Begin VB.PictureBox Picture11 
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
         Left            =   1080
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   35
         TabIndex        =   122
         Top             =   1880
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
            TabIndex        =   123
            Top             =   15
            Width           =   495
         End
      End
      Begin VB.PictureBox Picture10 
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
         Left            =   1080
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   35
         TabIndex        =   120
         Top             =   1280
         Width           =   555
         Begin VB.PictureBox BeltImage 
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
            TabIndex        =   121
            Top             =   15
            Width           =   495
         End
      End
      Begin VB.PictureBox Picture3 
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
         Left            =   1680
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   35
         TabIndex        =   118
         Top             =   1880
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
            TabIndex        =   119
            Top             =   15
            Width           =   495
         End
      End
      Begin VB.PictureBox Picture2 
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
         Left            =   480
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   35
         TabIndex        =   116
         Top             =   1280
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
            TabIndex        =   117
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
         Left            =   480
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   35
         TabIndex        =   105
         Top             =   680
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
            TabIndex        =   106
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
         Left            =   1080
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   35
         TabIndex        =   103
         Top             =   680
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
            TabIndex        =   104
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
         TabIndex        =   101
         Top             =   680
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
            TabIndex        =   102
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
         Index           =   0
         Left            =   1080
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   35
         TabIndex        =   99
         Top             =   80
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
            TabIndex        =   100
            Top             =   15
            Width           =   495
         End
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
         Left            =   2400
         Picture         =   "frmMirage.frx":2AFB
         ScaleHeight     =   2.23636e5
         ScaleMode       =   0  'User
         ScaleWidth      =   477.091
         TabIndex        =   44
         Top             =   2760
         Width           =   480
         Visible         =   0   'False
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
      Height          =   7185
      Left            =   120
      ScaleHeight     =   479
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   634
      TabIndex        =   23
      Top             =   90
      Width           =   9510
   End
   Begin VB.TextBox txtMyTextBox 
      Appearance      =   0  'Flat
      BackColor       =   &H00789298&
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
      Height          =   285
      Left            =   150
      MaxLength       =   255
      TabIndex        =   17
      Top             =   7395
      Width           =   9465
   End
   Begin MSWinsockLib.Winsock Socket 
      Left            =   9150
      Top             =   150
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox txtChat 
      Height          =   1170
      Left            =   150
      TabIndex        =   0
      Top             =   7680
      Width           =   9465
      _ExtentX        =   16695
      _ExtentY        =   2064
      _Version        =   393217
      BackColor       =   7901848
      BorderStyle     =   0
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmMirage.frx":16243D
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label LblSP 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   180
      Left            =   10200
      TabIndex        =   113
      Top             =   840
      Width           =   1890
   End
   Begin VB.Label lblEXP 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   180
      Left            =   10200
      TabIndex        =   111
      Top             =   1080
      Width           =   1890
   End
   Begin VB.Label lblPoints 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   180
      Left            =   10680
      TabIndex        =   98
      Top             =   1800
      Width           =   1215
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
      ForeColor       =   &H0000C0C0&
      Height          =   240
      Left            =   11880
      TabIndex        =   97
      Top             =   2760
      Width           =   165
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
      ForeColor       =   &H0000C0C0&
      Height          =   240
      Left            =   11880
      TabIndex        =   96
      Top             =   2520
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
      ForeColor       =   &H0000C0C0&
      Height          =   240
      Left            =   11880
      TabIndex        =   95
      Top             =   2235
      Width           =   165
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
      ForeColor       =   &H0000C0C0&
      Height          =   240
      Left            =   11880
      TabIndex        =   94
      Top             =   1980
      Width           =   165
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   165
      Left            =   9750
      TabIndex        =   79
      Top             =   6885
      Width           =   465
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   11685
      TabIndex        =   45
      Top             =   6300
      Width           =   855
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
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
      Height          =   165
      Left            =   9750
      TabIndex        =   36
      Top             =   7035
      Width           =   525
   End
   Begin VB.Label lblWhosOnline 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   9750
      TabIndex        =   26
      Top             =   6735
      Width           =   1110
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   9750
      TabIndex        =   18
      Top             =   7185
      Width           =   705
   End
   Begin VB.Shape Shape1 
      Height          =   180
      Left            =   10200
      Top             =   1080
      Width           =   1890
   End
   Begin VB.Label lblLevel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   180
      Left            =   10680
      TabIndex        =   16
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label lblSPEED 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   180
      Left            =   10680
      TabIndex        =   15
      Top             =   2550
      Width           =   1260
   End
   Begin VB.Label lblMAGI 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   180
      Left            =   10680
      TabIndex        =   14
      Top             =   2820
      Width           =   1245
   End
   Begin VB.Label lblDEF 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   180
      Left            =   10680
      TabIndex        =   13
      Top             =   2295
      Width           =   1215
   End
   Begin VB.Label lblSTR 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   180
      Left            =   10680
      TabIndex        =   12
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   10800
      TabIndex        =   11
      Top             =   6300
      Width           =   750
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   9750
      TabIndex        =   10
      Top             =   6300
      Width           =   915
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   11835
      TabIndex        =   9
      Top             =   8640
      Width           =   630
   End
   Begin VB.Label lblMP 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   180
      Left            =   10200
      TabIndex        =   8
      Top             =   600
      Width           =   1890
   End
   Begin VB.Label lblHP 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   180
      Left            =   10200
      TabIndex        =   7
      Top             =   360
      Width           =   1890
   End
   Begin VB.Shape shpHP 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   180
      Left            =   10200
      Top             =   360
      Width           =   1905
   End
   Begin VB.Shape shpMP 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   180
      Left            =   10200
      Top             =   600
      Width           =   1905
   End
   Begin VB.Shape shpTNL 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000FF&
      Height          =   180
      Left            =   10200
      Top             =   1080
      Width           =   1905
   End
   Begin VB.Shape ShpSP 
      BackColor       =   &H00FF00FF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   180
      Left            =   10200
      Top             =   840
      Width           =   1875
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

' Private Sub chkbubblebar_Click()
'    WriteINI "CONFIG", "SpeechBubbles", chkbubblebar.Value, App.Path & "\config.ini"
'End Sub

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

' Private Sub chkAutoScroll_Click()
'    WriteINI "CONFIG", "AutoScroll", chkAutoScroll.Value, App.Path & "\config.ini"
' End Sub

Private Sub chkplayername_Click()
    WriteINI "CONFIG", "PlayerName", chkplayername.Value, App.Path & "\config.ini"
End Sub

Private Sub chkmusic_Click()
    WriteINI "CONFIG", "Music", chkmusic.Value, App.Path & "\config.ini"
    If MyIndex <= 0 Then Exit Sub
    Call PlayMidi(Trim(Map(GetPlayerMap(MyIndex)).Music))
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

Private Sub KeepNotes_Click()
frmKeepNotes.Visible = True
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
picEquip.Visible = False
picWhosOnline.Visible = False
End Sub

Private Sub Label19_Click()
    picEquip.Visible = True
    picInv3.Visible = False
    picPlayerSpells.Visible = False
    picWhosOnline.Visible = False
    Picture1.Visible = False
    frmMirage.picGuildAdmin.Visible = False
    Call UpdateVisInv
End Sub

Private Sub label2_Click()
    picOptions.Visible = True
End Sub

Private Sub Label21_Click()
picEquip.Visible = False
End Sub

Private Sub Label3_Click()
    Call GameDestroy
End Sub

Private Sub Label6_Click()
    Call SendData("getstats" & SEP_CHAR & END_CHAR)
End Sub

Private Sub Label7_Click()
    Call UpdateVisInv
    picInv3.Visible = True
    Picture1.Visible = False
    frmMirage.picGuildAdmin.Visible = False
    picEquip.Visible = False
    picPlayerSpells.Visible = False
    picWhosOnline.Visible = False
End Sub

Private Sub Label8_Click()
    Call SendData("spells" & SEP_CHAR & END_CHAR)
    picInv3.Visible = False
    frmMirage.picGuildAdmin.Visible = False
    picEquip.Visible = False
    Picture1.Visible = False
    picWhosOnline.Visible = False
End Sub

Private Sub lblCloseOnline_Click()
Call SendOnlineList
picWhosOnline.Visible = False
End Sub

Private Sub lblClosePicGuildAdmin_Click()
picGuildAdmin.Visible = False
End Sub

Private Sub lblWhosOnline_Click()
Call SendOnlineList
picWhosOnline.Visible = True
picInv3.Visible = False
picEquip.Visible = False
Picture1.Visible = False
frmMirage.picGuildAdmin.Visible = False
picPlayerSpells.Visible = False
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
        If Item(GetPlayerInvItemNum(MyIndex, d + 1)).Type = ITEM_TYPE_CURRENCY Then
            If Trim(Item(GetPlayerInvItemNum(MyIndex, d + 1)).desc) = "" Then
                itmDesc.Height = 17
                itmDesc.Top = 224
            Else
                itmDesc.Height = 222
                itmDesc.Top = 8
            End If
        Else
            If Trim(Item(GetPlayerInvItemNum(MyIndex, d + 1)).desc) = "" Then
                itmDesc.Height = 222
                itmDesc.Top = 8
            Else
                itmDesc.Height = 222
                itmDesc.Top = 8
            End If
        End If
        If Item(GetPlayerInvItemNum(MyIndex, d + 1)).Type = ITEM_TYPE_CURRENCY Then
            descName.Caption = Trim(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Name) & " (" & GetPlayerInvItemValue(MyIndex, d + 1) & ")"
        Else
            If GetPlayerWeaponSlot(MyIndex) = d + 1 Then
                descName.Caption = Trim(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Name) & " (worn)"
            ElseIf GetPlayerArmorSlot(MyIndex) = d + 1 Then
                descName.Caption = Trim(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Name) & " (worn)"
            ElseIf GetPlayerHelmetSlot(MyIndex) = d + 1 Then
                descName.Caption = Trim(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Name) & " (worn)"
            ElseIf GetPlayerShieldSlot(MyIndex) = d + 1 Then
                descName.Caption = Trim(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Name) & " (worn)"
            ElseIf GetPlayerGlovesSlot(MyIndex) = d + 1 Then
                descName.Caption = Trim(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Name) & " (worn)"
            ElseIf GetPlayerBootsSlot(MyIndex) = d + 1 Then
                descName.Caption = Trim(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Name) & " (worn)"
            ElseIf GetPlayerBeltSlot(MyIndex) = d + 1 Then
                descName.Caption = Trim(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Name) & " (worn)"
            ElseIf GetPlayerLegsSlot(MyIndex) = d + 1 Then
                descName.Caption = Trim(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Name) & " (worn)"
            ElseIf GetPlayerRingSlot(MyIndex) = d + 1 Then
                descName.Caption = Trim(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Name) & " (worn)"
            ElseIf GetPlayerAmuletSlot(MyIndex) = d + 1 Then
                descName.Caption = Trim(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Name) & " (worn)"
            Else
                descName.Caption = Trim(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Name)
            End If
        End If
        descStr.Caption = "STR " & Item(GetPlayerInvItemNum(MyIndex, d + 1)).StrReq
        descDef.Caption = "CON " & Item(GetPlayerInvItemNum(MyIndex, d + 1)).DefReq
        descSpeed.Caption = "DEX " & Item(GetPlayerInvItemNum(MyIndex, d + 1)).SpeedReq
        descMagi.Caption = "INT " & Item(GetPlayerInvItemNum(MyIndex, d + 1)).MagiReq
        descLvl.Caption = "Level " & Item(GetPlayerInvItemNum(MyIndex, d + 1)).LvlReq
        descHpMp.Caption = "HP: " & Item(GetPlayerInvItemNum(MyIndex, d + 1)).AddHP & " MP: " & Item(GetPlayerInvItemNum(MyIndex, d + 1)).AddMP & " SP: " & Item(GetPlayerInvItemNum(MyIndex, d + 1)).AddSP
        descSD.Caption = "STR: " & Item(GetPlayerInvItemNum(MyIndex, d + 1)).AddStr & " CON: " & Item(GetPlayerInvItemNum(MyIndex, d + 1)).AddDef
        descMS.Caption = "DEX: " & Item(GetPlayerInvItemNum(MyIndex, d + 1)).AddSpeed & " INT: " & Item(GetPlayerInvItemNum(MyIndex, d + 1)).AddMagi
        descDUR.Caption = "Durability: " & GetPlayerInvItemDur(MyIndex, d + 1) & " / " & Item(GetPlayerInvItemNum(MyIndex, d + 1)).Data1
        desc.Caption = Trim(Item(GetPlayerInvItemNum(MyIndex, d + 1)).desc)
        
        itmDesc.Visible = True
    Else
        itmDesc.Visible = False
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
    
    ' The Guild Creator
    If KeyCode = vbKeyF4 Then
        If Player(MyIndex).Access > 0 Then
            frmGuild.Show vbModeless, frmMirage
        End If
    End If

    ' The Guild Maker
    If KeyCode = vbKeyF5 Then
        frmMirage.picGuildAdmin.Visible = True
        frmMirage.picInv3.Visible = False
        frmMirage.Picture1.Visible = False
        frmMirage.picEquip.Visible = False
        frmMirage.picPlayerSpells.Visible = False
        frmMirage.picWhosOnline.Visible = False
      End If
      
    If KeyCode = vbKeyInsert Then
        If SpellMemorized > 0 Then
            If GetTickCount > Player(MyIndex).AttackTimer + 100 Then
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
    
    frmMapEditor.Caption = "Map Editor - " & "X: " & Int((x + (NewPlayerX * PIC_X)) / 32) & " Y: " & Int((y + (NewPlayerY * PIC_Y)) / 32)
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
    
   ' MAX_BLT_LINE = scrlBltText.Value
    MAX_BLT_LINE = 4
    ReDim BattlePMsg(1 To MAX_BLT_LINE) As BattleMsgRec
    ReDim BattleMMsg(1 To MAX_BLT_LINE) As BattleMsgRec
 '   lblLines.Caption = "On Screen Text Line Amount: " & scrlBltText.Value
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
            GoldAmount = InputBox("How much " & Trim(Item(GetPlayerInvItemNum(MyIndex, InvNum)).Name) & "(" & GetPlayerInvItemValue(MyIndex, InvNum) & ") would you like to drop?", "Drop " & Trim(Item(GetPlayerInvItemNum(MyIndex, InvNum)).Name), 0, frmMirage.Left, frmMirage.Top)
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
        MsgBox "The variable can't handle that amount!"
    End If
End Sub


Private Sub lblCast_Click()
    If Player(MyIndex).Spell(lstSpells.ListIndex + 1) > 0 Then
        If GetTickCount > Player(MyIndex).AttackTimer + 100 Then
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

Private Sub LblCancel_Click()
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
If VScroll1.Value = 1 Then
Picture9.Top = (VScroll1.Value * -PIC_Y) - 88
Else
    Picture9.Top = VScroll1.Value * -PIC_Y
End If
End Sub

Private Sub Down_Click()
If VScroll1.Value = 2 Then Exit Sub
VScroll1.Value = VScroll1.Value + 1
If VScroll1.Value = 2 Then
Picture9.Top = (VScroll1.Value * -PIC_Y) - 176
Else
    Picture9.Top = (VScroll1.Value * -PIC_Y) - 88
End If
End Sub
