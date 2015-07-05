VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmMirage 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Elysium Diamond"
   ClientHeight    =   8985
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
   Picture         =   "frmMirage.frx":058A
   ScaleHeight     =   599
   ScaleMode       =   0  'User
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.PictureBox picPlayerSpells 
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
      Height          =   3795
      Left            =   0
      Picture         =   "frmMirage.frx":FB06
      ScaleHeight     =   253
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   229
      TabIndex        =   3
      Top             =   5160
      Visible         =   0   'False
      Width           =   3435
      Begin VB.ListBox lstSpells 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
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
         Height          =   2340
         ItemData        =   "frmMirage.frx":152CC
         Left            =   405
         List            =   "frmMirage.frx":152CE
         TabIndex        =   4
         Top             =   960
         Width           =   2655
      End
      Begin VB.Label lblForgetSpell 
         BackStyle       =   0  'Transparent
         Caption         =   "Forget Spell"
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
         Height          =   255
         Left            =   2040
         TabIndex        =   154
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   3000
         TabIndex        =   143
         Top             =   120
         Width           =   375
      End
      Begin VB.Label lblCast 
         BackStyle       =   0  'Transparent
         Caption         =   "Cast Spell"
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
         Left            =   600
         TabIndex        =   5
         Top             =   600
         Width           =   735
      End
   End
   Begin VB.PictureBox itmDesc 
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
      Height          =   3795
      Left            =   2160
      Picture         =   "frmMirage.frx":152D0
      ScaleHeight     =   253
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   231
      TabIndex        =   63
      Top             =   5160
      Visible         =   0   'False
      Width           =   3465
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   3000
         TabIndex        =   140
         Top             =   120
         Width           =   375
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
         Left            =   840
         TabIndex        =   103
         Top             =   1440
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
         Left            =   720
         TabIndex        =   74
         Top             =   2400
         Width           =   2055
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
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   840
         TabIndex        =   73
         Top             =   2640
         Width           =   1815
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
         TabIndex        =   72
         Top             =   2880
         Width           =   3015
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
         Left            =   720
         TabIndex        =   71
         Top             =   2160
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
         Left            =   720
         TabIndex        =   70
         Top             =   1920
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
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   840
         TabIndex        =   69
         Top             =   1680
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
         Left            =   840
         TabIndex        =   68
         Top             =   1200
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
         Left            =   840
         TabIndex        =   67
         Top             =   960
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
         Left            =   840
         TabIndex        =   66
         Top             =   720
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
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   840
         TabIndex        =   65
         Top             =   480
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
         Left            =   480
         TabIndex        =   64
         Top             =   240
         Width           =   2535
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
      Left            =   0
      Picture         =   "frmMirage.frx":20070
      ScaleHeight     =   249
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   229
      TabIndex        =   0
      Top             =   1680
      Visible         =   0   'False
      Width           =   3435
      Begin VB.VScrollBar VScroll1 
         Height          =   330
         Left            =   3480
         Max             =   3
         TabIndex        =   137
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
         Picture         =   "frmMirage.frx":25836
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   59
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
         Picture         =   "frmMirage.frx":25ACE
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   58
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
         TabIndex        =   30
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
            Height          =   3735
            Left            =   0
            ScaleHeight     =   249
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   169
            TabIndex        =   31
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
               Index           =   7
               Left            =   120
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
               Index           =   6
               Left            =   1920
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
               TabIndex        =   52
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
               TabIndex        =   51
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
               TabIndex        =   50
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
               TabIndex        =   49
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
               TabIndex        =   48
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
               TabIndex        =   47
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
               Left            =   1920
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   45
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
               Index           =   11
               Left            =   120
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   44
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
               Index           =   12
               Left            =   720
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   43
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
               Index           =   13
               Left            =   1320
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   42
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
               Index           =   14
               Left            =   1920
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   41
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
               Index           =   15
               Left            =   1320
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   40
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
               Index           =   16
               Left            =   1920
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   39
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
               Index           =   17
               Left            =   120
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   38
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
               Index           =   18
               Left            =   720
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   37
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
               Index           =   19
               Left            =   1320
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   36
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
               Index           =   20
               Left            =   1920
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   35
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
               Index           =   21
               Left            =   120
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   34
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
               Index           =   22
               Left            =   720
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   33
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
               Index           =   23
               Left            =   1320
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   32
               Top             =   1920
               Width           =   480
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
         Height          =   375
         Left            =   3000
         TabIndex        =   148
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
   Begin VB.PictureBox picStat 
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
      Height          =   1695
      Left            =   8520
      Picture         =   "frmMirage.frx":25D59
      ScaleHeight     =   113
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   229
      TabIndex        =   107
      Top             =   0
      Visible         =   0   'False
      Width           =   3435
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   3000
         TabIndex        =   146
         Top             =   120
         Width           =   375
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
         Left            =   2520
         TabIndex        =   117
         Top             =   720
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
         Left            =   2520
         TabIndex        =   116
         Top             =   960
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
         Left            =   2520
         TabIndex        =   115
         Top             =   1200
         Width           =   105
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
         Left            =   2520
         TabIndex        =   114
         Top             =   480
         Width           =   105
      End
      Begin VB.Label lblSTR 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         Height          =   180
         Left            =   720
         TabIndex        =   113
         Top             =   480
         Width           =   1905
      End
      Begin VB.Label lblDEF 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Defense"
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
         Left            =   720
         TabIndex        =   112
         Top             =   720
         Width           =   1905
      End
      Begin VB.Label lblMAGI 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         Height          =   180
         Left            =   720
         TabIndex        =   111
         Top             =   1200
         Width           =   1905
      End
      Begin VB.Label lblSPEED 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         Height          =   180
         Left            =   720
         TabIndex        =   110
         Top             =   960
         Width           =   1905
      End
      Begin VB.Label lblLevel 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   360
         TabIndex        =   109
         Top             =   240
         Width           =   1305
      End
      Begin VB.Label lblPoints 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Points"
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
         Left            =   1680
         TabIndex        =   108
         Top             =   240
         Width           =   1185
      End
   End
   Begin VB.PictureBox picStatus 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
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
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   0
      Picture         =   "frmMirage.frx":2A7A1
      ScaleHeight     =   113
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   229
      TabIndex        =   133
      Top             =   0
      Visible         =   0   'False
      Width           =   3435
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   3000
         TabIndex        =   138
         Top             =   120
         Width           =   375
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
         Left            =   840
         TabIndex        =   136
         Top             =   1155
         Width           =   2385
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
         Left            =   840
         TabIndex        =   135
         Top             =   855
         Width           =   2385
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
         Left            =   840
         TabIndex        =   134
         Top             =   555
         Width           =   2385
      End
      Begin VB.Shape shpMP 
         BackColor       =   &H00CB884B&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   165
         Left            =   855
         Top             =   870
         Width           =   2370
      End
      Begin VB.Shape shpHP 
         BackColor       =   &H0000C000&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   165
         Left            =   855
         Top             =   570
         Width           =   2370
      End
      Begin VB.Shape shpTNL 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000000FF&
         Height          =   165
         Left            =   855
         Top             =   1170
         Width           =   2370
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
      Height          =   3795
      Left            =   0
      Picture         =   "frmMirage.frx":2F050
      ScaleHeight     =   253
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   229
      TabIndex        =   28
      Top             =   1680
      Visible         =   0   'False
      Width           =   3435
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
         Left            =   2040
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   35
         TabIndex        =   95
         Top             =   840
         Visible         =   0   'False
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
            TabIndex        =   96
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
         TabIndex        =   93
         Top             =   2640
         Visible         =   0   'False
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
            TabIndex        =   94
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
         TabIndex        =   91
         Top             =   2040
         Visible         =   0   'False
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
            TabIndex        =   92
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
         TabIndex        =   89
         Top             =   2040
         Visible         =   0   'False
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
            TabIndex        =   90
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
         Left            =   2040
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   35
         TabIndex        =   87
         Top             =   2640
         Visible         =   0   'False
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
            TabIndex        =   88
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
         TabIndex        =   85
         Top             =   2040
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
            TabIndex        =   86
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
         Left            =   840
         ScaleHeight     =   35
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   35
         TabIndex        =   83
         Top             =   1440
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
            TabIndex        =   84
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
         TabIndex        =   81
         Top             =   1440
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
            TabIndex        =   82
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
         TabIndex        =   79
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
            TabIndex        =   80
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
         TabIndex        =   77
         Top             =   840
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
            TabIndex        =   78
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
         Height          =   2.25120e5
         Left            =   3480
         Picture         =   "frmMirage.frx":39DF0
         ScaleHeight     =   2.23755e5
         ScaleMode       =   0  'User
         ScaleWidth      =   2862.546
         TabIndex        =   29
         Top             =   0
         Visible         =   0   'False
         Width           =   2880
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   3000
         TabIndex        =   145
         Top             =   120
         Width           =   375
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Note: Many of these boxes are unused and are not visible."
         Height          =   495
         Left            =   600
         TabIndex        =   132
         Top             =   360
         Visible         =   0   'False
         Width           =   2415
      End
   End
   Begin VB.TextBox txtMyTextBox 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   6960
      MaxLength       =   255
      TabIndex        =   104
      Top             =   7635
      Width           =   4755
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
      TabIndex        =   62
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
   Begin RichTextLib.RichTextBox txtChat 
      Height          =   900
      Left            =   6960
      TabIndex        =   105
      Top             =   7965
      Width           =   4755
      _ExtentX        =   8387
      _ExtentY        =   1588
      _Version        =   393217
      BackColor       =   16744576
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmMirage.frx":878632
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
   Begin VB.PictureBox picGuild 
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
      Height          =   1695
      Left            =   8520
      Picture         =   "frmMirage.frx":8786AD
      ScaleHeight     =   113
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   229
      TabIndex        =   22
      Top             =   1680
      Visible         =   0   'False
      Width           =   3435
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   3000
         TabIndex        =   141
         Top             =   120
         Width           =   375
      End
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
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   1320
         TabIndex        =   27
         Top             =   1320
         Width           =   765
      End
      Begin VB.Label lblRank 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00FFFFFF&
         Height          =   165
         Left            =   720
         TabIndex        =   26
         Top             =   960
         Width           =   1920
      End
      Begin VB.Label lblGuild 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00FFFFFF&
         Height          =   165
         Left            =   720
         TabIndex        =   25
         Top             =   480
         Width           =   1905
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
         ForeColor       =   &H00FFFFFF&
         Height          =   165
         Left            =   810
         TabIndex        =   24
         Top             =   720
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
         ForeColor       =   &H00FFFFFF&
         Height          =   165
         Left            =   720
         TabIndex        =   23
         Top             =   240
         Width           =   825
      End
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
      Height          =   3795
      Left            =   8520
      Picture         =   "frmMirage.frx":87D0F5
      ScaleHeight     =   253
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   229
      TabIndex        =   6
      Top             =   3000
      Visible         =   0   'False
      Width           =   3435
      Begin VB.HScrollBar scrlMusicVol 
         Height          =   135
         Left            =   1560
         Max             =   100
         TabIndex        =   152
         Top             =   2160
         Value           =   100
         Width           =   1575
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
         Left            =   1560
         TabIndex        =   150
         Top             =   1800
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
         Left            =   1560
         TabIndex        =   102
         Top             =   2760
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
         TabIndex        =   101
         Top             =   3240
         Width           =   495
      End
      Begin VB.HScrollBar scrlBltText 
         Height          =   255
         Left            =   1320
         Max             =   20
         Min             =   4
         TabIndex        =   75
         Top             =   3240
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
         Left            =   2400
         TabIndex        =   61
         Top             =   1560
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
         TabIndex        =   60
         Top             =   1560
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
         Left            =   1560
         TabIndex        =   57
         Top             =   1080
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
         Left            =   1560
         TabIndex        =   56
         Top             =   480
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
         TabIndex        =   9
         Top             =   2520
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
         TabIndex        =   8
         Top             =   840
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
         TabIndex        =   7
         Top             =   240
         Value           =   1  'Checked
         Width           =   765
      End
      Begin VB.Label lblMusicVol 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "100"
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
         Height          =   135
         Left            =   1560
         TabIndex        =   153
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "Music Volume:"
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
         Height          =   255
         Left            =   360
         TabIndex        =   151
         Top             =   2100
         Width           =   1095
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   3000
         TabIndex        =   144
         Top             =   120
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
         TabIndex        =   100
         Top             =   2520
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
         TabIndex        =   99
         Top             =   1560
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
         TabIndex        =   98
         Top             =   960
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
         TabIndex        =   97
         Top             =   360
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
         Left            =   1320
         TabIndex        =   76
         Top             =   3060
         Width           =   1965
      End
   End
   Begin VB.PictureBox picWhosOnline 
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
      Height          =   3795
      Left            =   8520
      Picture         =   "frmMirage.frx":887E95
      ScaleHeight     =   253
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   229
      TabIndex        =   11
      Top             =   2040
      Visible         =   0   'False
      Width           =   3435
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
         ItemData        =   "frmMirage.frx":88D65B
         Left            =   410
         List            =   "frmMirage.frx":88D65D
         TabIndex        =   12
         Top             =   960
         Width           =   2630
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   3000
         TabIndex        =   147
         Top             =   120
         Width           =   375
      End
   End
   Begin VB.PictureBox picFriend 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
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
      ForeColor       =   &H80000008&
      Height          =   3795
      Left            =   8520
      Picture         =   "frmMirage.frx":88D65F
      ScaleHeight     =   253
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   229
      TabIndex        =   118
      Top             =   2400
      Visible         =   0   'False
      Width           =   3435
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
         ItemData        =   "frmMirage.frx":892E25
         Left            =   375
         List            =   "frmMirage.frx":892E27
         TabIndex        =   119
         Top             =   960
         Width           =   2700
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
         Left            =   1920
         TabIndex        =   122
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
         Left            =   360
         TabIndex        =   121
         Top             =   600
         Width           =   1215
      End
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
         Left            =   360
         TabIndex        =   120
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   3120
         TabIndex        =   139
         Top             =   120
         Width           =   255
      End
   End
   Begin VB.PictureBox picGuildAdmin 
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
      Height          =   3795
      Left            =   8520
      Picture         =   "frmMirage.frx":892E29
      ScaleHeight     =   3795
      ScaleWidth      =   3435
      TabIndex        =   13
      Top             =   1680
      Visible         =   0   'False
      Width           =   3435
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
         Left            =   720
         TabIndex        =   19
         Top             =   1200
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
         Left            =   720
         TabIndex        =   18
         Top             =   570
         Width           =   1815
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
         Left            =   780
         TabIndex        =   17
         Top             =   2160
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
         Left            =   780
         TabIndex        =   16
         Top             =   2490
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
         Left            =   780
         TabIndex        =   15
         Top             =   2835
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
         Left            =   780
         TabIndex        =   14
         Top             =   3165
         Width           =   1815
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   3000
         TabIndex        =   142
         Top             =   120
         Width           =   375
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
         Left            =   720
         TabIndex        =   21
         Top             =   945
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
         Left            =   720
         TabIndex        =   20
         Top             =   345
         Width           =   420
      End
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
      Left            =   1200
      ScaleHeight     =   443
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   637
      TabIndex        =   10
      Top             =   15
      Width           =   9585
      Begin VB.Shape shpSelect 
         BorderColor     =   &H00808080&
         Height          =   480
         Left            =   0
         Top             =   0
         Width           =   480
      End
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   1080
      TabIndex        =   149
      Top             =   6960
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00808080&
      Height          =   375
      Left            =   240
      TabIndex        =   131
      Top             =   6960
      Width           =   735
   End
   Begin VB.Label lblEqp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00808080&
      Height          =   375
      Left            =   3120
      TabIndex        =   130
      Top             =   6960
      Width           =   1095
   End
   Begin VB.Label lblGld 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00808080&
      Height          =   375
      Left            =   5160
      TabIndex        =   129
      Top             =   6960
      Width           =   855
   End
   Begin VB.Label lblOpt 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00808080&
      Height          =   375
      Left            =   9720
      TabIndex        =   128
      Top             =   6960
      Width           =   855
   End
   Begin VB.Label lblCht 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00808080&
      Height          =   375
      Left            =   6120
      TabIndex        =   127
      Top             =   6960
      Width           =   735
   End
   Begin VB.Label lblWho 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00808080&
      Height          =   375
      Left            =   6960
      TabIndex        =   126
      Top             =   6960
      Width           =   1455
   End
   Begin VB.Label lblFriend 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00808080&
      Height          =   375
      Left            =   8520
      TabIndex        =   125
      Top             =   6960
      Width           =   855
   End
   Begin VB.Label lblSpl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00808080&
      Height          =   375
      Left            =   4320
      TabIndex        =   124
      Top             =   6960
      Width           =   735
   End
   Begin VB.Label lblInv 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00808080&
      Height          =   375
      Left            =   1920
      TabIndex        =   123
      Top             =   6960
      Width           =   1095
   End
   Begin VB.Label lblQit 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   10680
      TabIndex        =   106
      Top             =   6840
      Width           =   975
   End
End
Attribute VB_Name = "frmMirage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Copyright (c) 2007-2008 Elysium Source. All rights reserved.
' This code is licensed under the Elysium General License.

Option Explicit

Dim SpellMemorized As Long

Private Sub Close_Click()
    Call GameDestroy
End Sub

Private Sub AddDef_Click()
    Call SendData(USESTATPOINT_CHAR & SEP_CHAR & 1 & END_CHAR)
End Sub

Private Sub AddMagi_Click()
    Call SendData(USESTATPOINT_CHAR & SEP_CHAR & 2 & END_CHAR)
End Sub

Private Sub AddSpeed_Click()
    Call SendData(USESTATPOINT_CHAR & SEP_CHAR & 3 & END_CHAR)
End Sub

Private Sub AddStr_Click()
    Call SendData(USESTATPOINT_CHAR & SEP_CHAR & 0 & END_CHAR)
End Sub

Private Sub chksound_Click()
    Call PutVar(App.Path & "\config.ini", "CONFIG", "Sound", chkSound.Value)
    SoundOn = chkSound.Value
End Sub

Private Sub chkbubblebar_Click()
    Call PutVar(App.Path & "\config.ini", "CONFIG", "SpeechBubbles", chkbubblebar.Value)
    SpeechBubblesOn = chkbubblebar.Value
End Sub

Private Sub chkEmoSound_Click()
    Call PutVar(App.Path & "\config.ini", "CONFIG", "EmoticonSound", chkEmoSound.Value)
    EmoticonSoundOn = chkEmoSound.Value
End Sub

Private Sub chknpcdamage_Click()
    Call PutVar(App.Path & "\config.ini", "CONFIG", "NPCDamage", chknpcdamage.Value)
    NPCDamageOn = chknpcdamage.Value
End Sub

Private Sub chknpcname_Click()
    Call PutVar(App.Path & "\config.ini", "CONFIG", "NPCName", chknpcname.Value)
    NPCNameOn = chknpcname.Value
End Sub

Private Sub chkplayerdamage_Click()
    Call PutVar(App.Path & "\config.ini", "CONFIG", "PlayerDamage", chkplayerdamage.Value)
    PlayerDamageOn = chkplayerdamage.Value
End Sub

Private Sub chkAutoScroll_Click()
    Call PutVar(App.Path & "\config.ini", "CONFIG", "AutoScroll", chkAutoScroll.Value)
End Sub

Private Sub chkplayername_Click()
    Call PutVar(App.Path & "\config.ini", "CONFIG", "PlayerName", chkplayername.Value)
    PlayerNameOn = chkplayername.Value
End Sub

Private Sub chkmusic_Click()
    Call PutVar(App.Path & "\config.ini", "CONFIG", "Music", chkmusic.Value)
    MusicOn = chkmusic.Value
    If MyIndex <= 0 Then Exit Sub
    Call PlayMidi(Trim$(Map(GetPlayerMap(MyIndex)).Music))
End Sub

Private Sub cmdAdd_Click()
Dim Packet As String
    If txtPlayerName.Text <> vbNullString Then
        Packet = ADDFRIEND_CHAR & SEP_CHAR & txtPlayerName.Text & END_CHAR
        Call SendData(Packet)
    End If
End Sub

Private Sub cmdLeave_Click()
Dim Packet As String
    Packet = GUILDLEAVE_CHAR & END_CHAR
    Call SendData(Packet)
End Sub

Private Sub cmdMember_Click()
Dim Packet As String
    Packet = GUILDMEMBER_CHAR & SEP_CHAR & txtName.Text & END_CHAR
    Call SendData(Packet)
End Sub

Private Sub cmdRemove_Click()
Dim Packet As String
    If txtPlayerName.Text <> vbNullString Then
        Packet = REMOVEFRIEND_CHAR & SEP_CHAR & txtPlayerName.Text & END_CHAR
        Call SendData(Packet)
    End If
End Sub

Private Sub Command1_Click()
    picOptions.Visible = False
End Sub

Private Sub Form_Load()
Dim I As Long
Dim Ending As String
    For I = 1 To 3
        If I = 1 Then Ending = ".gif"
        If I = 2 Then Ending = ".jpg"
        If I = 3 Then Ending = ".png"
 
        If FileExist("GUI\game" & Ending) Then frmMirage.Picture = LoadPicture(App.Path & "\GUI\game" & Ending)
    Next I
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

'Private Sub KeepNotes_Click()
'frmKeepNotes.Visible = True
'End Sub

Private Sub Label13_Click()
    picGuildAdmin.Visible = False
End Sub

Private Sub Label19_Click()
    picPlayerSpells.Visible = False
End Sub

Private Sub Label21_Click()
    picEquip.Visible = False
End Sub

Private Sub Label1_Click()
    picStat.Visible = True
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
    picStatus.Visible = True
End Sub

Private Sub Label3_Click()
    picStatus.Visible = False
End Sub

Private Sub Label4_Click()
    picFriend.Visible = False
End Sub

Private Sub Label6_Click()
    Call SendData(GETSTATS_CHAR & END_CHAR)
End Sub

Private Sub Label7_Click()
    itmDesc.Visible = False
End Sub

Private Sub Label8_Click()
    picGuild.Visible = False
End Sub

Private Sub lblCloseOnline_Click()
Call SendOnlineList
picWhosOnline.Visible = False
End Sub

Private Sub lblClosePicGuildAdmin_Click()
picGuildAdmin.Visible = False
End Sub

Private Sub lblCht_Click()
Dim I As Long

    For I = 1 To MAX_PLAYERS
        If IsPlaying(I) = True Then
            If MouseDownX = GetPlayerX(I) And MouseDownY = GetPlayerY(I) Then
                Call SendData(PLAYERCHAT_CHAR & SEP_CHAR & GetPlayerName(I) & END_CHAR)
                Exit Sub
            End If
        End If
    Next I
    
    Call AddText("Target a player to chat with first!", Red)
End Sub

Private Sub lblEqp_Click()
    picEquip.Visible = True
    Call UpdateVisInv
End Sub

Private Sub lblForgetSpell_Click()

    If Player(MyIndex).Spell(lstSpells.ListIndex + 1) > 0 Then
        If GetTickCount > Player(MyIndex).AttackTimer + 1000 Then
            If MsgBox("Are you sure you want to forget this spell?", vbYesNo, "Forget Spell") = vbNo Then Exit Sub
            Call SendData(FORGETSPELL_CHAR & SEP_CHAR & lstSpells.ListIndex + 1 & END_CHAR)
            picPlayerSpells.Visible = False
        End If
    Else
        Call AddText("No spell here.", BrightRed)
    End If

End Sub

Private Sub lblFriend_Click()
    picFriend.Visible = True
End Sub

Private Sub lblGld_Click()
    frmMirage.lblGuild.Caption = GetPlayerGuild(MyIndex)
    frmMirage.lblRank.Caption = GetPlayerGuildAccess(MyIndex)
    picGuild.Visible = True
End Sub

Private Sub lblInv_Click()
    Call UpdateVisInv
    picInv3.Visible = True
End Sub

Private Sub lblOpt_Click()
    picOptions.Visible = True
End Sub

Private Sub lblQit_Click()
    Call GameDestroy
End Sub

Private Sub lblSpl_Click()
    Call SendData(SPELLS_CHAR & END_CHAR)
    picPlayerSpells.Visible = True
End Sub

Private Sub lblWho_Click()
    Call SendOnlineList
    picWhosOnline.Visible = True
End Sub

Private Sub lstFriend_DblClick()
    Call SendData(PLAYERCHAT_CHAR & SEP_CHAR & Trim$(lstFriend.Text) & END_CHAR)
End Sub

Private Sub lstOnline_DblClick()
    Call SendData(PLAYERCHAT_CHAR & SEP_CHAR & Trim$(lstOnline.Text) & END_CHAR)
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
            Else
                descName.Caption = Trim$(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Name)
            End If
        End If
        descStr.Caption = Item(GetPlayerInvItemNum(MyIndex, d + 1)).StrReq & " Strength"
        descDef.Caption = Item(GetPlayerInvItemNum(MyIndex, d + 1)).DefReq & " Defence"
        descSpeed.Caption = Item(GetPlayerInvItemNum(MyIndex, d + 1)).SpeedReq & " Speed"
        descMagic.Caption = Item(GetPlayerInvItemNum(MyIndex, d + 1)).MagicReq & " Magic"
        descHpMp.Caption = "HP: " & Item(GetPlayerInvItemNum(MyIndex, d + 1)).AddHP & " MP: " & Item(GetPlayerInvItemNum(MyIndex, d + 1)).AddMP & " SP: " & Item(GetPlayerInvItemNum(MyIndex, d + 1)).AddSP
        descSD.Caption = "Str: " & Item(GetPlayerInvItemNum(MyIndex, d + 1)).AddStr & " Def: " & Item(GetPlayerInvItemNum(MyIndex, d + 1)).AddDef
        descMS.Caption = "Magi: " & Item(GetPlayerInvItemNum(MyIndex, d + 1)).AddMagi & " Speed: " & Item(GetPlayerInvItemNum(MyIndex, d + 1)).AddSpeed
        desc.Caption = Trim$(Item(GetPlayerInvItemNum(MyIndex, d + 1)).desc)
        
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
        If Player(MyIndex).Guildaccess = 5 Then
            frmMirage.picGuildAdmin.Visible = True
            frmMirage.picInv3.Visible = False
            frmMirage.picGuild.Visible = False
            frmMirage.picEquip.Visible = False
            frmMirage.picPlayerSpells.Visible = False
            frmMirage.picWhosOnline.Visible = False
        End If
    End If
      
    If KeyCode = vbKeyInsert Then
        If SpellMemorized > 0 Then
            If GetTickCount > Player(MyIndex).AttackTimer + 1000 Then
                If Player(MyIndex).Moving = 0 Then
                    Call SendData(CAST_CHAR & SEP_CHAR & SpellMemorized & END_CHAR)
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
        I = 0
        ii = 0
        If LCase(Dir(App.Path & "\Screenshots", vbDirectory)) <> "screenshots" Then
            Call MkDir(App.Path & "\Screenshots")
        End If
        Do
            If ScreenFileExist("Screenshot" & I & ".bmp") = True Then
                I = I + 1
            Else
                Call SavePicture(ScreenShot.Picture, App.Path & "\Screenshots\Screenshot" & I & ".bmp")
                ii = 1
            End If
           
            DoEvents
        Loop Until ii = 1
    ElseIf KeyCode = vbKeyF12 Then
        ScreenShot.Picture = CaptureArea(frmMirage, picScreen.Left, picScreen.Top, picScreen.Width, picScreen.Height)
        I = 0
        ii = 0
        If LCase(Dir(App.Path & "\Screenshots", vbDirectory)) <> "screenshots" Then
            Call MkDir(App.Path & "\Screenshots")
        End If
        Do
            If ScreenFileExist("Screenshot" & I & ".bmp") = True Then
                I = I + 1
            Else
                Call SavePicture(ScreenShot.Picture, App.Path & "\Screenshots\Screenshot" & I & ".bmp")
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

Private Sub picStatus_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    SOffsetX = x
    SOffsetY = y
    Call picStatus.ZOrder(0)
End Sub

Private Sub picStatus_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call MovePicture(frmMirage.picStatus, Button, Shift, x, y)
End Sub

Private Sub picGuild_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    SOffsetX = x
    SOffsetY = y
    Call picGuild.ZOrder(0)
End Sub

Private Sub picStat_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    SOffsetX = x
    SOffsetY = y
    Call picStat.ZOrder(0)
End Sub

Private Sub picWhosOnline_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    SOffsetX = x
    SOffsetY = y
    Call picWhosOnline.ZOrder(0)
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
    SOffsetX = x
    SOffsetY = y
    Call picFriend.ZOrder(0)
End Sub

Private Sub picFriend_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call MovePicture(frmMirage.picFriend, Button, Shift, x, y)
End Sub

Private Sub itmDesc_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    SOffsetX = x
    SOffsetY = y
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

Private Sub picGuild_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call MovePicture(frmMirage.picGuild, Button, Shift, x, y)
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
    
    If (Button = 1 Or Button = 2) And InEditor = False Then
        If Button = 1 And Player(MyIndex).Pet.Alive = YES Then
            Call PetMove(Button, Shift, (x + (NewPlayerX * PIC_X)), (y + (NewPlayerY * PIC_Y)))
        Else
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

Private Sub scrlMusicVol_Change()
    lblMusicVol.Caption = scrlMusicVol.Value
    Call PutVar(App.Path & "\config.ini", "CONFIG", "MusicVol", scrlMusicVol.Value)
    MusicVolume = scrlMusicVol.Value
    perf.SetMasterVolume (MusicVolume * 42 - 3000)
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
On Error GoTo done
If Inventory <= 0 Then Exit Sub

    InvNum = Inventory
    If GetPlayerInvItemNum(MyIndex, InvNum) > 0 And GetPlayerInvItemNum(MyIndex, InvNum) <= MAX_ITEMS Then
        If Item(GetPlayerInvItemNum(MyIndex, InvNum)).Type = ITEM_TYPE_CURRENCY Then
            GoldAmount = InputBox("How much " & Trim$(Item(GetPlayerInvItemNum(MyIndex, InvNum)).Name) & "(" & GetPlayerInvItemValue(MyIndex, InvNum) & ") would you like to drop?", "Drop " & Trim$(Item(GetPlayerInvItemNum(MyIndex, InvNum)).Name), 0, frmMirage.Left, frmMirage.Top)
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
done:
    If Item(GetPlayerInvItemNum(MyIndex, InvNum)).Type = ITEM_TYPE_CURRENCY Then
        MsgBox "The variable can't handle that amount!"
    End If
End Sub

Private Sub lblCast_Click()
    If Player(MyIndex).Spell(lstSpells.ListIndex + 1) > 0 Then
        If GetTickCount > Player(MyIndex).AttackTimer + 1000 Then
            If Player(MyIndex).Moving = 0 Then
                Call SendData(CAST_CHAR & SEP_CHAR & lstSpells.ListIndex + 1 & END_CHAR)
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
    Call SendData(SPELLS_CHAR & END_CHAR)
End Sub

Private Sub picStats_Click()
    Call SendData(GETSTATS_CHAR & END_CHAR)
End Sub

Private Sub picTrade_Click()
    Call SendData(TRADE_CHAR & END_CHAR)
End Sub

Private Sub picQuit_Click()
    Call GameDestroy
End Sub

Private Sub cmdAccess_Click()
Dim Packet As String

    If txtName.Text = "" Then
       Call AddText("Type the player name.", BrightRed)
       Exit Sub
    ElseIf txtAccess.Text = "" Then
       Call AddText("Type the player access level.", BrightRed)
       Exit Sub
    End If

    Packet = GUILDCHANGEACCESS_CHAR & SEP_CHAR & txtName.Text & SEP_CHAR & txtAccess.Text & END_CHAR
    Call SendData(Packet)
End Sub

Private Sub cmdDisown_Click()
Dim Packet As String

    Packet = GUILDDISOWN_CHAR & SEP_CHAR & txtName.Text & END_CHAR
    Call SendData(Packet)
End Sub

Private Sub cmdTrainee_Click()
Dim Packet As String
    
    Packet = GUILDTRAINEE_CHAR & SEP_CHAR & txtName.Text & END_CHAR
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
