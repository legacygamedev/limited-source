VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMirage 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bevrias Engine"
   ClientHeight    =   9105
   ClientLeft      =   300
   ClientTop       =   630
   ClientWidth     =   10650
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
   Picture         =   "frmMirage.frx":038A
   ScaleHeight     =   607
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   710
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.PictureBox picBanker 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   3945
      Left            =   2040
      Picture         =   "frmMirage.frx":2230
      ScaleHeight     =   3915
      ScaleWidth      =   6150
      TabIndex        =   189
      Top             =   1920
      Width           =   6180
      Visible         =   0   'False
      Begin VB.PictureBox picAmmountBank 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   120
         Picture         =   "frmMirage.frx":2925
         ScaleHeight     =   705
         ScaleWidth      =   1785
         TabIndex        =   240
         Top             =   3120
         Width           =   1815
         Visible         =   0   'False
         Begin VB.TextBox txtAmmountBank 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   0
            TabIndex        =   241
            Text            =   "1"
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label lblWithdraw 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Withdraw"
            Height          =   255
            Left            =   0
            TabIndex        =   243
            Top             =   495
            Width           =   1815
         End
         Begin VB.Label Label66 
            BackStyle       =   0  'Transparent
            Caption         =   "How much to Withdraw"
            Height          =   255
            Left            =   0
            TabIndex        =   242
            Top             =   0
            Width           =   1815
         End
      End
      Begin VB.PictureBox picBank 
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
         Index           =   48
         Left            =   5280
         Picture         =   "frmMirage.frx":301A
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   239
         Top             =   2520
         Width           =   480
      End
      Begin VB.PictureBox picBank 
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
         Index           =   47
         Left            =   4680
         Picture         =   "frmMirage.frx":33AD
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   238
         Top             =   2520
         Width           =   480
      End
      Begin VB.PictureBox picBank 
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
         Index           =   46
         Left            =   4080
         Picture         =   "frmMirage.frx":3740
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   237
         Top             =   2520
         Width           =   480
      End
      Begin VB.PictureBox picBank 
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
         Index           =   45
         Left            =   3480
         Picture         =   "frmMirage.frx":3AD3
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   236
         Top             =   2520
         Width           =   480
      End
      Begin VB.PictureBox picBank 
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
         Index           =   44
         Left            =   2830
         Picture         =   "frmMirage.frx":3E66
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   235
         Top             =   2520
         Width           =   480
      End
      Begin VB.PictureBox picBank 
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
         Index           =   43
         Left            =   2160
         Picture         =   "frmMirage.frx":41F9
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   234
         Top             =   2520
         Width           =   480
      End
      Begin VB.PictureBox picBank 
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
         Index           =   42
         Left            =   1560
         Picture         =   "frmMirage.frx":458C
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   233
         Top             =   2520
         Width           =   480
      End
      Begin VB.PictureBox picBank 
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
         Index           =   41
         Left            =   960
         Picture         =   "frmMirage.frx":491F
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   232
         Top             =   2520
         Width           =   480
      End
      Begin VB.PictureBox picBank 
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
         Index           =   40
         Left            =   360
         Picture         =   "frmMirage.frx":4CB2
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   231
         Top             =   2520
         Width           =   480
      End
      Begin VB.PictureBox picBank 
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
         Index           =   39
         Left            =   5520
         Picture         =   "frmMirage.frx":5045
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   230
         Top             =   1920
         Width           =   480
      End
      Begin VB.PictureBox picBank 
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
         Index           =   38
         Left            =   4920
         Picture         =   "frmMirage.frx":53D8
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   229
         Top             =   1920
         Width           =   480
      End
      Begin VB.PictureBox picBank 
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
         Index           =   37
         Left            =   4320
         Picture         =   "frmMirage.frx":576B
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   228
         Top             =   1920
         Width           =   480
      End
      Begin VB.PictureBox picBank 
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
         Index           =   36
         Left            =   3720
         Picture         =   "frmMirage.frx":5AFE
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   227
         Top             =   1920
         Width           =   480
      End
      Begin VB.PictureBox picBank 
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
         Left            =   3120
         Picture         =   "frmMirage.frx":5E91
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   226
         Top             =   1920
         Width           =   480
      End
      Begin VB.PictureBox picBank 
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
         Left            =   2520
         Picture         =   "frmMirage.frx":6224
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   225
         Top             =   1920
         Width           =   480
      End
      Begin VB.PictureBox picBank 
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
         Left            =   1920
         Picture         =   "frmMirage.frx":65B7
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   224
         Top             =   1920
         Width           =   480
      End
      Begin VB.PictureBox picBank 
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
         Left            =   1320
         Picture         =   "frmMirage.frx":694A
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   223
         Top             =   1920
         Width           =   480
      End
      Begin VB.PictureBox picBank 
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
         Left            =   720
         Picture         =   "frmMirage.frx":6CDD
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   222
         Top             =   1920
         Width           =   480
      End
      Begin VB.PictureBox picBank 
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
         Left            =   120
         Picture         =   "frmMirage.frx":7070
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   221
         Top             =   1920
         Width           =   480
      End
      Begin VB.PictureBox picBank 
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
         Left            =   5520
         Picture         =   "frmMirage.frx":7403
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   220
         Top             =   1320
         Width           =   480
      End
      Begin VB.PictureBox picBank 
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
         Left            =   4920
         Picture         =   "frmMirage.frx":7796
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   219
         Top             =   1320
         Width           =   480
      End
      Begin VB.PictureBox picBank 
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
         Left            =   4320
         Picture         =   "frmMirage.frx":7B29
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   218
         Top             =   1320
         Width           =   480
      End
      Begin VB.PictureBox picBank 
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
         Left            =   3720
         Picture         =   "frmMirage.frx":7EBC
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   217
         Top             =   1320
         Width           =   480
      End
      Begin VB.PictureBox picBank 
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
         Left            =   3120
         Picture         =   "frmMirage.frx":824F
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   216
         Top             =   1320
         Width           =   480
      End
      Begin VB.PictureBox picBank 
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
         Picture         =   "frmMirage.frx":85E2
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   215
         Top             =   1320
         Width           =   480
      End
      Begin VB.PictureBox picBank 
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
         Picture         =   "frmMirage.frx":8975
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   214
         Top             =   1320
         Width           =   480
      End
      Begin VB.PictureBox picBank 
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
         Picture         =   "frmMirage.frx":8D08
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   213
         Top             =   1320
         Width           =   480
      End
      Begin VB.PictureBox picBank 
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
         Picture         =   "frmMirage.frx":909B
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   212
         Top             =   1320
         Width           =   480
      End
      Begin VB.PictureBox picBank 
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
         Left            =   5520
         Picture         =   "frmMirage.frx":942E
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   211
         Top             =   720
         Width           =   480
      End
      Begin VB.PictureBox picBank 
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
         Left            =   4920
         Picture         =   "frmMirage.frx":97C1
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   210
         Top             =   720
         Width           =   480
      End
      Begin VB.PictureBox picBank 
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
         Left            =   4320
         Picture         =   "frmMirage.frx":9B54
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   209
         Top             =   720
         Width           =   480
      End
      Begin VB.PictureBox picBank 
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
         Left            =   3720
         Picture         =   "frmMirage.frx":9EE7
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   208
         Top             =   720
         Width           =   480
      End
      Begin VB.PictureBox picBank 
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
         Left            =   3120
         Picture         =   "frmMirage.frx":A27A
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   207
         Top             =   720
         Width           =   480
      End
      Begin VB.PictureBox picBank 
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
         Left            =   2520
         Picture         =   "frmMirage.frx":A60D
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   206
         Top             =   720
         Width           =   480
      End
      Begin VB.PictureBox picBank 
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
         Left            =   1920
         Picture         =   "frmMirage.frx":A9A0
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   205
         Top             =   720
         Width           =   480
      End
      Begin VB.PictureBox picBank 
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
         Left            =   1320
         Picture         =   "frmMirage.frx":AD33
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   204
         Top             =   720
         Width           =   480
      End
      Begin VB.PictureBox picBank 
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
         Left            =   720
         Picture         =   "frmMirage.frx":B0C6
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   203
         Top             =   720
         Width           =   480
      End
      Begin VB.PictureBox picBank 
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
         Left            =   120
         Picture         =   "frmMirage.frx":B459
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   202
         Top             =   720
         Width           =   480
      End
      Begin VB.PictureBox picBank 
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
         Left            =   5520
         Picture         =   "frmMirage.frx":B7EC
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   201
         Top             =   120
         Width           =   480
      End
      Begin VB.PictureBox picBank 
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
         Left            =   4920
         Picture         =   "frmMirage.frx":BB7F
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   200
         Top             =   120
         Width           =   480
      End
      Begin VB.PictureBox picBank 
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
         Left            =   4320
         Picture         =   "frmMirage.frx":BF12
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   199
         Top             =   120
         Width           =   480
      End
      Begin VB.PictureBox picBank 
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
         Left            =   3720
         Picture         =   "frmMirage.frx":C2A5
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   198
         Top             =   120
         Width           =   480
      End
      Begin VB.PictureBox picBank 
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
         Left            =   3120
         Picture         =   "frmMirage.frx":C638
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   197
         Top             =   120
         Width           =   480
      End
      Begin VB.PictureBox picBank 
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
         Left            =   2520
         Picture         =   "frmMirage.frx":C9CB
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   196
         Top             =   120
         Width           =   480
      End
      Begin VB.PictureBox picBank 
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
         Picture         =   "frmMirage.frx":CD5E
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   195
         Top             =   120
         Width           =   480
      End
      Begin VB.PictureBox picBank 
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
         Picture         =   "frmMirage.frx":D0F1
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   194
         Top             =   120
         Width           =   480
      End
      Begin VB.PictureBox picBank 
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
         Picture         =   "frmMirage.frx":D484
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   193
         Top             =   120
         Width           =   480
      End
      Begin VB.PictureBox picBank 
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
         Picture         =   "frmMirage.frx":D817
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   192
         Top             =   120
         Width           =   480
      End
      Begin VB.PictureBox picBank 
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
         Left            =   2520
         Picture         =   "frmMirage.frx":DBAA
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   191
         Top             =   1320
         Width           =   480
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "Close"
         Height          =   255
         Left            =   4320
         TabIndex        =   190
         Top             =   3600
         Width           =   1335
      End
      Begin VB.Timer Timer1 
         Left            =   3120
         Top             =   3240
      End
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
      Height          =   1935
      Left            =   2520
      Picture         =   "frmMirage.frx":DF3D
      ScaleHeight     =   127
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   335
      TabIndex        =   176
      Top             =   3960
      Width           =   5055
      Visible         =   0   'False
      Begin VB.Label descMS 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Wis: XXXXX Agi: XXXX"
         ForeColor       =   &H00E0E0E0&
         Height          =   210
         Left            =   2280
         TabIndex        =   188
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "-Description-"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   2760
         TabIndex        =   187
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label desc 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00E0E0E0&
         Height          =   1095
         Left            =   2520
         TabIndex        =   186
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label descSD 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Str: XXXX Def: XXXXX"
         ForeColor       =   &H00E0E0E0&
         Height          =   210
         Left            =   0
         TabIndex        =   185
         Top             =   1680
         Width           =   2655
      End
      Begin VB.Label descHpMp 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "HP: XXXX MP: XXXX SP: XXXX"
         ForeColor       =   &H00E0E0E0&
         Height          =   210
         Left            =   0
         TabIndex        =   184
         Top             =   1440
         Width           =   2655
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "-Add-"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   360
         TabIndex        =   183
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label descSpeed 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Agility"
         ForeColor       =   &H00E0E0E0&
         Height          =   210
         Left            =   360
         TabIndex        =   182
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label descDef 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Defence"
         ForeColor       =   &H00E0E0E0&
         Height          =   210
         Left            =   360
         TabIndex        =   181
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label descStr 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Strength"
         ForeColor       =   &H00E0E0E0&
         Height          =   210
         Left            =   360
         TabIndex        =   180
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "-Requirements-"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   360
         TabIndex        =   179
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label descName 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         ForeColor       =   &H00E0E0E0&
         Height          =   210
         Left            =   360
         TabIndex        =   178
         Top             =   0
         Width           =   1815
      End
      Begin VB.Label Label65 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "-Add-"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   2760
         TabIndex        =   177
         Top             =   0
         Width           =   1815
      End
   End
   Begin VB.PictureBox Picture35 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   1665
      Picture         =   "frmMirage.frx":E632
      ScaleHeight     =   1665
      ScaleWidth      =   7305
      TabIndex        =   153
      Top             =   7395
      Width           =   7335
      Visible         =   0   'False
      Begin VB.Label lblMP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0000C000&
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
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   2640
         TabIndex        =   267
         Top             =   480
         Width           =   1890
      End
      Begin VB.Label lblHP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0000C000&
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
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   2640
         TabIndex        =   266
         Top             =   240
         Width           =   1890
      End
      Begin VB.Shape Shape1 
         BorderStyle     =   0  'Transparent
         Height          =   180
         Left            =   2880
         Top             =   750
         Width           =   1890
      End
      Begin VB.Label lblEXP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0000C000&
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
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   2640
         TabIndex        =   265
         Top             =   750
         Width           =   1890
      End
      Begin VB.Label Label74 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Options"
         Height          =   255
         Left            =   6120
         TabIndex        =   264
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label lblname 
         BackStyle       =   0  'Transparent
         Caption         =   "None"
         Height          =   255
         Left            =   5280
         TabIndex        =   263
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label Label73 
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         Height          =   255
         Left            =   4680
         TabIndex        =   262
         Top             =   120
         Width           =   615
      End
      Begin VB.Label lblclass 
         BackStyle       =   0  'Transparent
         Caption         =   "None"
         Height          =   255
         Left            =   5280
         TabIndex        =   175
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label dsdaf 
         BackStyle       =   0  'Transparent
         Caption         =   "Class:"
         Height          =   255
         Left            =   4680
         TabIndex        =   174
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblLevel 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Left            =   1080
         TabIndex        =   173
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label64 
         BackStyle       =   0  'Transparent
         Caption         =   "Level:"
         Height          =   255
         Left            =   120
         TabIndex        =   172
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label63 
         BackStyle       =   0  'Transparent
         Caption         =   "Points:"
         Height          =   255
         Left            =   120
         TabIndex        =   171
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lblpoints 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Left            =   1080
         TabIndex        =   170
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label62 
         BackStyle       =   0  'Transparent
         Caption         =   "Tnl:"
         Height          =   255
         Left            =   2280
         TabIndex        =   169
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label61 
         BackStyle       =   0  'Transparent
         Caption         =   "Mana:"
         Height          =   255
         Left            =   2280
         TabIndex        =   168
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label60 
         BackStyle       =   0  'Transparent
         Caption         =   "Health:"
         Height          =   255
         Left            =   2280
         TabIndex        =   167
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label59 
         BackStyle       =   0  'Transparent
         Caption         =   "Wisdom:"
         Height          =   255
         Left            =   120
         TabIndex        =   166
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label58 
         BackStyle       =   0  'Transparent
         Caption         =   "Agility:"
         Height          =   255
         Left            =   120
         TabIndex        =   165
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label57 
         BackStyle       =   0  'Transparent
         Caption         =   "Defence:"
         Height          =   255
         Left            =   120
         TabIndex        =   164
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label55 
         BackStyle       =   0  'Transparent
         Caption         =   "Strength:"
         Height          =   255
         Left            =   120
         TabIndex        =   162
         Top             =   120
         Width           =   855
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
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   1920
         TabIndex        =   161
         Top             =   840
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
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   1920
         TabIndex        =   160
         Top             =   600
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
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   1920
         TabIndex        =   159
         Top             =   360
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
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   1920
         TabIndex        =   158
         Top             =   120
         Width           =   165
      End
      Begin VB.Label lblSPEED 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   180
         Left            =   1080
         TabIndex        =   157
         Top             =   600
         Width           =   660
      End
      Begin VB.Label lblMAGI 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   180
         Left            =   1080
         TabIndex        =   156
         Top             =   840
         Width           =   690
      End
      Begin VB.Label lblDEF 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   180
         Left            =   1080
         TabIndex        =   155
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblSTR 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   180
         Left            =   1080
         TabIndex        =   154
         Top             =   120
         Width           =   720
      End
   End
   Begin VB.PictureBox Picture8 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   1665
      Picture         =   "frmMirage.frx":ED27
      ScaleHeight     =   1665
      ScaleWidth      =   7305
      TabIndex        =   146
      Top             =   7395
      Width           =   7335
      Visible         =   0   'False
      Begin VB.ListBox lstOnline 
         Appearance      =   0  'Flat
         BackColor       =   &H0084ADB3&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1470
         ItemData        =   "frmMirage.frx":F41C
         Left            =   120
         List            =   "frmMirage.frx":F41E
         TabIndex        =   147
         Top             =   120
         Width           =   7020
      End
   End
   Begin VB.PictureBox Picture31 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   1665
      Picture         =   "frmMirage.frx":F420
      ScaleHeight     =   1665
      ScaleWidth      =   7305
      TabIndex        =   128
      Top             =   7395
      Width           =   7335
      Visible         =   0   'False
      Begin VB.CommandButton cmdTrainee 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   5040
         Style           =   1  'Graphical
         TabIndex        =   145
         Top             =   960
         Width           =   1575
      End
      Begin VB.CommandButton cmdMember 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   5040
         Style           =   1  'Graphical
         TabIndex        =   144
         Top             =   720
         Width           =   1575
      End
      Begin VB.CommandButton cmdDisown 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   5040
         Style           =   1  'Graphical
         TabIndex        =   143
         Top             =   480
         Width           =   1575
      End
      Begin VB.CommandButton cmdAccess 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   5040
         Style           =   1  'Graphical
         TabIndex        =   142
         Top             =   240
         Width           =   1575
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
         Left            =   120
         Picture         =   "frmMirage.frx":FB15
         ScaleHeight     =   2505
         ScaleWidth      =   2145
         TabIndex        =   135
         Top             =   0
         Width           =   2145
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            Caption         =   "Guild"
            Height          =   375
            Left            =   600
            TabIndex        =   141
            Top             =   75
            Width           =   495
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
            Height          =   180
            Left            =   720
            TabIndex        =   140
            Top             =   1200
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
            Left            =   1080
            TabIndex        =   139
            Top             =   840
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
            Left            =   1080
            TabIndex        =   138
            Top             =   480
            Width           =   1065
         End
         Begin VB.Label Label17 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Guild Access :"
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
            Left            =   120
            TabIndex        =   137
            Top             =   840
            Width           =   870
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
            Left            =   120
            TabIndex        =   136
            Top             =   480
            Width           =   825
         End
      End
      Begin VB.PictureBox picGuildAdmin 
         Appearance      =   0  'Flat
         BackColor       =   &H00004080&
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
         Height          =   2265
         Left            =   1800
         Picture         =   "frmMirage.frx":FC28
         ScaleHeight     =   2265
         ScaleWidth      =   2985
         TabIndex        =   129
         Top             =   0
         Width           =   2985
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
            Height          =   240
            Left            =   1080
            TabIndex        =   131
            Top             =   930
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
            Height          =   240
            Left            =   1080
            TabIndex        =   130
            Top             =   580
            Width           =   1575
         End
         Begin VB.Label Label21 
            BackStyle       =   0  'Transparent
            Caption         =   "Options"
            Height          =   255
            Left            =   1440
            TabIndex        =   134
            Top             =   60
            Width           =   855
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
            Left            =   480
            TabIndex        =   133
            Top             =   960
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
            Left            =   480
            TabIndex        =   132
            Top             =   600
            Width           =   420
         End
      End
   End
   Begin VB.PictureBox Picture30 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   2055
      Left            =   7920
      Picture         =   "frmMirage.frx":1031D
      ScaleHeight     =   2025
      ScaleWidth      =   2145
      TabIndex        =   103
      Top             =   5280
      Width           =   2175
      Visible         =   0   'False
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
         Height          =   2025
         Left            =   0
         Picture         =   "frmMirage.frx":10A12
         ScaleHeight     =   2025
         ScaleWidth      =   2025
         TabIndex        =   104
         Top             =   0
         Width           =   2025
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
            Left            =   3240
            Picture         =   "frmMirage.frx":10B25
            ScaleHeight     =   2.23636e5
            ScaleMode       =   0  'User
            ScaleWidth      =   477.091
            TabIndex        =   126
            Top             =   720
            Width           =   480
            Visible         =   0   'False
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
            Left            =   840
            ScaleHeight     =   35
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   35
            TabIndex        =   124
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
               TabIndex        =   125
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
            Left            =   1320
            ScaleHeight     =   35
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   35
            TabIndex        =   122
            Top             =   720
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
               TabIndex        =   123
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
            Left            =   1320
            ScaleHeight     =   35
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   35
            TabIndex        =   120
            Top             =   1320
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
               TabIndex        =   121
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
            Left            =   0
            ScaleHeight     =   35
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   35
            TabIndex        =   118
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
               TabIndex        =   119
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
            Left            =   2760
            ScaleHeight     =   35
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   35
            TabIndex        =   116
            Top             =   720
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
               TabIndex        =   117
               Top             =   0
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
            Left            =   2760
            ScaleHeight     =   35
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   35
            TabIndex        =   114
            Top             =   1920
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
               TabIndex        =   115
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
            Left            =   1800
            ScaleHeight     =   35
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   35
            TabIndex        =   112
            Top             =   2520
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
               TabIndex        =   113
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
            Left            =   2760
            ScaleHeight     =   35
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   35
            TabIndex        =   110
            Top             =   1320
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
               TabIndex        =   111
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
            Left            =   2760
            ScaleHeight     =   35
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   35
            TabIndex        =   108
            Top             =   2520
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
               TabIndex        =   109
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
            Left            =   2760
            ScaleHeight     =   35
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   35
            TabIndex        =   106
            Top             =   120
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
               TabIndex        =   107
               Top             =   15
               Width           =   495
            End
         End
         Begin VB.PictureBox Picture11 
            BorderStyle     =   0  'None
            Height          =   967
            Left            =   720
            Picture         =   "frmMirage.frx":170467
            ScaleHeight     =   960
            ScaleWidth      =   480
            TabIndex        =   105
            Top             =   720
            Width           =   487
         End
         Begin VB.Label Label20 
            BackStyle       =   0  'Transparent
            Caption         =   "Close"
            Height          =   255
            Left            =   0
            TabIndex        =   127
            Top             =   1680
            Width           =   495
         End
      End
   End
   Begin VB.PictureBox Picture34 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   1665
      ScaleHeight     =   1665
      ScaleWidth      =   7290
      TabIndex        =   77
      Top             =   7395
      Width           =   7320
      Visible         =   0   'False
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
         Height          =   2175
         Left            =   0
         Picture         =   "frmMirage.frx":1707CA
         ScaleHeight     =   145
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   489
         TabIndex        =   78
         Top             =   0
         Width           =   7335
         Begin VB.PictureBox picAmmount 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   930
            Left            =   4920
            Picture         =   "frmMirage.frx":170EBF
            ScaleHeight     =   900
            ScaleWidth      =   2265
            TabIndex        =   256
            Top             =   120
            Width           =   2295
            Visible         =   0   'False
            Begin VB.TextBox txtAmmount 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               Height          =   285
               Left            =   240
               TabIndex        =   257
               Text            =   "1"
               Top             =   240
               Width           =   1815
            End
            Begin VB.Label Label72 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Double Click To Deposit a Item"
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
               Left            =   0
               TabIndex        =   261
               Top             =   720
               Width           =   2280
            End
            Begin VB.Label lblDeposit 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "- Deposit -"
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
               Left            =   720
               TabIndex        =   260
               Top             =   530
               Width           =   720
            End
            Begin VB.Label Label71 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "How much to Deposit"
               Height          =   255
               Left            =   240
               TabIndex        =   258
               Top             =   0
               Width           =   1815
            End
         End
         Begin VB.PictureBox Picture40 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   4920
            Picture         =   "frmMirage.frx":1715B4
            ScaleHeight     =   225
            ScaleWidth      =   2265
            TabIndex        =   250
            Top             =   1080
            Width           =   2295
            Visible         =   0   'False
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
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   1440
               TabIndex        =   252
               Top             =   0
               Width           =   720
            End
            Begin VB.Label Label35 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "*Protection Spells*"
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
               Left            =   0
               TabIndex        =   251
               Top             =   0
               Width           =   1335
            End
         End
         Begin VB.PictureBox Picture39 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   4920
            Picture         =   "frmMirage.frx":171CA9
            ScaleHeight     =   225
            ScaleWidth      =   2265
            TabIndex        =   249
            Top             =   1080
            Width           =   2295
            Visible         =   0   'False
            Begin VB.Label Label33 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Paperdoll"
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
               Left            =   1440
               TabIndex        =   255
               Top             =   0
               Width           =   675
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
               ForeColor       =   &H00000000&
               Height          =   210
               Left            =   690
               TabIndex        =   254
               Top             =   0
               Width           =   675
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
               Left            =   0
               TabIndex        =   253
               Top             =   0
               Width           =   690
            End
         End
         Begin VB.PictureBox Picture38 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   4920
            Picture         =   "frmMirage.frx":17239E
            ScaleHeight     =   225
            ScaleWidth      =   2265
            TabIndex        =   245
            Top             =   1320
            Width           =   2295
            Begin VB.Label Label70 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "None"
               Height          =   255
               Left            =   1680
               TabIndex        =   259
               Top             =   0
               Width           =   615
            End
            Begin VB.Label Label69 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Spells"
               Height          =   255
               Left            =   1200
               TabIndex        =   248
               Top             =   0
               Width           =   495
            End
            Begin VB.Label Label68 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Bank"
               Height          =   255
               Left            =   600
               TabIndex        =   247
               Top             =   0
               Width           =   615
            End
            Begin VB.Label Label67 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Items"
               Height          =   255
               Left            =   0
               TabIndex        =   246
               Top             =   0
               Width           =   615
            End
         End
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
            Height          =   930
            ItemData        =   "frmMirage.frx":172A93
            Left            =   4920
            List            =   "frmMirage.frx":172A95
            TabIndex        =   244
            Top             =   120
            Width           =   2295
            Visible         =   0   'False
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
            Picture         =   "frmMirage.frx":172A97
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   102
            Top             =   600
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
            Left            =   4320
            Picture         =   "frmMirage.frx":172E2A
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   101
            Top             =   75
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
            Left            =   3720
            Picture         =   "frmMirage.frx":1731BD
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   100
            Top             =   75
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
            Left            =   3120
            Picture         =   "frmMirage.frx":173550
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   99
            Top             =   75
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
            Left            =   2520
            Picture         =   "frmMirage.frx":1738E3
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   98
            Top             =   75
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
            Picture         =   "frmMirage.frx":173C76
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   97
            Top             =   75
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
            Picture         =   "frmMirage.frx":174009
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   96
            Top             =   75
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
            Picture         =   "frmMirage.frx":17439C
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   95
            Top             =   75
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
            Picture         =   "frmMirage.frx":17472F
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   94
            Top             =   75
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
            Picture         =   "frmMirage.frx":174AC2
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   93
            Top             =   600
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
            Picture         =   "frmMirage.frx":174E55
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   92
            Top             =   600
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
            Picture         =   "frmMirage.frx":1751E8
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   91
            Top             =   600
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
            Left            =   2520
            Picture         =   "frmMirage.frx":17557B
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   90
            Top             =   600
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
            Left            =   3120
            Picture         =   "frmMirage.frx":17590E
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   89
            Top             =   600
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
            Left            =   3720
            Picture         =   "frmMirage.frx":175CA1
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   88
            Top             =   600
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
            Left            =   4320
            Picture         =   "frmMirage.frx":176034
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   87
            Top             =   600
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
            Picture         =   "frmMirage.frx":1763C7
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   86
            Top             =   1125
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
            Picture         =   "frmMirage.frx":17675A
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   85
            Top             =   1125
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
            Picture         =   "frmMirage.frx":176AED
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   84
            Top             =   1125
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
            Picture         =   "frmMirage.frx":176E80
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   83
            Top             =   1125
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
            Left            =   2520
            Picture         =   "frmMirage.frx":177213
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   82
            Top             =   1125
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
            Left            =   3120
            Picture         =   "frmMirage.frx":1775A6
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   81
            Top             =   1125
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
            Left            =   3720
            Picture         =   "frmMirage.frx":177939
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   80
            Top             =   1125
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
            Left            =   4320
            Picture         =   "frmMirage.frx":177CCC
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   79
            Top             =   1125
            Width           =   480
         End
         Begin VB.Shape SelectedItem 
            BorderColor     =   &H000000FF&
            BorderWidth     =   2
            Height          =   525
            Left            =   105
            Top             =   75
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
   Begin VB.PictureBox Picture32 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   1665
      Picture         =   "frmMirage.frx":17805F
      ScaleHeight     =   1665
      ScaleWidth      =   7290
      TabIndex        =   74
      Top             =   7395
      Width           =   7320
      Visible         =   0   'False
      Begin VB.TextBox txtMyTextBox 
         Appearance      =   0  'Flat
         BackColor       =   &H00789298&
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
         Left            =   0
         MaxLength       =   255
         TabIndex        =   75
         Top             =   0
         Width           =   7290
      End
      Begin RichTextLib.RichTextBox txtChat 
         Height          =   1350
         Left            =   0
         TabIndex        =   76
         Top             =   315
         Width           =   7290
         _ExtentX        =   12859
         _ExtentY        =   2381
         _Version        =   393217
         BackColor       =   7901848
         BorderStyle     =   0
         ScrollBars      =   2
         Appearance      =   0
         TextRTF         =   $"frmMirage.frx":178754
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
   End
   Begin VB.PictureBox Picture29 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5280
      Left            =   1440
      Picture         =   "frmMirage.frx":178818
      ScaleHeight     =   5250
      ScaleWidth      =   7500
      TabIndex        =   69
      Top             =   720
      Width           =   7530
      Visible         =   0   'False
      Begin VB.CommandButton Command4s 
         Caption         =   "Close"
         Height          =   255
         Left            =   6240
         TabIndex        =   72
         Top             =   4740
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   3375
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   71
         Top             =   1320
         Width           =   6975
      End
      Begin VB.CommandButton Command3s 
         Caption         =   "Dont Show This Again"
         Height          =   255
         Left            =   240
         TabIndex        =   70
         Top             =   4740
         Width           =   1815
      End
      Begin VB.Label Label2s 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   73
         Top             =   960
         Width           =   6975
      End
   End
   Begin VB.PictureBox Picture21 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   9360
      Picture         =   "frmMirage.frx":179B4F
      ScaleHeight     =   750
      ScaleWidth      =   750
      TabIndex        =   68
      Top             =   4680
      Width           =   750
      Visible         =   0   'False
   End
   Begin VB.Timer str 
      Enabled         =   0   'False
      Interval        =   30000
      Left            =   11040
      Top             =   7440
   End
   Begin VB.Timer Timer0 
      Interval        =   60000
      Left            =   10560
      Top             =   7440
   End
   Begin VB.Timer Timer2 
      Interval        =   60000
      Left            =   10080
      Top             =   7440
   End
   Begin VB.Timer Timer3 
      Interval        =   22
      Left            =   12120
      Top             =   8040
   End
   Begin VB.PictureBox Picture15 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6375
      Left            =   4560
      Picture         =   "frmMirage.frx":179D43
      ScaleHeight     =   6345
      ScaleWidth      =   4785
      TabIndex        =   42
      Top             =   0
      Width           =   4815
      Visible         =   0   'False
      Begin VB.PictureBox Picture28 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Left            =   240
         Picture         =   "frmMirage.frx":17A0ED
         ScaleHeight     =   750
         ScaleWidth      =   750
         TabIndex        =   65
         Top             =   5280
         Width           =   750
      End
      Begin VB.PictureBox Picture27 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Left            =   240
         Picture         =   "frmMirage.frx":17A2E1
         ScaleHeight     =   750
         ScaleWidth      =   750
         TabIndex        =   64
         Top             =   4440
         Width           =   750
      End
      Begin VB.PictureBox Picture26 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Left            =   240
         Picture         =   "frmMirage.frx":17A494
         ScaleHeight     =   750
         ScaleWidth      =   750
         TabIndex        =   63
         Top             =   3600
         Width           =   750
      End
      Begin VB.PictureBox Picture25 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Left            =   240
         Picture         =   "frmMirage.frx":17A636
         ScaleHeight     =   750
         ScaleWidth      =   750
         TabIndex        =   62
         Top             =   2760
         Width           =   750
      End
      Begin VB.PictureBox Picture24 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Left            =   240
         Picture         =   "frmMirage.frx":17A7CA
         ScaleHeight     =   750
         ScaleWidth      =   750
         TabIndex        =   61
         Top             =   1920
         Width           =   750
      End
      Begin VB.PictureBox Picture23 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Left            =   240
         Picture         =   "frmMirage.frx":17A9D1
         ScaleHeight     =   750
         ScaleWidth      =   750
         TabIndex        =   60
         Top             =   1080
         Width           =   750
      End
      Begin VB.PictureBox Picture13 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Left            =   240
         Picture         =   "frmMirage.frx":17ABA6
         ScaleHeight     =   750
         ScaleWidth      =   750
         TabIndex        =   59
         Top             =   240
         Width           =   750
      End
      Begin VB.Label Label51 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Close"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3960
         TabIndex        =   58
         Top             =   6000
         Width           =   735
      End
      Begin VB.Label Label42 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Magical"
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
         Height          =   255
         Left            =   1080
         TabIndex        =   57
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label41 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tranquility"
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
         Height          =   255
         Left            =   1080
         TabIndex        =   56
         Top             =   4080
         Width           =   735
      End
      Begin VB.Label Label40 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nimbleness"
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
         Height          =   255
         Left            =   1080
         TabIndex        =   55
         Top             =   3240
         Width           =   735
      End
      Begin VB.Label Label39 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Berserk"
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
         Height          =   255
         Left            =   1080
         TabIndex        =   54
         Top             =   5760
         Width           =   735
      End
      Begin VB.Label Label38 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Stoneskin"
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
         Height          =   255
         Left            =   1080
         TabIndex        =   53
         Top             =   4920
         Width           =   735
      End
      Begin VB.Label Label36 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Bless"
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
         Height          =   255
         Left            =   1080
         TabIndex        =   52
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label43 
         BackStyle       =   0  'Transparent
         Caption         =   "Add: 25 Max HP Duration: 1min Requirement: 25 MP"
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
         Left            =   1080
         TabIndex        =   51
         Top             =   480
         Width           =   3975
      End
      Begin VB.Label Label44 
         BackStyle       =   0  'Transparent
         Caption         =   "Add: 5 Def Duration: ~0.3min Requirement: 5 MP"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   1080
         TabIndex        =   50
         Top             =   5520
         Width           =   3975
      End
      Begin VB.Label Label45 
         BackStyle       =   0  'Transparent
         Caption         =   "Add: 5 Str Duration: ~0.5min Requirement: 5 MP"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   255
         Left            =   1080
         TabIndex        =   49
         Top             =   4680
         Width           =   3975
      End
      Begin VB.Label Label46 
         BackStyle       =   0  'Transparent
         Caption         =   "Add: 5 Agility Duration: ~0.5min Requirement: 5 MP"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   1080
         TabIndex        =   48
         Top             =   3000
         Width           =   3975
      End
      Begin VB.Label Label47 
         BackStyle       =   0  'Transparent
         Caption         =   "Add: 5 Wisdom Duration: 1min Requirement: 10 MP"
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
         Height          =   255
         Left            =   1080
         TabIndex        =   47
         Top             =   3840
         Width           =   3975
      End
      Begin VB.Label Label48 
         BackStyle       =   0  'Transparent
         Caption         =   "Add: 25 Max MP Duration: 2min Requirement: 15 MP"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   1080
         TabIndex        =   46
         Top             =   1320
         Width           =   3975
      End
      Begin VB.Label Label49 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Haste"
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
         Height          =   255
         Left            =   1080
         TabIndex        =   45
         Top             =   2400
         Width           =   735
      End
      Begin VB.Label Label50 
         BackStyle       =   0  'Transparent
         Caption         =   "Add: 25 Max SP Duration: ~1min Requirement: 10 MP"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   1080
         TabIndex        =   44
         Top             =   2160
         Width           =   3975
      End
      Begin VB.Label Label37 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Spells"
         Height          =   255
         Left            =   1200
         TabIndex        =   43
         Top             =   0
         Width           =   2535
      End
   End
   Begin VB.PictureBox Picture22 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   9360
      Picture         =   "frmMirage.frx":17AD76
      ScaleHeight     =   750
      ScaleWidth      =   750
      TabIndex        =   41
      Top             =   3960
      Width           =   750
      Visible         =   0   'False
   End
   Begin VB.PictureBox Picture20 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   9360
      Picture         =   "frmMirage.frx":17AF29
      ScaleHeight     =   750
      ScaleWidth      =   750
      TabIndex        =   40
      Top             =   2520
      Width           =   750
      Visible         =   0   'False
   End
   Begin VB.PictureBox Picture19 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   9360
      Picture         =   "frmMirage.frx":17B0BD
      ScaleHeight     =   750
      ScaleWidth      =   750
      TabIndex        =   39
      Top             =   3240
      Width           =   750
      Visible         =   0   'False
   End
   Begin VB.PictureBox Picture18 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   9360
      Picture         =   "frmMirage.frx":17B25F
      ScaleHeight     =   750
      ScaleWidth      =   750
      TabIndex        =   38
      Top             =   1080
      Width           =   750
      Visible         =   0   'False
   End
   Begin VB.PictureBox Picture16 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   9360
      Picture         =   "frmMirage.frx":17B434
      ScaleHeight     =   750
      ScaleWidth      =   750
      TabIndex        =   37
      Top             =   360
      Width           =   750
      Visible         =   0   'False
   End
   Begin VB.PictureBox Picture17 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   9360
      Picture         =   "frmMirage.frx":17B604
      ScaleHeight     =   750
      ScaleWidth      =   750
      TabIndex        =   36
      Top             =   1800
      Width           =   750
      Visible         =   0   'False
   End
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
      Height          =   4305
      Left            =   570
      Picture         =   "frmMirage.frx":17B80B
      ScaleHeight     =   285
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   173
      TabIndex        =   9
      Top             =   90
      Width           =   2625
      Visible         =   0   'False
      Begin VB.CommandButton Command1 
         Caption         =   "Close"
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
         TabIndex        =   67
         Top             =   3960
         Width           =   495
      End
      Begin VB.CommandButton Command2 
         Caption         =   "User Panel"
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
         Left            =   960
         TabIndex        =   66
         Top             =   3960
         Width           =   975
      End
      Begin VB.CheckBox chkplayerbar 
         BackColor       =   &H00789298&
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
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   720
         Value           =   1  'Checked
         Width           =   210
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
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Value           =   1  'Checked
         Width           =   210
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
         Height          =   225
         Left            =   120
         TabIndex        =   18
         Top             =   1200
         Value           =   1  'Checked
         Width           =   210
      End
      Begin VB.CheckBox chkbubblebar 
         BackColor       =   &H00789298&
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
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   2880
         Value           =   1  'Checked
         Width           =   195
      End
      Begin VB.CheckBox chknpcbar 
         BackColor       =   &H00789298&
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
         Height          =   225
         Left            =   120
         TabIndex        =   16
         Top             =   1680
         Value           =   1  'Checked
         Width           =   210
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
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   480
         Value           =   1  'Checked
         Width           =   210
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
         Height          =   225
         Left            =   120
         TabIndex        =   14
         Top             =   1440
         Value           =   1  'Checked
         Width           =   210
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
         Height          =   225
         Left            =   120
         TabIndex        =   13
         Top             =   2160
         Value           =   1  'Checked
         Width           =   210
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
         Height          =   225
         Left            =   120
         TabIndex        =   12
         Top             =   2400
         Value           =   1  'Checked
         Width           =   210
      End
      Begin VB.HScrollBar scrlBltText 
         Height          =   255
         Left            =   120
         Max             =   20
         Min             =   4
         TabIndex        =   11
         Top             =   3360
         Value           =   6
         Width           =   2295
      End
      Begin VB.CheckBox chkAutoScroll 
         BackColor       =   &H00789298&
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
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   3720
         Value           =   1  'Checked
         Width           =   210
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
         Height          =   165
         Left            =   120
         TabIndex        =   35
         Top             =   3180
         Width           =   1965
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
         ForeColor       =   &H00E0E0E0&
         Height          =   210
         Left            =   0
         TabIndex        =   34
         Top             =   0
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
         ForeColor       =   &H00E0E0E0&
         Height          =   210
         Left            =   0
         TabIndex        =   33
         Top             =   960
         Width           =   2655
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
         ForeColor       =   &H00E0E0E0&
         Height          =   210
         Left            =   0
         TabIndex        =   32
         Top             =   1920
         Width           =   2655
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
         ForeColor       =   &H00E0E0E0&
         Height          =   210
         Left            =   0
         TabIndex        =   31
         Top             =   2640
         Width           =   2655
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "Speech Bubbles"
         Height          =   255
         Left            =   360
         TabIndex        =   30
         Top             =   2880
         Width           =   1695
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "Auto Scroll"
         Height          =   255
         Left            =   360
         TabIndex        =   29
         Top             =   3720
         Width           =   1455
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "Sound"
         Height          =   255
         Left            =   360
         TabIndex        =   28
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "Music"
         Height          =   375
         Left            =   360
         TabIndex        =   27
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "Show NPC HP Bars"
         Height          =   375
         Left            =   360
         TabIndex        =   26
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         Caption         =   "Damage Above Heads"
         Height          =   375
         Left            =   360
         TabIndex        =   25
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         Caption         =   "Names"
         Height          =   255
         Left            =   360
         TabIndex        =   24
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   "Mini HP Bar"
         Height          =   255
         Left            =   360
         TabIndex        =   23
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label31 
         BackStyle       =   0  'Transparent
         Caption         =   "Damage Above Head"
         Height          =   375
         Left            =   360
         TabIndex        =   22
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label32 
         BackStyle       =   0  'Transparent
         Caption         =   "Names"
         Height          =   255
         Left            =   360
         TabIndex        =   21
         Top             =   240
         Width           =   1335
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
      Left            =   585
      Picture         =   "frmMirage.frx":17B972
      ScaleHeight     =   479
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   632
      TabIndex        =   8
      Top             =   105
      Width           =   9480
   End
   Begin VB.Timer tmrSnowDrop 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   12000
      Top             =   7440
   End
   Begin VB.Timer tmrRainDrop 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   11520
      Top             =   7440
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
      Left            =   12120
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   7
      Top             =   8520
      Width           =   615
      Visible         =   0   'False
   End
   Begin MSWinsockLib.Winsock Socket 
      Left            =   9150
      Top             =   150
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label56 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   480
      TabIndex        =   163
      Top             =   8640
      Width           =   1095
   End
   Begin VB.Label Label54 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   9120
      TabIndex        =   152
      Top             =   8640
      Width           =   1095
   End
   Begin VB.Label Label53 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   9120
      TabIndex        =   151
      Top             =   8040
      Width           =   1095
   End
   Begin VB.Label Label52 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   480
      TabIndex        =   150
      Top             =   7440
      Width           =   1095
   End
   Begin VB.Label Label34 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   9120
      TabIndex        =   149
      Top             =   7440
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   480
      TabIndex        =   148
      Top             =   8040
      Width           =   1095
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
      Height          =   195
      Left            =   9960
      TabIndex        =   6
      Top             =   8040
      Width           =   1065
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
      Height          =   285
      Left            =   9915
      TabIndex        =   5
      Top             =   8400
      Width           =   1185
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
      Height          =   285
      Left            =   11160
      TabIndex        =   4
      Top             =   7800
      Width           =   1380
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
      Height          =   195
      Left            =   11160
      TabIndex        =   3
      Top             =   8040
      Width           =   1260
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
      Height          =   180
      Left            =   9840
      TabIndex        =   2
      Top             =   8280
      Width           =   1185
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
      Height          =   300
      Left            =   9870
      TabIndex        =   1
      Top             =   7800
      Width           =   1155
   End
   Begin VB.Label Label3 
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
      Left            =   11160
      TabIndex        =   0
      Top             =   8280
      Width           =   1305
   End
   Begin VB.Menu stFile 
      Caption         =   "File"
      Begin VB.Menu stLogout 
         Caption         =   "Logout"
      End
      Begin VB.Menu stExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu stPlayer 
      Caption         =   "Player"
      Begin VB.Menu stInventory 
         Caption         =   "Inventory"
      End
      Begin VB.Menu stChat 
         Caption         =   "Chat"
      End
      Begin VB.Menu stStatss 
         Caption         =   "Stats"
      End
      Begin VB.Menu stGuild 
         Caption         =   "Guild"
      End
      Begin VB.Menu stWhos 
         Caption         =   "Who's Online"
      End
   End
   Begin VB.Menu stOption 
      Caption         =   "Options"
      Begin VB.Menu stFPS 
         Caption         =   "FPS"
      End
      Begin VB.Menu stRefresh 
         Caption         =   "Refresh"
      End
      Begin VB.Menu stStats 
         Caption         =   "Stats"
      End
      Begin VB.Menu stPrivateMessage 
         Caption         =   "Private Message"
      End
   End
   Begin VB.Menu stHelp 
      Caption         =   "Help"
      Begin VB.Menu stUserPanel 
         Caption         =   "User Panel"
      End
      Begin VB.Menu stBevriasWebsite 
         Caption         =   "Bevrias ORPGE Website"
      End
   End
End
Attribute VB_Name = "frmMirage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SpellMemorized As Long
Dim IndexClicked As Long

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

Private Sub Command2_Click()
frmUserPanel.Visible = True
End Sub

Private Sub Command3_Click()
Picture13.Visible = False
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
'lblname.Caption = GetPlayerName(MyIndex)
Label2s.Caption = ReadINI("STORY", "Headline", App.Path & "\config.ini")
Text1.Text = ReadINI("STORY", "Story", App.Path & "\config.ini")
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    itmDesc.Visible = False
End Sub

Private Sub Form_Terminate()
'If StrenghtUse = 1 Then
'Picture22.Visible = False
'Dim Packet As String
'Packet = ""
'Packet = "strength" & SEP_CHAR & "-5" & SEP_CHAR & END_CHAR
'Call SendData(Packet)
'str.Enabled = False
'End If
'If DefenceUse = 1 Then
'End If
'If AgilityUse = 1 Then
'End If
'If WisdomUse = 1 Then
'End If
Call GameDestroy
End Sub

Private Sub Form_Unload(Cancel As Integer)
'If StrenghtUse = 1 Then
'Picture22.Visible = False
'Dim Packet As String
'Packet = ""
'Packet = "strength" & SEP_CHAR & "-5" & SEP_CHAR & END_CHAR
'Call SendData(Packet)
'str.Enabled = False
'End If
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
'picInv3.Visible = False
'picPlayerSpells.Visible = False
'Picture15.Visible = False
'picEquip.Visible = False
'picWhosOnline.Visible = False
End Sub

Private Sub Label19_Click()
    picEquip.Visible = True
    Picture30.Visible = False
    'picPlayerSpells.Visible = False
    'picWhosOnline.Visible = False
    'Picture15.Visible = False
    'Picture1.Visible = False
    frmMirage.picGuildAdmin.Visible = False
    Call UpdateVisInv
End Sub

Private Sub Label2_Click()
chkplayername.Value = Trim(ReadINI("CONFIG", "playername", App.Path & "\config.ini"))
chkplayerdamage.Value = Trim(ReadINI("CONFIG", "playerdamage", App.Path & "\config.ini"))
chkplayerbar.Value = Trim(ReadINI("CONFIG", "playerbar", App.Path & "\config.ini"))
chknpcname.Value = Trim(ReadINI("CONFIG", "npcname", App.Path & "\config.ini"))
chknpcdamage.Value = Trim(ReadINI("CONFIG", "npcdamage", App.Path & "\config.ini"))
chknpcbar.Value = Trim(ReadINI("CONFIG", "npcbar", App.Path & "\config.ini"))
chkmusic.Value = Trim(ReadINI("CONFIG", "music", App.Path & "\config.ini"))
chksound.Value = Trim(ReadINI("CONFIG", "sound", App.Path & "\config.ini"))
chkbubblebar.Value = Trim(ReadINI("CONFIG", "speechbubbles", App.Path & "\config.ini"))
chkAutoScroll.Value = Trim(ReadINI("CONFIG", "AutoScroll", App.Path & "\config.ini"))
picOptions.Visible = True
End Sub

Private Sub Label20_Click()
Picture30.Visible = False
End Sub

Private Sub Label21_Click()
picEquip.Visible = False
End Sub

Private Sub Label3_Click()
If StrenghtUse = 1 Then
Picture22.Visible = False
Dim Packet As String
Packet = ""
Packet = "strength" & SEP_CHAR & "-5" & SEP_CHAR & END_CHAR
Call SendData(Packet)
str.Enabled = False
End If
Call GameDestroy
End Sub

Private Sub Label33_Click()
    Call UpdateVisInv
Picture30.Visible = True
End Sub

Private Sub Label34_Click()
Picture34.Visible = False
Picture35.Visible = False
Picture32.Visible = False
Picture8.Visible = False
Picture31.Visible = True
End Sub

Private Sub Label35_Click()
Picture15.Visible = True
End Sub

Private Sub Label36_Click()
'Timer5.Enabled = True
'Picture16.Visible = True
'Picture15.Visible = False
End Sub

Private Sub Label38_Click()
If StrenghtUse = 0 Then
If GetPlayerMP(MyIndex) > 5 Then
Call SetPlayerMP(MyIndex, GetPlayerMP(MyIndex) - 5)
Picture22.Visible = True
StrenghtUse = 1
Dim Packet As String
Packet = ""
Packet = "strength" & SEP_CHAR & "5" & SEP_CHAR & END_CHAR
Call SendData(Packet)
str.Enabled = True
End If
End If
End Sub

Private Sub Label4_Click()
Picture34.Visible = False
Picture35.Visible = False
Picture8.Visible = False
Picture31.Visible = False
Picture32.Visible = True
End Sub

Private Sub Label41_Click()
'Timer3.Enabled = True
'Picture19.Visible = True
'Picture15.Visible = False
End Sub

Private Sub Label42_Click()
'Label33.Caption = "+25"
'Timer5.Enabled = True
'Picture18.Visible = True
'Picture15.Visible = False
End Sub

Private Sub Label49_Click()
'Label33.Caption = "+25"
'Timer2.Enabled = True
'Picture17.Visible = True
'Picture15.Visible = False
End Sub

Private Sub Label51_Click()
Picture15.Visible = False
End Sub

Private Sub Label52_Click()
Call SendData("spells" & SEP_CHAR & END_CHAR)
Call UpdateVisInv
Picture30.Visible = False
Picture32.Visible = False
Picture35.Visible = False
Picture31.Visible = False
Picture8.Visible = False
Picture34.Visible = True
End Sub

Private Sub Label53_Click()
Picture34.Visible = False
Picture31.Visible = False
Picture32.Visible = False
Picture35.Visible = False
Picture8.Visible = True
End Sub

Private Sub Label54_Click()
Call GameDestroy
End Sub

Private Sub Label56_Click()
Picture34.Visible = False
Picture31.Visible = False
Picture32.Visible = False
Picture8.Visible = False
Picture35.Visible = True
End Sub

Private Sub Label6_Click()
    Call SendData("getstats" & SEP_CHAR & END_CHAR)
End Sub

Private Sub Label67_Click()
lstSpells.Visible = False
Picture40.Visible = False
picAmmount.Visible = False
Picture39.Visible = True
End Sub

Private Sub Label68_Click()
lstSpells.Visible = False
Picture39.Visible = False
Picture40.Visible = False
picAmmount.Visible = True
End Sub

Private Sub Label69_Click()
Picture39.Visible = False
picAmmount.Visible = False
lstSpells.Visible = True
Picture40.Visible = True
End Sub

Private Sub Label7_Click()
    Call UpdateVisInv
    Picture30.Visible = True
    Picture1.Visible = False
    frmMirage.picGuildAdmin.Visible = False
    'picPlayerSpells.Visible = False
    Picture15.Visible = False
    'picWhosOnline.Visible = False
End Sub

Private Sub Label70_Click()
lstSpells.Visible = False
Picture39.Visible = False
Picture40.Visible = False
picAmmount.Visible = False
End Sub

Private Sub Label74_Click()
picOptions.Visible = True
End Sub

Private Sub Label8_Click()
    Call SendData("spells" & SEP_CHAR & END_CHAR)
    Picture30.Visible = False
    frmMirage.picGuildAdmin.Visible = False
    picEquip.Visible = False
    Picture1.Visible = False
    Picture15.Visible = False
    'picWhosOnline.Visible = False
End Sub

Private Sub lblCloseOnline_Click()
Call SendOnlineList
'picWhosOnline.Visible = False
End Sub

Private Sub lblClosePicGuildAdmin_Click()
picGuildAdmin.Visible = False
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

Private Sub picInv_DblClick(index As Integer)
Dim d As Long
If Player(MyIndex).Inv(index + 1).Num = 0 Then Exit Sub
If picBanker.Visible = True Then
   If Player(MyIndex).Inv(index + 1).Value > 1 Then
       IndexClicked = index
       picAmmount.Visible = True
   Else
       Call BankItem(index + 1)
   End If
   Exit Sub
End If

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

Private Sub picInv_MouseDown(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Inventory = index + 1
    frmMirage.SelectedItem.Top = frmMirage.picInv(Inventory - 1).Top - 1
    frmMirage.SelectedItem.Left = frmMirage.picInv(Inventory - 1).Left - 1
    
    If Button = 1 Then
        Call UpdateVisInv
    ElseIf Button = 2 Then
        Call DropItems
    End If
End Sub

Private Sub picInv_MouseMove(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim d As Long
d = index

    If Player(MyIndex).Inv(d + 1).Num > 0 Then
        If Item(GetPlayerInvItemNum(MyIndex, d + 1)).Type = ITEM_TYPE_CURRENCY Then
            If Trim(Item(GetPlayerInvItemNum(MyIndex, d + 1)).desc) = "" Then
                itmDesc.Height = 17
                itmDesc.Top = 424
                 itmDesc.Left = 310
            Else
                itmDesc.Height = 233
                itmDesc.Top = 208
                 itmDesc.Left = 310
            End If
        Else
            If Trim(Item(GetPlayerInvItemNum(MyIndex, d + 1)).desc) = "" Then
                itmDesc.Height = 145
                itmDesc.Top = 296
                itmDesc.Left = 310
            Else
                itmDesc.Height = 233
                itmDesc.Top = 208
                 itmDesc.Left = 310
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
            Else
                descName.Caption = Trim(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Name)
            End If
        End If
        descStr.Caption = Item(GetPlayerInvItemNum(MyIndex, d + 1)).StrReq & " Strength"
        descDef.Caption = Item(GetPlayerInvItemNum(MyIndex, d + 1)).DefReq & " Defence"
        descSpeed.Caption = Item(GetPlayerInvItemNum(MyIndex, d + 1)).SpeedReq & " Speed"
        descHpMp.Caption = "HP: " & Item(GetPlayerInvItemNum(MyIndex, d + 1)).AddHP & " MP: " & Item(GetPlayerInvItemNum(MyIndex, d + 1)).AddMP & " SP: " & Item(GetPlayerInvItemNum(MyIndex, d + 1)).AddSP
        descSD.Caption = "Str: " & Item(GetPlayerInvItemNum(MyIndex, d + 1)).AddStr & " Def: " & Item(GetPlayerInvItemNum(MyIndex, d + 1)).AddDef
        descMS.Caption = "Magi: " & Item(GetPlayerInvItemNum(MyIndex, d + 1)).AddMagi & " Speed: " & Item(GetPlayerInvItemNum(MyIndex, d + 1)).AddSpeed
        desc.Caption = Trim(Item(GetPlayerInvItemNum(MyIndex, d + 1)).desc)
        
        itmDesc.Visible = True
    Else
        itmDesc.Visible = False
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Dim d As Long, i As Long
Dim ii As Long
Dim Value As Long
Call CheckInput(0, KeyCode, Shift)
    If KeyCode = vbKeyF1 Then
        If Player(MyIndex).Access > 0 Then
            frmadmin.Visible = False
            frmadmin.Visible = True
        Else
            frmUserPanel.Visible = True
        End If
    End If
    
        If KeyCode = vbKeyF2 Then
        For i = 1 To MAX_INV
        If GetPlayerInvItemNum(MyIndex, i) > 0 And GetPlayerInvItemNum(MyIndex, i) <= MAX_ITEMS Then
        If Item(GetPlayerInvItemNum(MyIndex, i)).Type = ITEM_TYPE_POTIONADDHP Then
        Call SendUseItem(i)
        Call AddText("You used a healing potion!", Yellow)
        Exit Sub
        End If
        End If
        Next i
        Call AddText("You don't have any potions!", Red)
        End If
        
                If KeyCode = vbKeyF3 Then
        For i = 1 To MAX_INV
        If GetPlayerInvItemNum(MyIndex, i) > 0 And GetPlayerInvItemNum(MyIndex, i) <= MAX_ITEMS Then
        If Item(GetPlayerInvItemNum(MyIndex, i)).Type = ITEM_TYPE_POTIONADDMP Then
        Call SendUseItem(i)
        Call AddText("You used a mana potion!", Yellow)
        Exit Sub
        End If
        End If
        Next i
        Call AddText("You don't have any potions!", Red)
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
frmMirage.lblclass.Caption = Trim(Class(GetPlayerClass(MyIndex)).Name)
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
        BattlePMsg(i).index = 1
        BattlePMsg(i).Time = i
        BattleMMsg(i).index = 1
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
Private Sub Command4s_Click()
Picture29.Visible = False
End Sub

Private Sub Command3s_Click()
Call WriteINI("STORY", "DontShowAgain", "1", (App.Path & "\config.ini"))
End Sub

Private Sub stBevriasWebsite_Click()
Shell ("explorer http://www.Bevrias.com"), vbNormalNoFocus
End Sub

Private Sub stEquipment_Click()
    picEquip.Visible = True
    'picInv3.Visible = False
    'picPlayerSpells.Visible = False
    'picWhosOnline.Visible = False
    Picture15.Visible = False
    Picture1.Visible = False
    frmMirage.picGuildAdmin.Visible = False
    Call UpdateVisInv
End Sub

Private Sub stChat_Click()
Picture34.Visible = False
Picture35.Visible = False
Picture8.Visible = False
Picture31.Visible = False
Picture32.Visible = True
End Sub

Private Sub stExit_Click()
If StrenghtUse = 1 Then
Picture22.Visible = False
Dim Packet As String
Packet = ""
Packet = "strength" & SEP_CHAR & "-5" & SEP_CHAR & END_CHAR
Call SendData(Packet)
str.Enabled = False
End If
Call GameDestroy
End Sub

Private Sub stFPS_Click()
            Call AddText("FPS: " & GameFPS, Pink)
            MyText = ""
End Sub

Private Sub stGuild_Click()
' Set Their Guild Name and Their Rank
frmMirage.lblGuild.Caption = GetPlayerGuild(MyIndex)
frmMirage.lblRank.Caption = GetPlayerGuildAccess(MyIndex)
Picture34.Visible = False
Picture35.Visible = False
Picture32.Visible = False
Picture8.Visible = False
Picture31.Visible = True
End Sub

Private Sub stInventory_Click()
Call UpdateVisInv
Picture30.Visible = False
Picture32.Visible = False
Picture35.Visible = False
Picture31.Visible = False
Picture8.Visible = False
Picture34.Visible = True
End Sub

Private Sub stLogout_Click()
frmSendGetData.Visible = False
   InGame = False
   frmMirage.Socket.Close
   frmMainMenu.Visible = True
End Sub

Private Sub stOptions_Click()
chkplayername.Value = Trim(ReadINI("CONFIG", "playername", App.Path & "\config.ini"))
chkplayerdamage.Value = Trim(ReadINI("CONFIG", "playerdamage", App.Path & "\config.ini"))
chkplayerbar.Value = Trim(ReadINI("CONFIG", "playerbar", App.Path & "\config.ini"))
chknpcname.Value = Trim(ReadINI("CONFIG", "npcname", App.Path & "\config.ini"))
chknpcdamage.Value = Trim(ReadINI("CONFIG", "npcdamage", App.Path & "\config.ini"))
chknpcbar.Value = Trim(ReadINI("CONFIG", "npcbar", App.Path & "\config.ini"))
chkmusic.Value = Trim(ReadINI("CONFIG", "music", App.Path & "\config.ini"))
chksound.Value = Trim(ReadINI("CONFIG", "sound", App.Path & "\config.ini"))
chkbubblebar.Value = Trim(ReadINI("CONFIG", "speechbubbles", App.Path & "\config.ini"))
chkAutoScroll.Value = Trim(ReadINI("CONFIG", "AutoScroll", App.Path & "\config.ini"))
picOptions.Visible = True
End Sub

Private Sub stPrivateMessage_Click()
txtMyTextBox.Text = "!<playername> <message>"
End Sub

Private Sub str_Timer()
Picture22.Visible = False
Dim Packet As String
Packet = ""
Packet = "strength" & SEP_CHAR & "-5" & SEP_CHAR & END_CHAR
Call SendData(Packet)
str.Enabled = False
StrenghtUse = 0
End Sub

Private Sub stRefresh_Click()
            Call SendData("refresh" & SEP_CHAR & END_CHAR)
            MyText = ""
End Sub

Private Sub stSpells_Click()
    Call SendData("spells" & SEP_CHAR & END_CHAR)
Picture31.Visible = False
Picture30.Visible = True
End Sub

Private Sub stStats_Click()
            Call SendData("getstats" & SEP_CHAR & END_CHAR)
            MyText = ""
End Sub

Private Sub stStatss_Click()
Picture34.Visible = False
Picture31.Visible = False
Picture32.Visible = False
Picture8.Visible = False
Picture35.Visible = True
End Sub

Private Sub stUserPanel_Click()
frmUserPanel.Visible = True
End Sub

Private Sub stWhos_Click()
Call SendOnlineList
Picture34.Visible = False
Picture31.Visible = False
Picture32.Visible = False
Picture35.Visible = False
Picture8.Visible = True
End Sub

Private Sub Timer0_Timer()
Picture18.Visible = False
Timer0.Enabled = False
End Sub

Private Sub Timer2_Timer()
Picture17.Visible = False
Timer2.Enabled = False
End Sub

Private Sub Timer3_Timer()
frmMirage.lblname.Caption = GetPlayerName(MyIndex)
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
Dim invNum As Long
Dim GoldAmount As String
On Error GoTo Done
If Inventory <= 0 Then Exit Sub

    invNum = Inventory
    If GetPlayerInvItemNum(MyIndex, invNum) > 0 And GetPlayerInvItemNum(MyIndex, invNum) <= MAX_ITEMS Then
        If Item(GetPlayerInvItemNum(MyIndex, invNum)).Type = ITEM_TYPE_CURRENCY Then
            GoldAmount = InputBox("How much " & Trim(Item(GetPlayerInvItemNum(MyIndex, invNum)).Name) & "(" & GetPlayerInvItemValue(MyIndex, invNum) & ") would you like to drop?", "Drop " & Trim(Item(GetPlayerInvItemNum(MyIndex, invNum)).Name), 0, frmMirage.Left, frmMirage.Top)
            If IsNumeric(GoldAmount) Then
                Call SendDropItem(invNum, GoldAmount)
            End If
        Else
            Call SendDropItem(invNum, 0)
        End If
    End If
   
    picInv(invNum - 1).Picture = LoadPicture()
    Call UpdateVisInv
    Exit Sub
Done:
    If Item(GetPlayerInvItemNum(MyIndex, invNum)).Type = ITEM_TYPE_CURRENCY Then
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
If StrenghtUse = 1 Then
Picture22.Visible = False
Dim Packet As String
Packet = ""
Packet = "strength" & SEP_CHAR & "-5" & SEP_CHAR & END_CHAR
Call SendData(Packet)
str.Enabled = False
End If
'If DefenceUse = 1 Then
'End If
'If AgilityUse = 1 Then
'End If
'If WisdomUse = 1 Then
'End If
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
Private Sub lstSpells_GotFocus()
picScreen.SetFocus
End Sub
Private Sub lblWithdraw_Click()
  If Player(MyIndex).bank(IndexClicked + 1).Value < txtAmmountBank.Text Then Exit Sub
Call InvItem(IndexClicked + 1, txtAmmountBank.Text)
picAmmountBank.Visible = False
End Sub

Private Sub Timer1_Timer()
          ' Visual Inventory
        Dim Q As Long
        Dim Qq As Long
               
            For Q = 0 To 48
                Qq = Player(MyIndex).bank(Q + 1).Num
                If picBank(Q).Picture <> LoadPicture() Then
                    picBank(Q).Picture = LoadPicture()
                Else
                    If Qq = 0 Then
                        picBank(Q).Picture = LoadPicture()
                    Else
                        Call BitBlt(picBank(Q).hDC, 0, 0, PIC_X, PIC_Y, frmMirage.picItems.hDC, (Item(Qq).Pic - Int(Item(Qq).Pic / 6) * 6) * PIC_X, Int(Item(Qq).Pic / 6) * PIC_Y, SRCCOPY)
                    End If
                End If
            Next Q
End Sub

Private Sub picBank_Click(index As Integer)
      If Player(MyIndex).bank(index + 1).Num = 0 Then Exit Sub
   If Player(MyIndex).bank(index + 1).Value > 1 Then
       IndexClicked = index
       picAmmountBank.Visible = True
       Exit Sub
   Else
       Call InvItem(index + 1)
   End If
End Sub

Private Sub picBank_MouseMove(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim dhh As Long
dhh = index

    If Player(MyIndex).bank(dhh + 1).Num > 0 Then
        If Item(GetPlayerBankItemNum(dhh + 1)).Type = ITEM_TYPE_CURRENCY Then
            If Trim(Item(GetPlayerBankItemNum(dhh + 1)).desc) = "" Then
                itmDesc.Height = 17
                itmDesc.Top = 224
            Else
                itmDesc.Height = 233
                itmDesc.Top = 8
            End If
        Else
            If Trim(Item(GetPlayerBankItemNum(dhh + 1)).desc) = "" Then
                itmDesc.Height = 145
                itmDesc.Top = 96
            Else
                itmDesc.Height = 233
                itmDesc.Top = 8
            End If
        End If
        If Item(GetPlayerBankItemNum(dhh + 1)).Type = ITEM_TYPE_CURRENCY Then
            descName.Caption = Trim(Item(GetPlayerBankItemNum(dhh + 1)).Name) & " (" & GetPlayerBankItemValue(dhh + 1) & ")"
            End If
        descStr.Caption = Item(GetPlayerBankItemNum(dhh + 1)).StrReq & " Strength"
        descDef.Caption = Item(GetPlayerBankItemNum(dhh + 1)).DefReq & " Defence"
        descSpeed.Caption = Item(GetPlayerBankItemNum(dhh + 1)).SpeedReq & " Speed"
        descHpMp.Caption = "HP: " & Item(GetPlayerBankItemNum(dhh + 1)).AddHP & " MP: " & Item(GetPlayerBankItemNum(dhh + 1)).AddMP & " SP: " & Item(GetPlayerBankItemNum(dhh + 1)).AddSP
        descSD.Caption = "Str: " & Item(GetPlayerBankItemNum(dhh + 1)).AddStr & " Def: " & Item(GetPlayerBankItemNum(dhh + 1)).AddDef
        descMS.Caption = "Magi: " & Item(GetPlayerBankItemNum(dhh + 1)).AddMagi & " Speed: " & Item(GetPlayerBankItemNum(dhh + 1)).AddSpeed
        desc.Caption = Trim(Item(GetPlayerBankItemNum(dhh + 1)).desc)
        
        itmDesc.Visible = True
    Else
       itmDesc.Visible = False
    End If
End Sub

Private Sub picBanker_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    SOffsetX = x
    SOffsetY = y
End Sub

Private Sub picBanker_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call MovePicture(frmMirage.picBanker, Button, Shift, x, y)
End Sub

Private Sub cmdOk_Click()
picBanker.Visible = False
End Sub

Private Sub lblDeposit_Click()
If txtAmmount < 1 Then
txtAmmount.Text = "1"
End If
If Player(MyIndex).Inv(IndexClicked + 1).Value < txtAmmount.Text Then Exit Sub
Call BankItem(IndexClicked + 1, txtAmmount.Text)
picAmmount.Visible = False
End Sub

