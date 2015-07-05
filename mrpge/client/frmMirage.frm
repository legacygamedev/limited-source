VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "Mswinsck.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmMirage 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "M:RPGe"
   ClientHeight    =   9015
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   12030
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
   Picture         =   "frmMirage.frx":0442
   ScaleHeight     =   601
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   802
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.PictureBox picQuestMsg 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5850
      Left            =   11880
      Picture         =   "frmMirage.frx":190FE
      ScaleHeight     =   390
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   283
      TabIndex        =   258
      Top             =   120
      Visible         =   0   'False
      Width           =   4245
      Begin VB.PictureBox cmdQuestOK 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   1395
         Picture         =   "frmMirage.frx":1FF26
         ScaleHeight     =   23
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   109
         TabIndex        =   260
         Top             =   5250
         Width           =   1635
      End
      Begin VB.Label lblQuestMsg 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   4710
         Left            =   195
         TabIndex        =   259
         Top             =   150
         Width           =   3855
      End
   End
   Begin VB.TextBox txtSend 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   105
      TabIndex        =   256
      Top             =   5955
      Width           =   7590
   End
   Begin VB.PictureBox picChannelAllImg 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   390
      Index           =   1
      Left            =   12375
      Picture         =   "frmMirage.frx":227D9
      ScaleHeight     =   360
      ScaleWidth      =   1365
      TabIndex        =   255
      Top             =   9780
      Width           =   1395
   End
   Begin VB.PictureBox picChannelPrivateImg 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   390
      Index           =   1
      Left            =   10845
      Picture         =   "frmMirage.frx":2501F
      ScaleHeight     =   360
      ScaleWidth      =   1365
      TabIndex        =   254
      Top             =   9765
      Width           =   1395
   End
   Begin VB.PictureBox picChannelGuildImg 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   390
      Index           =   1
      Left            =   9435
      Picture         =   "frmMirage.frx":27C3C
      ScaleHeight     =   360
      ScaleWidth      =   1365
      TabIndex        =   253
      Top             =   9630
      Width           =   1395
   End
   Begin VB.PictureBox picChannelGlobalImg 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   390
      Index           =   1
      Left            =   7770
      Picture         =   "frmMirage.frx":2A689
      ScaleHeight     =   360
      ScaleWidth      =   1365
      TabIndex        =   252
      Top             =   9780
      Width           =   1395
   End
   Begin VB.PictureBox picChannelGeneralImg 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   390
      Index           =   1
      Left            =   6225
      Picture         =   "frmMirage.frx":2D1D6
      ScaleHeight     =   360
      ScaleWidth      =   1365
      TabIndex        =   251
      Top             =   9780
      Width           =   1395
   End
   Begin VB.PictureBox picChannelAllImg 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   390
      Index           =   0
      Left            =   12360
      Picture         =   "frmMirage.frx":2FD79
      ScaleHeight     =   360
      ScaleWidth      =   1365
      TabIndex        =   250
      Top             =   9255
      Width           =   1395
   End
   Begin VB.PictureBox picChannelPrivateImg 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   390
      Index           =   0
      Left            =   10830
      Picture         =   "frmMirage.frx":325BE
      ScaleHeight     =   360
      ScaleWidth      =   1365
      TabIndex        =   249
      Top             =   9255
      Width           =   1395
   End
   Begin VB.PictureBox picChannelGuildImg 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   390
      Index           =   0
      Left            =   9330
      Picture         =   "frmMirage.frx":35190
      ScaleHeight     =   360
      ScaleWidth      =   1365
      TabIndex        =   248
      Top             =   9255
      Width           =   1395
   End
   Begin VB.PictureBox picChannelGlobalImg 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   390
      Index           =   0
      Left            =   7755
      Picture         =   "frmMirage.frx":37BB6
      ScaleHeight     =   360
      ScaleWidth      =   1365
      TabIndex        =   247
      Top             =   9255
      Width           =   1395
   End
   Begin VB.PictureBox picChannelGeneralImg 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   390
      Index           =   0
      Left            =   6210
      Picture         =   "frmMirage.frx":3A6FA
      ScaleHeight     =   360
      ScaleWidth      =   1365
      TabIndex        =   246
      Top             =   9255
      Width           =   1395
   End
   Begin RichTextLib.RichTextBox txtChat 
      Height          =   1950
      Left            =   12495
      TabIndex        =   1
      Top             =   15
      Visible         =   0   'False
      Width           =   7620
      _ExtentX        =   13441
      _ExtentY        =   3440
      _Version        =   393217
      BackColor       =   16777215
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmMirage.frx":3D269
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picChannelGeneral 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   11880
      Picture         =   "frmMirage.frx":3D2E0
      ScaleHeight     =   360
      ScaleWidth      =   1365
      TabIndex        =   240
      Top             =   11310
      Width           =   1395
   End
   Begin VB.PictureBox picChannelGlobal 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   12195
      Picture         =   "frmMirage.frx":3FE4F
      ScaleHeight     =   360
      ScaleWidth      =   1365
      TabIndex        =   239
      Top             =   11565
      Width           =   1395
   End
   Begin VB.PictureBox picChannelGuild 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   11775
      Picture         =   "frmMirage.frx":4299C
      ScaleHeight     =   360
      ScaleWidth      =   1365
      TabIndex        =   238
      Top             =   10815
      Width           =   1395
   End
   Begin VB.PictureBox picChannelPrivate 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   11535
      Picture         =   "frmMirage.frx":453E9
      ScaleHeight     =   360
      ScaleWidth      =   1365
      TabIndex        =   237
      Top             =   10935
      Width           =   1395
   End
   Begin VB.PictureBox picChannelAll 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   11625
      Picture         =   "frmMirage.frx":48006
      ScaleHeight     =   360
      ScaleWidth      =   1365
      TabIndex        =   236
      Top             =   11370
      Width           =   1395
   End
   Begin VB.PictureBox picPrayerPane 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   12735
      Picture         =   "frmMirage.frx":4A84C
      ScaleHeight     =   435
      ScaleWidth      =   1695
      TabIndex        =   235
      Top             =   11910
      Width           =   1695
   End
   Begin VB.PictureBox picWhite 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5760
      Left            =   16845
      ScaleHeight     =   5760
      ScaleWidth      =   7665
      TabIndex        =   230
      Top             =   2205
      Width           =   7665
   End
   Begin VB.Timer timWeather 
      Interval        =   1000
      Left            =   12000
      Top             =   465
   End
   Begin VB.Timer timMusic 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   12000
      Top             =   0
   End
   Begin VB.PictureBox picBankWindow 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      ForeColor       =   &H80000008&
      Height          =   4710
      Left            =   165
      ScaleHeight     =   312
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   460
      TabIndex        =   155
      Top             =   9090
      Visible         =   0   'False
      Width           =   6930
      Begin VB.PictureBox picBank 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         FillColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   3450
         Left            =   3765
         ScaleHeight     =   3450
         ScaleWidth      =   2955
         TabIndex        =   193
         Top             =   375
         Width           =   2955
         Begin VB.PictureBox picbankItem 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   540
            Index           =   29
            Left            =   2340
            ScaleHeight     =   540
            ScaleWidth      =   540
            TabIndex        =   223
            Top             =   2835
            Width           =   540
         End
         Begin VB.PictureBox picbankItem 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   540
            Index           =   28
            Left            =   1770
            ScaleHeight     =   540
            ScaleWidth      =   540
            TabIndex        =   222
            Top             =   2835
            Width           =   540
         End
         Begin VB.PictureBox picbankItem 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   540
            Index           =   27
            Left            =   1200
            ScaleHeight     =   540
            ScaleWidth      =   540
            TabIndex        =   221
            Top             =   2835
            Width           =   540
         End
         Begin VB.PictureBox picbankItem 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   540
            Index           =   26
            Left            =   630
            ScaleHeight     =   540
            ScaleWidth      =   540
            TabIndex        =   220
            Top             =   2835
            Width           =   540
         End
         Begin VB.PictureBox picbankItem 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   540
            Index           =   25
            Left            =   60
            ScaleHeight     =   540
            ScaleWidth      =   540
            TabIndex        =   219
            Top             =   2835
            Width           =   540
         End
         Begin VB.PictureBox picbankItem 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   540
            Index           =   24
            Left            =   2340
            ScaleHeight     =   540
            ScaleWidth      =   540
            TabIndex        =   218
            Top             =   2280
            Width           =   540
         End
         Begin VB.PictureBox picbankItem 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   540
            Index           =   23
            Left            =   1770
            ScaleHeight     =   540
            ScaleWidth      =   540
            TabIndex        =   217
            Top             =   2280
            Width           =   540
         End
         Begin VB.PictureBox picbankItem 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   540
            Index           =   22
            Left            =   1200
            ScaleHeight     =   540
            ScaleWidth      =   540
            TabIndex        =   216
            Top             =   2280
            Width           =   540
         End
         Begin VB.PictureBox picbankItem 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   540
            Index           =   21
            Left            =   630
            ScaleHeight     =   540
            ScaleWidth      =   540
            TabIndex        =   215
            Top             =   2280
            Width           =   540
         End
         Begin VB.PictureBox picbankItem 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   540
            Index           =   20
            Left            =   60
            ScaleHeight     =   540
            ScaleWidth      =   540
            TabIndex        =   214
            Top             =   2280
            Width           =   540
         End
         Begin VB.PictureBox picbankItem 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   540
            Index           =   19
            Left            =   2340
            ScaleHeight     =   540
            ScaleWidth      =   540
            TabIndex        =   213
            Top             =   1725
            Width           =   540
         End
         Begin VB.PictureBox picbankItem 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   540
            Index           =   18
            Left            =   1770
            ScaleHeight     =   540
            ScaleWidth      =   540
            TabIndex        =   212
            Top             =   1725
            Width           =   540
         End
         Begin VB.PictureBox picbankItem 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   540
            Index           =   17
            Left            =   1200
            ScaleHeight     =   540
            ScaleWidth      =   540
            TabIndex        =   211
            Top             =   1725
            Width           =   540
         End
         Begin VB.PictureBox picbankItem 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   540
            Index           =   16
            Left            =   630
            ScaleHeight     =   540
            ScaleWidth      =   540
            TabIndex        =   210
            Top             =   1725
            Width           =   540
         End
         Begin VB.PictureBox picbankItem 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   540
            Index           =   15
            Left            =   60
            ScaleHeight     =   540
            ScaleWidth      =   540
            TabIndex        =   209
            Top             =   1725
            Width           =   540
         End
         Begin VB.PictureBox picbankItem 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   540
            Index           =   14
            Left            =   2340
            ScaleHeight     =   540
            ScaleWidth      =   540
            TabIndex        =   208
            Top             =   1170
            Width           =   540
         End
         Begin VB.PictureBox picbankItem 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   540
            Index           =   13
            Left            =   1770
            ScaleHeight     =   540
            ScaleWidth      =   540
            TabIndex        =   207
            Top             =   1170
            Width           =   540
         End
         Begin VB.PictureBox picbankItem 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   540
            Index           =   12
            Left            =   1200
            ScaleHeight     =   540
            ScaleWidth      =   540
            TabIndex        =   206
            Top             =   1170
            Width           =   540
         End
         Begin VB.PictureBox picbankItem 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   540
            Index           =   11
            Left            =   630
            ScaleHeight     =   540
            ScaleWidth      =   540
            TabIndex        =   205
            Top             =   1170
            Width           =   540
         End
         Begin VB.PictureBox picbankItem 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   540
            Index           =   10
            Left            =   60
            ScaleHeight     =   540
            ScaleWidth      =   540
            TabIndex        =   204
            Top             =   1170
            Width           =   540
         End
         Begin VB.PictureBox picbankItem 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   540
            Index           =   9
            Left            =   2340
            ScaleHeight     =   540
            ScaleWidth      =   540
            TabIndex        =   203
            Top             =   615
            Width           =   540
         End
         Begin VB.PictureBox picbankItem 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   540
            Index           =   8
            Left            =   1770
            ScaleHeight     =   540
            ScaleWidth      =   540
            TabIndex        =   202
            Top             =   615
            Width           =   540
         End
         Begin VB.PictureBox picbankItem 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   540
            Index           =   7
            Left            =   1200
            ScaleHeight     =   540
            ScaleWidth      =   540
            TabIndex        =   201
            Top             =   615
            Width           =   540
         End
         Begin VB.PictureBox picbankItem 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   540
            Index           =   6
            Left            =   630
            ScaleHeight     =   540
            ScaleWidth      =   540
            TabIndex        =   200
            Top             =   615
            Width           =   540
         End
         Begin VB.PictureBox picbankItem 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   540
            Index           =   5
            Left            =   60
            ScaleHeight     =   540
            ScaleWidth      =   540
            TabIndex        =   199
            Top             =   615
            Width           =   540
         End
         Begin VB.PictureBox picbankItem 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   540
            Index           =   4
            Left            =   2340
            ScaleHeight     =   540
            ScaleWidth      =   540
            TabIndex        =   198
            Top             =   60
            Width           =   540
         End
         Begin VB.PictureBox picbankItem 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   540
            Index           =   3
            Left            =   1770
            ScaleHeight     =   540
            ScaleWidth      =   540
            TabIndex        =   197
            Top             =   60
            Width           =   540
         End
         Begin VB.PictureBox picbankItem 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   540
            Index           =   2
            Left            =   1200
            ScaleHeight     =   540
            ScaleWidth      =   540
            TabIndex        =   196
            Top             =   60
            Width           =   540
         End
         Begin VB.PictureBox picbankItem 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   540
            Index           =   1
            Left            =   630
            ScaleHeight     =   540
            ScaleWidth      =   540
            TabIndex        =   195
            Top             =   60
            Width           =   540
         End
         Begin VB.PictureBox picbankItem 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   540
            Index           =   0
            Left            =   60
            ScaleHeight     =   540
            ScaleWidth      =   540
            TabIndex        =   194
            Top             =   60
            Width           =   540
         End
      End
      Begin VB.PictureBox picLocalItems 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         FillColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   3450
         Left            =   150
         ScaleHeight     =   3450
         ScaleWidth      =   2955
         TabIndex        =   162
         Top             =   375
         Width           =   2955
         Begin VB.PictureBox picItemLocal 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   540
            Index           =   29
            Left            =   2340
            ScaleHeight     =   540
            ScaleWidth      =   540
            TabIndex        =   192
            Top             =   2850
            Width           =   540
         End
         Begin VB.PictureBox picItemLocal 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   540
            Index           =   28
            Left            =   1770
            ScaleHeight     =   540
            ScaleWidth      =   540
            TabIndex        =   191
            Top             =   2850
            Width           =   540
         End
         Begin VB.PictureBox picItemLocal 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   540
            Index           =   27
            Left            =   1200
            ScaleHeight     =   540
            ScaleWidth      =   540
            TabIndex        =   190
            Top             =   2850
            Width           =   540
         End
         Begin VB.PictureBox picItemLocal 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   540
            Index           =   26
            Left            =   630
            ScaleHeight     =   540
            ScaleWidth      =   540
            TabIndex        =   189
            Top             =   2850
            Width           =   540
         End
         Begin VB.PictureBox picItemLocal 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   540
            Index           =   25
            Left            =   60
            ScaleHeight     =   540
            ScaleWidth      =   540
            TabIndex        =   188
            Top             =   2850
            Width           =   540
         End
         Begin VB.PictureBox picItemLocal 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   540
            Index           =   24
            Left            =   2340
            ScaleHeight     =   540
            ScaleWidth      =   540
            TabIndex        =   187
            Top             =   2295
            Width           =   540
         End
         Begin VB.PictureBox picItemLocal 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   540
            Index           =   23
            Left            =   1770
            ScaleHeight     =   540
            ScaleWidth      =   540
            TabIndex        =   186
            Top             =   2295
            Width           =   540
         End
         Begin VB.PictureBox picItemLocal 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   540
            Index           =   22
            Left            =   1200
            ScaleHeight     =   540
            ScaleWidth      =   540
            TabIndex        =   185
            Top             =   2295
            Width           =   540
         End
         Begin VB.PictureBox picItemLocal 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   540
            Index           =   21
            Left            =   630
            ScaleHeight     =   540
            ScaleWidth      =   540
            TabIndex        =   184
            Top             =   2295
            Width           =   540
         End
         Begin VB.PictureBox picItemLocal 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   540
            Index           =   20
            Left            =   60
            ScaleHeight     =   540
            ScaleWidth      =   540
            TabIndex        =   183
            Top             =   2295
            Width           =   540
         End
         Begin VB.PictureBox picItemLocal 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   540
            Index           =   19
            Left            =   2340
            ScaleHeight     =   540
            ScaleWidth      =   540
            TabIndex        =   182
            Top             =   1740
            Width           =   540
         End
         Begin VB.PictureBox picItemLocal 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   540
            Index           =   18
            Left            =   1770
            ScaleHeight     =   540
            ScaleWidth      =   540
            TabIndex        =   181
            Top             =   1740
            Width           =   540
         End
         Begin VB.PictureBox picItemLocal 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   540
            Index           =   17
            Left            =   1200
            ScaleHeight     =   540
            ScaleWidth      =   540
            TabIndex        =   180
            Top             =   1740
            Width           =   540
         End
         Begin VB.PictureBox picItemLocal 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   540
            Index           =   16
            Left            =   630
            ScaleHeight     =   540
            ScaleWidth      =   540
            TabIndex        =   179
            Top             =   1740
            Width           =   540
         End
         Begin VB.PictureBox picItemLocal 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   540
            Index           =   15
            Left            =   60
            ScaleHeight     =   540
            ScaleWidth      =   540
            TabIndex        =   178
            Top             =   1740
            Width           =   540
         End
         Begin VB.PictureBox picItemLocal 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   540
            Index           =   14
            Left            =   2340
            ScaleHeight     =   540
            ScaleWidth      =   540
            TabIndex        =   177
            Top             =   1185
            Width           =   540
         End
         Begin VB.PictureBox picItemLocal 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   540
            Index           =   13
            Left            =   1770
            ScaleHeight     =   540
            ScaleWidth      =   540
            TabIndex        =   176
            Top             =   1185
            Width           =   540
         End
         Begin VB.PictureBox picItemLocal 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   540
            Index           =   12
            Left            =   1200
            ScaleHeight     =   540
            ScaleWidth      =   540
            TabIndex        =   175
            Top             =   1185
            Width           =   540
         End
         Begin VB.PictureBox picItemLocal 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   540
            Index           =   11
            Left            =   630
            ScaleHeight     =   540
            ScaleWidth      =   540
            TabIndex        =   174
            Top             =   1185
            Width           =   540
         End
         Begin VB.PictureBox picItemLocal 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   540
            Index           =   10
            Left            =   60
            ScaleHeight     =   540
            ScaleWidth      =   540
            TabIndex        =   173
            Top             =   1185
            Width           =   540
         End
         Begin VB.PictureBox picItemLocal 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   540
            Index           =   9
            Left            =   2340
            ScaleHeight     =   540
            ScaleWidth      =   540
            TabIndex        =   172
            Top             =   630
            Width           =   540
         End
         Begin VB.PictureBox picItemLocal 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   540
            Index           =   8
            Left            =   1770
            ScaleHeight     =   540
            ScaleWidth      =   540
            TabIndex        =   171
            Top             =   630
            Width           =   540
         End
         Begin VB.PictureBox picItemLocal 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   540
            Index           =   7
            Left            =   1200
            ScaleHeight     =   540
            ScaleWidth      =   540
            TabIndex        =   170
            Top             =   630
            Width           =   540
         End
         Begin VB.PictureBox picItemLocal 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   540
            Index           =   6
            Left            =   630
            ScaleHeight     =   540
            ScaleWidth      =   540
            TabIndex        =   169
            Top             =   630
            Width           =   540
         End
         Begin VB.PictureBox picItemLocal 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   540
            Index           =   5
            Left            =   60
            ScaleHeight     =   540
            ScaleWidth      =   540
            TabIndex        =   168
            Top             =   630
            Width           =   540
         End
         Begin VB.PictureBox picItemLocal 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   540
            Index           =   4
            Left            =   2340
            ScaleHeight     =   540
            ScaleWidth      =   540
            TabIndex        =   167
            Top             =   75
            Width           =   540
         End
         Begin VB.PictureBox picItemLocal 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   540
            Index           =   3
            Left            =   1770
            ScaleHeight     =   540
            ScaleWidth      =   540
            TabIndex        =   166
            Top             =   75
            Width           =   540
         End
         Begin VB.PictureBox picItemLocal 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   540
            Index           =   2
            Left            =   1200
            ScaleHeight     =   540
            ScaleWidth      =   540
            TabIndex        =   165
            Top             =   75
            Width           =   540
         End
         Begin VB.PictureBox picItemLocal 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   540
            Index           =   1
            Left            =   630
            ScaleHeight     =   540
            ScaleWidth      =   540
            TabIndex        =   164
            Top             =   75
            Width           =   540
         End
         Begin VB.PictureBox picItemLocal 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   540
            Index           =   0
            Left            =   60
            ScaleHeight     =   540
            ScaleWidth      =   540
            TabIndex        =   163
            Top             =   75
            Width           =   540
         End
      End
      Begin VB.TextBox txtDeposit 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   360
         Left            =   135
         TabIndex        =   161
         Top             =   4140
         Width           =   2310
      End
      Begin VB.TextBox txtWithdraw 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   360
         Left            =   4410
         TabIndex        =   160
         Top             =   4140
         Width           =   2310
      End
      Begin VB.CommandButton cmdDeposit 
         Caption         =   "D"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2565
         TabIndex        =   159
         Top             =   4125
         Width           =   540
      End
      Begin VB.CommandButton cmdWithdraw 
         Caption         =   "W"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3780
         TabIndex        =   158
         Top             =   4125
         Width           =   540
      End
      Begin VB.CommandButton cmdDepItem 
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3210
         TabIndex        =   157
         Top             =   1305
         Width           =   465
      End
      Begin VB.CommandButton cmdWithItem 
         Caption         =   "<<"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3195
         TabIndex        =   156
         Top             =   1740
         Width           =   465
      End
      Begin VB.Label lblCloseBank 
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6600
         TabIndex        =   228
         Top             =   0
         Width           =   375
      End
      Begin VB.Label lblBankTitle1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Inventory"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   135
         TabIndex        =   227
         Top             =   90
         Width           =   2955
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Bank"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   3765
         TabIndex        =   226
         Top             =   105
         Width           =   2955
      End
      Begin VB.Label lblInvGold 
         BackStyle       =   0  'Transparent
         Caption         =   "1,000,547"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   150
         TabIndex        =   225
         Top             =   3855
         Width           =   2880
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   405
         Left            =   120
         Top             =   4125
         Width           =   2355
      End
      Begin VB.Shape Shape2 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   405
         Left            =   4395
         Top             =   4125
         Width           =   2355
      End
      Begin VB.Label lblBankGold 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "1,000,547"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   3810
         TabIndex        =   224
         Top             =   3840
         Width           =   2880
      End
   End
   Begin VB.PictureBox picBlack 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5760
      Left            =   16155
      ScaleHeight     =   5760
      ScaleWidth      =   7665
      TabIndex        =   154
      Top             =   600
      Width           =   7665
   End
   Begin VB.PictureBox picEquiped 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2715
      Left            =   7305
      Picture         =   "frmMirage.frx":5105D
      ScaleHeight     =   181
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   280
      TabIndex        =   133
      Top             =   10335
      Visible         =   0   'False
      Width           =   4200
      Begin VB.PictureBox picEquip 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   660
         Index           =   3
         Left            =   2925
         ScaleHeight     =   42
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   42
         TabIndex        =   137
         Top             =   855
         Width           =   660
      End
      Begin VB.PictureBox picEquip 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   660
         Index           =   2
         Left            =   2145
         ScaleHeight     =   42
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   42
         TabIndex        =   136
         Top             =   855
         Width           =   660
      End
      Begin VB.PictureBox picEquip 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   660
         Index           =   1
         Left            =   1380
         ScaleHeight     =   42
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   42
         TabIndex        =   135
         Top             =   855
         Width           =   660
      End
      Begin VB.PictureBox picEquip 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   660
         Index           =   0
         Left            =   615
         ScaleHeight     =   42
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   42
         TabIndex        =   134
         Top             =   855
         Width           =   660
      End
      Begin VB.Label lblEquipClose 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3360
         TabIndex        =   138
         Top             =   4950
         Width           =   135
      End
   End
   Begin VB.PictureBox picMapEditor 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   6810
      Left            =   12015
      ScaleHeight     =   452
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   496
      TabIndex        =   2
      Top             =   450
      Visible         =   0   'False
      Width           =   7470
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear layer/atributes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   180
         TabIndex        =   234
         Top             =   5265
         Width           =   2160
      End
      Begin VB.ComboBox cmbAtributes 
         Height          =   360
         ItemData        =   "frmMirage.frx":5792F
         Left            =   180
         List            =   "frmMirage.frx":57957
         Style           =   2  'Dropdown List
         TabIndex        =   233
         Top             =   4455
         Width           =   2160
      End
      Begin VB.ComboBox cmbLayers 
         Height          =   360
         ItemData        =   "frmMirage.frx":579C2
         Left            =   180
         List            =   "frmMirage.frx":579D2
         Style           =   2  'Dropdown List
         TabIndex        =   232
         Top             =   4095
         Width           =   2160
      End
      Begin VB.ComboBox cmbTilePack 
         Height          =   360
         Left            =   180
         Style           =   2  'Dropdown List
         TabIndex        =   231
         Top             =   3720
         Width           =   2160
      End
      Begin VB.Frame frmTileSelect 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Tile Packs"
         ForeColor       =   &H00000000&
         Height          =   2775
         Left            =   7230
         TabIndex        =   48
         Top             =   3735
         Width           =   2295
         Begin VB.OptionButton opt1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Pack20"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   19
            Left            =   1200
            TabIndex        =   68
            Top             =   2400
            Width           =   975
         End
         Begin VB.OptionButton opt1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Pack19"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   18
            Left            =   1200
            TabIndex        =   67
            Top             =   2160
            Width           =   975
         End
         Begin VB.OptionButton opt1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Pack18"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   17
            Left            =   1200
            TabIndex        =   66
            Top             =   1920
            Width           =   975
         End
         Begin VB.OptionButton opt1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Pack17"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   16
            Left            =   1200
            TabIndex        =   65
            Top             =   1680
            Width           =   975
         End
         Begin VB.OptionButton opt1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Pack16"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   15
            Left            =   1200
            TabIndex        =   64
            Top             =   1440
            Width           =   975
         End
         Begin VB.OptionButton opt1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Pack15"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   14
            Left            =   1200
            TabIndex        =   63
            Top             =   1200
            Width           =   975
         End
         Begin VB.OptionButton opt1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Pack14"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   13
            Left            =   1200
            TabIndex        =   62
            Top             =   960
            Width           =   975
         End
         Begin VB.OptionButton opt1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Pack11"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   12
            Left            =   1200
            TabIndex        =   61
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton opt1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Pack12"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   11
            Left            =   1200
            TabIndex        =   60
            Top             =   480
            Width           =   975
         End
         Begin VB.OptionButton opt1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Pack13"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   10
            Left            =   1200
            TabIndex        =   59
            Top             =   720
            Width           =   975
         End
         Begin VB.OptionButton opt1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Pack10"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   9
            Left            =   120
            TabIndex        =   58
            Top             =   2400
            Width           =   1695
         End
         Begin VB.OptionButton opt1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Pack9"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   57
            Top             =   2160
            Width           =   1695
         End
         Begin VB.OptionButton opt1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Pack8"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   56
            Top             =   1920
            Width           =   1695
         End
         Begin VB.OptionButton opt1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Pack7"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   55
            Top             =   1680
            Width           =   1695
         End
         Begin VB.OptionButton opt1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Pack6"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   54
            Top             =   1440
            Width           =   1695
         End
         Begin VB.OptionButton opt1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Pack5"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   53
            Top             =   1200
            Width           =   1695
         End
         Begin VB.OptionButton opt1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Pack4"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   52
            Top             =   960
            Width           =   1695
         End
         Begin VB.OptionButton opt1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Pack3"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   51
            Top             =   720
            Width           =   1695
         End
         Begin VB.OptionButton opt1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Pack2"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   50
            Top             =   480
            Width           =   1695
         End
         Begin VB.OptionButton opt1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Pack1"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   49
            Top             =   240
            Value           =   -1  'True
            Width           =   1695
         End
      End
      Begin VB.Frame fraLayers 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Layers"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   2055
         Left            =   7230
         TabIndex        =   19
         Top             =   3705
         Width           =   1575
         Begin VB.OptionButton optGround 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Ground"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   135
            TabIndex        =   24
            Top             =   360
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton optMask 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Mask"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   105
            TabIndex        =   23
            Top             =   600
            Width           =   1215
         End
         Begin VB.OptionButton optAnim 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Animation"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   840
            Width           =   1215
         End
         Begin VB.OptionButton optFringe 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Fringe"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   1080
            Width           =   1215
         End
         Begin VB.CommandButton cmdClear4 
            Caption         =   "Clear"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   120
            TabIndex        =   20
            Top             =   1440
            Width           =   1215
         End
      End
      Begin VB.Frame fraAttribs 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Attributes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   3255
         Left            =   7140
         TabIndex        =   12
         Top             =   3615
         Visible         =   0   'False
         Width           =   1575
         Begin VB.OptionButton optNpcSpawn 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Spawn NPC"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   229
            Top             =   2640
            Width           =   1335
         End
         Begin VB.OptionButton optWarpDoor 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Warp (Level)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   47
            Top             =   720
            Width           =   1215
         End
         Begin VB.OptionButton optSign 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Sign"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   46
            Top             =   2400
            Width           =   1095
         End
         Begin VB.OptionButton optHeal 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Heal"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   45
            Top             =   2160
            Width           =   1215
         End
         Begin VB.OptionButton optDamage 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Damage"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   120
            TabIndex        =   44
            Top             =   1920
            Width           =   1095
         End
         Begin VB.OptionButton optKeyOpen 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Key Open"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   120
            TabIndex        =   34
            Top             =   1680
            Width           =   1215
         End
         Begin VB.OptionButton optBlocked 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Blocked"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   240
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton optWarp 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Warp"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   480
            Width           =   1215
         End
         Begin VB.CommandButton cmdClear2 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Clear"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   120
            TabIndex        =   16
            Top             =   2880
            Width           =   1215
         End
         Begin VB.OptionButton optItem 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Item"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   270
            Left            =   120
            TabIndex        =   15
            Top             =   960
            Width           =   1215
         End
         Begin VB.OptionButton optNpcAvoid 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Npc Avoid"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   270
            Left            =   120
            TabIndex        =   14
            Top             =   1200
            Width           =   1215
         End
         Begin VB.OptionButton optKey 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Key"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   270
            Left            =   120
            TabIndex        =   13
            Top             =   1440
            Width           =   1215
         End
      End
      Begin VB.OptionButton optLayers 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Layers"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2895
         TabIndex        =   11
         Top             =   3930
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton optAttribs 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Attributes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2895
         TabIndex        =   10
         Top             =   4185
         Width           =   1575
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2370
         TabIndex        =   9
         Top             =   4845
         Width           =   1695
      End
      Begin VB.CommandButton cmdSend 
         Caption         =   "Send"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2370
         TabIndex        =   8
         Top             =   4455
         Width           =   1695
      End
      Begin VB.CommandButton cmdProperties 
         Caption         =   "Properties"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   180
         TabIndex        =   7
         Top             =   4845
         Width           =   2160
      End
      Begin VB.VScrollBar scrlPicture 
         Height          =   3375
         Left            =   6750
         Max             =   10
         TabIndex        =   6
         Top             =   120
         Width           =   255
      End
      Begin VB.PictureBox picBack 
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
         Height          =   3360
         Left            =   120
         ScaleHeight     =   224
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   224
         TabIndex        =   4
         Top             =   120
         Width           =   3360
         Begin VB.PictureBox picBackSelect 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
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
            Height          =   960
            Left            =   0
            ScaleHeight     =   64
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   64
            TabIndex        =   5
            Top             =   0
            Width           =   960
         End
      End
      Begin VB.PictureBox picSelect 
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
         Height          =   480
         Left            =   2370
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   3
         Top             =   3960
         Width           =   480
      End
      Begin VB.Label lblMoveTilesRight 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "->"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   480
         TabIndex        =   126
         Top             =   3480
         Width           =   375
      End
      Begin VB.Label lblMoveTilesLeft 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "<-"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   125
         Top             =   3480
         Width           =   375
      End
   End
   Begin RichTextLib.RichTextBox txtGlobalChat 
      Height          =   1950
      Left            =   570
      TabIndex        =   132
      Top             =   14250
      Width           =   4020
      _ExtentX        =   7091
      _ExtentY        =   3440
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      Appearance      =   0
      TextRTF         =   $"frmMirage.frx":579F7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picSpellPane 
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2715
      Left            =   13065
      Picture         =   "frmMirage.frx":57A6E
      ScaleHeight     =   2715
      ScaleWidth      =   4200
      TabIndex        =   127
      Top             =   9960
      Visible         =   0   'False
      Width           =   4200
      Begin VB.ListBox lstSpells 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   2130
         ItemData        =   "frmMirage.frx":5E18C
         Left            =   75
         List            =   "frmMirage.frx":5E18E
         TabIndex        =   128
         Top             =   525
         Width           =   4095
      End
      Begin VB.Label lblCastP 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Cast"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2910
         TabIndex        =   141
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label picSpellClose 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5865
         TabIndex        =   131
         Top             =   1215
         Width           =   135
      End
      Begin VB.Label lblcast 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Cast"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2910
         TabIndex        =   130
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblCastType 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Cast"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4755
         TabIndex        =   129
         Top             =   585
         Width           =   1935
      End
   End
   Begin VB.PictureBox picScreen 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Left            =   0
      ScaleHeight     =   382
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   509
      TabIndex        =   0
      Top             =   45
      Width           =   7665
   End
   Begin VB.PictureBox picInvPane 
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2715
      Left            =   13410
      Picture         =   "frmMirage.frx":5E190
      ScaleHeight     =   2715
      ScaleWidth      =   4200
      TabIndex        =   71
      Top             =   10080
      Visible         =   0   'False
      Width           =   4200
      Begin VB.PictureBox picInvItem 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H00000000&
         Height          =   540
         Index           =   0
         Left            =   135
         ScaleHeight     =   34
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   121
         Top             =   510
         Width           =   510
      End
      Begin VB.PictureBox picInvItem 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   540
         Index           =   1
         Left            =   630
         ScaleHeight     =   34
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   120
         Top             =   510
         Width           =   510
      End
      Begin VB.PictureBox picInvItem 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   540
         Index           =   2
         Left            =   1125
         ScaleHeight     =   34
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   119
         Top             =   510
         Width           =   510
      End
      Begin VB.PictureBox picInvItem 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   540
         Index           =   3
         Left            =   1620
         ScaleHeight     =   34
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   118
         Top             =   510
         Width           =   510
      End
      Begin VB.PictureBox picInvItem 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   540
         Index           =   4
         Left            =   2115
         ScaleHeight     =   34
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   117
         Top             =   510
         Width           =   510
      End
      Begin VB.PictureBox picInvItem 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   540
         Index           =   5
         Left            =   2610
         ScaleHeight     =   34
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   116
         Top             =   510
         Width           =   510
      End
      Begin VB.PictureBox picInvItem 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   540
         Index           =   6
         Left            =   3105
         ScaleHeight     =   34
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   115
         Top             =   510
         Width           =   510
      End
      Begin VB.PictureBox picInvItem 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   540
         Index           =   7
         Left            =   3600
         ScaleHeight     =   34
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   114
         Top             =   510
         Width           =   510
      End
      Begin VB.PictureBox picInvItem 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   540
         Index           =   8
         Left            =   135
         ScaleHeight     =   34
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   113
         Top             =   1050
         Width           =   510
      End
      Begin VB.PictureBox picInvItem 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   540
         Index           =   9
         Left            =   630
         ScaleHeight     =   34
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   112
         Top             =   1050
         Width           =   510
      End
      Begin VB.PictureBox picInvItem 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   540
         Index           =   10
         Left            =   1125
         ScaleHeight     =   34
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   111
         Top             =   1050
         Width           =   510
      End
      Begin VB.PictureBox picInvItem 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   540
         Index           =   11
         Left            =   1620
         ScaleHeight     =   34
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   110
         Top             =   1050
         Width           =   510
      End
      Begin VB.PictureBox picInvItem 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   540
         Index           =   12
         Left            =   2115
         ScaleHeight     =   34
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   109
         Top             =   1050
         Width           =   510
      End
      Begin VB.PictureBox picInvItem 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   540
         Index           =   13
         Left            =   2610
         ScaleHeight     =   34
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   108
         Top             =   1050
         Width           =   510
      End
      Begin VB.PictureBox picInvItem 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   540
         Index           =   14
         Left            =   3105
         ScaleHeight     =   34
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   107
         Top             =   1050
         Width           =   510
      End
      Begin VB.PictureBox picInvItem 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   540
         Index           =   15
         Left            =   3600
         ScaleHeight     =   34
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   106
         Top             =   1050
         Width           =   510
      End
      Begin VB.PictureBox picInvItem 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   540
         Index           =   16
         Left            =   135
         ScaleHeight     =   34
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   105
         Top             =   1590
         Width           =   510
      End
      Begin VB.PictureBox picInvItem 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   540
         Index           =   17
         Left            =   630
         ScaleHeight     =   34
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   104
         Top             =   1590
         Width           =   510
      End
      Begin VB.PictureBox picInvItem 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   540
         Index           =   18
         Left            =   1125
         ScaleHeight     =   34
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   103
         Top             =   1590
         Width           =   510
      End
      Begin VB.PictureBox picInvItem 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   540
         Index           =   19
         Left            =   1620
         ScaleHeight     =   34
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   102
         Top             =   1590
         Width           =   510
      End
      Begin VB.PictureBox picInvItem 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   540
         Index           =   20
         Left            =   2115
         ScaleHeight     =   34
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   101
         Top             =   1590
         Width           =   510
      End
      Begin VB.PictureBox picInvItem 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   540
         Index           =   21
         Left            =   2610
         ScaleHeight     =   34
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   100
         Top             =   1590
         Width           =   510
      End
      Begin VB.PictureBox picInvItem 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   540
         Index           =   22
         Left            =   3105
         ScaleHeight     =   34
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   99
         Top             =   1590
         Width           =   510
      End
      Begin VB.PictureBox picInvItem 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   540
         Index           =   23
         Left            =   3600
         ScaleHeight     =   34
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   98
         Top             =   1590
         Width           =   510
      End
      Begin VB.PictureBox picInvItem 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   540
         Index           =   24
         Left            =   135
         ScaleHeight     =   34
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   97
         Top             =   2130
         Width           =   510
      End
      Begin VB.PictureBox picInvItem 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   540
         Index           =   25
         Left            =   630
         ScaleHeight     =   34
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   96
         Top             =   2130
         Width           =   510
      End
      Begin VB.PictureBox picInvItem 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   540
         Index           =   26
         Left            =   1125
         ScaleHeight     =   34
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   95
         Top             =   2130
         Width           =   510
      End
      Begin VB.PictureBox picInvItem 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   540
         Index           =   27
         Left            =   1620
         ScaleHeight     =   34
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   94
         Top             =   2130
         Width           =   510
      End
      Begin VB.PictureBox picInvItem 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   540
         Index           =   28
         Left            =   2115
         ScaleHeight     =   34
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   93
         Top             =   2130
         Width           =   510
      End
      Begin VB.PictureBox picInvItem 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   540
         Index           =   29
         Left            =   2610
         ScaleHeight     =   34
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   92
         Top             =   2130
         Width           =   510
      End
      Begin VB.PictureBox picInvItem 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   540
         Index           =   30
         Left            =   3105
         ScaleHeight     =   34
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   91
         Top             =   2130
         Width           =   510
      End
      Begin VB.PictureBox picInvItem 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   540
         Index           =   31
         Left            =   3600
         ScaleHeight     =   34
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   90
         Top             =   2130
         Width           =   510
      End
      Begin VB.PictureBox picInvItem 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   540
         Index           =   32
         Left            =   5445
         ScaleHeight     =   34
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   89
         Top             =   3150
         Width           =   510
      End
      Begin VB.PictureBox picInvItem 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   540
         Index           =   33
         Left            =   6045
         ScaleHeight     =   34
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   88
         Top             =   3150
         Width           =   510
      End
      Begin VB.PictureBox picInvItem 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   540
         Index           =   34
         Left            =   6645
         ScaleHeight     =   34
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   87
         Top             =   3150
         Width           =   510
      End
      Begin VB.PictureBox picInvItem 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   540
         Index           =   35
         Left            =   7245
         ScaleHeight     =   34
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   86
         Top             =   3150
         Width           =   510
      End
      Begin VB.PictureBox picInvItem 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   540
         Index           =   36
         Left            =   4245
         ScaleHeight     =   34
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   85
         Top             =   3705
         Width           =   510
      End
      Begin VB.PictureBox picInvItem 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   540
         Index           =   37
         Left            =   4845
         ScaleHeight     =   34
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   84
         Top             =   3705
         Width           =   510
      End
      Begin VB.PictureBox picInvItem 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   540
         Index           =   38
         Left            =   5445
         ScaleHeight     =   34
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   83
         Top             =   3705
         Width           =   510
      End
      Begin VB.PictureBox picInvItem 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   540
         Index           =   39
         Left            =   6045
         ScaleHeight     =   34
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   82
         Top             =   3705
         Width           =   510
      End
      Begin VB.PictureBox picInvItem 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   540
         Index           =   40
         Left            =   6645
         ScaleHeight     =   34
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   81
         Top             =   3705
         Width           =   510
      End
      Begin VB.PictureBox picInvItem 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   540
         Index           =   41
         Left            =   7245
         ScaleHeight     =   34
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   80
         Top             =   3705
         Width           =   510
      End
      Begin VB.PictureBox picInvItem 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   540
         Index           =   42
         Left            =   4245
         ScaleHeight     =   34
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   79
         Top             =   4260
         Width           =   510
      End
      Begin VB.PictureBox picInvItem 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   540
         Index           =   43
         Left            =   4845
         ScaleHeight     =   34
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   78
         Top             =   4260
         Width           =   510
      End
      Begin VB.PictureBox picInvItem 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   540
         Index           =   44
         Left            =   5445
         ScaleHeight     =   34
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   77
         Top             =   4260
         Width           =   510
      End
      Begin VB.PictureBox picInvItem 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   540
         Index           =   45
         Left            =   6045
         ScaleHeight     =   34
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   76
         Top             =   4260
         Width           =   510
      End
      Begin VB.PictureBox picInvItem 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   540
         Index           =   46
         Left            =   6645
         ScaleHeight     =   34
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   75
         Top             =   4260
         Width           =   510
      End
      Begin VB.PictureBox picInvItem 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   540
         Index           =   47
         Left            =   7245
         ScaleHeight     =   34
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   74
         Top             =   4260
         Width           =   510
      End
      Begin VB.PictureBox picInvItem 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   510
         Index           =   48
         Left            =   4245
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   73
         Top             =   4830
         Width           =   510
      End
      Begin VB.PictureBox picInvItem 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   510
         Index           =   49
         Left            =   4845
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   72
         Top             =   4830
         Width           =   510
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "i"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   5970
         TabIndex        =   153
         Top             =   4065
         Width           =   165
      End
      Begin VB.Label lblInvUse 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Use Item"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2085
         TabIndex        =   124
         Top             =   225
         Width           =   1095
      End
      Begin VB.Label lblInvDrop 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Drop Item"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3165
         TabIndex        =   123
         Top             =   225
         Width           =   1095
      End
      Begin VB.Label lblInvClose 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6360
         TabIndex        =   122
         Top             =   4125
         Width           =   135
      End
   End
   Begin VB.PictureBox picItems 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   13740
      ScaleHeight     =   345
      ScaleWidth      =   1665
      TabIndex        =   70
      Top             =   12360
      Width           =   1695
   End
   Begin VB.PictureBox picQuit 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
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
      Height          =   345
      Left            =   11640
      Picture         =   "frmMirage.frx":64A27
      ScaleHeight     =   345
      ScaleWidth      =   360
      TabIndex        =   40
      ToolTipText     =   "Quit"
      Top             =   8610
      Width           =   360
   End
   Begin VB.PictureBox picTrain 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
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
      Height          =   345
      Left            =   9420
      Picture         =   "frmMirage.frx":66E5A
      ScaleHeight     =   345
      ScaleWidth      =   360
      TabIndex        =   39
      ToolTipText     =   "Train"
      Top             =   8625
      Width           =   360
   End
   Begin VB.PictureBox picSpells 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
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
      Height          =   345
      Left            =   9030
      Picture         =   "frmMirage.frx":692CC
      ScaleHeight     =   345
      ScaleWidth      =   360
      TabIndex        =   38
      ToolTipText     =   "Spells"
      Top             =   8625
      Width           =   360
   End
   Begin VB.PictureBox picPlayerSpells 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
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
      Height          =   3285
      Left            =   18255
      ScaleHeight     =   217
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   255
      TabIndex        =   31
      Top             =   8505
      Visible         =   0   'False
      Width           =   3855
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Spells"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   165
         TabIndex        =   33
         Top             =   -30
         Width           =   1935
      End
      Begin VB.Label lblSpellsCancel 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   32
         Top             =   2880
         Width           =   1935
      End
   End
   Begin VB.PictureBox picInv 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
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
      Height          =   3285
      Left            =   18315
      ScaleHeight     =   217
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   255
      TabIndex        =   25
      Top             =   8475
      Visible         =   0   'False
      Width           =   3855
      Begin VB.ListBox lstInv 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1920
         ItemData        =   "frmMirage.frx":6B6C6
         Left            =   0
         List            =   "frmMirage.frx":6B6C8
         TabIndex        =   27
         Top             =   240
         Width           =   3855
      End
      Begin VB.Label lblCancel 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   30
         Top             =   2880
         Width           =   1935
      End
      Begin VB.Label lblDropItem 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Drop Item"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   29
         Top             =   2520
         Width           =   1935
      End
      Begin VB.Label lblUseItem 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Use Item"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   28
         Top             =   2160
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Inventory"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   26
         Top             =   0
         Width           =   1935
      End
   End
   Begin MSWinsockLib.Winsock Socket 
      Left            =   120
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.PictureBox PicInventory 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
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
      Height          =   345
      Index           =   0
      Left            =   7860
      Picture         =   "frmMirage.frx":6B6CA
      ScaleHeight     =   345
      ScaleWidth      =   360
      TabIndex        =   69
      ToolTipText     =   "Inventory"
      Top             =   8625
      Width           =   360
   End
   Begin VB.PictureBox picPrayer 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   8640
      Picture         =   "frmMirage.frx":6DA82
      ScaleHeight     =   345
      ScaleWidth      =   360
      TabIndex        =   142
      ToolTipText     =   "Prayers"
      Top             =   8625
      Width           =   360
   End
   Begin VB.PictureBox picPaperDoll 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   8250
      Picture         =   "frmMirage.frx":6FEB0
      ScaleHeight     =   345
      ScaleWidth      =   360
      TabIndex        =   143
      ToolTipText     =   "Equipment"
      Top             =   8625
      Width           =   360
   End
   Begin RichTextLib.RichTextBox txtChannelGeneral 
      Height          =   1950
      Left            =   12615
      TabIndex        =   241
      Top             =   15
      Width           =   7620
      _ExtentX        =   13441
      _ExtentY        =   3440
      _Version        =   393217
      BackColor       =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmMirage.frx":7227E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox txtChannelGlobal 
      Height          =   1950
      Left            =   12555
      TabIndex        =   242
      Top             =   -45
      Width           =   7620
      _ExtentX        =   13441
      _ExtentY        =   3440
      _Version        =   393217
      BackColor       =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmMirage.frx":722F5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox txtChannelGuild 
      Height          =   1950
      Left            =   12525
      TabIndex        =   243
      Top             =   30
      Width           =   7620
      _ExtentX        =   13441
      _ExtentY        =   3440
      _Version        =   393217
      BackColor       =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmMirage.frx":7236C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox txtChannelPrivate 
      Height          =   1950
      Left            =   12525
      TabIndex        =   244
      Top             =   -15
      Width           =   7620
      _ExtentX        =   13441
      _ExtentY        =   3440
      _Version        =   393217
      BackColor       =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmMirage.frx":723E3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox txtChannelAll 
      Height          =   2385
      Left            =   90
      TabIndex        =   245
      Top             =   6495
      Width           =   7605
      _ExtentX        =   13414
      _ExtentY        =   4207
      _Version        =   393217
      BackColor       =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmMirage.frx":7245A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblstreetname 
      Alignment       =   2  'Center
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
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   7920
      TabIndex        =   261
      Top             =   360
      Width           =   3885
   End
   Begin VB.Label lblMapName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7920
      TabIndex        =   257
      Top             =   120
      Width           =   3885
   End
   Begin VB.Label lblLevel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "LEVEL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   9195
      TabIndex        =   152
      Top             =   3675
      Width           =   1455
   End
   Begin VB.Label lblblock 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0%"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   9435
      TabIndex        =   151
      Top             =   4635
      Width           =   1500
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "block"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8955
      TabIndex        =   150
      Top             =   4635
      Width           =   615
   End
   Begin VB.Shape shpblock 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   135
      Left            =   8955
      Top             =   4875
      Width           =   1950
   End
   Begin VB.Label lblcrit 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0%"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   9435
      TabIndex        =   149
      Top             =   4275
      Width           =   1500
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "critical"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8955
      TabIndex        =   148
      Top             =   4275
      Width           =   1095
   End
   Begin VB.Shape shpcrit 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   135
      Left            =   8955
      Top             =   4515
      Width           =   1950
   End
   Begin VB.Label lblexp 
      BackStyle       =   0  'Transparent
      Caption         =   "0%"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   9435
      TabIndex        =   147
      Top             =   3915
      Width           =   2340
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "exp"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8955
      TabIndex        =   146
      Top             =   3915
      Width           =   375
   End
   Begin VB.Shape shpexp 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   135
      Left            =   8955
      Top             =   4155
      Width           =   1950
   End
   Begin VB.Label lblStats2 
      BackStyle       =   0  'Transparent
      Caption         =   "STATS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   9960
      TabIndex        =   145
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label lblStats1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "STATS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   8355
      TabIndex        =   144
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label lblPP 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0%"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   9435
      TabIndex        =   140
      Top             =   2250
      Width           =   1500
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Grace"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8955
      TabIndex        =   139
      Top             =   2250
      Width           =   675
   End
   Begin VB.Shape shpPP 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   135
      Left            =   8955
      Top             =   2490
      Width           =   1950
   End
   Begin VB.Shape shpHP 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   135
      Left            =   8955
      Top             =   1050
      Width           =   1950
   End
   Begin VB.Shape shpMP 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   135
      Left            =   8955
      Top             =   1530
      Width           =   1950
   End
   Begin VB.Shape shpSP 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   135
      Left            =   8970
      Top             =   2010
      Width           =   1950
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Stamina"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8955
      TabIndex        =   43
      Top             =   1770
      Width           =   840
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "MP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8955
      TabIndex        =   42
      Top             =   1290
      Width           =   375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "HP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8955
      TabIndex        =   41
      Top             =   810
      Width           =   615
   End
   Begin VB.Label lblSP 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0%"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   9435
      TabIndex        =   37
      Top             =   1770
      Width           =   1500
   End
   Begin VB.Label lblMP 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0%"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   9435
      TabIndex        =   36
      Top             =   1290
      Width           =   1500
   End
   Begin VB.Label lblHP 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0%"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   9435
      TabIndex        =   35
      Top             =   810
      Width           =   1500
   End
   Begin VB.Shape Shape9 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   135
      Left            =   8955
      Top             =   4875
      Width           =   1950
   End
   Begin VB.Shape Shape8 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   135
      Left            =   8955
      Top             =   4515
      Width           =   1950
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   135
      Left            =   8955
      Top             =   2490
      Width           =   1950
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   135
      Left            =   8970
      Top             =   2010
      Width           =   1950
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   135
      Left            =   8955
      Top             =   1530
      Width           =   1950
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   135
      Left            =   8955
      Top             =   1035
      Width           =   1935
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   135
      Left            =   8955
      Top             =   4155
      Width           =   1950
   End
   Begin VB.Menu charmnu 
      Caption         =   "Character"
      Begin VB.Menu editBio 
         Caption         =   "Edit Bio"
      End
      Begin VB.Menu mnuShowEquipment 
         Caption         =   "Show Equipment"
      End
      Begin VB.Menu mnusavetxt 
         Caption         =   "Save text"
      End
   End
   Begin VB.Menu mnutrade 
      Caption         =   "Trading"
      Begin VB.Menu mnubank 
         Caption         =   "Bank"
      End
      Begin VB.Menu mnushop 
         Caption         =   "Shop"
      End
   End
   Begin VB.Menu mnuOpts 
      Caption         =   "Game Options"
      Begin VB.Menu mnuSound 
         Caption         =   "Sound"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnumusic 
         Caption         =   "Music"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuhlp 
      Caption         =   "Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
      Begin VB.Menu mnuwebsite 
         Caption         =   "Website"
      End
      Begin VB.Menu mnuforum 
         Caption         =   "Support Forum"
      End
   End
End
Attribute VB_Name = "frmMirage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tileNo As Long
Dim t As Long
Dim lightningCount As Long
Dim weatherCounter As Long











Private Sub cmdPrayers_Click()
    Call SendData("Prayers" & SEP_CHAR & END_CHAR)
End Sub














Private Sub cmbAtributes_Click()
    Select Case cmbAtributes.ListIndex
        Case Is = 0
            'Blocked
        Case Is = 1
            'Warp
            frmMapWarp.Show vbModal
        Case Is = 2
            'Warp (Level)
            frmMapWarp.Show vbModal
        Case Is = 3
            'Item
            frmMapItem.Show vbModal
        Case Is = 4
            'NpcAvoid
        Case Is = 5
            'key
            frmMapKey.Show vbModal
        Case Is = 6
            'Key Open
            frmKeyOpen.Show vbModal
        Case Is = 7
            'Damage
            EditorDamage = Val(InputBox("Please enter the number of HP to remove."))
        Case Is = 8
            'Heal
            EditorDamage = Val(InputBox("Please enter the number of HP to regen."))
        Case Is = 9
            'Sign
            frmMapSign.Show vbModal
        Case Is = 10
            'Spawn Npc
            frmMapNPCSpawn.Show vbModal
    End Select
End Sub

Private Sub cmbTilePack_Click()
Dim Index As Long
'scrlPicture.Value = 0
Index = cmbTilePack.ListIndex
picBackSelect.Picture = LoadPicture(App.Path + "\data\bmp\tiles" & Index & ".bmp")
tileNo = Index
scrlPicture.Max = picBackSelect.Height / PIC_Y
End Sub



Private Sub cmdDepItem_Click()
Dim i As Long
        i = selectedBankInvLocalItem + 1
            Call SendData("bankitem" & SEP_CHAR & i & SEP_CHAR & END_CHAR)
End Sub

Private Sub cmdDeposit_Click()
Call SendData("bankgold" & SEP_CHAR & Val(txtDeposit.text) & SEP_CHAR & END_CHAR)
End Sub


Private Sub cmdQuestOK_Click()
picQuestMsg.Visible = False
End Sub

Private Sub cmdWithdraw_Click()
Call SendData("unbankgold" & SEP_CHAR & Val(txtWithdraw.text) & SEP_CHAR & END_CHAR)

End Sub

Private Sub cmdWithItem_Click()
Dim i As Long
i = selectedBankItem + 1
Call SendData("unbankitem" & SEP_CHAR & i & SEP_CHAR & END_CHAR)
End Sub






Private Sub editBio_Click()
    SendData ("REQUESTEDITBIO" & SEP_CHAR & END_CHAR)
End Sub

Private Sub Form_Activate()

    picItems.Picture = LoadPicture(App.Path & "\data\bmp\Items.bmp")
    If Not InEditor Then
        Me.Width = FORM_X
        Me.Height = FORM_Y
        Me.Refresh
    End If
        Call initPanes
        Me.Refresh
        
        Me.txtChannelAll.Visible = True
        Me.txtChannelGeneral.Visible = False
        Me.txtChannelGlobal.Visible = False
        Me.txtChannelGuild.Visible = False
        Me.txtChannelPrivate.Visible = False
        
        
        MuteMusic = LoadMusicOption
        MuteSound = LoadSoundOption
        mnumusic.Checked = Not MuteMusic
        mnuSound.Checked = Not MuteSound

End Sub

Private Sub Form_Click()
    'Call SendData("getstats" & SEP_CHAR & END_CHAR)
    txtChat.Refresh
End Sub

Private Sub Form_GotFocus()
    'frmMirage.txtSend.SetFocus
End Sub

Private Sub Form_Terminate()
    Call GameDestroy
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call GameDestroy
End Sub





Private Sub lblblock_Click()
    On Local Error Resume Next
    shpcrit.Width = (Val(Left(lblcrit, Len(lblcrit) - 1)) / 100) * 130
End Sub

Private Sub lblCastP_Click()
    If Player(MyIndex).Prayer(lstSpells.ListIndex + 1) > 0 Then
            If Player(MyIndex).moving = 0 Then
                Call SendData("castp" & SEP_CHAR & lstSpells.ListIndex + 1 & SEP_CHAR & END_CHAR)
                Player(MyIndex).Attacking = 1
                Player(MyIndex).AttackTimer = GetTickCount
                Player(MyIndex).CastedSpell = YES
                
            Else
                Call AddTextNew(frmMirage.txtChannelAll, "Cannot cast while walking!", RGB_AlertColor)
            End If
    Else
        Call AddTextNew(frmMirage.txtChannelAll, "No prayer here.", RGB_AlertColor)
    End If
End Sub

Private Sub lblCloseBank_Click()
picBankWindow.Visible = False
InBank = False
End Sub

Private Sub lblcrit_Click()
    On Local Error Resume Next
    shpcrit.Width = (Val(Left(lblcrit, Len(lblcrit) - 1)) / 100) * 130
End Sub

Private Sub lblEquipClose_Click()
picEquiped.Visible = False
End Sub

Private Sub lblexp_Click()
    On Local Error Resume Next
    Dim pec() As String
    pec = Split(lblexp.Caption, "/")
    shpexp.Width = (CLng(pec(0)) / CLng(pec(1))) * 130
End Sub

Private Sub lblHP_Change()
    On Local Error Resume Next
    Dim pec() As String
    pec = Split(lblHP.Caption, "/")
    shpHP.Width = (CLng(pec(0)) / CLng(pec(1))) * 130
End Sub




Private Sub lblInfo_Click()
On Error Resume Next
    Call SendData("ITEMLIB" & SEP_CHAR & Item(GetPlayerInvItemNum(MyIndex, selectedInvItem + 1)).Name & SEP_CHAR & END_CHAR)
    
End Sub

Private Sub lblInvClose_Click()
Me.picInvPane.Visible = False
End Sub

Private Sub lblInvDrop_Click()
Dim value As Long
Dim InvNum As Long

    InvNum = selectedInvItem + 1

    If GetPlayerInvItemNum(MyIndex, InvNum) > 0 And GetPlayerInvItemNum(MyIndex, InvNum) <= MAX_ITEMS Then
        If Item(GetPlayerInvItemNum(MyIndex, InvNum)).type = ITEM_TYPE_CURRENCY Then
            ' Show them the drop dialog
            frmDrop.Show vbModal
        Else
            Call SendDropItem(selectedInvItem + 1, 0)
        End If
    End If
    Call UpdateInventory
End Sub

Private Sub lblInvUse_Click()
    Call SendUseItem(selectedInvItem + 1)
End Sub




Private Sub lblMoveTilesLeft_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If picBack.Left < 8 Then
    picBack.Left = picBack.Left + PIC_X
End If
End Sub

Private Sub lblMoveTilesLeft_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then
    If picBack.Left < 8 Then
        'picBack.Left = picBack.Left + 1
    End If
End If
End Sub

Private Sub lblMoveTilesRight_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

'picBack.Left = picBack.Left - 1
If picBack.Left > -224 Then
    picBack.Left = picBack.Left - PIC_X
End If
End Sub

Private Sub lblMoveTilesRight_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

If Button = 1 Then
    If picBack.Left > -224 Then
       ' picBack.Left = picBack.Left - 1
    End If
End If
End Sub

Private Sub lblMP_Change()
    On Local Error Resume Next
    Dim pec() As String
    pec = Split(lblMP.Caption, "/")
    shpMP.Width = (CLng(pec(0)) / CLng(pec(1))) * 130
End Sub

Private Sub lblPP_Change()
    On Local Error Resume Next
    Dim pec() As String
    pec = Split(lblPP.Caption, "/")
    shpPP.Width = (CLng(pec(0)) / CLng(pec(1))) * 130
End Sub


Private Sub lblSP_Change()
    On Local Error Resume Next
    Dim pec() As String
    pec = Split(lblSP.Caption, "/")
    shpSP.Width = (CLng(pec(0)) / CLng(pec(1))) * 130
End Sub












Private Sub mnuAbout_Click()
Dim x
x = ShellExecute(Me.hwnd, "Open", "http://www.afterdarkness.squiggleuk.com/", 0&, 0&, 0&)
End Sub

Private Sub mnubank_Click()
    Call SendData("playerbank" & SEP_CHAR & END_CHAR)
    DoEvents
    Call SendData("playerbank" & SEP_CHAR & END_CHAR)
    DoEvents
End Sub



Private Sub mnuforum_Click()
Dim x
x = ShellExecute(Me.hwnd, "Open", "http://www.afterdarkness.squiggleuk.com/community/", 0&, 0&, 0&)

End Sub

Private Sub mnumusic_Click()
If mnumusic.Checked = True Then
    MuteMusic = True
    mnumusic.Checked = False
    StopMidi
Else
    MuteMusic = False
    mnumusic.Checked = True
    Call PlayMidi(MusicPlaying, , True)
End If
End Sub

Private Sub mnusavetxt_Click()
    Dim intFile As Long
    intFile = FreeFile()
    
    Open App.Path & "\chatLog.txt" For Output As #intFile
        Print #intFile, Me.txtChannelAll.text
    Close #intFile

End Sub

Private Sub mnushop_Click()
    Call SendData("trade" & SEP_CHAR & END_CHAR)
End Sub

Private Sub mnuShowEquipment_Click()
Call clearPanes
picEquiped.Visible = True
DoEvents
blitEquip
End Sub



Private Sub mnuSound_Click()
If mnuSound.Checked = True Then
    MuteSound = True
    mnuSound.Checked = False
    modDatabase.setMconfig MuteMusic, MuteSound
Else
    MuteSound = False
    mnuSound.Checked = True
    modDatabase.setMconfig MuteMusic, MuteSound
End If

End Sub



Private Sub mnuwebsite_Click()
Dim x
x = ShellExecute(Me.hwnd, "Open", "http://www.afterdarkness.cjb.net", 0&, 0&, 0&)
End Sub

Private Sub opt1_Click(Index As Integer)
picBackSelect.Picture = LoadPicture(App.Path + "\data\bmp\tiles" & Index & ".bmp")
tileNo = Index
scrlPicture.Max = picBackSelect.Height / PIC_Y
End Sub






Private Sub optDamage_Click()
    EditorDamage = Val(InputBox("Please enter the number of HP to remove."))
End Sub

Private Sub optHeal_Click()
    EditorDamage = Val(InputBox("Please enter the number of HP to regen."))
End Sub

Private Sub optNpcSpawn_Click()
    frmMapNPCSpawn.Show vbModal
End Sub

Private Sub optSign_Click()
    frmMapSign.Show vbModal
End Sub

Private Sub optWarpDoor_Click()
    frmMapWarp.Show vbModal
End Sub












Private Sub picbankItem_Click(Index As Integer)
selectedBankItem = Index
Call SendData("playerbank" & SEP_CHAR & END_CHAR)
DoEvents
End Sub







Private Sub picChannelAll_Click()
Me.txtChannelAll.Visible = True
Me.txtChannelGeneral.Visible = False
Me.txtChannelGlobal.Visible = False
Me.txtChannelGuild.Visible = False
Me.txtChannelPrivate.Visible = False
Me.picChannelAll.Picture = Me.picChannelAllImg(0).Picture
Me.picChannelGeneral.Picture = Me.picChannelGeneralImg(1).Picture
Me.picChannelGlobal.Picture = Me.picChannelGlobalImg(1).Picture
Me.picChannelGuild.Picture = Me.picChannelGuildImg(1).Picture
Me.picChannelPrivate.Picture = Me.picChannelPrivateImg(1).Picture
End Sub

Private Sub picChannelGeneral_Click()
Me.txtChannelAll.Visible = False
Me.txtChannelGeneral.Visible = True
Me.txtChannelGlobal.Visible = False
Me.txtChannelGuild.Visible = False
Me.txtChannelPrivate.Visible = False
Me.picChannelAll.Picture = Me.picChannelAllImg(1).Picture
Me.picChannelGeneral.Picture = Me.picChannelGeneralImg(0).Picture
Me.picChannelGlobal.Picture = Me.picChannelGlobalImg(1).Picture
Me.picChannelGuild.Picture = Me.picChannelGuildImg(1).Picture
Me.picChannelPrivate.Picture = Me.picChannelPrivateImg(1).Picture
End Sub

Private Sub picChannelGlobal_Click()
Me.txtChannelAll.Visible = False
Me.txtChannelGeneral.Visible = False
Me.txtChannelGlobal.Visible = True
Me.txtChannelGuild.Visible = False
Me.txtChannelPrivate.Visible = False
Me.picChannelAll.Picture = Me.picChannelAllImg(1).Picture
Me.picChannelGeneral.Picture = Me.picChannelGeneralImg(1).Picture
Me.picChannelGlobal.Picture = Me.picChannelGlobalImg(0).Picture
Me.picChannelGuild.Picture = Me.picChannelGuildImg(1).Picture
Me.picChannelPrivate.Picture = Me.picChannelPrivateImg(1).Picture
End Sub

Private Sub picChannelGuild_Click()
Me.txtChannelAll.Visible = False
Me.txtChannelGeneral.Visible = False
Me.txtChannelGlobal.Visible = False
Me.txtChannelGuild.Visible = True
Me.txtChannelPrivate.Visible = False
Me.picChannelAll.Picture = Me.picChannelAllImg(1).Picture
Me.picChannelGeneral.Picture = Me.picChannelGeneralImg(1).Picture
Me.picChannelGlobal.Picture = Me.picChannelGlobalImg(1).Picture
Me.picChannelGuild.Picture = Me.picChannelGuildImg(0).Picture
Me.picChannelPrivate.Picture = Me.picChannelPrivateImg(1).Picture
End Sub

Private Sub picChannelPrivate_Click()
Me.txtChannelAll.Visible = False
Me.txtChannelGeneral.Visible = False
Me.txtChannelGlobal.Visible = False
Me.txtChannelGuild.Visible = False
Me.txtChannelPrivate.Visible = True
Me.picChannelAll.Picture = Me.picChannelAllImg(1).Picture
Me.picChannelGeneral.Picture = Me.picChannelGeneralImg(1).Picture
Me.picChannelGlobal.Picture = Me.picChannelGlobalImg(1).Picture
Me.picChannelGuild.Picture = Me.picChannelGuildImg(1).Picture
Me.picChannelPrivate.Picture = Me.picChannelPrivateImg(0).Picture
End Sub

Private Sub picInventory_Click(Index As Integer)
Call clearPanes
    picInvPane.Visible = True
    Call UpdateInventory
End Sub

Private Sub picInvItem_Click(Index As Integer)
Dim i As Long
selectedInvItem = Index
Call UpdateInventory
End Sub





Private Sub picItemLocal_Click(Index As Integer)
selectedBankInvLocalItem = Index
Call SendData("playerbank" & SEP_CHAR & END_CHAR)
DoEvents
End Sub



Private Sub picPaperDoll_Click()
Call clearPanes
picEquiped.Visible = True
DoEvents
blitEquip
End Sub

Private Sub picPrayer_Click()
    Call clearPanes
    Call SendData("Prayers" & SEP_CHAR & END_CHAR)
    lblcast.Visible = False
    lblCastP.Visible = True
End Sub

Private Sub picScreen_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call EditorMouseDown(Button, Shift, x, y, tileNo)
    Call PlayerSearch(Button, Shift, x, y)
    frmMirage.txtSend.SetFocus

End Sub

Private Sub picScreen_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call EditorMouseDown(Button, Shift, x, y, tileNo)
End Sub





Private Sub picSpellClose_Click()
Me.picSpellPane.Visible = False
End Sub







Private Sub Socket_DataArrival(ByVal bytesTotal As Long)
    If IsConnected Then
        Call IncomingData(bytesTotal)
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Call HandleKeypresses(KeyAscii)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
    If InBank = False And frmSprite.Visible = False Then
        txtSend.SetFocus
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Call CheckInput(1, KeyCode, Shift)
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Call CheckInput(0, KeyCode, Shift)
End Sub




Private Sub timMusic_Timer()
If MuteMusic = False Then
    If MP3.MP3Playing = False Then
        Call PlayMidi(MusicPlaying, True)
    End If
End If
End Sub



Private Sub timWeather_Timer()
If blnNight Then
    If weatherCounter >= lightningCount Then
        lightningCount = Int(Rnd() * 100)
        weatherCounter = 0
    End If
    weatherCounter = weatherCounter + 1
End If

End Sub

Private Sub lblUseItem_Click()
    Call SendUseItem(frmMirage.lstInv.ListIndex + 1)
End Sub

Private Sub lblDropItem_Click()
Dim value As Long
Dim InvNum As Long

    InvNum = frmMirage.lstInv.ListIndex + 1

    If GetPlayerInvItemNum(MyIndex, InvNum) > 0 And GetPlayerInvItemNum(MyIndex, InvNum) <= MAX_ITEMS Then
        If Item(GetPlayerInvItemNum(MyIndex, InvNum)).type = ITEM_TYPE_CURRENCY Then
            ' Show them the drop dialog
            frmDrop.Show vbModal
        Else
            Call SendDropItem(frmMirage.lstInv.ListIndex + 1, 0)
        End If
    End If
End Sub

Private Sub lblCast_Click()
    If Player(MyIndex).Spell(lstSpells.ListIndex + 1) > 0 Then
        If GetTickCount > Player(MyIndex).AttackTimer + 1000 Then
            If Player(MyIndex).moving = 0 Then
                Call SendData("cast" & SEP_CHAR & lstSpells.ListIndex + 1 & SEP_CHAR & END_CHAR)
                Player(MyIndex).Attacking = 1
                Player(MyIndex).AttackTimer = GetTickCount
                Player(MyIndex).CastedSpell = YES
                
            Else
                Call AddTextNew(frmMirage.txtChannelAll, "Cannot cast while walking!", RGB_AlertColor)
            End If
        End If
    Else
        Call AddTextNew(frmMirage.txtChannelAll, "No spell here.", RGB_AlertColor)
    End If
End Sub

Private Sub lblCancel_Click()
    picInv.Visible = False
End Sub

Private Sub lblSpellsCancel_Click()
    picPlayerSpells.Visible = False
End Sub

Private Sub picSpells_Click()
Call clearPanes
    Call SendData("spells" & SEP_CHAR & END_CHAR)
    lblcast.Visible = True
    lblCastP.Visible = False
End Sub

Private Sub picTrain_Click()
    Call clearPanes
    frmTraining.Show vbModal
End Sub



Private Sub picQuit_Click()
    Call GameDestroy
End Sub

' // MAP EDITOR STUFF //

Private Sub optLayers_Click()
    If optLayers.value = True Then
        cmbTilePack.Enabled = True
        cmbLayers.Enabled = True
        cmbAtributes.Enabled = False
    End If
End Sub

Private Sub optAttribs_Click()
    If optAttribs.value = True Then
        cmbTilePack.Enabled = False
        cmbLayers.Enabled = False
        cmbAtributes.Enabled = True
    End If
End Sub

Private Sub picBackSelect_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call EditorChooseTile(Button, Shift, x, y)
End Sub

Private Sub picBackSelect_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call EditorChooseTile(Button, Shift, x, y)
End Sub

Private Sub cmdSend_Click()
    Call EditorSend
    Me.Width = FORM_X
End Sub

Private Sub cmdCancel_Click()
    Call EditorCancel
    If Me.Width = FORM_EDITOR_X Then
        For t = FORM_EDITOR_X To FORM_X Step -100
            Me.Width = t
            Me.Refresh
            DoEvents
        Next t
    End If
    Me.Width = FORM_X
End Sub

Private Sub cmdProperties_Click()
    frmMapProperties.Show vbModal
End Sub

Private Sub optItem_Click()
    frmMapItem.Show vbModal
End Sub

Private Sub optKey_Click()
    frmMapKey.Show vbModal
End Sub

Private Sub optKeyOpen_Click()
    frmKeyOpen.Show vbModal
End Sub

Private Sub scrlPicture_Change()
    Call EditorTileScroll
End Sub

Private Sub cmdClear_Click()
    If Me.cmbAtributes.Enabled = True Then
        Call EditorClearAttribs
    Else
        Call EditorClearLayer
    End If
End Sub






Private Sub txtSend_Change()
    MyText = txtSend.text
End Sub

Private Sub initPanes()
    picEquiped.Left = 520
    picEquiped.top = 391
    
    picInvPane.Left = 520
    picInvPane.top = 391
    
    picSpellPane.Left = 520
    picSpellPane.top = 391
    
    picPrayerPane.Left = 520
    picPrayerPane.top = 391
    
    picMapEditor.Left = 521
    picMapEditor.top = 0
    
    picBankWindow.top = 37
    picBankWindow.Left = 29
    
    Me.txtChannelAll.Visible = False
    Me.txtChannelGlobal.Visible = False
    Me.txtChannelGuild.Visible = False
    Me.txtChannelPrivate.Visible = False
    
    
    DoEvents
End Sub

Public Sub clearPanes()
    picEquiped.Visible = False
    picMapEditor.Visible = False
    picInvPane.Visible = False
    picSpellPane.Visible = False
    picBankWindow.Visible = False
    picPrayerPane.Visible = False
    InBank = False
End Sub


