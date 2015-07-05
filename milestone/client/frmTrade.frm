VERSION 5.00
Begin VB.Form frmTrade 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Trade"
   ClientHeight    =   6000
   ClientLeft      =   30
   ClientTop       =   -90
   ClientWidth     =   6000
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmTrade.frx":0000
   ScaleHeight     =   400
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox itmDesc 
      Appearance      =   0  'Flat
      BackColor       =   &H00789298&
      ForeColor       =   &H80000008&
      Height          =   3495
      Left            =   120
      ScaleHeight     =   231
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   175
      TabIndex        =   36
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
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   210
         Left            =   360
         TabIndex        =   47
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   360
         TabIndex        =   46
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label descStr 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Strength"
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
         Height          =   210
         Left            =   360
         TabIndex        =   45
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label descDef 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Defence"
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
         Height          =   210
         Left            =   360
         TabIndex        =   44
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label descSpeed 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Speed"
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
         Height          =   210
         Left            =   360
         TabIndex        =   43
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   360
         TabIndex        =   42
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label descHpMp 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "HP: XXXX MP: XXXX SP: XXXX"
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
         Height          =   210
         Left            =   0
         TabIndex        =   41
         Top             =   1440
         Width           =   2655
      End
      Begin VB.Label descSD 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Str: XXXX Def: XXXXX"
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
         Height          =   210
         Left            =   0
         TabIndex        =   40
         Top             =   1680
         Width           =   2655
      End
      Begin VB.Label desc 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
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
         Height          =   975
         Left            =   120
         TabIndex        =   39
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   360
         TabIndex        =   38
         Top             =   2160
         Width           =   1815
      End
      Begin VB.Label descMS 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Magi: XXXXX Speed: XXXX"
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
         Height          =   210
         Left            =   0
         TabIndex        =   37
         Top             =   1920
         Width           =   2655
      End
   End
   Begin VB.PictureBox picShopItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   7
      Left            =   3000
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   32
      Top             =   3960
      Width           =   480
   End
   Begin VB.PictureBox picShopItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   6
      Left            =   360
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   28
      Top             =   3960
      Width           =   480
   End
   Begin VB.PictureBox picShopItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   5
      Left            =   3000
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   24
      Top             =   3120
      Width           =   480
   End
   Begin VB.PictureBox picShopItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   4
      Left            =   360
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   20
      Top             =   3120
      Width           =   480
   End
   Begin VB.PictureBox picShopItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   3
      Left            =   3000
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   16
      Top             =   2280
      Width           =   480
   End
   Begin VB.PictureBox picShopItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   2
      Left            =   360
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   12
      Top             =   2280
      Width           =   480
   End
   Begin VB.PictureBox picShopItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   1
      Left            =   3000
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   8
      Top             =   1440
      Width           =   480
   End
   Begin VB.PictureBox picShopItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   0
      Left            =   360
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   2
      Top             =   1440
      Width           =   480
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   6240
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   1
      Top             =   4890
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   0
      Top             =   0
   End
   Begin VB.ListBox lstTrade 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      ItemData        =   "frmTrade.frx":231C
      Left            =   6615
      List            =   "frmTrade.frx":231E
      TabIndex        =   0
      Top             =   345
      Visible         =   0   'False
      Width           =   5445
   End
   Begin VB.Label Deal 
      BackStyle       =   0  'Transparent
      Caption         =   "Deal"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   3000
      TabIndex        =   35
      Top             =   4440
      Width           =   375
   End
   Begin VB.Label Price 
      BackStyle       =   0  'Transparent
      Caption         =   "Price"
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
      Index           =   7
      Left            =   3600
      TabIndex        =   34
      Top             =   4200
      Width           =   1845
   End
   Begin VB.Label ItemS 
      BackStyle       =   0  'Transparent
      Caption         =   "Item"
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
      Index           =   7
      Left            =   3600
      TabIndex        =   33
      Top             =   3960
      Width           =   1845
   End
   Begin VB.Label Deal 
      BackStyle       =   0  'Transparent
      Caption         =   "Deal"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   360
      TabIndex        =   31
      Top             =   4440
      Width           =   375
   End
   Begin VB.Label Price 
      BackStyle       =   0  'Transparent
      Caption         =   "Price"
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
      Index           =   6
      Left            =   960
      TabIndex        =   30
      Top             =   4200
      Width           =   1845
   End
   Begin VB.Label ItemS 
      BackStyle       =   0  'Transparent
      Caption         =   "Item"
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
      Index           =   6
      Left            =   960
      TabIndex        =   29
      Top             =   3960
      Width           =   1845
   End
   Begin VB.Label Deal 
      BackStyle       =   0  'Transparent
      Caption         =   "Deal"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   3000
      TabIndex        =   27
      Top             =   3600
      Width           =   375
   End
   Begin VB.Label Price 
      BackStyle       =   0  'Transparent
      Caption         =   "Price"
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
      Index           =   5
      Left            =   3600
      TabIndex        =   26
      Top             =   3360
      Width           =   1845
   End
   Begin VB.Label ItemS 
      BackStyle       =   0  'Transparent
      Caption         =   "Item"
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
      Index           =   5
      Left            =   3600
      TabIndex        =   25
      Top             =   3120
      Width           =   1845
   End
   Begin VB.Label Deal 
      BackStyle       =   0  'Transparent
      Caption         =   "Deal"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   360
      TabIndex        =   23
      Top             =   3600
      Width           =   375
   End
   Begin VB.Label Price 
      BackStyle       =   0  'Transparent
      Caption         =   "Price"
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
      Index           =   4
      Left            =   960
      TabIndex        =   22
      Top             =   3360
      Width           =   1845
   End
   Begin VB.Label ItemS 
      BackStyle       =   0  'Transparent
      Caption         =   "Item"
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
      Index           =   4
      Left            =   960
      TabIndex        =   21
      Top             =   3120
      Width           =   1845
   End
   Begin VB.Label Deal 
      BackStyle       =   0  'Transparent
      Caption         =   "Deal"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   3000
      TabIndex        =   19
      Top             =   2760
      Width           =   375
   End
   Begin VB.Label Price 
      BackStyle       =   0  'Transparent
      Caption         =   "Price"
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
      Index           =   3
      Left            =   3600
      TabIndex        =   18
      Top             =   2520
      Width           =   1845
   End
   Begin VB.Label ItemS 
      BackStyle       =   0  'Transparent
      Caption         =   "Item"
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
      Index           =   3
      Left            =   3600
      TabIndex        =   17
      Top             =   2280
      Width           =   1845
   End
   Begin VB.Label Deal 
      BackStyle       =   0  'Transparent
      Caption         =   "Deal"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   15
      Top             =   2760
      Width           =   375
   End
   Begin VB.Label Price 
      BackStyle       =   0  'Transparent
      Caption         =   "Price"
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
      Index           =   2
      Left            =   960
      TabIndex        =   14
      Top             =   2520
      Width           =   1845
   End
   Begin VB.Label ItemS 
      BackStyle       =   0  'Transparent
      Caption         =   "Item"
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
      Index           =   2
      Left            =   960
      TabIndex        =   13
      Top             =   2280
      Width           =   1845
   End
   Begin VB.Label Deal 
      BackStyle       =   0  'Transparent
      Caption         =   "Deal"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   3000
      TabIndex        =   11
      Top             =   1920
      Width           =   375
   End
   Begin VB.Label Price 
      BackStyle       =   0  'Transparent
      Caption         =   "Price"
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
      Index           =   1
      Left            =   3600
      TabIndex        =   10
      Top             =   1680
      Width           =   1845
   End
   Begin VB.Label ItemS 
      BackStyle       =   0  'Transparent
      Caption         =   "Item"
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
      Index           =   1
      Left            =   3600
      TabIndex        =   9
      Top             =   1440
      Width           =   1845
   End
   Begin VB.Label picCancel 
      BackStyle       =   0  'Transparent
      Caption         =   "Return To Game"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   4440
      TabIndex        =   7
      Top             =   5400
      Width           =   1215
   End
   Begin VB.Label picFixItems 
      BackStyle       =   0  'Transparent
      Caption         =   "Fix Selected Item"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2280
      TabIndex        =   6
      Top             =   5160
      Width           =   1230
   End
   Begin VB.Label Deal 
      BackStyle       =   0  'Transparent
      Caption         =   "Deal"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   5
      Top             =   1920
      Width           =   375
   End
   Begin VB.Label Price 
      BackStyle       =   0  'Transparent
      Caption         =   "Price"
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
      Index           =   0
      Left            =   960
      TabIndex        =   4
      Top             =   1680
      Width           =   1845
   End
   Begin VB.Label ItemS 
      BackStyle       =   0  'Transparent
      Caption         =   "Item"
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
      Index           =   0
      Left            =   960
      TabIndex        =   3
      Top             =   1440
      Width           =   1845
   End
End
Attribute VB_Name = "frmTrade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Deal_Click(Index As Integer)
    If lstTrade.ListCount > 0 Then
        Call SendData("traderequest" & SEP_CHAR & Index + 1 & SEP_CHAR & END_CHAR)
    End If
End Sub

Private Sub Deal_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    itmDesc.Visible = False
End Sub

Private Sub Form_Load()
    Picture2.Picture = LoadPicture(App.Path & "\GFX\Items.bmp")
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    itmDesc.Visible = False
End Sub

Private Sub ItemS_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    itmDesc.Visible = False
End Sub

Private Sub picFixItems_Click()
Dim i As Long

    For i = 1 To MAX_INV
        If GetPlayerInvItemNum(MyIndex, i) > 0 Then
            frmFixItem.cmbItem.AddItem i & ": " & Trim(Item(GetPlayerInvItemNum(MyIndex, i)).Name)
        Else
            frmFixItem.cmbItem.AddItem "Unused Slot"
        End If
    Next i
    frmFixItem.cmbItem.ListIndex = 0
    frmFixItem.Show vbModal
End Sub

Private Sub picCancel_Click()
    Unload Me
End Sub

Private Sub picShopItem_Click(Index As Integer)
    If ItemGetS(Index + 1) > 0 Then
        lstTrade.ListIndex = Index
    End If
End Sub

Private Sub picShopItem_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index = 0 Or Index = 2 Or Index = 4 Or Index = 6 Then
        itmDesc.Left = 216
    Else
        itmDesc.Left = 8
    End If

    If Trim(Item(ItemGetS(Index + 1)).desc) <> "" And Trim(Item(ItemGetS(Index + 1)).Name) <> "" Then
        itmDesc.Visible = True
    Else
        itmDesc.Visible = False
    End If

    If Item(ItemGetS(Index + 1)).Type = ITEM_TYPE_CURRENCY Then
        descName.Caption = Trim(Item(ItemGetS(Index + 1)).Name) & " (" & ItemGetSS(Index + 1) & ")"
    Else
        descName.Caption = Trim(Item(ItemGetS(Index + 1)).Name)
    End If
    descStr.Caption = Item(ItemGetS(Index + 1)).StrReq & " Strength"
    descDef.Caption = Item(ItemGetS(Index + 1)).DefReq & " Defence"
    descSpeed.Caption = Item(ItemGetS(Index + 1)).SpeedReq & " Speed"
    descHpMp.Caption = "HP: " & Item(ItemGetS(Index + 1)).AddHP & " MP: " & Item(ItemGetS(Index + 1)).AddMP & " SP: " & Item(ItemGetS(Index + 1)).AddSP
    descSD.Caption = "Str: " & Item(ItemGetS(Index + 1)).AddStr & " Def: " & Item(ItemGetS(Index + 1)).AddDef
    descMS.Caption = "Magi: " & Item(ItemGetS(Index + 1)).AddMagi & " Speed: " & Item(ItemGetS(Index + 1)).AddSpeed
    desc.Caption = Trim(Item(ItemGetS(Index + 1)).desc)
End Sub

Private Sub Price_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    itmDesc.Visible = False
End Sub

Private Sub Timer1_Timer()
Dim i As Long

For i = 1 To MAX_TRADES
    If ItemGetS(i) > 0 Then
        ItemS(i - 1).Visible = True
        Price(i - 1).Visible = True
        picShopItem(i - 1).Visible = True
        Deal(i - 1).Visible = True
        ItemS(i - 1).Caption = "Item: " & Trim(Item(ItemGetS(i)).Name) & "  (Qnt " & ItemGetSS(i) & ")"
        Price(i - 1).Caption = "Give: " & Trim(Item(ItemGiveS(i)).Name) & " (Qnt " & ItemGiveSS(i) & ")"
        Call BitBlt(picShopItem(i - 1).hDC, 0, 0, PIC_X, PIC_Y, Picture2.hDC, (Item(ItemGetS(i)).Pic - Int(Item(ItemGetS(i)).Pic / 6) * 6) * PIC_X, Int(Item(ItemGetS(i)).Pic / 6) * PIC_Y, SRCCOPY)
    Else
        ItemS(i - 1).Visible = False
        Price(i - 1).Visible = False
        picShopItem(i - 1).Visible = False
        Deal(i - 1).Visible = False
        ItemS(i - 1).Caption = ""
        Price(i - 1).Caption = ""
        picShopItem(i - 1).Picture = LoadPicture()
    End If
Next i
End Sub
