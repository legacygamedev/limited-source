VERSION 5.00
Begin VB.Form frmTrade 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Trade"
   ClientHeight    =   6000
   ClientLeft      =   30
   ClientTop       =   -90
   ClientWidth     =   11250
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
   ScaleWidth      =   750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrItems 
      Interval        =   50
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox picItems 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   10800
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   92
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   66
      Left            =   495
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   88
      Top             =   5250
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   65
      Left            =   7260
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   86
      Top             =   4350
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   64
      Left            =   6585
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   85
      Top             =   4350
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   63
      Left            =   5910
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   84
      Top             =   4350
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   62
      Left            =   5235
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   83
      Top             =   4350
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   61
      Left            =   4560
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   82
      Top             =   4350
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   60
      Left            =   3885
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   81
      Top             =   4350
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   59
      Left            =   3210
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   80
      Top             =   4350
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   58
      Left            =   2535
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   79
      Top             =   4350
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   57
      Left            =   1860
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   78
      Top             =   4350
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   56
      Left            =   1185
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   77
      Top             =   4350
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   55
      Left            =   510
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   76
      Top             =   4350
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   54
      Left            =   7260
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   75
      Top             =   3705
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   53
      Left            =   6585
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   74
      Top             =   3705
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   52
      Left            =   5910
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   73
      Top             =   3705
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   51
      Left            =   5235
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   72
      Top             =   3705
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   50
      Left            =   4560
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   71
      Top             =   3705
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   49
      Left            =   3885
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   70
      Top             =   3705
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   48
      Left            =   3210
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   69
      Top             =   3705
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   47
      Left            =   2535
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   68
      Top             =   3705
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   46
      Left            =   1860
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   67
      Top             =   3705
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   45
      Left            =   1185
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   66
      Top             =   3705
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   44
      Left            =   510
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   65
      Top             =   3705
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   43
      Left            =   7260
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   64
      Top             =   3060
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   42
      Left            =   6585
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   63
      Top             =   3060
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   41
      Left            =   5910
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   62
      Top             =   3060
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   40
      Left            =   5235
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   61
      Top             =   3060
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   39
      Left            =   4560
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   60
      Top             =   3060
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   38
      Left            =   3885
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   59
      Top             =   3060
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   37
      Left            =   3210
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   58
      Top             =   3060
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   36
      Left            =   2535
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   57
      Top             =   3060
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   35
      Left            =   1860
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   56
      Top             =   3060
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   34
      Left            =   1185
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   55
      Top             =   3060
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   33
      Left            =   510
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   54
      Top             =   3060
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   32
      Left            =   7260
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   53
      Top             =   2415
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   31
      Left            =   6585
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   52
      Top             =   2415
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   30
      Left            =   5910
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   51
      Top             =   2415
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   29
      Left            =   5235
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   50
      Top             =   2415
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   28
      Left            =   4560
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   49
      Top             =   2415
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   27
      Left            =   3885
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   48
      Top             =   2415
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   26
      Left            =   3210
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   47
      Top             =   2415
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   25
      Left            =   2535
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   46
      Top             =   2415
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   24
      Left            =   1860
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   45
      Top             =   2415
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   23
      Left            =   1185
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   44
      Top             =   2415
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   22
      Left            =   510
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   43
      Top             =   2415
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   21
      Left            =   7260
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   42
      Top             =   1770
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   20
      Left            =   6585
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   41
      Top             =   1770
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   19
      Left            =   5910
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   40
      Top             =   1770
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   18
      Left            =   5235
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   39
      Top             =   1770
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   17
      Left            =   4560
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   38
      Top             =   1770
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   16
      Left            =   3885
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   37
      Top             =   1770
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   15
      Left            =   3210
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   36
      Top             =   1770
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   14
      Left            =   2535
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   35
      Top             =   1770
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   13
      Left            =   1860
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   34
      Top             =   1770
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   12
      Left            =   1185
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   33
      Top             =   1770
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   11
      Left            =   510
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   32
      Top             =   1770
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   10
      Left            =   7260
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   31
      Top             =   1125
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   9
      Left            =   6585
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   30
      Top             =   1125
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   8
      Left            =   5910
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   29
      Top             =   1125
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   7
      Left            =   5235
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   28
      Top             =   1125
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   6
      Left            =   4560
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   27
      Top             =   1125
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   5
      Left            =   3885
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   26
      Top             =   1125
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   4
      Left            =   3210
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   25
      Top             =   1125
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   3
      Left            =   2535
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   24
      Top             =   1125
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   2
      Left            =   1860
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   23
      Top             =   1125
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   1
      Left            =   1185
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   22
      Top             =   1125
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   0
      Left            =   510
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   2
      Top             =   1125
      Width           =   480
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-Bonus Stats-"
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
      Left            =   8280
      TabIndex        =   95
      Top             =   2640
      Width           =   2535
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
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
      Height          =   255
      Left            =   8280
      TabIndex        =   94
      Top             =   1320
      Width           =   2535
   End
   Begin VB.Label descMagi 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Magic: 0"
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
      Height          =   165
      Left            =   8280
      TabIndex        =   93
      Top             =   2280
      Width           =   2565
   End
   Begin VB.Label descQuantity 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Quantity: 0"
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
      Height          =   165
      Left            =   9180
      TabIndex        =   91
      Top             =   960
      Width           =   735
   End
   Begin VB.Label lblQuantity 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Quantity:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1200
      TabIndex        =   90
      Top             =   5520
      Width           =   1770
   End
   Begin VB.Label lblTradeFor 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Trade for: "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1200
      TabIndex        =   89
      Top             =   5280
      Width           =   1740
   End
   Begin VB.Shape shopType 
      Height          =   255
      Left            =   240
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label lblDeal 
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
      Height          =   315
      Left            =   5640
      TabIndex        =   87
      Top             =   5040
      Width           =   525
   End
   Begin VB.Shape shpSelect 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Height          =   525
      Left            =   495
      Top             =   1110
      Width           =   525
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Height          =   360
      Left            =   4560
      TabIndex        =   21
      Top             =   690
      Width           =   810
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Height          =   360
      Left            =   3795
      TabIndex        =   20
      Top             =   690
      Width           =   780
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Height          =   360
      Left            =   2880
      TabIndex        =   19
      Top             =   690
      Width           =   930
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Height          =   360
      Left            =   2175
      TabIndex        =   18
      Top             =   690
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   360
      Left            =   1290
      TabIndex        =   17
      Top             =   690
      Width           =   900
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   360
      Left            =   360
      TabIndex        =   16
      Top             =   720
      Width           =   930
   End
   Begin VB.Label descName 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      ForeColor       =   &H00E0E0E0&
      Height          =   165
      Left            =   9300
      TabIndex        =   15
      Top             =   720
      Width           =   405
   End
   Begin VB.Label descAExp 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "EXP: 0%"
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
      Height          =   165
      Left            =   9600
      TabIndex        =   14
      Top             =   4080
      Width           =   1170
   End
   Begin VB.Label desc 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
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
      Height          =   975
      Left            =   8280
      TabIndex        =   13
      Top             =   4440
      Width           =   2490
   End
   Begin VB.Label descASpeed 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Speed: 0"
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
      Height          =   165
      Left            =   8370
      TabIndex        =   12
      Top             =   3360
      Width           =   2385
   End
   Begin VB.Label descAMagi 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Magic: 0"
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
      Height          =   165
      Left            =   8340
      TabIndex        =   11
      Top             =   3600
      Width           =   2475
   End
   Begin VB.Label descADef 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Defense: 0"
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
      Height          =   165
      Left            =   8280
      TabIndex        =   10
      Top             =   3120
      Width           =   2475
   End
   Begin VB.Label descAStr 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Strength: 0"
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
      Height          =   165
      Left            =   8280
      TabIndex        =   9
      Top             =   2880
      Width           =   2490
   End
   Begin VB.Label descMp 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "MP: 0"
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
      Height          =   165
      Left            =   9480
      TabIndex        =   8
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label descSp 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "SP: 0"
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
      Height          =   165
      Left            =   8280
      TabIndex        =   7
      Top             =   4080
      Width           =   1305
   End
   Begin VB.Label descHp 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "HP: 0"
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
      Height          =   165
      Left            =   8280
      TabIndex        =   6
      Top             =   3840
      Width           =   1320
   End
   Begin VB.Label descSpeed 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Speed: 0"
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
      Height          =   165
      Left            =   8400
      TabIndex        =   5
      Top             =   2040
      Width           =   2355
   End
   Begin VB.Label descDef 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Defence: 0"
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
      Height          =   165
      Left            =   8280
      TabIndex        =   4
      Top             =   1800
      Width           =   2475
   End
   Begin VB.Label descStr 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Strength: 0"
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
      Height          =   165
      Left            =   8280
      TabIndex        =   3
      Top             =   1560
      Width           =   2490
   End
   Begin VB.Label picCancel 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   7170
      TabIndex        =   1
      Top             =   5040
      Width           =   855
   End
   Begin VB.Label picFixItems 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6180
      TabIndex        =   0
      Top             =   4995
      Width           =   975
   End
End
Attribute VB_Name = "frmTrade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Copyright (c) 2007-2008 Elysium Source. All rights reserved.
' This code is licensed under the Elysium General License.

Private Sub Deal_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    itmDesc.Visible = False
End Sub

Private Sub Form_Load()
Dim I As Long
Dim Ending As String
    For I = 1 To 3
        If I = 1 Then Ending = ".gif"
        If I = 2 Then Ending = ".jpg"
        If I = 3 Then Ending = ".png"
 
        If FileExist("GUI\shop" & Ending) Then frmTrade.Picture = LoadPicture(App.Path & "\GUI\shop" & Ending)
    Next I
    
    picItems.Picture = LoadPicture(App.Path & "\GFX\items.bmp")
End Sub

Private Sub Label1_Click()
Dim xx As Long
For xx = 1 To 6
    Trade(xx).Selected = NO
Next xx

Trade(1).Selected = YES
    
    shopType.Left = Label1.Left
    shopType.Height = Label1.Height
    shopType.Width = Label1.Width
    
    Trade(1).SelectedItem = 1
    Call ItemSelected(1, 1)
End Sub

Private Sub Label2_Click()
Dim xx As Long
For xx = 1 To 6
    Trade(xx).Selected = NO
Next xx

Trade(2).Selected = YES

    shopType.Left = Label2.Left
    shopType.Height = Label2.Height
    shopType.Width = Label2.Width
    
    Trade(2).SelectedItem = 1
    Call ItemSelected(1, 2)
End Sub

Private Sub Label3_Click()
Dim xx As Long
For xx = 1 To 6
    Trade(xx).Selected = NO
Next xx

Trade(3).Selected = YES
    
    shopType.Left = Label3.Left
    shopType.Height = Label3.Height
    shopType.Width = Label3.Width
    Trade(3).SelectedItem = 1
    Call ItemSelected(1, 3)
End Sub

Private Sub Label4_Click()
Dim xx As Long
For xx = 1 To 6
    Trade(xx).Selected = NO
Next xx

Trade(4).Selected = YES

    shopType.Left = Label4.Left
    shopType.Height = Label4.Height
    shopType.Width = Label4.Width
    
    Trade(4).SelectedItem = 1
    Call ItemSelected(1, 4)
End Sub

Private Sub Label5_Click()
Dim xx As Long
For xx = 1 To 6
    Trade(xx).Selected = NO
Next xx

Trade(5).Selected = YES
    
    shopType.Left = Label5.Left
    shopType.Height = Label5.Height
    shopType.Width = Label5.Width
    Trade(5).SelectedItem = 1
    Call ItemSelected(1, 5)
End Sub

Private Sub Label6_Click()
Dim xx As Long
For xx = 1 To 6
    Trade(xx).Selected = NO
Next xx

Trade(6).Selected = YES
    
    shopType.Left = Label6.Left
    shopType.Height = Label6.Height
    shopType.Width = Label6.Width
    Trade(6).SelectedItem = 1
    Call ItemSelected(1, 6)
End Sub

Private Sub lblDeal_Click()
Dim I As Long
Dim Selected As Long

    For I = 1 To 6
        If Trade(I).Selected = YES Then
            Selected = I
            Exit For
        End If
    Next I
    
    If Trade(Selected).Items(Trade(Selected).SelectedItem).ItemGetNum > 0 Then
        Call SendData(TRADEREQUEST_CHAR & SEP_CHAR & Selected & SEP_CHAR & Trade(Selected).SelectedItem & SEP_CHAR & ShopNum & END_CHAR)
    End If
End Sub

Private Sub picFixItems_Click()
Dim I As Long
    For I = 1 To MAX_INV
        If GetPlayerInvItemNum(MyIndex, I) > 0 Then
            frmFixItem.cmbItem.AddItem I & ": " & Trim$(Item(GetPlayerInvItemNum(MyIndex, I)).Name)
        Else
            frmFixItem.cmbItem.AddItem "Unused Slot"
        End If
    Next I
    frmFixItem.cmbItem.ListIndex = 0
    frmFixItem.Show vbModal
End Sub

Private Sub picCancel_Click()
    Unload Me
End Sub

Private Sub picItem_Click(Index As Integer)
Dim I As Long
Dim Selected As Long

If Index = 66 Then Exit Sub

For I = 1 To 6
    If Trade(I).Selected = YES Then
        Selected = I
        Exit For
    End If
Next I

Trade(Selected).SelectedItem = Index + 1

Call ItemSelected(Index + 1, Selected)
End Sub

Private Sub tmrItems_Timer()
On Error Resume Next
Dim I As Long
Dim Selected As Byte
Dim Pic As Long
Selected = 0

    For I = 1 To 6
        If Trade(I).Selected = YES Then
            Selected = I
            Exit For
        End If
    Next I
    
    If Selected = 0 Then Exit Sub
    
    For I = 1 To MAX_TRADES
        Pic = STR(Item(Trade(Selected).Items(I).ItemGetNum).Pic)
        If Trade(Selected).Items(I).ItemGetNum > 0 Then
            Call BitBlt(picItem(I - 1).hDC, 0, 0, PIC_X, PIC_Y, picItems.hDC, (Pic - Int(Pic / 6) * 6) * PIC_X, Int(Pic / 6) * PIC_Y, SRCCOPY)
        Else
            picItem(I - 1).Picture = LoadPicture()
        End If
    Next I

    picItem(66).Picture = LoadPicture()
    Pic = STR(Item(Trade(Selected).Items(Trade(Selected).SelectedItem).ItemGiveNum).Pic)
    If Trade(Selected).Items(Trade(Selected).SelectedItem).ItemGiveNum > 0 Then
        Call BitBlt(picItem(66).hDC, 0, 0, PIC_X, PIC_Y, picItems.hDC, (Pic - Int(Pic / 6) * 6) * PIC_X, Int(Pic / 6) * PIC_Y, SRCCOPY)
    End If
End Sub
