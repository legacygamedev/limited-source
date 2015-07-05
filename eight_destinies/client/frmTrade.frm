VERSION 5.00
Begin VB.Form frmTrade 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   ClientHeight    =   5070
   ClientLeft      =   30
   ClientTop       =   -90
   ClientWidth     =   9270
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
   ScaleHeight     =   338
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   618
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrItems 
      Interval        =   50
      Left            =   8760
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
      Left            =   8760
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   81
      Top             =   0
      Width           =   480
      Visible         =   0   'False
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   66
      Left            =   225
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   77
      Top             =   4455
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   59
      Left            =   6300
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   75
      Top             =   3750
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   58
      Left            =   5625
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   74
      Top             =   3750
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   57
      Left            =   4950
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   73
      Top             =   3750
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   56
      Left            =   4275
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   72
      Top             =   3750
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   55
      Left            =   3600
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   71
      Top             =   3750
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   54
      Left            =   2925
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   70
      Top             =   3750
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   53
      Left            =   2250
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   69
      Top             =   3750
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   52
      Left            =   1575
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   68
      Top             =   3750
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   51
      Left            =   900
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   67
      Top             =   3750
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   50
      Left            =   225
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   66
      Top             =   3750
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   49
      Left            =   6300
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   65
      Top             =   3105
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   48
      Left            =   5625
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   64
      Top             =   3105
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   47
      Left            =   4950
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   63
      Top             =   3105
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   46
      Left            =   4275
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   62
      Top             =   3105
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   45
      Left            =   3600
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   61
      Top             =   3105
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   44
      Left            =   2925
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   60
      Top             =   3105
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   43
      Left            =   2250
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   59
      Top             =   3105
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   42
      Left            =   1575
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   58
      Top             =   3105
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   41
      Left            =   900
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   57
      Top             =   3105
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   40
      Left            =   225
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   56
      Top             =   3105
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   39
      Left            =   6300
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   55
      Top             =   2460
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   38
      Left            =   5625
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   54
      Top             =   2460
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   37
      Left            =   4950
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   53
      Top             =   2460
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   36
      Left            =   4275
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   52
      Top             =   2460
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   35
      Left            =   3600
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   51
      Top             =   2460
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   34
      Left            =   2925
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   50
      Top             =   2460
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   33
      Left            =   2250
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   49
      Top             =   2460
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   32
      Left            =   1575
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   48
      Top             =   2460
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   31
      Left            =   900
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   47
      Top             =   2460
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   30
      Left            =   225
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   46
      Top             =   2460
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   29
      Left            =   6300
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   45
      Top             =   1815
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   28
      Left            =   5625
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   44
      Top             =   1815
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   27
      Left            =   4950
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   43
      Top             =   1815
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   26
      Left            =   4275
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   42
      Top             =   1815
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   25
      Left            =   3600
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   41
      Top             =   1815
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   24
      Left            =   2925
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   40
      Top             =   1815
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   23
      Left            =   2250
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   39
      Top             =   1815
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   22
      Left            =   1575
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   38
      Top             =   1815
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   21
      Left            =   900
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   37
      Top             =   1815
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   20
      Left            =   225
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   36
      Top             =   1815
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   19
      Left            =   6300
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   35
      Top             =   1170
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   18
      Left            =   5625
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   34
      Top             =   1170
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   17
      Left            =   4950
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   33
      Top             =   1170
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   16
      Left            =   4275
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   32
      Top             =   1170
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   15
      Left            =   3600
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   31
      Top             =   1170
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   14
      Left            =   2925
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   30
      Top             =   1170
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   13
      Left            =   2250
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   29
      Top             =   1170
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   12
      Left            =   1575
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   28
      Top             =   1170
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   11
      Left            =   900
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   27
      Top             =   1170
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   10
      Left            =   225
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   26
      Top             =   1170
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   9
      Left            =   6300
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   25
      Top             =   525
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   8
      Left            =   5625
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   24
      Top             =   525
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   7
      Left            =   4950
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   23
      Top             =   525
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   6
      Left            =   4275
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   22
      Top             =   525
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   5
      Left            =   3600
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   21
      Top             =   525
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   4
      Left            =   2925
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   20
      Top             =   525
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   3
      Left            =   2250
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   19
      Top             =   525
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   2
      Left            =   1575
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   18
      Top             =   525
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   1
      Left            =   900
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   17
      Top             =   525
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   0
      Left            =   225
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   2
      Top             =   525
      Width           =   480
   End
   Begin VB.Label descLvl 
      BackStyle       =   0  'Transparent
      Caption         =   "Level"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   7605
      TabIndex        =   83
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label descMagi 
      BackStyle       =   0  'Transparent
      Caption         =   "Magic"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   7605
      TabIndex        =   82
      Top             =   1155
      Width           =   495
   End
   Begin VB.Label descQuantity 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Quantity"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   7830
      TabIndex        =   80
      Top             =   255
      Width           =   630
   End
   Begin VB.Label lblQuantity 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "XXX"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   1575
      TabIndex        =   79
      Top             =   4560
      Width           =   270
   End
   Begin VB.Label lblTradeFor 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "XXX"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   1575
      TabIndex        =   78
      Top             =   4395
      Width           =   270
   End
   Begin VB.Shape shopType 
      Height          =   255
      Left            =   6000
      Top             =   120
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
      Height          =   195
      Left            =   5295
      TabIndex        =   76
      Top             =   4380
      Width           =   525
   End
   Begin VB.Shape shpSelect 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Height          =   525
      Left            =   210
      Top             =   510
      Width           =   525
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   165
      Left            =   9960
      TabIndex        =   16
      Top             =   5040
      Width           =   135
   End
   Begin VB.Label descName 
      AutoSize        =   -1  'True
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
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   7605
      TabIndex        =   15
      Top             =   90
      Width           =   405
   End
   Begin VB.Label descAExp 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "XXX%"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   7680
      TabIndex        =   14
      Top             =   2835
      Width           =   435
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
      ForeColor       =   &H0000FFFF&
      Height          =   1095
      Left            =   7110
      TabIndex        =   13
      Top             =   3210
      Width           =   2025
   End
   Begin VB.Label descASpeed 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "XXXX"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   7605
      TabIndex        =   12
      Top             =   2010
      Width           =   375
   End
   Begin VB.Label descAMagi 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "XXXX"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   7605
      TabIndex        =   11
      Top             =   2175
      Width           =   375
   End
   Begin VB.Label descADef 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "XXXX"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   7605
      TabIndex        =   10
      Top             =   1845
      Width           =   375
   End
   Begin VB.Label descAStr 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "XXXX"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   7605
      TabIndex        =   9
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label descMp 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "XXXX"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   7605
      TabIndex        =   8
      Top             =   2505
      Width           =   360
   End
   Begin VB.Label descSp 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "XXXX"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   7605
      TabIndex        =   7
      Top             =   2670
      Width           =   360
   End
   Begin VB.Label descHp 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "XXXX"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   7605
      TabIndex        =   6
      Top             =   2340
      Width           =   360
   End
   Begin VB.Label descSpeed 
      AutoSize        =   -1  'True
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
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   7605
      TabIndex        =   5
      Top             =   990
      Width           =   450
   End
   Begin VB.Label descDef 
      AutoSize        =   -1  'True
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
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   7605
      TabIndex        =   4
      Top             =   825
      Width           =   600
   End
   Begin VB.Label descStr 
      AutoSize        =   -1  'True
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
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   7605
      TabIndex        =   3
      Top             =   660
      Width           =   630
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
      Height          =   225
      Left            =   6600
      TabIndex        =   1
      Top             =   4380
      Width           =   360
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
      Height          =   210
      Left            =   5925
      TabIndex        =   0
      Top             =   4380
      Width           =   570
   End
End
Attribute VB_Name = "frmTrade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Deal_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    itmDesc.Visible = False
End Sub

Private Sub Form_Load()
Dim i As Long
Dim Ending As String
    For i = 1 To 3
        If i = 1 Then Ending = ".gif"
        If i = 2 Then Ending = ".jpg"
        If i = 3 Then Ending = ".png"
 
        If FileExist("GUI\Shop" & Ending) Then frmTrade.Picture = LoadPicture(App.Path & "\GUI\Shop" & Ending)
    Next i
    
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
Dim i As Long
Dim Selected As Long

For i = 1 To 6
    If Trade(i).Selected = YES Then
        Selected = i
        Exit For
    End If
Next i

    If Trade(Selected).Items(Trade(Selected).SelectedItem).ItemGetNum > 0 Then
        Call SendData("traderequest" & SEP_CHAR & Selected & SEP_CHAR & Trade(Selected).SelectedItem & SEP_CHAR & END_CHAR)
    End If
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

Private Sub picItem_Click(Index As Integer)
Dim i As Long
Dim Selected As Long

If Index = 66 Then Exit Sub

For i = 1 To 6
    If Trade(i).Selected = YES Then
        Selected = i
        Exit For
    End If
Next i

Trade(Selected).SelectedItem = Index + 1

Call ItemSelected(Index + 1, Selected)
End Sub

Private Sub tmrItems_Timer()
On Error Resume Next
Dim i As Long
Dim Selected As Byte
Dim Pic As Long
Selected = 0

    For i = 1 To 6
        If Trade(i).Selected = YES Then
            Selected = i
            Exit For
        End If
    Next i
    
    If Selected = 0 Then Exit Sub
    
    For i = 1 To MAX_TRADES
        Pic = STR(Item(Trade(Selected).Items(i).ItemGetNum).Pic)
        If Trade(Selected).Items(i).ItemGetNum > 0 Then
            Call BitBlt(picItem(i - 1).hDC, 0, 0, PIC_X, PIC_Y, picItems.hDC, (Pic - Int(Pic / 6) * 6) * PIC_X, Int(Pic / 6) * PIC_Y, SRCCOPY)
        Else
            picItem(i - 1).Picture = LoadPicture()
        End If
    Next i

    picItem(66).Picture = LoadPicture()
    Pic = STR(Item(Trade(Selected).Items(Trade(Selected).SelectedItem).ItemGiveNum).Pic)
    If Trade(Selected).Items(Trade(Selected).SelectedItem).ItemGiveNum > 0 Then
        Call BitBlt(picItem(66).hDC, 0, 0, PIC_X, PIC_Y, picItems.hDC, (Pic - Int(Pic / 6) * 6) * PIC_X, Int(Pic / 6) * PIC_Y, SRCCOPY)
    End If
End Sub
