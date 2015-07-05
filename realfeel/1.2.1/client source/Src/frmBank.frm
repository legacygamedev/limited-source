VERSION 5.00
Begin VB.Form frmBank 
   BorderStyle     =   0  'None
   ClientHeight    =   3150
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6570
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmBank.frx":0000
   ScaleHeight     =   210
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   438
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picWithdraw 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   3300
      Picture         =   "frmBank.frx":437CA
      ScaleHeight     =   255
      ScaleWidth      =   1500
      TabIndex        =   84
      Top             =   300
      Width           =   1530
   End
   Begin VB.PictureBox picDeposit 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   3300
      Picture         =   "frmBank.frx":44BF8
      ScaleHeight     =   255
      ScaleWidth      =   1500
      TabIndex        =   83
      Top             =   30
      Width           =   1530
   End
   Begin VB.PictureBox picInv 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   19
      Left            =   5985
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   42
      Top             =   2520
      Width           =   480
   End
   Begin VB.PictureBox picInv 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   18
      Left            =   5385
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   41
      Top             =   2520
      Width           =   480
   End
   Begin VB.PictureBox picInv 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   17
      Left            =   4785
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   40
      Top             =   2520
      Width           =   480
   End
   Begin VB.PictureBox picInv 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   13
      Left            =   5385
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   39
      Top             =   1920
      Width           =   480
   End
   Begin VB.PictureBox picInv 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   12
      Left            =   4785
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   38
      Top             =   1920
      Width           =   480
   End
   Begin VB.PictureBox picInv 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   16
      Left            =   4185
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   37
      Top             =   2520
      Width           =   480
   End
   Begin VB.PictureBox picInv 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   15
      Left            =   3585
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   36
      Top             =   2520
      Width           =   480
   End
   Begin VB.PictureBox picInv 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   14
      Left            =   5985
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   35
      Top             =   1920
      Width           =   480
   End
   Begin VB.PictureBox picInv 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   11
      Left            =   4185
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   34
      Top             =   1920
      Width           =   480
   End
   Begin VB.PictureBox picInv 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   10
      Left            =   3585
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   33
      Top             =   1920
      Width           =   480
   End
   Begin VB.PictureBox picInv 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   9
      Left            =   5985
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   32
      Top             =   1320
      Width           =   480
   End
   Begin VB.PictureBox picInv 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   8
      Left            =   5385
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   31
      Top             =   1320
      Width           =   480
   End
   Begin VB.PictureBox picInv 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   7
      Left            =   4785
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   30
      Top             =   1320
      Width           =   480
   End
   Begin VB.PictureBox picInv 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   6
      Left            =   4185
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   29
      Top             =   1320
      Width           =   480
   End
   Begin VB.PictureBox picInv 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   5
      Left            =   3585
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   28
      Top             =   1320
      Width           =   480
   End
   Begin VB.PictureBox picInv 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   4
      Left            =   5985
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   27
      Top             =   720
      Width           =   480
   End
   Begin VB.PictureBox picInv 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   3
      Left            =   5385
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   26
      Top             =   720
      Width           =   480
   End
   Begin VB.PictureBox picInv 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   2
      Left            =   4785
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   25
      Top             =   720
      Width           =   480
   End
   Begin VB.PictureBox picInv 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   1
      Left            =   4185
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   24
      Top             =   720
      Width           =   480
   End
   Begin VB.PictureBox picInv 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Index           =   0
      Left            =   3585
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   23
      Top             =   720
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   19
      Left            =   2505
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   22
      Top             =   2520
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   18
      Left            =   1905
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   21
      Top             =   2520
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   17
      Left            =   1305
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   20
      Top             =   2520
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   16
      Left            =   705
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   19
      Top             =   2520
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   15
      Left            =   105
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   18
      Top             =   2520
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   14
      Left            =   2505
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   17
      Top             =   1920
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   13
      Left            =   1905
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   16
      Top             =   1920
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   12
      Left            =   1305
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   15
      Top             =   1920
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   11
      Left            =   705
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   14
      Top             =   1920
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   10
      Left            =   105
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   13
      Top             =   1920
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   9
      Left            =   2505
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   12
      Top             =   1320
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   8
      Left            =   1905
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   11
      Top             =   1320
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   7
      Left            =   1305
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   10
      Top             =   1320
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   6
      Left            =   705
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   9
      Top             =   1320
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   5
      Left            =   105
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   8
      Top             =   1320
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   4
      Left            =   2505
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   7
      Top             =   720
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   3
      Left            =   1905
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   6
      Top             =   720
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   2
      Left            =   1305
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   5
      Top             =   720
      Width           =   480
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   1
      Left            =   705
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   4
      Top             =   720
      Width           =   480
   End
   Begin VB.VScrollBar scrlBank 
      Height          =   2295
      Left            =   3165
      Max             =   4
      TabIndex        =   3
      Top             =   720
      Width           =   255
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   0
      Left            =   105
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   2
      Top             =   720
      Width           =   480
   End
   Begin VB.TextBox txtWithdraw 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4815
      TabIndex        =   1
      Text            =   "Withdraw amount"
      Top             =   300
      Width           =   1725
   End
   Begin VB.TextBox txtDeposit 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4815
      TabIndex        =   0
      Text            =   "Deposit amount"
      Top             =   30
      Width           =   1725
   End
   Begin VB.Shape shpSelected 
      BorderColor     =   &H00FF0000&
      Height          =   510
      Left            =   90
      Top             =   705
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Label lblBValue 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   135
      Index           =   19
      Left            =   2505
      TabIndex        =   82
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label lblBValue 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   135
      Index           =   18
      Left            =   1905
      TabIndex        =   81
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label lblBValue 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   135
      Index           =   17
      Left            =   1305
      TabIndex        =   80
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label lblBValue 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   135
      Index           =   16
      Left            =   705
      TabIndex        =   79
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label lblBValue 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   135
      Index           =   15
      Left            =   105
      TabIndex        =   78
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label lblBValue 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   135
      Index           =   14
      Left            =   2505
      TabIndex        =   77
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label lblBValue 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   135
      Index           =   13
      Left            =   1905
      TabIndex        =   76
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label lblBValue 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   135
      Index           =   12
      Left            =   1305
      TabIndex        =   75
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label lblBValue 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   135
      Index           =   11
      Left            =   705
      TabIndex        =   74
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label lblBValue 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   135
      Index           =   10
      Left            =   105
      TabIndex        =   73
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label lblBValue 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   135
      Index           =   9
      Left            =   2505
      TabIndex        =   72
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label lblBValue 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   135
      Index           =   8
      Left            =   1905
      TabIndex        =   71
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label lblBValue 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   135
      Index           =   7
      Left            =   1305
      TabIndex        =   70
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label lblBValue 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   135
      Index           =   6
      Left            =   705
      TabIndex        =   69
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label lblBValue 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   135
      Index           =   5
      Left            =   105
      TabIndex        =   68
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label lblBValue 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   135
      Index           =   4
      Left            =   2505
      TabIndex        =   67
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label lblBValue 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   135
      Index           =   3
      Left            =   1905
      TabIndex        =   66
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label lblBValue 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   135
      Index           =   2
      Left            =   1305
      TabIndex        =   65
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label lblBValue 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   135
      Index           =   1
      Left            =   705
      TabIndex        =   64
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label lblBValue 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   135
      Index           =   0
      Left            =   105
      TabIndex        =   63
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label lblValue 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   135
      Index           =   19
      Left            =   5985
      TabIndex        =   62
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label lblValue 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   135
      Index           =   18
      Left            =   5385
      TabIndex        =   61
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label lblValue 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   135
      Index           =   17
      Left            =   4785
      TabIndex        =   60
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label lblValue 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   135
      Index           =   16
      Left            =   4185
      TabIndex        =   59
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label lblValue 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   135
      Index           =   15
      Left            =   3585
      TabIndex        =   58
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label lblValue 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   135
      Index           =   14
      Left            =   5985
      TabIndex        =   57
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label lblValue 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   135
      Index           =   13
      Left            =   5385
      TabIndex        =   56
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label lblValue 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   135
      Index           =   12
      Left            =   4785
      TabIndex        =   55
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label lblValue 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   135
      Index           =   11
      Left            =   4185
      TabIndex        =   54
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label lblValue 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   135
      Index           =   10
      Left            =   3585
      TabIndex        =   53
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label lblValue 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   135
      Index           =   9
      Left            =   5985
      TabIndex        =   52
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label lblValue 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   135
      Index           =   8
      Left            =   5385
      TabIndex        =   51
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label lblValue 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   135
      Index           =   7
      Left            =   4785
      TabIndex        =   50
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label lblValue 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   135
      Index           =   6
      Left            =   4185
      TabIndex        =   49
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label lblValue 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   135
      Index           =   5
      Left            =   3585
      TabIndex        =   48
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label lblValue 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   135
      Index           =   4
      Left            =   5985
      TabIndex        =   47
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label lblValue 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   135
      Index           =   3
      Left            =   5385
      TabIndex        =   46
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label lblValue 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   135
      Index           =   2
      Left            =   4785
      TabIndex        =   45
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label lblValue 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   135
      Index           =   1
      Left            =   4185
      TabIndex        =   44
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label lblValue 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   135
      Index           =   0
      Left            =   3585
      TabIndex        =   43
      Top             =   1200
      Width           =   495
   End
End
Attribute VB_Name = "frmBank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub picDeposit_Click()
If Player(MyIndex).Inv(InvSelected).Num > 0 Then
    Call SendData("BANKDEPOSIT" & SEP_CHAR & CStr(Player(MyIndex).Inv(InvSelected).Num) & SEP_CHAR & Trim$(txtDeposit.Text) & SEP_CHAR & CStr(InvSelected) & SEP_CHAR & END_CHAR)
Else
    Call AddText("There is nothing there!", Red)
End If
End Sub

Private Sub picWithdraw_Click()
If Player(MyIndex).BankInv(BankSelected).Num > 0 Then
    Call SendData("BANKWITHDRAW" & SEP_CHAR & CStr(Player(MyIndex).BankInv(BankSelected).Num) & SEP_CHAR & Trim$(txtWithdraw.Text) & SEP_CHAR & CStr(BankSelected) & SEP_CHAR & END_CHAR)
Else
    Call AddText("There is nothing there!", Red)
End If
End Sub

Private Sub Form_Load()
frmBank.txtDeposit.Text = "1"
frmBank.txtWithdraw.Text = "1"
InvSelected = 1
BankSelected = 1
End Sub

Private Sub picBankInv_Click(Index As Integer)
picWithdraw.Enabled = False
txtWithdraw.Enabled = False
If shpBank.Visible = False Then shpBank.Visible = True
shpBank.Left = picInv(Index).Left - 1
shpBank.top = picInv(Index).top - 1
InvSelected = Index + 1
If Player(MyIndex).Inv(InvSelected).Num = 0 Then
    picDeposit.Enabled = False
    txtDeposit.Enabled = False
    Exit Sub
Else
    picDeposit.Enabled = True
    If Item(Player(MyIndex).Inv(InvSelected).Num).Type <> ITEM_TYPE_CURRENCY Then
        txtDeposit.Text = "1"
        txtDeposit.Enabled = False
    Else
        txtDeposit.Enabled = True
    End If
End If
End Sub

Private Sub picBankItem_Click(Index As Integer)
picDeposit.Enabled = False
txtDeposit.Enabled = False
If shpBank.Visible = False Then shpBank.Visible = True
shpBank.Left = picItem(Index).Left - 1
shpBank.top = picItem(Index).top - 1
BankSelected = Index + 1
If Player(MyIndex).BankInv(BankSelected + (5 * scrlBank.Value)).Num = 0 Then
    picWithdraw.Enabled = False
    txtWithdraw.Enabled = False
    Exit Sub
Else
    picWithdraw.Enabled = True
    If Item(Player(MyIndex).BankInv(BankSelected + (5 * scrlBank.Value)).Num).Type <> ITEM_TYPE_CURRENCY Then
        txtWithdraw.Text = "1"
        txtWithdraw.Enabled = False
    Else
        txtWithdraw.Enabled = True
    End If
End If
End Sub

Private Sub scrlBank_Change()
TopBank = scrlBank.Value
BottomBank = TopRow + 3
Call SendData("BANKSCROLL" & SEP_CHAR & TopBank & SEP_CHAR & BottomBank & SEP_CHAR & END_CHAR)
End Sub
