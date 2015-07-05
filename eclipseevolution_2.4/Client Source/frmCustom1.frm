VERSION 5.00
Begin VB.Form frmCustom1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   8505
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   12105
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   567
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   807
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBackground 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8175
      Left            =   120
      ScaleHeight     =   545
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   785
      TabIndex        =   0
      Top             =   120
      Width           =   11775
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   6000
         Top             =   4680
      End
      Begin VB.CommandButton txtcustomOK 
         Caption         =   "OK"
         Height          =   300
         Index           =   4
         Left            =   2160
         TabIndex        =   70
         Top             =   7560
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton txtcustomOK 
         Caption         =   "OK"
         Height          =   300
         Index           =   3
         Left            =   2160
         TabIndex        =   69
         Top             =   7200
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton txtcustomOK 
         Caption         =   "OK"
         Height          =   300
         Index           =   2
         Left            =   2160
         TabIndex        =   68
         Top             =   6840
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton txtcustomOK 
         Caption         =   "OK"
         Height          =   300
         Index           =   1
         Left            =   2160
         TabIndex        =   67
         Top             =   6480
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton txtcustomOK 
         Caption         =   "OK"
         Height          =   300
         Index           =   0
         Left            =   2160
         TabIndex        =   66
         Top             =   6120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.PictureBox picCustom 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   39
         Left            =   1920
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   33
         TabIndex        =   65
         Top             =   5520
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.PictureBox picCustom 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   38
         Left            =   1920
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   33
         TabIndex        =   64
         Top             =   4920
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.PictureBox picCustom 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   37
         Left            =   1920
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   33
         TabIndex        =   63
         Top             =   4320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.PictureBox picCustom 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   36
         Left            =   1920
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   33
         TabIndex        =   62
         Top             =   3720
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.PictureBox picCustom 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   35
         Left            =   1920
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   33
         TabIndex        =   61
         Top             =   3120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.PictureBox picCustom 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   34
         Left            =   1920
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   33
         TabIndex        =   60
         Top             =   2520
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.PictureBox picCustom 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   33
         Left            =   1920
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   33
         TabIndex        =   59
         Top             =   1920
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.PictureBox picCustom 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   32
         Left            =   1920
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   33
         TabIndex        =   58
         Top             =   1320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.PictureBox picCustom 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   31
         Left            =   1920
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   33
         TabIndex        =   57
         Top             =   720
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.PictureBox picCustom 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   30
         Left            =   1920
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   33
         TabIndex        =   56
         Top             =   120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.PictureBox picCustom 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   29
         Left            =   1320
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   33
         TabIndex        =   55
         Top             =   5520
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.PictureBox picCustom 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   28
         Left            =   1320
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   33
         TabIndex        =   54
         Top             =   4920
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.PictureBox picCustom 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   27
         Left            =   1320
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   33
         TabIndex        =   53
         Top             =   4320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.PictureBox picCustom 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   26
         Left            =   1320
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   33
         TabIndex        =   52
         Top             =   3720
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.PictureBox picCustom 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   25
         Left            =   1320
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   33
         TabIndex        =   51
         Top             =   3120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.PictureBox picCustom 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   24
         Left            =   1320
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   33
         TabIndex        =   50
         Top             =   2520
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.PictureBox picCustom 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   23
         Left            =   1320
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   33
         TabIndex        =   49
         Top             =   1920
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.PictureBox picCustom 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   22
         Left            =   1320
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   33
         TabIndex        =   48
         Top             =   1320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.PictureBox picCustom 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   21
         Left            =   1320
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   33
         TabIndex        =   47
         Top             =   720
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.PictureBox picCustom 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   20
         Left            =   1320
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   33
         TabIndex        =   46
         Top             =   120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtCustom 
         Height          =   300
         Index           =   4
         Left            =   120
         TabIndex        =   45
         Text            =   "Text1"
         Top             =   7560
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.TextBox txtCustom 
         Height          =   300
         Index           =   3
         Left            =   120
         TabIndex        =   44
         Text            =   "Text1"
         Top             =   7200
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.TextBox txtCustom 
         Height          =   300
         Index           =   2
         Left            =   120
         TabIndex        =   43
         Text            =   "Text1"
         Top             =   6840
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.TextBox txtCustom 
         Height          =   300
         Index           =   1
         Left            =   120
         TabIndex        =   42
         Text            =   "Text1"
         Top             =   6480
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.TextBox txtCustom 
         Height          =   300
         Index           =   0
         Left            =   120
         TabIndex        =   41
         Text            =   "Text1"
         Top             =   6120
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.PictureBox picCustom 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   19
         Left            =   720
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   33
         TabIndex        =   21
         Top             =   5520
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.PictureBox picCustom 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   18
         Left            =   720
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   33
         TabIndex        =   20
         Top             =   4920
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.PictureBox picCustom 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   17
         Left            =   720
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   33
         TabIndex        =   19
         Top             =   4320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.PictureBox picCustom 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   16
         Left            =   720
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   33
         TabIndex        =   18
         Top             =   3720
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.PictureBox picCustom 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   15
         Left            =   720
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   33
         TabIndex        =   17
         Top             =   3120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.PictureBox picCustom 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   14
         Left            =   720
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   33
         TabIndex        =   16
         Top             =   2520
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.PictureBox picCustom 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   13
         Left            =   720
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   33
         TabIndex        =   15
         Top             =   1920
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.PictureBox picCustom 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   12
         Left            =   720
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   33
         TabIndex        =   14
         Top             =   1320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.PictureBox picCustom 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   11
         Left            =   720
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   33
         TabIndex        =   13
         Top             =   720
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.PictureBox picCustom 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   10
         Left            =   720
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   33
         TabIndex        =   12
         Top             =   120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.PictureBox picCustom 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   9
         Left            =   120
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   33
         TabIndex        =   11
         Top             =   5520
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.PictureBox picCustom 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   8
         Left            =   120
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   33
         TabIndex        =   10
         Top             =   4920
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.PictureBox picCustom 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   7
         Left            =   120
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   33
         TabIndex        =   9
         Top             =   4320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.PictureBox picCustom 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   6
         Left            =   120
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   33
         TabIndex        =   8
         Top             =   3720
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.PictureBox picCustom 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   5
         Left            =   120
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   33
         TabIndex        =   7
         Top             =   3120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.PictureBox picCustom 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   4
         Left            =   120
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   33
         TabIndex        =   6
         Top             =   2520
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.PictureBox picCustom 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   2
         Left            =   120
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   33
         TabIndex        =   5
         Top             =   1920
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.PictureBox picCustom 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   3
         Left            =   120
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   33
         TabIndex        =   3
         Top             =   1320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.PictureBox picCustom 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   1
         Left            =   120
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   33
         TabIndex        =   2
         Top             =   720
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.PictureBox picCustom 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   0
         Left            =   120
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   33
         TabIndex        =   4
         Top             =   120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label BtnCustom 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   19
         Left            =   2760
         TabIndex        =   40
         Top             =   5760
         Visible         =   0   'False
         Width           =   2820
      End
      Begin VB.Label BtnCustom 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   18
         Left            =   2760
         TabIndex        =   39
         Top             =   5520
         Visible         =   0   'False
         Width           =   2820
      End
      Begin VB.Label BtnCustom 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   17
         Left            =   2760
         TabIndex        =   38
         Top             =   5160
         Visible         =   0   'False
         Width           =   2820
      End
      Begin VB.Label BtnCustom 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   16
         Left            =   2760
         TabIndex        =   37
         Top             =   4920
         Visible         =   0   'False
         Width           =   2820
      End
      Begin VB.Label BtnCustom 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   15
         Left            =   2760
         TabIndex        =   36
         Top             =   4560
         Visible         =   0   'False
         Width           =   2820
      End
      Begin VB.Label BtnCustom 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   14
         Left            =   2760
         TabIndex        =   35
         Top             =   4320
         Visible         =   0   'False
         Width           =   2820
      End
      Begin VB.Label BtnCustom 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   13
         Left            =   2760
         TabIndex        =   34
         Top             =   3960
         Visible         =   0   'False
         Width           =   2820
      End
      Begin VB.Label BtnCustom 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   12
         Left            =   2760
         TabIndex        =   33
         Top             =   3720
         Visible         =   0   'False
         Width           =   2820
      End
      Begin VB.Label BtnCustom 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   11
         Left            =   2760
         TabIndex        =   32
         Top             =   3360
         Visible         =   0   'False
         Width           =   2820
      End
      Begin VB.Label BtnCustom 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   10
         Left            =   2760
         TabIndex        =   31
         Top             =   3120
         Visible         =   0   'False
         Width           =   2820
      End
      Begin VB.Label BtnCustom 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   9
         Left            =   2760
         TabIndex        =   30
         Top             =   2760
         Visible         =   0   'False
         Width           =   2820
      End
      Begin VB.Label BtnCustom 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   8
         Left            =   2760
         TabIndex        =   29
         Top             =   2520
         Visible         =   0   'False
         Width           =   2820
      End
      Begin VB.Label BtnCustom 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   7
         Left            =   2760
         TabIndex        =   28
         Top             =   2160
         Visible         =   0   'False
         Width           =   2820
      End
      Begin VB.Label BtnCustom 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   6
         Left            =   2760
         TabIndex        =   27
         Top             =   1920
         Visible         =   0   'False
         Width           =   2820
      End
      Begin VB.Label BtnCustom 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   5
         Left            =   2760
         TabIndex        =   26
         Top             =   1560
         Visible         =   0   'False
         Width           =   2820
      End
      Begin VB.Label BtnCustom 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   2760
         TabIndex        =   25
         Top             =   1320
         Visible         =   0   'False
         Width           =   2820
      End
      Begin VB.Label BtnCustom 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   2760
         TabIndex        =   24
         Top             =   960
         Visible         =   0   'False
         Width           =   2820
      End
      Begin VB.Label BtnCustom 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   2760
         TabIndex        =   23
         Top             =   720
         Visible         =   0   'False
         Width           =   2820
      End
      Begin VB.Label BtnCustom 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   2760
         TabIndex        =   22
         Top             =   360
         Visible         =   0   'False
         Width           =   2820
      End
      Begin VB.Label BtnCustom 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   0
         Left            =   2760
         TabIndex        =   1
         Top             =   120
         Visible         =   0   'False
         Width           =   2820
      End
   End
End
Attribute VB_Name = "frmCustom1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub BtnCustom_Click(Index As Integer)

  Dim packet
  Dim Custom_Type As Long
  Dim custom_string As String

    Custom_Type = 3
    custom_string = " "

    packet = PacketID.CustomMenuClick & SEP_CHAR & MyIndex & SEP_CHAR & Index & SEP_CHAR & CUSTOM_TITLE & SEP_CHAR & Custom_Type & SEP_CHAR & custom_string & SEP_CHAR & END_CHAR
    Call SendData(packet)

End Sub

Private Sub Form_Load()

    Timer1.Enabled = True

End Sub

Private Sub Form_LostFocus()

    If Me.Visible = True Then
        Me.SetFocus
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    If CUSTOM_IS_CLOSABLE = 1 Then
     Else
        Cancel = 1
        frmCustom1.Visible = True
    End If

    Timer1.Enabled = False

End Sub

Private Sub picCustom_Click(Index As Integer)

  Dim packet
  Dim Custom_Type As Long
  Dim custom_string As String

    Custom_Type = 1
    custom_string = " "

    packet = PacketID.CustomMenuClick & SEP_CHAR & MyIndex & SEP_CHAR & Index & SEP_CHAR & CUSTOM_TITLE & SEP_CHAR & Custom_Type & SEP_CHAR & custom_string & SEP_CHAR & END_CHAR
    Call SendData(packet)

End Sub

Private Sub Timer1_Timer()

    If Me.Visible = True Then
        Me.SetFocus
        Call AlwaysOnTop(Me, True)
    End If

End Sub

Private Sub txtcustomOK_Click(Index As Integer)

  Dim packet
  Dim Custom_Type As Long
  Dim custom_string As String

    Custom_Type = 2
    custom_string = frmCustom1.txtCustom(Index).Text

    packet = PacketID.CustomMenuClick & SEP_CHAR & MyIndex & SEP_CHAR & Index & SEP_CHAR & CUSTOM_TITLE & SEP_CHAR & Custom_Type & SEP_CHAR & custom_string & SEP_CHAR & END_CHAR
    Call SendData(packet)

End Sub

