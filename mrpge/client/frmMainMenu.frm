VERSION 5.00
Begin VB.Form frmMainMenu 
   BackColor       =   &H000000FF&
   BorderStyle     =   0  'None
   Caption         =   "M:RPGe"
   ClientHeight    =   6540
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   7665
   ControlBox      =   0   'False
   Icon            =   "frmMainMenu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   7665
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer mp3Timer 
      Interval        =   10
      Left            =   405
      Top             =   720
   End
   Begin VB.Image imgMinimize 
      Height          =   195
      Left            =   7080
      Picture         =   "frmMainMenu.frx":0442
      Top             =   120
      Width           =   195
   End
   Begin VB.Image imgClose 
      Height          =   195
      Left            =   7320
      Picture         =   "frmMainMenu.frx":25B5
      Top             =   120
      Width           =   195
   End
   Begin VB.Image picQuit 
      Height          =   480
      Left            =   5520
      Picture         =   "frmMainMenu.frx":475B
      Top             =   4440
      Width           =   1320
   End
   Begin VB.Image picLogin 
      Height          =   480
      Left            =   3960
      Picture         =   "frmMainMenu.frx":4C3B
      Top             =   4440
      Width           =   1320
   End
   Begin VB.Image picDeleteAccount 
      Height          =   480
      Left            =   2400
      Picture         =   "frmMainMenu.frx":50E8
      Top             =   4440
      Width           =   1320
   End
   Begin VB.Image picNewAccount 
      Height          =   480
      Left            =   840
      Picture         =   "frmMainMenu.frx":55CB
      Top             =   4440
      Width           =   1320
   End
   Begin VB.Image imgBackground 
      Height          =   6195
      Left            =   0
      Picture         =   "frmMainMenu.frx":5AC5
      Top             =   360
      Width           =   7680
   End
   Begin VB.Image imgTitle 
      Height          =   375
      Left            =   0
      Picture         =   "frmMainMenu.frx":12740
      Top             =   0
      Width           =   7680
   End
End
Attribute VB_Name = "frmMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub imgClose_Click()
    Call GameDestroy
End Sub

Private Sub imgMinimize_Click()
 frmMainMenu.WindowState = vbMinimized
End Sub

Private Sub mp3Timer_Timer()
If notInGame = True Then
If MP3.MP3Playing = False Then
    PlayMP3 ("splash")
End If
End If
End Sub

Private Sub picNewAccount_Click()
    frmNewAccount.Visible = True
    Me.Visible = False
End Sub

Private Sub picDeleteAccount_Click()
Dim YesNo As Long

    YesNo = MsgBox("You are on the path for a character deletion, are you sure you want to go through with this?", vbYesNo, GAME_NAME)
    If YesNo = vbYes Then
        frmDeleteAccount.Visible = True
        Me.Visible = False
    End If
End Sub

Private Sub picLogin_Click()
    frmLogin.Visible = True
    Me.Visible = False
End Sub
Private Sub picQuit_Click()
    Call GameDestroy
End Sub
