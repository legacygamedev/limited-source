VERSION 5.00
Begin VB.Form frmMainMenu 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Main Menu"
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3750
   ControlBox      =   0   'False
   Icon            =   "frmMainMenu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   3750
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblquit 
      BackStyle       =   0  'Transparent
      Caption         =   "Exit game"
      Height          =   210
      Left            =   1515
      TabIndex        =   3
      Top             =   2955
      Width           =   720
   End
   Begin VB.Label picCredits 
      BackStyle       =   0  'Transparent
      Caption         =   "Credits"
      Height          =   210
      Left            =   1605
      TabIndex        =   2
      Top             =   2175
      Width           =   480
   End
   Begin VB.Label picLogin 
      BackStyle       =   0  'Transparent
      Caption         =   "Login with account"
      Height          =   255
      Left            =   1185
      TabIndex        =   1
      Top             =   1395
      Width           =   1365
   End
   Begin VB.Label picNewAccount 
      BackStyle       =   0  'Transparent
      Caption         =   "New account"
      Height          =   225
      Left            =   1380
      TabIndex        =   0
      Top             =   1785
      Width           =   960
   End
End
Attribute VB_Name = "frmMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Load()
Dim i As Long
Dim Ending As String
    For i = 1 To 3
        If i = 1 Then Ending = ".gif"
        If i = 2 Then Ending = ".jpg"
        If i = 3 Then Ending = ".png"

        If FileExist("\core files\interface\mainmenu" & Ending) Then frmMainMenu.Picture = LoadPicture(App.Path & "\core files\interface\mainmenu" & Ending)
    Next i
End Sub

Private Sub lblquit_Click()
    Call GameDestroy
End Sub

Private Sub picNewAccount_Click()
    frmNewAccount.Visible = True
    Me.Visible = False
End Sub

Private Sub picLogin_Click()
    frmLogin.Visible = True
    Me.Visible = False
End Sub

Private Sub picCredits_Click()
    frmCredits.Visible = True
    Me.Visible = False
End Sub

