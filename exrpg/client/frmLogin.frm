VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3750
   ControlBox      =   0   'False
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   3750
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   105
      MaxLength       =   20
      TabIndex        =   0
      Top             =   2010
      Width           =   1875
   End
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   255
      IMEMode         =   3  'DISABLE
      Left            =   105
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2610
      Width           =   1875
   End
   Begin VB.Label picCancel 
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel"
      Height          =   225
      Left            =   1620
      TabIndex        =   6
      Top             =   3675
      Width           =   510
   End
   Begin VB.Label picConnect 
      BackStyle       =   0  'Transparent
      Caption         =   "Proceed into game"
      Height          =   225
      Left            =   1200
      TabIndex        =   5
      Top             =   3285
      Width           =   1365
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   90
      TabIndex        =   4
      Top             =   1770
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter a account name and password.  "
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   75
      TabIndex        =   3
      Top             =   1215
      Width           =   2685
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   90
      TabIndex        =   2
      Top             =   2370
      Width           =   1215
   End
End
Attribute VB_Name = "frmLogin"
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

        If FileExist("\core files\interface\login" & Ending) Then frmLogin.Picture = LoadPicture(App.Path & "\core files\interface\login" & Ending)
    Next i
End Sub

Private Sub picCancel_Click()
    frmMainMenu.Visible = True
    frmLogin.Visible = False
End Sub

Private Sub picConnect_Click()
    If Trim(txtName.Text) <> "" And Trim(txtPassword.Text) <> "" Then
        Call MenuState(MENU_STATE_LOGIN)
    End If
End Sub

