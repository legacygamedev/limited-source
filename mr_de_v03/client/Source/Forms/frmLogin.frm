VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   2130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2505
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2130
   ScaleWidth      =   2505
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   165
      Left            =   2040
      ScaleHeight     =   165
      ScaleWidth      =   165
      TabIndex        =   3
      Top             =   1200
      Width           =   165
      Begin VB.CheckBox chkRemember 
         Appearance      =   0  'Flat
         BackColor       =   &H00577E8F&
         Caption         =   "Check1"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   -15
         TabIndex        =   4
         Top             =   -15
         Width           =   195
      End
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   240
      MaxLength       =   20
      TabIndex        =   1
      Top             =   360
      Width           =   1995
   End
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      IMEMode         =   3  'DISABLE
      Left            =   240
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   840
      Width           =   1995
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   240
      TabIndex        =   8
      Top             =   600
      Width           =   1155
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Username"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   120
      Width           =   1155
   End
   Begin VB.Label lblRemember 
      BackStyle       =   0  'Transparent
      Caption         =   "Remember Me?"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   840
      TabIndex        =   6
      Top             =   1200
      Width           =   1155
   End
   Begin VB.Label picCancel 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Back"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   960
      TabIndex        =   5
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label picConnect 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Login"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   960
      TabIndex        =   0
      Top             =   1440
      Width           =   1335
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Deactivate()
    SaveConfiguration
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMainMenu.Visible = True
    frmLogin.Visible = False
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtPassword.SetFocus
        KeyAscii = 0
    End If
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
         Call picConnect_Click
         KeyAscii = 0
     End If
End Sub

Private Sub picCancel_Click()
    frmMainMenu.Visible = True
    frmLogin.Visible = False
End Sub

Private Sub picConnect_Click()
    If Trim$(txtName.Text) <> vbNullString Then
        If Trim$(txtPassword.Text) <> vbNullString Then
            frmLogin.Visible = False
            If ConnectToServer = True Then
                Call SetStatus("Connected, sending login information...")
                Call SendLogin(frmLogin.txtName.Text, frmLogin.txtPassword.Text)
            End If
        End If
    End If
End Sub
