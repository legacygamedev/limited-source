VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Login"
   ClientHeight    =   6000
   ClientLeft      =   120
   ClientTop       =   -45
   ClientWidth     =   6000
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLogin.frx":0000
   ScaleHeight     =   6000
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrInfo 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   5520
      Top             =   0
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Save Password"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2280
      TabIndex        =   2
      Top             =   2640
      Width           =   195
   End
   Begin VB.TextBox txtPassword 
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
      Height          =   225
      IMEMode         =   3  'DISABLE
      Left            =   1920
      MaxLength       =   20
      TabIndex        =   1
      Top             =   2280
      Width           =   2355
   End
   Begin VB.TextBox txtName 
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
      Height          =   225
      Left            =   1920
      MaxLength       =   20
      TabIndex        =   0
      Top             =   1920
      Width           =   2355
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Save Password"
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
      Left            =   2640
      TabIndex        =   7
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label lblOnOff 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Offline"
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
      Left            =   3600
      TabIndex        =   6
      Top             =   4200
      Width           =   855
   End
   Begin VB.Label lblPlayers 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Unknown Players Online"
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
      Left            =   2520
      TabIndex        =   5
      Top             =   4440
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label picConnect 
      Alignment       =   2  'Center
      BackColor       =   &H00789298&
      BackStyle       =   0  'Transparent
      Caption         =   "Login "
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
      Height          =   345
      Left            =   2760
      TabIndex        =   4
      Top             =   3600
      Width           =   705
   End
   Begin VB.Label picCancel 
      Alignment       =   2  'Center
      BackColor       =   &H00789298&
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
      ForeColor       =   &H00808080&
      Height          =   345
      Left            =   4320
      TabIndex        =   3
      Top             =   1200
      Width           =   330
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Copyright (c) 2007-2008 Elysium Source. All rights reserved.
' This code is licensed under the Elysium General License.

Option Explicit

Private Sub Form_Load()
Dim I As Long
Dim Ending As String
Dim Packet As String

    For I = 1 To 3
        If I = 1 Then Ending = ".gif"
        If I = 2 Then Ending = ".jpg"
        If I = 3 Then Ending = ".png"

        If FileExist("GUI\medium" & Ending) Then frmLogin.Picture = LoadPicture(App.Path & "\GUI\medium" & Ending)
    Next I
End Sub

Private Sub picCancel_Click()
    frmMainMenu.Visible = True
    frmLogin.Visible = False
End Sub

Private Sub picConnect_Click()
    If Trim$(txtName.Text) <> vbNullString And Trim$(txtPassword.Text) <> vbNullString Then
        If Len(Trim$(txtName.Text)) < 3 Or Len(Trim$(txtPassword.Text)) < 3 Then
            MsgBox "Your name and password must be at least three characters in length"
            Exit Sub
        End If
        Call MenuState(MENU_STATE_LOGIN)
        Call PutVar(App.Path & "\config.ini", "CONFIG", "Account", txtName.Text)
        If Check1.Value = Checked Then
            Call PutVar(App.Path & "\config.ini", "CONFIG", "Password", txtPassword.Text)
        Else
            Call PutVar(App.Path & "\config.ini", "CONFIG", "Password", vbNullString)
        End If
    End If
End Sub

Private Sub tmrInfo_Timer()
    lblOnOff.Caption = "Offline"
    lblPlayers.Visible = False
    tmrInfo.Enabled = False
End Sub

Private Sub txtName_Click()
    If Trim$(txtName.Text) = "Username" Then txtName.Text = vbNullString
End Sub

Private Sub txtPassword_Change()
    frmLogin.txtPassword.PasswordChar = "*"
End Sub

Private Sub txtPassword_Click()
    If Trim$(txtPassword.Text) = "Password" Then txtPassword.Text = vbNullString
    txtPassword.PasswordChar = "*"
End Sub
