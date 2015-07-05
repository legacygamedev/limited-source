VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mirage Source Engine (Login)"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8190
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   8190
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkRemember 
      BackColor       =   &H00000000&
      Caption         =   "Remember user info"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   255
      Left            =   4680
      TabIndex        =   2
      Top             =   2880
      Width           =   2295
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4635
      Left            =   0
      Picture         =   "frmLogin.frx":0000
      ScaleHeight     =   4635
      ScaleWidth      =   3015
      TabIndex        =   6
      Top             =   0
      Width           =   3015
   End
   Begin VB.PictureBox picNewAccount 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   825
      Left            =   3240
      Picture         =   "frmLogin.frx":2D88
      ScaleHeight     =   825
      ScaleWidth      =   4800
      TabIndex        =   5
      Top             =   120
      Width           =   4800
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   390
      Left            =   4680
      MaxLength       =   20
      TabIndex        =   0
      Top             =   1800
      Width           =   3375
   End
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   390
      IMEMode         =   3  'DISABLE
      Left            =   4680
      MaxLength       =   20
      PasswordChar    =   "•"
      TabIndex        =   1
      Top             =   2400
      Width           =   3375
   End
   Begin VB.PictureBox picConnect 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   4080
      Picture         =   "frmLogin.frx":46BD
      ScaleHeight     =   510
      ScaleWidth      =   3000
      TabIndex        =   3
      Top             =   3480
      Width           =   3000
   End
   Begin VB.PictureBox picCancel 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   4080
      Picture         =   "frmLogin.frx":53DA
      ScaleHeight     =   510
      ScaleWidth      =   3000
      TabIndex        =   4
      Top             =   3960
      Width           =   3000
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   3360
      TabIndex        =   9
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter a account name and password. "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   3360
      TabIndex        =   8
      Top             =   1080
      Width           =   4575
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   3360
      TabIndex        =   7
      Top             =   2400
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
    If Val(GetVar(App.Path & "\config.ini", "GENERAL", "Remember")) = vbChecked Then
        chkRemember.Value = vbChecked
        txtName = GetVar(App.Path & "\config.ini", "GENERAL", "User")
        txtPassword = GetVar(App.Path & "\config.ini", "GENERAL", "Pass")
    End If
End Sub

Private Sub picCancel_Click()
    frmMainMenu.Visible = True
    frmLogin.Visible = False
End Sub

Private Sub picConnect_Click()
    If Trim$(txtName.Text) <> vbNullString And Trim$(txtPassword.Text) <> vbNullString Then
        Call MenuState(MENU_STATE_LOGIN)
        If chkRemember.Value = vbChecked Then
            Call PutVar(App.Path & "\config.ini", "GENERAL", "User", Trim$(txtName))
            Call PutVar(App.Path & "\config.ini", "GENERAL", "Pass", Trim$(txtPassword))
            Call PutVar(App.Path & "\config.ini", "GENERAL", "Remember", Trim$(chkRemember.Value))
        Else
            Kill App.Path & "\config.ini"
        End If
    End If
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then picConnect_Click
End Sub
