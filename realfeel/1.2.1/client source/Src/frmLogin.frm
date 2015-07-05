VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00004080&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2895
   ClientLeft      =   15
   ClientTop       =   -90
   ClientWidth     =   3000
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLogin.frx":0000
   ScaleHeight     =   2895
   ScaleWidth      =   3000
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picCancel 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1505
      Picture         =   "frmLogin.frx":1C49A
      ScaleHeight     =   375
      ScaleWidth      =   1500
      TabIndex        =   4
      Top             =   2520
      Width           =   1500
   End
   Begin VB.PictureBox picConnect 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      Picture         =   "frmLogin.frx":1E228
      ScaleHeight     =   375
      ScaleWidth      =   1500
      TabIndex        =   3
      Top             =   2520
      Width           =   1500
   End
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      BackColor       =   &H00F5763F&
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
      ForeColor       =   &H000AD3CE&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1200
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1800
      Width           =   1695
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      BackColor       =   &H00F5763F&
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
      ForeColor       =   &H000AD3CE&
      Height          =   285
      Left            =   1200
      MaxLength       =   20
      TabIndex        =   0
      Top             =   1425
      Width           =   1695
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   780
      Left            =   0
      Picture         =   "frmLogin.frx":1FFB6
      ScaleHeight     =   780
      ScaleWidth      =   3000
      TabIndex        =   2
      Top             =   0
      Width           =   3000
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00F5763F&
      Height          =   255
      Left            =   60
      TabIndex        =   7
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Login to an Existing account"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00F5763F&
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   840
      Width           =   2415
   End
   Begin VB.Label Label1 
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
      ForeColor       =   &H00F5763F&
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1440
      Width           =   735
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub picCancel_Click()
    Call LoadMenu
    frmLogin.Visible = False
End Sub

Private Sub picConnect_Click()
    If Trim$(txtName.Text) <> "" And Trim$(txtPassword.Text) <> "" Then
        Call MenuState(MENU_STATE_LOGIN)
        Call PutVar(App.Path & "\data.dat", "Account", "Name", txtName.Text)
        Call PutVar(App.Path & "\data.dat", "Account", "Password", txtPassword.Text)
    End If
End Sub

Private Sub Form_Load()
'See if the account data is being loaded or not
If GetVar(App.Path & "\data.dat", "Account", "Enable") = "1" Then
'Check data to prevent error, load accordingly
If GetVar(App.Path & "\data.dat", "Account", "Name") = "" Then
    txtName.Text = ""
Else
    txtName.Text = GetVar(App.Path & "\data.dat", "Account", "Name")
End If
If GetVar(App.Path & "\data.dat", "Account", "Password") = "" Then
    txtPassword.Text = ""
Else
    txtPassword.Text = GetVar(App.Path & "\data.dat", "Account", "Password")
End If
Else
    Exit Sub
End If

End Sub
