VERSION 5.00
Begin VB.Form frmMainMenu 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Powered by Konfuze ORPG Creation - www.Konfuze.com"
   ClientHeight    =   6000
   ClientLeft      =   195
   ClientTop       =   405
   ClientWidth     =   6000
   ControlBox      =   0   'False
   Icon            =   "frmMainMenu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMainMenu.frx":030A
   ScaleHeight     =   6000
   ScaleWidth      =   6000
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label picIpConfig 
      BackStyle       =   0  'Transparent
      Caption         =   "Configure IP"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4920
      TabIndex        =   5
      Top             =   5520
      Width           =   855
   End
   Begin VB.Label picLogin 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   4
      Top             =   1800
      Width           =   555
   End
   Begin VB.Label picNewAccount 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "New Account"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2370
      TabIndex        =   3
      Top             =   2400
      Width           =   1125
   End
   Begin VB.Label picDeleteAccount 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Delete Account"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   2280
      TabIndex        =   2
      Top             =   3000
      Width           =   1305
   End
   Begin VB.Label picCredits 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Credits"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2565
      TabIndex        =   1
      Top             =   3600
      Width           =   690
   End
   Begin VB.Label picQuit 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2670
      TabIndex        =   0
      Top             =   4440
      Width           =   465
   End
End
Attribute VB_Name = "frmMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub picIpConfig_Click()
    frmIpconfig.Visible = True
    Me.Visible = False
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
    frmLogin.txtName.Text = Trim(ReadINI("CONFIG", "Account", App.Path & "\config.ini"))
    frmLogin.txtPassword.Text = Trim(ReadINI("CONFIG", "Password", App.Path & "\config.ini"))
    If Trim(frmLogin.txtPassword.Text) <> "" Then
        frmLogin.Check1.Value = Checked
    Else
        frmLogin.Check1.Value = Unchecked
    End If
    frmLogin.Visible = True
    Me.Visible = False
    frmLogin.txtName.SetFocus
    frmLogin.txtName.SelStart = Len(frmLogin.txtName.Text)
End Sub

Private Sub picCredits_Click()
    frmCredits.Visible = True
    Me.Visible = False
End Sub

Private Sub picQuit_Click()
    Call GameDestroy
End Sub
