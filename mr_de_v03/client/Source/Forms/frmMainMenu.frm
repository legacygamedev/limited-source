VERSION 5.00
Begin VB.Form frmMainMenu 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Main Menu"
   ClientHeight    =   1140
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3240
   Icon            =   "frmMainMenu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1140
   ScaleWidth      =   3240
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtFocus 
      Height          =   300
      Left            =   480
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   6495
      Width           =   465
   End
   Begin VB.Label lblOnline 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Realm: Offline (Click to refresh)"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label lblQuit 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Quit"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label lblLogin 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Login"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label lblNewAccount 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "New Account"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label lblOptions 
      BackStyle       =   0  'Transparent
      Caption         =   "options"
      Height          =   225
      Left            =   2850
      TabIndex        =   0
      Top             =   6135
      Width           =   570
   End
End
Attribute VB_Name = "frmMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Unload(Cancel As Integer)
    Call GameDestroy
End Sub

Private Sub Form_Activate()
    CheckConnection
End Sub

Private Sub lblOnline_Click()
    CheckConnection
End Sub

Private Sub lblQuit_Click()
    Call GameDestroy
End Sub

Private Sub CheckConnection()

    lblOnline.Caption = "Realm: Offline(Click To Refresh)"
    lblLogin.Enabled = False
    lblNewAccount.Enabled = False
    
    If ConnectToServer Then
        lblOnline.Caption = "Realm: Online"
        lblLogin.Enabled = True
        lblNewAccount.Enabled = True
    End If
End Sub

Private Sub lblNewAccount_Click()
    frmNewAccount.Visible = True
    frmNewAccount.txtName.Text = vbNullString
    frmNewAccount.txtPassword.Text = vbNullString
    frmNewAccount.txtName.SetFocus
    
    Me.Visible = False
End Sub

Private Sub lblLogin_Click()
    frmLogin.Visible = True
    Me.Visible = False
End Sub

