VERSION 5.00
Begin VB.Form frmDeleteAccount 
   BackColor       =   &H000000FF&
   BorderStyle     =   0  'None
   Caption         =   "M:RPGe (Delete Account)"
   ClientHeight    =   6540
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   7680
   ControlBox      =   0   'False
   Icon            =   "frmDeleteAccount.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   7680
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   390
      IMEMode         =   3  'DISABLE
      Left            =   2160
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   4080
      Width           =   3375
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   390
      Left            =   2160
      MaxLength       =   20
      TabIndex        =   0
      Top             =   3120
      Width           =   3375
   End
   Begin VB.Image Image4 
      Height          =   195
      Left            =   7080
      Picture         =   "frmDeleteAccount.frx":0442
      Top             =   120
      Width           =   195
   End
   Begin VB.Image Image3 
      Height          =   195
      Left            =   7320
      Picture         =   "frmDeleteAccount.frx":25B5
      Top             =   120
      Width           =   195
   End
   Begin VB.Image picCancel 
      Height          =   480
      Left            =   480
      Picture         =   "frmDeleteAccount.frx":475B
      Top             =   3360
      Width           =   1320
   End
   Begin VB.Image picConnect 
      Height          =   480
      Left            =   480
      Picture         =   "frmDeleteAccount.frx":4C3B
      Top             =   2760
      Width           =   1320
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H8000000C&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000C&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H8000000C&
      Height          =   465
      Left            =   2130
      Top             =   4050
      Width           =   3435
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000C&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000C&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H8000000C&
      Height          =   465
      Left            =   2130
      Top             =   3085
      Width           =   3435
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   0
      Picture         =   "frmDeleteAccount.frx":510F
      Top             =   0
      Width           =   7680
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2160
      TabIndex        =   3
      Top             =   3720
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2160
      TabIndex        =   2
      Top             =   2760
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   6195
      Left            =   0
      Picture         =   "frmDeleteAccount.frx":7CDE
      Top             =   360
      Width           =   7680
   End
End
Attribute VB_Name = "frmDeleteAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Label2_Click()
End Sub

Private Sub Image4_Click()
frmDeleteAccount.WindowState = vbMinimized
End Sub

Private Sub picCancel_Click()
    frmMainMenu.Visible = True
    frmDeleteAccount.Visible = False
End Sub

Private Sub Image3_Click()
Call GameDestroy
End Sub

Private Sub picConnect_Click()
    If Trim(txtName.text) <> "" And Trim(txtPassword.text) <> "" Then
        Call MenuState(MENU_STATE_DELACCOUNT)
    End If
End Sub

