VERSION 5.00
Begin VB.Form frmDeleteAccount 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Delete Account"
   ClientHeight    =   6000
   ClientLeft      =   150
   ClientTop       =   -30
   ClientWidth     =   5985
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmDeleteAccount.frx":0000
   ScaleHeight     =   6000
   ScaleWidth      =   5985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPassword 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   270
      IMEMode         =   3  'DISABLE
      Left            =   450
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2400
      Width           =   2460
   End
   Begin VB.TextBox txtName 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   270
      Left            =   450
      MaxLength       =   20
      TabIndex        =   0
      Top             =   1755
      Width           =   2460
   End
   Begin VB.Label picCancel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "[BACK TO MAIN MENU]"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3600
      TabIndex        =   5
      Top             =   5520
      Width           =   2130
   End
   Begin VB.Label Label1 
      BackColor       =   &H00789298&
      BackStyle       =   0  'Transparent
      Caption         =   "Account Name"
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   480
      TabIndex        =   4
      Top             =   1560
      Width           =   1185
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Account Password"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   480
      TabIndex        =   3
      Top             =   2205
      Width           =   1545
   End
   Begin VB.Label picConnect 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "[DELETE ACCOUNT]"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   2910
      Width           =   1785
   End
End
Attribute VB_Name = "frmDeleteAccount"
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
 
        If FileExist("GUI\DeleteAccount" & Ending) Then frmDeleteAccount.Picture = LoadPicture(App.Path & "\GUI\DeleteAccount" & Ending)
    Next i
End Sub

Private Sub picCancel_Click()
    frmMainMenu.Visible = True
    frmDeleteAccount.Visible = False
End Sub

Private Sub picConnect_Click()
    If Trim(txtName.Text) <> "" And Trim(txtPassword.Text) <> "" Then
        If Len(Trim(txtName.Text)) < 3 Or Len(Trim(txtPassword.Text)) < 3 Then
            MsgBox "Your name and password must be at least three characters in length"
            Exit Sub
        End If
        Call MenuState(MENU_STATE_DELACCOUNT)
    End If
End Sub

