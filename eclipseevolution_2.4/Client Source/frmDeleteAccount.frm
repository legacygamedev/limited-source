VERSION 5.00
Begin VB.Form frmDeleteAccount 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Delete Account"
   ClientHeight    =   5985
   ClientLeft      =   195
   ClientTop       =   345
   ClientWidth     =   5985
   ControlBox      =   0   'False
   Icon            =   "frmDeleteAccount.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   399
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   399
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      IMEMode         =   3  'DISABLE
      Left            =   240
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2640
      Width           =   2805
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   240
      MaxLength       =   20
      TabIndex        =   0
      Top             =   1680
      Width           =   2805
   End
   Begin VB.Label picCancel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   240
      TabIndex        =   5
      Top             =   5520
      Width           =   5565
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   240
      TabIndex        =   4
      Top             =   1320
      Width           =   2805
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   240
      TabIndex        =   3
      Top             =   2280
      Width           =   2805
   End
   Begin VB.Label picConnect 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   240
      TabIndex        =   2
      Top             =   5040
      Width           =   2805
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

    If Trim$(txtName.Text) <> vbNullString And Trim$(txtPassword.Text) <> vbNullString Then
        If Len(Trim$(txtName.Text)) < 3 Or Len(Trim$(txtPassword.Text)) < 3 Then
            MsgBox "Your name and password must be at least three characters in length"
            Exit Sub
        End If

        Call MenuState(MENU_STATE_DELACCOUNT)
    End If

End Sub

