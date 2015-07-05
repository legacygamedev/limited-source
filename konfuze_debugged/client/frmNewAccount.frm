VERSION 5.00
Begin VB.Form frmNewAccount 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "New Account"
   ClientHeight    =   6000
   ClientLeft      =   90
   ClientTop       =   -60
   ClientWidth     =   6000
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmNewAccount.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmNewAccount.frx":0FC2
   ScaleHeight     =   6000
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPassword2 
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
      Height          =   225
      IMEMode         =   3  'DISABLE
      Left            =   360
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   2640
      Width           =   2685
   End
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
      Height          =   225
      IMEMode         =   3  'DISABLE
      Left            =   360
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2130
      Width           =   2685
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
      Height          =   225
      Left            =   360
      MaxLength       =   20
      TabIndex        =   0
      Top             =   1665
      Width           =   2685
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Retype Password:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   360
      TabIndex        =   7
      Top             =   2400
      Width           =   1590
   End
   Begin VB.Label Label1 
      BackColor       =   &H00789298&
      BackStyle       =   0  'Transparent
      Caption         =   "Desired Account Name:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   360
      TabIndex        =   5
      Top             =   1440
      Width           =   1770
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Account Password:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   360
      TabIndex        =   4
      Top             =   1920
      Width           =   1590
   End
   Begin VB.Label picConnect 
      BackStyle       =   0  'Transparent
      Caption         =   "Create New Account"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   360
      TabIndex        =   3
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Label picCancel 
      BackStyle       =   0  'Transparent
      Caption         =   "Back To Main Menu"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2400
      TabIndex        =   2
      Top             =   5400
      Width           =   1485
   End
End
Attribute VB_Name = "frmNewAccount"
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
 
        If FileExist("GUI\NewAccount" & Ending) Then frmNewAccount.Picture = LoadPicture(App.Path & "\GUI\NewAccount" & Ending)
    Next i
End Sub

Private Sub picCancel_Click()
    frmMainMenu.Visible = True
    frmNewAccount.Visible = False
End Sub

Private Sub picConnect_Click()
Dim Msg As String
Dim i As Long
    
    If Trim(txtName.Text) <> "" And Trim(txtPassword.Text) <> "" And Trim(txtPassword2.Text) <> "" Then
        Msg = Trim(txtName.Text)
        
        If Trim(txtPassword.Text) <> Trim(txtPassword2.Text) Then
            MsgBox "Passwords dont match!"
            Exit Sub
        End If
        
        If Len(Trim(txtName.Text)) < 3 Or Len(Trim(txtPassword.Text)) < 3 Then
            MsgBox "Your name and password must be at least three characters in length"
            Exit Sub
        End If
        
        ' Prevent high ascii chars
        For i = 1 To Len(Msg)
            If Asc(Mid(Msg, i, 1)) < 32 Or Asc(Mid(Msg, i, 1)) > 126 Then
                Call MsgBox("You cannot use high ascii chars in your name, please reenter.", vbOKOnly, GAME_NAME)
                txtName.Text = ""
                Exit Sub
            End If
        Next i
    
        Call MenuState(MENU_STATE_NEWACCOUNT)
    End If
End Sub

