VERSION 5.00
Begin VB.Form frmNewAccount 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Account"
   ClientHeight    =   9465
   ClientLeft      =   135
   ClientTop       =   375
   ClientWidth     =   12855
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   631
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   857
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picAccCreate 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1935
      Left            =   4440
      ScaleHeight     =   127
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   295
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   6000
      Width           =   4455
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
         Left            =   2160
         MaxLength       =   20
         TabIndex        =   1
         Top             =   120
         Width           =   2085
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
         Left            =   2160
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   480
         Width           =   2085
      End
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
         Left            =   2160
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   840
         Width           =   2085
      End
      Begin VB.Label picCancel 
         BackStyle       =   0  'Transparent
         Caption         =   "Back"
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
         Left            =   3840
         TabIndex        =   8
         Top             =   1560
         Width           =   405
      End
      Begin VB.Label picConnect 
         BackStyle       =   0  'Transparent
         Caption         =   "Create"
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
         Left            =   120
         TabIndex        =   7
         Top             =   1560
         Width           =   495
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
         Left            =   240
         TabIndex        =   6
         Top             =   480
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
         Left            =   240
         TabIndex        =   5
         Top             =   120
         Width           =   1770
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
         Left            =   240
         TabIndex        =   4
         Top             =   840
         Width           =   1590
      End
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

Private Sub picAccCreate_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call MovePicture(frmNewAccount.picAccCreate, Button, Shift, x, y)
End Sub

Private Sub picCancel_Click()
    frmLogin.Visible = True
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
                Call MsgBox("Some characters in your account name are not allowed.", vbOKOnly, GAME_NAME)
                txtName.Text = ""
                Exit Sub
            End If
        Next i
    
        Call MenuState(MENU_STATE_NEWACCOUNT)
    End If
End Sub

Private Sub picAccCreate_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    SOffsetX = x
    SOffsetY = y
End Sub

