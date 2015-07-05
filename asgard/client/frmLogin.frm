VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   9465
   ClientLeft      =   165
   ClientTop       =   390
   ClientWidth     =   12855
   ControlBox      =   0   'False
   DrawMode        =   5  'Not Copy Pen
   FillStyle       =   0  'Solid
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9465
   ScaleWidth      =   12855
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picLoginWin 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      DrawMode        =   1  'Blackness
      DrawStyle       =   5  'Transparent
      ForeColor       =   &H80000008&
      Height          =   1935
      Left            =   4440
      ScaleHeight     =   129
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   297
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
         TabIndex        =   2
         Top             =   600
         Width           =   1995
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
         TabIndex        =   1
         Top             =   1080
         Width           =   1995
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00789298&
         BackStyle       =   0  'Transparent
         Height          =   225
         Left            =   1680
         TabIndex        =   5
         Top             =   1560
         Width           =   1290
      End
      Begin VB.Label picCancel 
         Alignment       =   2  'Center
         BackColor       =   &H00789298&
         BackStyle       =   0  'Transparent
         Height          =   225
         Left            =   3600
         TabIndex        =   4
         Top             =   1560
         Width           =   690
      End
      Begin VB.Label picConnect 
         Alignment       =   2  'Center
         BackColor       =   &H00789298&
         BackStyle       =   0  'Transparent
         Height          =   225
         Left            =   120
         TabIndex        =   3
         Top             =   1560
         Width           =   1425
      End
   End
End
Attribute VB_Name = "frmLogin"
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

        If FileExist("GUI\Login" & Ending) Then frmLogin.Picture = LoadPicture(App.Path & "\GUI\Login" & Ending)
        If FileExist("GUI\LoginWin" & Ending) Then frmLogin.picLoginWin.Picture = LoadPicture(App.Path & "\GUI\LoginWin" & Ending)
    Next i
        frmLogin.txtName.Text = ReadINI("CONFIG", "Account", (App.Path & "\config.ini"))
End Sub

Private Sub Label3_Click()
    frmNewAccount.Visible = True
    frmLogin.Visible = False
End Sub

Private Sub picCancel_Click()
    Call GameDestroy
End Sub

Private Sub picConnect_Click()
    If Trim(txtName.Text) <> "" And Trim(txtPassword.Text) <> "" Then
        If Len(Trim(txtName.Text)) < 3 Or Len(Trim(txtPassword.Text)) < 3 Then
            MsgBox "Your name and password must be at least three characters in length"
            Exit Sub
        End If
        Call MenuState(MENU_STATE_LOGIN)
        Call WriteINI("CONFIG", "Account", txtName.Text, (App.Path & "\config.ini"))
    End If
End Sub

Private Sub Picture1_Click()

End Sub

Private Sub picLoginWin_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    SOffsetX = x
    SOffsetY = y
End Sub

Private Sub picLoginWin_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call MovePicture(frmLogin.picLoginWin, Button, Shift, x, y)
End Sub

