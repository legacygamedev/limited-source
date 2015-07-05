VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   5985
   ClientLeft      =   195
   ClientTop       =   420
   ClientWidth     =   5985
   ControlBox      =   0   'False
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLogin.frx":0FC2
   ScaleHeight     =   399
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   399
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Save Password"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   4440
      TabIndex        =   5
      Top             =   3360
      Width           =   195
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Save Password"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2040
      TabIndex        =   2
      Top             =   3360
      Width           =   195
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
      Height          =   270
      IMEMode         =   3  'DISABLE
      Left            =   1800
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2040
      Width           =   3735
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
      Left            =   1800
      MaxLength       =   20
      TabIndex        =   0
      Top             =   1080
      Width           =   3735
   End
   Begin VB.Label picConnect 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Label picCancel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   2160
      TabIndex        =   3
      Top             =   5280
      Width           =   1650
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check1_Click()
    If Check1.Value = 0 Then
        Check2.Value = 0
    End If
End Sub

Private Sub Check2_Click()
    If Check2.Value = 1 Then
        Check1.Value = 1
    End If
End Sub

Private Sub Form_Load()
    Dim I As Long
    Dim Ending As String

    For I = 1 To 3
        If I = 1 Then
            Ending = ".gif"
        End If
        If I = 2 Then
            Ending = ".jpg"
        End If
        If I = 3 Then
            Ending = ".png"
        End If

        If FileExists("GUI\Login" & Ending) Then
            frmLogin.Picture = LoadPicture(App.Path & "\GUI\Login" & Ending)
        End If
    Next I

    frmLogin.txtName.Text = Trim$(ReadINI("CONFIG", "Account", App.Path & "\config.ini"))
    frmLogin.txtPassword.Text = Trim$(ReadINI("CONFIG", "Password", App.Path & "\config.ini"))

    If AutoLogin = 1 Then
        Check2.Value = Checked
        Check1.Value = Checked
    End If

    If LenB(frmLogin.txtPassword.Text) <> 0 Then
        Check1.Value = Checked
    Else
        Check1.Value = Unchecked
    End If
End Sub

Private Sub picCancel_Click()
    frmMainMenu.Visible = True
    frmLogin.Visible = False
End Sub

Private Sub picConnect_Click()
    If AllDataReceived Then
        If LenB(txtName.Text) < 6 Then
            Call MsgBox("Your username must be at least three characters in length.")
            Exit Sub
        End If
    
        If LenB(txtPassword.Text) < 6 Then
            Call MsgBox("Your password must be at least three characters in length.")
            Exit Sub
        End If

        Call WriteINI("CONFIG", "Account", txtName.Text, (App.Path & "\config.ini"))

        If Check1.Value = Checked Then
            Call WriteINI("CONFIG", "Password", txtPassword.Text, (App.Path & "\config.ini"))
        Else
            Call WriteINI("CONFIG", "Password", vbNullString, (App.Path & "\config.ini"))
        End If

        If Check2.Value = Checked Then
            Call WriteINI("CONFIG", "Auto", 1, (App.Path & "\config.ini"))
        Else
            Call WriteINI("CONFIG", "Auto", 0, (App.Path & "\config.ini"))
        End If

        Call MenuState(MENU_STATE_LOGIN)
    End If
End Sub
