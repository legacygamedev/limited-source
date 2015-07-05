VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Login"
   ClientHeight    =   6000
   ClientLeft      =   120
   ClientTop       =   -45
   ClientWidth     =   6000
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLogin.frx":0000
   ScaleHeight     =   6000
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      BackColor       =   &H00404040&
      Caption         =   "Save Password"
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   1200
      MaskColor       =   &H00404040&
      TabIndex        =   2
      Top             =   2520
      Width           =   195
   End
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
      Left            =   1200
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1920
      Width           =   2355
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
      Left            =   1200
      MaxLength       =   20
      TabIndex        =   0
      Top             =   1440
      Width           =   2355
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Save Password"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1560
      TabIndex        =   7
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00789298&
      BackStyle       =   0  'Transparent
      Caption         =   "Account:"
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   435
      TabIndex        =   6
      Top             =   1560
      Width           =   705
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00789298&
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   360
      TabIndex        =   5
      Top             =   2040
      Width           =   780
   End
   Begin VB.Label picConnect 
      BackColor       =   &H00789298&
      BackStyle       =   0  'Transparent
      Caption         =   "[LOGIN]"
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
      Height          =   225
      Left            =   1200
      TabIndex        =   4
      Top             =   3240
      Width           =   735
   End
   Begin VB.Label picCancel 
      Alignment       =   2  'Center
      BackColor       =   &H00789298&
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
      Height          =   225
      Left            =   3480
      TabIndex        =   3
      Top             =   5520
      Width           =   2370
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
    Next i
End Sub

Private Sub picCancel_Click()
    frmMainMenu.Visible = True
    frmLogin.Visible = False
End Sub

Private Sub picConnect_Click()
    If Trim(txtName.Text) <> "" And Trim(txtPassword.Text) <> "" Then
        If Len(Trim(txtName.Text)) < 3 Or Len(Trim(txtPassword.Text)) < 3 Then
            MsgBox "Your name and password must be at least three characters in length"
            Exit Sub
        End If
        Call MenuState(MENU_STATE_LOGIN)
        Call WriteINI("CONFIG", "Account", txtName.Text, (App.Path & "\config.ini"))
        If Check1.Value = Checked Then
            Call WriteINI("CONFIG", "Password", txtPassword.Text, (App.Path & "\config.ini"))
        Else
            Call WriteINI("CONFIG", "Password", "", (App.Path & "\config.ini"))
        End If
    End If
End Sub

