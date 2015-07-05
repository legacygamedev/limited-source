VERSION 5.00
Begin VB.Form frmMainMenu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Main Menu"
   ClientHeight    =   5250
   ClientLeft      =   195
   ClientTop       =   405
   ClientWidth     =   4260
   ControlBox      =   0   'False
   Icon            =   "frmMainMenu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMainMenu.frx":000C
   ScaleHeight     =   5250
   ScaleWidth      =   4260
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   3120
      Top             =   4320
   End
   Begin VB.Label lblserverstatus 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   4800
      Width           =   2175
   End
   Begin VB.Label picIpConfig 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   525
      Left            =   960
      TabIndex        =   5
      Top             =   3120
      Width           =   1995
   End
   Begin VB.Label picLogin 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Top             =   1440
      Width           =   1035
   End
   Begin VB.Label picNewAccount 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   960
      TabIndex        =   3
      Top             =   1800
      Width           =   2205
   End
   Begin VB.Label picDeleteAccount 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   840
      TabIndex        =   2
      Top             =   2280
      Width           =   2505
   End
   Begin VB.Label picCredits 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   1200
      TabIndex        =   1
      Top             =   2640
      Width           =   1530
   End
   Begin VB.Label picQuit 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   3600
      Width           =   1065
   End
End
Attribute VB_Name = "frmMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Dim i As Long
Dim Ending As String
If ReadINI("CONFIG", "TitleMusic", App.Path & "\config.ini") = 0 Then
    For i = 1 To 3
        If i = 1 Then Ending = ".gif"
        If i = 2 Then Ending = ".jpg"
        If i = 3 Then Ending = ".png"
        
        If FileExist("GUI\MainMenu" & Ending) Then frmMainMenu.Picture = LoadPicture(App.Path & "\GUI\MainMenu" & Ending)
    Next i
End If
If ReadINI("CONFIG", "TitleMusic", App.Path & "\config.ini") = 1 Then
    For i = 1 To 3
        If i = 1 Then Ending = ".gif"
        If i = 2 Then Ending = ".jpg"
        If i = 3 Then Ending = ".png"
        
        If FileExist("GUI\MainMenu" & Ending) Then frmMainMenu.Picture = LoadPicture(App.Path & "\GUI\MainMenu" & Ending)
    Next i
        Call PlayMidi(ReadINI("CONFIG", "MusicName", App.Path & "\config.ini"))
End If
End Sub
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

Private Sub Timer1_Timer()
    If ConnectToServer = True Then
        lblserverstatus.Caption = "Server Status: Online"
        lblserverstatus.ForeColor = &H4000&
    Else
        lblserverstatus.Caption = "Server Status: Offline"
        lblserverstatus.ForeColor = &H80&
    End If
End Sub
