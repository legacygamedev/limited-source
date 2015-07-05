VERSION 5.00
Begin VB.Form frmMainMenu 
   BorderStyle     =   0  'None
   Caption         =   "Main Menu"
   ClientHeight    =   6000
   ClientLeft      =   150
   ClientTop       =   -30
   ClientWidth     =   6000
   ControlBox      =   0   'False
   Icon            =   "frmMainMenu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMainMenu.frx":0ECA
   ScaleHeight     =   6000
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
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
      Height          =   285
      Left            =   3600
      TabIndex        =   5
      Top             =   2640
      Width           =   915
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
      Height          =   255
      Left            =   1560
      TabIndex        =   4
      Top             =   2640
      Width           =   555
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
      Height          =   255
      Left            =   1560
      TabIndex        =   3
      Top             =   2880
      Width           =   1245
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
      Height          =   210
      Left            =   1560
      TabIndex        =   2
      Top             =   3200
      Width           =   1425
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
      Height          =   255
      Left            =   3840
      TabIndex        =   1
      Top             =   2880
      Width           =   690
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
      Height          =   255
      Left            =   3960
      TabIndex        =   0
      Top             =   3120
      Width           =   585
   End
End
Attribute VB_Name = "frmMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Copyright (c) 2007-2008 Elysium Source. All rights reserved.
' This code is licensed under the Elysium General License.

Option Explicit

Private Sub Form_Load()
Dim I As Long
Dim Ending As String
Dim Packet As String
    For I = 1 To 3
        If I = 1 Then Ending = ".gif"
        If I = 2 Then Ending = ".jpg"
        If I = 3 Then Ending = ".png"
        
        If FileExist("GUI\mainmenu" & Ending) Then frmMainMenu.Picture = LoadPicture(App.Path & "\GUI\mainmenu" & Ending)
    Next I
    
    frmLogin.lblPlayers.Visible = True
    frmLogin.lblPlayers.Caption = "Getting info..."
    
    If ConnectToServer = True Then
        Packet = GETINFO_CHAR & END_CHAR
        Call SendData(Packet)
    Else
        frmLogin.lblOnOff.Caption = "Offline"
        frmLogin.lblPlayers.Visible = False
    End If
End Sub

Private Sub picIpConfig_Click()
    frmIpconfig.Visible = True
    Me.Visible = False
End Sub

Private Sub picNewAccount_Click()

    If frmLogin.lblOnOff.Caption <> "Online" Then
        MsgBox "Sorry, but the server seems to be down.", vbOKOnly, GAME_NAME
        Exit Sub
    End If

    frmNewAccount.Visible = True
    Me.Visible = False
End Sub

Private Sub picDeleteAccount_Click()
Dim YesNo As Long

    If frmLogin.lblOnOff.Caption <> "Online" Then
        MsgBox "Sorry, but the server seems to be down.", vbOKOnly, GAME_NAME
        Exit Sub
    End If

    YesNo = MsgBox("You are on the path for a character deletion, are you sure you want to go through with this?", vbYesNo, GAME_NAME)
    If YesNo = vbYes Then
        frmDeleteAccount.Visible = True
        Me.Visible = False
    End If
End Sub

Private Sub picLogin_Click()
Dim Packet As String

    If frmLogin.lblOnOff.Caption <> "Online" Then
        MsgBox "Sorry, but the server seems to be down.", vbOKOnly, GAME_NAME
        Exit Sub
    End If

    frmLogin.txtName.Text = Trim$(GetVar(App.Path & "\config.ini", "CONFIG", "Account"))
    frmLogin.txtPassword.Text = Trim$(GetVar(App.Path & "\config.ini", "CONFIG", "Password"))
    If Trim$(frmLogin.txtPassword.Text) <> vbNullString Then
        frmLogin.Check1.Value = Checked
        frmLogin.txtPassword.PasswordChar = "*"
    Else
        frmLogin.Check1.Value = Unchecked
        frmLogin.txtPassword = "Password"
        frmLogin.txtPassword.PasswordChar = vbNullString
    End If
    If Trim$(frmLogin.txtName.Text) = vbNullString Then
        frmLogin.txtName = "Username"
    End If
    frmLogin.Visible = True
    Me.Visible = False
    frmLogin.txtName.SetFocus
    frmLogin.txtName.SelStart = Len(frmLogin.txtName.Text)
        
    frmLogin.lblPlayers.Visible = True
    frmLogin.lblPlayers.Caption = "Getting info..."
    
    If ConnectToServer = True Then
        frmLogin.tmrInfo.Enabled = True
        Packet = GETINFO_CHAR & END_CHAR
        Call SendData(Packet)
    Else
        frmLogin.lblOnOff.Caption = "Offline"
        frmLogin.lblPlayers.Visible = False
    End If
End Sub

Private Sub picCredits_Click()
    frmCredits.Visible = True
    Me.Visible = False
End Sub

Private Sub picQuit_Click()
    Call GameDestroy
End Sub
