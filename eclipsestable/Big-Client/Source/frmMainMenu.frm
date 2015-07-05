VERSION 5.00
Begin VB.Form frmMainMenu 
   Appearance      =   0  'Flat
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Main Menu"
   ClientHeight    =   5985
   ClientLeft      =   225
   ClientTop       =   435
   ClientWidth     =   8940
   ControlBox      =   0   'False
   FillColor       =   &H000000FF&
   Icon            =   "frmMainMenu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMainMenu.frx":57E2
   ScaleHeight     =   5985
   ScaleWidth      =   8940
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Status 
      Interval        =   2000
      Left            =   8760
      Top             =   0
   End
   Begin VB.Label LblTotalOnline 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Players Online:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   6360
      TabIndex        =   11
      Top             =   4920
      Width           =   2175
   End
   Begin VB.Label lblWebsite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   435
      Left            =   6000
      TabIndex        =   10
      Top             =   5280
      Width           =   885
   End
   Begin VB.Label lblOnline 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Checking..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3240
      TabIndex        =   9
      Top             =   4440
      Width           =   855
   End
   Begin VB.Label lblVersion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   240
      Width           =   2535
   End
   Begin VB.Label picNews 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Receiving News..."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2775
      Left            =   2040
      TabIndex        =   7
      Top             =   1560
      Width           =   4935
   End
   Begin VB.Label picAutoLogin 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   975
      Left            =   1920
      TabIndex        =   6
      Top             =   4800
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.Label picIpConfig 
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
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   7080
      TabIndex        =   5
      Top             =   5280
      Width           =   1020
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
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   5280
      Width           =   900
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
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   5280
      Width           =   1620
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
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   5280
      Width           =   1740
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
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   4920
      TabIndex        =   1
      Top             =   5280
      Width           =   900
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
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   8160
      TabIndex        =   0
      Top             =   5280
      Width           =   540
   End
End
Attribute VB_Name = "frmMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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
            Ending = ".bmp"
        End If

        If FileExists("GUI\MainMenu" & Ending) Then
            frmMainMenu.Picture = LoadPicture(App.Path & "\GUI\MainMenu" & Ending)
        End If
    Next I

    Ending = ReadINI("CONFIG", "MenuMusic", App.Path & "\config.ini")
    If LenB(Ending) <> 0 Then
        MapSound = BASS_StreamCreateFile(BASSFALSE, StrPtr(App.Path & "\Music\" & Ending), 0, 0, 0)
        Call BASS_ChannelPlay(MapSound, BASSFALSE)
    End If

    Call MainMenuInit
End Sub

Private Sub Form_GotFocus()
    If frmStable.Socket.State = 0 Then
        frmStable.Socket.Connect
    End If
End Sub

Private Sub lblWebsite_Click()
    If Not ReadINI("CONFIG", "WebSite", App.Path & "\config.ini") = "" Then
        Call OpenURL(ReadINI("CONFIG", "WebSite", App.Path & "\config.ini"))
    End If
End Sub

Private Sub picAutoLogin_Click()
    If ConnectToServer = False Or (ConnectToServer = True And AutoLogin = 1 And AllDataReceived) Then
        Call MenuState(MENU_STATE_AUTO_LOGIN)
    End If
End Sub

Private Sub picIpConfig_Click()
    Me.Visible = False
    frmIpconfig.Visible = True
End Sub

Private Sub picNewAccount_Click()
    Me.Visible = False
    frmNewAccount.Visible = True
End Sub

Private Sub picDeleteAccount_Click()
    frmDeleteAccount.Visible = True
    Me.Visible = False
End Sub

Private Sub picLogin_Click()
If ReadINI("CONFIG", "Auto", App.Path & "\config.ini") = 0 Then
    If LenB(frmLogin.txtPassword.Text) <> 0 Then
        frmLogin.Check1.value = Checked
    Else
        frmLogin.Check1.value = Unchecked
    End If
    frmLogin.Visible = True
    Me.Visible = False
Else
        If AllDataReceived Then
        If LenB(frmLogin.txtName.Text) < 6 Then
            Call MsgBox("Your username must be at least three characters in length.")
            Exit Sub
        End If
    
        If LenB(frmLogin.txtPassword.Text) < 6 Then
            Call MsgBox("Your password must be at least three characters in length.")
            Exit Sub
        End If

        Call WriteINI("CONFIG", "Account", frmLogin.txtName.Text, (App.Path & "\config.ini"))
        Call MenuState(MENU_STATE_LOGIN)
        Me.Visible = False
    End If
End If
End Sub

Private Sub picCredits_Click()
    Me.Visible = False
    frmCredits.Visible = True
End Sub

Private Sub picQuit_Click()
    Call GameDestroy
End Sub

Private Sub Status_Timer()
    If ConnectToServer = True Then
        If Not AllDataReceived Then
            Call SendData("givemethemax" & END_CHAR)
        Else
            lblOnline.Caption = "Online"
            lblOnline.ForeColor = vbGreen
        End If
    Else
        picNews.Caption = "Could not connect. The server may be down."

        lblOnline.Caption = "Offline"
        lblOnline.ForeColor = vbRed
    End If
    Call SendData("getonline" & END_CHAR)
End Sub

Private Sub OpenURL(strURL As String) 'Just call the OpenURL sub with the URL of the webpage to open as its parameter.
    ShellExecute Me.hWnd, "open", strURL, vbNullString, "C:\", ByVal 1&
End Sub
