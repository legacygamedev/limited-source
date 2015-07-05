VERSION 5.00
Object = "{96366485-4AD2-4BC8-AFBF-B1FC132616A5}#2.0#0"; "VBMP.ocx"
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   8940
   StartUpPosition =   2  'CenterScreen
   Begin VBMP.VBMPlayer MenuMusic 
      Height          =   1095
      Left            =   3840
      Top             =   2160
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1931
   End
   Begin VB.Timer Status 
      Interval        =   2000
      Left            =   5280
      Top             =   720
   End
   Begin VB.Label picNews 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Recieving News..."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3375
      Left            =   1920
      TabIndex        =   8
      Top             =   1320
      Width           =   6855
   End
   Begin VB.Label Label1 
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
      Height          =   255
      Left            =   1920
      TabIndex        =   7
      Top             =   5520
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.Label lblss 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Checking Server Status ... "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6480
      TabIndex        =   6
      Top             =   5520
      Width           =   2295
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
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   4200
      Width           =   1500
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
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   1320
      Width           =   1500
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
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   2040
      Width           =   1500
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
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   2760
      Width           =   1500
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
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   3480
      Width           =   1500
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
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   4920
      Width           =   1500
   End
End
Attribute VB_Name = "frmMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_GotFocus()
If frmMirage.Socket.State = 0 Then
frmMirage.Socket.Connect
End If
End Sub

Private Sub Form_Load()
Dim i As Long
Dim Ending As String

    For i = 1 To 3
        If i = 1 Then Ending = ".gif"
        If i = 2 Then Ending = ".jpg"
        If i = 3 Then Ending = ".png"
        
        If FileExist("GUI\MainMenu" & Ending) Then frmMainMenu.Picture = LoadPicture(App.Path & "\GUI\MainMenu" & Ending)
    Next i
        If Not ReadINI("CONFIG", "MenuMusic", App.Path & "\config.ini") = vbNullString Then
            Call MenuMusic.PlayMedia(App.Path & "\music\" & ReadINI("CONFIG", "MenuMusic", App.Path & "\config.ini"), True)
        End If
        Call MainMenuInit
    
End Sub

Private Sub Label1_Click()
    ' This will display the message if it's not connected or connect if it is
    If ConnectToServer = False Or (ConnectToServer = True And AutoLogin = 1 And AllDataReceived) Then
        Call MenuState(MENU_STATE_AUTO_LOGIN)
    End If
End Sub


Private Sub picIpConfig_Click()
    frmIpconfig.Visible = True
    Me.Visible = False
End Sub

Private Sub picNewAccount_Click()
    frmNewAccount.Visible = True
    Me.Visible = False
    Call mail
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
    If Trim$(frmLogin.txtPassword.Text) <> vbNullString Then
        frmLogin.Check1.Value = Checked
    Else
        frmLogin.Check1.Value = Unchecked
    End If
    frmLogin.Visible = True
    Unload Me
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

Private Sub Status_Timer()
Dim packet As String
    
    If ConnectToServer = True Then
        
        If AllDataReceived = False Then
            packet = "givemethemax" & SEP_CHAR & END_CHAR
            Call SendData(packet)
            lblss.Caption = "Getting server data..."
            picNews.Caption = "Receiving News..."
        Else
            lblss.Caption = "Server Status : Online"
        End If
    Else
        picNews.Caption = "Could not connect. The server may be down."
        lblss.Caption = "Server Status : Offline"
        
    End If
    
End Sub

