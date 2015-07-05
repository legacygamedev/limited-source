VERSION 5.00
Begin VB.Form frmMainMenu 
   BackColor       =   &H00F5763F&
   BorderStyle     =   0  'None
   ClientHeight    =   3075
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   5880
   ControlBox      =   0   'False
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMainMenu.frx":0000
   ScaleHeight     =   3075
   ScaleWidth      =   5880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H00404080&
      BorderStyle     =   0  'None
      Height          =   2145
      Left            =   3200
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Top             =   280
      Width           =   2640
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1285
      MaxLength       =   20
      TabIndex        =   10
      Top             =   1550
      Width           =   1695
   End
   Begin VB.PictureBox picRefresh 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   4840
      Picture         =   "frmMainMenu.frx":3ADFA
      ScaleHeight     =   375
      ScaleWidth      =   990
      TabIndex        =   9
      Top             =   2590
      Width           =   990
   End
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1285
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   1930
      Width           =   1695
   End
   Begin VB.PictureBox picConnect 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   380
      Left            =   30
      Picture         =   "frmMainMenu.frx":3C1C4
      ScaleHeight     =   375
      ScaleLeft       =   50
      ScaleMode       =   0  'User
      ScaleWidth      =   1500
      TabIndex        =   8
      Top             =   2280
      Width           =   1500
   End
   Begin VB.PictureBox picSettings 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   30
      Picture         =   "frmMainMenu.frx":3DF52
      ScaleHeight     =   375
      ScaleWidth      =   1500
      TabIndex        =   6
      Top             =   2670
      Width           =   1500
   End
   Begin VB.PictureBox picCredits 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   2040
      Picture         =   "frmMainMenu.frx":3FCE0
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   66
      TabIndex        =   5
      Top             =   830
      Width           =   990
   End
   Begin VB.PictureBox picDelAccount 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1030
      Picture         =   "frmMainMenu.frx":410AA
      ScaleHeight     =   375
      ScaleWidth      =   990
      TabIndex        =   4
      Top             =   830
      Width           =   990
   End
   Begin VB.PictureBox picNewAccount 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   30
      Picture         =   "frmMainMenu.frx":42474
      ScaleHeight     =   375
      ScaleWidth      =   990
      TabIndex        =   3
      Top             =   830
      Width           =   990
   End
   Begin VB.PictureBox picQuit 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1550
      Picture         =   "frmMainMenu.frx":4383E
      ScaleHeight     =   25
      ScaleMode       =   0  'User
      ScaleWidth      =   98.01
      TabIndex        =   2
      Top             =   2670
      Width           =   1485
   End
   Begin VB.PictureBox picHeading 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   780
      Left            =   30
      Picture         =   "frmMainMenu.frx":455CC
      ScaleHeight     =   52
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   200
      TabIndex        =   7
      Top             =   30
      Width           =   3000
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BackColor       =   &H0009E7F2&
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   3600
      TabIndex        =   1
      Top             =   2685
      Width           =   735
   End
End
Attribute VB_Name = "frmMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim PicArray() As VB.PictureBox

Public Sub MakePic(ByVal n As Long)
    Set PicArray(n) = Controls.Add("VB.PictureBox", "PicArray" & CStr(n), Me)
    PicArray(n).Appearance = 0
    PicArray(n).BorderStyle = 0
    PicArray(n).AutoRedraw = True
    PicArray(n).AutoSize = True
    PicArray(n).Picture = LoadPicture(App.Path & GUI_PATH & GetVar(App.Path & GUI_PATH & "config.ini", "Main Menu", "PicFrame" & n & "Target"))
    PicArray(n).Left = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Main Menu", "PicFrame" & n & "X"))
    PicArray(n).top = CLng(GetVar(App.Path & GUI_PATH & "config.ini", "Main Menu", "PicFrame" & n & "Y"))
    PicArray(n).Visible = True
End Sub

Public Sub SetArray(ByVal Num As Long)
    ReDim PicArray(1 To Num) As VB.PictureBox
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SOffsetX = X
    SOffsetY = Y
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MoveForm(frmMainMenu, Button, Shift, X, Y)
End Sub

Private Sub picCredits_Click()
    frmCredits.Visible = True
    Me.Visible = False
End Sub

Private Sub picDelAccount_Click()
Dim YesNo As Long

    YesNo = MsgBox("You are on the path for a character deletion, are you sure you want to go through with this?", vbYesNo, GAME_NAME)
    If YesNo = vbYes Then
        frmDeleteAccount.Visible = True
        Me.Visible = False
    End If
End Sub

Private Sub picNewAccount_Click()
    frmNewAccount.Visible = True
    Me.Visible = False
End Sub

Private Sub picQuit_Click()
    Call GameDestroy
End Sub

Private Sub picRefresh_Click()
If ConnectToServer = True Then
    lblStatus.Caption = "Online"
    Call SendData("GETINFO" & SEP_CHAR & END_CHAR)
Else
    lblStatus.Caption = "Offline"
    txtInfo.Text = "Couldn't retrieve data!"
End If
End Sub

Private Sub picSettings_Click()
    frmSettings.Visible = True
    Me.Visible = False
End Sub

Private Sub Form_Load()
If ConnectToServer = True Then
    lblStatus.Caption = "Online"
    Call SendData("GETINFO" & SEP_CHAR & END_CHAR)
Else
    lblStatus.Caption = "Offline"
    txtInfo.Text = "Couldn't retrieve data!"
End If

'See if the account data is being loaded or not
If GetVar(App.Path & "\data.dat", "Account", "Enable") = "1" Then
'Check data to prevent error, load accordingly
If GetVar(App.Path & "\data.dat", "Account", "Name") = "" Then
    txtName.Text = ""
Else
    txtName.Text = GetVar(App.Path & "\data.dat", "Account", "Name")
End If
If GetVar(App.Path & "\data.dat", "Account", "Password") = "" Then
    txtPassword.Text = ""
Else
    txtPassword.Text = GetVar(App.Path & "\data.dat", "Account", "Password")
End If
Else
    Exit Sub
End If

End Sub

Private Sub picConnect_Click()
    If Trim$(txtName.Text) <> "" And Trim$(txtPassword.Text) <> "" Then
        frmMainMenu.Visible = False
        Call MenuState(MENU_STATE_LOGIN)
        Call PutVar(App.Path & "\data.dat", "Account", "Name", txtName.Text)
        Call PutVar(App.Path & "\data.dat", "Account", "Password", txtPassword.Text)
    End If
End Sub

Private Sub picHeading_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SOffsetX = X
    SOffsetY = Y
End Sub

Private Sub picHeading_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MoveForm(frmMainMenu, Button, Shift, X, Y)
End Sub

Private Sub txtInfo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SOffsetX = X
    SOffsetY = Y
End Sub

Private Sub txtInfo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MoveForm(frmMainMenu, Button, Shift, X, Y)
End Sub
