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
      BackColor       =   &H002F3336&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2280
      Left            =   3255
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Top             =   135
      Width           =   2490
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      BackColor       =   &H002F3336&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1305
      MaxLength       =   20
      TabIndex        =   10
      Top             =   1440
      Width           =   1695
   End
   Begin VB.PictureBox picRefresh 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   4840
      Picture         =   "frmMainMenu.frx":3ADFC
      ScaleHeight     =   390
      ScaleWidth      =   1035
      TabIndex        =   9
      Top             =   2685
      Width           =   1035
   End
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      BackColor       =   &H002F3336&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1305
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   1800
      Width           =   1695
   End
   Begin VB.PictureBox picConnect 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   45
      Picture         =   "frmMainMenu.frx":3C35E
      ScaleHeight     =   345
      ScaleLeft       =   50
      ScaleMode       =   0  'User
      ScaleWidth      =   1500
      TabIndex        =   8
      Top             =   2310
      Width           =   1500
   End
   Begin VB.PictureBox picSettings 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   60
      Picture         =   "frmMainMenu.frx":3DE94
      ScaleHeight     =   390
      ScaleWidth      =   1485
      TabIndex        =   6
      Top             =   2640
      Width           =   1485
   End
   Begin VB.PictureBox picCredits 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   1530
      Picture         =   "frmMainMenu.frx":3FD4E
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   104
      TabIndex        =   5
      Top             =   2325
      Width           =   1560
   End
   Begin VB.PictureBox picDelAccount 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   1350
      Picture         =   "frmMainMenu.frx":41728
      ScaleHeight     =   360
      ScaleWidth      =   1770
      TabIndex        =   4
      Top             =   840
      Width           =   1770
   End
   Begin VB.PictureBox picNewAccount 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   0
      Picture         =   "frmMainMenu.frx":438CA
      ScaleHeight     =   345
      ScaleWidth      =   1350
      TabIndex        =   3
      Top             =   840
      Width           =   1350
   End
   Begin VB.PictureBox picQuit 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   1560
      Picture         =   "frmMainMenu.frx":4517C
      ScaleHeight     =   27
      ScaleMode       =   0  'User
      ScaleWidth      =   93.06
      TabIndex        =   2
      Top             =   2640
      Width           =   1410
   End
   Begin VB.PictureBox picHeading 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   780
      Left            =   0
      Picture         =   "frmMainMenu.frx":46FB2
      ScaleHeight     =   52
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   208
      TabIndex        =   7
      Top             =   0
      Width           =   3120
   End
   Begin VB.Timer tmrIntroMusic 
      Interval        =   1000
      Left            =   3660
      Top             =   1560
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BackColor       =   &H0009E7F2&
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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

Public Sub SetArray(ByVal num As Long)
    ReDim PicArray(1 To num) As VB.PictureBox
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
    lblStatus.ForeColor = &H8000
    lblStatus.Caption = "Online"
    Call SendData("GETINFO" & SEP_CHAR & END_CHAR)
Else
    lblStatus.ForeColor = &H9E7F2
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

Private Sub tmrIntroMusic_Timer()
Dim FSys As Object, Folder As Object, FolderFiles As Object, File As Object, FileName As String
Dim MusicList() As String, n As Byte, MusicSelected As String
    ' Check if we need to loop and if a loop is needed
    If LoopIntro = True Then
        Debug.Print DirectShow.GetTimeDur - DirectShow.GetTimePos
        If DirectShow.GetTimeDur - DirectShow.GetTimePos < 1 Then
            If UCase$(GetVar(App.Path & "\data.dat", "IntroMusic", "INTRO_STYLE")) <> "OFF" Then
                Call DirectShow.StopMP3
                Call DirectShow.SetPlayBackBalance(0)
                Call DirectShow.SetPlayBackVolume(0)
    
                If UCase$(GetVar(App.Path & "\data.dat", "IntroMusic", "INTRO_STYLE")) = "FIXED" Then
                    If Not DirectShow.LoadMP3(App.Path & "\" & GetVar(App.Path & "\data.dat", "IntroMusic", "INTRO_POINTER")) Then
                        MsgBox ("Error loading mp3!")
                        Call GameDestroy
                    End If
                ElseIf UCase$(GetVar(App.Path & "\data.dat", "IntroMusic", "INTRO_STYLE")) = "RANDOM" Then
                    ' Create the file
                    Set FSys = CreateObject("Scripting.FileSystemObject")

                    'Set the folder objects
                    Set Folder = FSys.GetFolder(App.Path & MUSIC_PATH)
                    Set FolderFiles = Folder.Files
        
                    n = 1
        
                    For Each File In FolderFiles
                        ' Make sure it is a music file
                        If UCase$(Right$(Mid(File, Len(App.Path & MUSIC_PATH) + 1, (Len(File) - Len(App.Path & MUSIC_PATH))), 3)) = "MP3" Or UCase$(Right$(Mid(File, Len(App.Path & MUSIC_PATH) + 1, (Len(File) - Len(App.Path & MUSIC_PATH))), 3)) = "MID" Or UCase$(Right$(Mid(File, Len(App.Path & MUSIC_PATH) + 1, (Len(File) - Len(App.Path & MUSIC_PATH))), 3)) = "MIDI" Then
                            ReDim Preserve MusicList(1 To n)
                            MusicList(n) = Mid(File, Len(App.Path & MUSIC_PATH) + 1, (Len(File) - Len(App.Path & MUSIC_PATH)))
                            Debug.Print MusicList(n)
                            n = n + 1
                        End If
                    Next File
        
                    'Destroy the folder objects
                    Set File = Nothing
                    Set FolderFiles = Nothing
                    Set Folder = Nothing
                    Set FSys = Nothing
        
                    ' Set up the random seed generator
                    Randomize
        
                    ' Choose a random number
                    ' Find the music
                    MusicSelected = MusicList(Int(Rnd * (n - 1)) + 1)
                    If Not DirectShow.LoadMP3(App.Path & "\" & GetVar(App.Path & "\data.dat", "IntroMusic", "INTRO_POINTER") & MusicSelected) Then
                        MsgBox ("Error loading mp3!")
                        Call GameDestroy
                    End If
                Else
                    If Not DirectShow.LoadMP3(App.Path & "\" & GetVar(App.Path & "\data.dat", "IntroMusic", "INTRO_POINTER")) Then
                        MsgBox ("Error loading mp3!")
                        Call GameDestroy
                    End If
                End If
    
                If UCase$(Trim$(GetVar(App.Path & "\data.dat", "IntroMusic", "INTRO_PLAYMODE"))) = "NORMAL" Then
                    Call DirectShow.SetPlayBackSpeed(1)
                ElseIf UCase$(Trim$(GetVar(App.Path & "\data.dat", "IntroMusic", "INTRO_PLAYMODE"))) = "MODERATELY_SLOW" Then
                    Call DirectShow.SetPlayBackSpeed(0.75)
                ElseIf UCase$(Trim$(GetVar(App.Path & "\data.dat", "IntroMusic", "INTRO_PLAYMODE"))) = "SLOW" Then
                    Call DirectShow.SetPlayBackSpeed(0.5)
                ElseIf UCase$(Trim$(GetVar(App.Path & "\data.dat", "IntroMusic", "INTRO_PLAYMODE"))) = "EXTREMELY_SLOW" Then
                    Call DirectShow.SetPlayBackSpeed(0.25)
                ElseIf UCase$(Trim$(GetVar(App.Path & "\data.dat", "IntroMusic", "INTRO_PLAYMODE"))) = "MAX_SLOW" Then
                    Call DirectShow.SetPlayBackSpeed(0.05)
                ElseIf UCase$(Trim$(GetVar(App.Path & "\data.dat", "IntroMusic", "INTRO_PLAYMODE"))) = "MODERATELY_FAST" Then
                    Call DirectShow.SetPlayBackSpeed(1.25)
                ElseIf UCase$(Trim$(GetVar(App.Path & "\data.dat", "IntroMusic", "INTRO_PLAYMODE"))) = "FAST" Then
                    Call DirectShow.SetPlayBackSpeed(1.5)
                ElseIf UCase$(Trim$(GetVar(App.Path & "\data.dat", "IntroMusic", "INTRO_PLAYMODE"))) = "EXTREMELY_FAST" Then
                    Call DirectShow.SetPlayBackSpeed(1.75)
                ElseIf UCase$(Trim$(GetVar(App.Path & "\data.dat", "IntroMusic", "INTRO_PLAYMODE"))) = "MAX_FAST" Then
                    Call DirectShow.SetPlayBackSpeed(2)
                Else
                    Call DirectShow.SetPlayBackSpeed(1)
                End If
    
                If UCase$(GetVar(App.Path & "\data.dat", "IntroMusic", "INTRO_PLAYLOOP")) = "TRUE" Then
                    LoopIntro = True
                ElseIf UCase$(GetVar(App.Path & "\data.dat", "IntroMusic", "INTRO_PLAYLOOP")) = "False" Then
                    LoopIntro = False
                Else
                    LoopIntro = True
                End If
    
                'Play the mp3
                Call DirectShow.PlayMP3
            Else
                Call DirectShow.StopMP3
                Call DirectShow.SetPlayBackBalance(0)
                Call DirectShow.SetPlayBackVolume(0)
            End If
        End If
    End If
End Sub

Private Sub txtInfo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SOffsetX = X
    SOffsetY = Y
End Sub

Private Sub txtInfo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MoveForm(frmMainMenu, Button, Shift, X, Y)
End Sub
