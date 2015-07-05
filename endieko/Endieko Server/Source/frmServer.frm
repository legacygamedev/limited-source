VERSION 5.00
Begin VB.Form frmServer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Endieko Server"
   ClientHeight    =   2895
   ClientLeft      =   765
   ClientTop       =   975
   ClientWidth     =   8535
   Icon            =   "frmServer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   8535
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrStartMsgQueue 
      Interval        =   100
      Left            =   7440
      Top             =   0
   End
   Begin VB.Timer tmrMain 
      Interval        =   1
      Left            =   7920
      Top             =   0
   End
   Begin VB.TextBox txtText 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   2655
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Text            =   "frmServer.frx":030A
      Top             =   0
      Width           =   8535
   End
   Begin VB.Timer PlayerTimer 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   840
      Top             =   4800
   End
   Begin VB.Timer tmrShutdown 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   1320
      Top             =   4800
   End
   Begin VB.Timer tmrGameAI 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   1800
      Top             =   4800
   End
   Begin VB.Timer tmrSpawnMapItems 
      Interval        =   1000
      Left            =   2280
      Top             =   4800
   End
   Begin VB.Timer tmrPlayerSave 
      Interval        =   60000
      Left            =   360
      Top             =   4800
   End
   Begin VB.FileListBox File1 
      Height          =   1260
      Hidden          =   -1  'True
      Left            =   2400
      System          =   -1  'True
      TabIndex        =   4
      Top             =   720
      Width           =   1455
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label lblPort 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Port: 4000"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7200
      TabIndex        =   6
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label lblIpAddress 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "IP Address: 192.168.1.3"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4800
      TabIndex        =   2
      Top             =   2640
      Width           =   2415
   End
   Begin VB.Label lblStatus 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Status: Online"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2400
      TabIndex        =   1
      Top             =   2640
      Width           =   2415
   End
   Begin VB.Label lblPlayersOnline 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Current Players Online:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   2640
      Width           =   2415
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuShutdown 
         Caption         =   "Shutdown"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuDatabase 
      Caption         =   "&Database"
      Begin VB.Menu mnuReloadClasses 
         Caption         =   "Reload Classes"
      End
      Begin VB.Menu mnuReloadScript 
         Caption         =   "Reload Scripts"
      End
      Begin VB.Menu mnuBreak 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBlank 
         Caption         =   "Temp"
      End
   End
   Begin VB.Menu mnuLog 
      Caption         =   "&Log"
      Begin VB.Menu mnuServerLog 
         Caption         =   "Server Log"
      End
   End
   Begin VB.Menu mnuCommands 
      Caption         =   "&Commands"
      Begin VB.Menu mnuSetOwner 
         Caption         =   "Make Owner"
      End
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TextBuffer As Long
Dim UserInput As String, CurrentP As String
Dim TmpP As String, LastCommand As String
Option Explicit

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim lmsg As Long
    
    lmsg = x / Screen.TwipsPerPixelX
    Select Case lmsg
        Case WM_LBUTTONDBLCLK
            frmServer.WindowState = vbNormal
            frmServer.Show
    End Select
End Sub

Private Sub Form_Resize()
    If frmServer.WindowState = vbMinimized Then
        frmServer.Hide
    End If
End Sub

Private Sub Form_Terminate()
    Call DestroyServer
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call DestroyServer
End Sub

Private Sub mnuReloadScript_Click()
If Scripting = 1 Then
    Set MyScript = Nothing
    Set clsScriptCommands = Nothing
    
    Set MyScript = New clsSadScript
    Set clsScriptCommands = New clsCommands
    MyScript.ReadInCode App.Path & "\Scripts\Main.txt", "Scripts\Main.txt", MyScript.SControl, False
    MyScript.SControl.AddObject "ScriptHardCode", clsScriptCommands, True
    Call TextAdd(frmServer.txtText, "Scripts reloaded.", True)
End If
End Sub

Private Sub mnuServerLog_Click()
    If mnuServerLog.Checked = True Then
        mnuServerLog.Checked = False
        ServerLog = False
    Else
        mnuServerLog.Checked = True
        ServerLog = True
    End If
End Sub


Private Sub mnuSetOwner_Click()
Dim Name As String
Dim i As Long

Name = InputBox("What is the player name?", "Endieko", "Obsidian")
i = FindPlayer(Name)

        ' sloppy... but whatever
        Player(i).Char(Player(i).CharNum).Access = 5
        Call PlayerMsg(i, "You have been deemed an owner. Please Login again for this to take effect.", BrightRed)
End Sub

Private Sub PlayerTimer_Timer()
Dim i As Long

If PlayerI <= MAX_PLAYERS Then
    If IsPlaying(PlayerI) Then
        Call SavePlayer(PlayerI)
        Call PlayerMsg(PlayerI, GetPlayerName(PlayerI) & " is now saved.", Yellow)
    End If
    PlayerI = PlayerI + 1
End If
    
If PlayerI >= MAX_PLAYERS Then
    PlayerI = 1
    PlayerTimer.Enabled = False
    tmrPlayerSave.Enabled = True
End If
End Sub

Private Sub tmrGameAI_Timer()
    Call ServerLogic
End Sub

Private Sub tmrMain_Timer()
    txtText.SelStart = Len(txtText.Text)
End Sub

Private Sub tmrPlayerSave_Timer()
    Call PlayerSaveTimer
End Sub

Private Sub tmrSpawnMapItems_Timer()
    Call CheckSpawnMapItems
End Sub

Private Sub txtChat_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And Trim$(txtChat.Text) <> vbNullString Then
        Call GlobalMsg(txtChat.Text, White)
        Call TextAdd(frmServer.txtText, "Server: " & txtChat.Text, True)
        txtChat.Text = vbNullString
    End If
End Sub

Private Sub tmrShutdown_Timer()
Static Secs As Long

    If Secs <= 0 Then Secs = 30
    Call GlobalMsg("Server Shutdown in " & Secs & " seconds.", BrightBlue)
    Call TextAdd(frmServer.txtText, "Automated Server Shutdown in " & Secs & " seconds.", True)
    Secs = Secs - 2
    If Secs <= 0 Then
        tmrShutdown.Enabled = False
        Call DestroyServer
    End If
End Sub

Private Sub mnuShutdown_Click()
    tmrShutdown.Enabled = True
End Sub

Private Sub mnuExit_Click()
    Call DestroyServer
End Sub

Private Sub mnuReloadClasses_Click()
    Call LoadClasses
    Call TextAdd(frmServer.txtText, "All classes reloaded.", True)
End Sub

' Pre-IOCP
'Private Sub Socket_ConnectionRequest(index As Integer, ByVal requestID As Long)
'    Call AcceptConnection(index, requestID)
'End Sub
'
'Private Sub Socket_Accept(index As Integer, SocketId As Integer)
'    Call AcceptConnection(index, SocketId)
'End Sub
'
'Private Sub Socket_DataArrival(index As Integer, ByVal bytesTotal As Long)
'    If IsConnected(index) Then
'        Call IncomingData(index, bytesTotal)
'    End If
'End Sub
'
'Private Sub Socket_Close(index As Integer)
'    Call CloseSocket(index)
'End Sub

Sub AddR()
    'Sub to add comments on the same line, not used at the moment
    'Note: Don't forget 'TextBuffer = 0' at the completion of each command
    'else the user for exemple will be able to delete 'C:\>'
    txtText.Text = txtText.Text & Chr(13) & Chr(10)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'If the up button is pressed, repeat last command
    'known bug: if you press up twice, the command will come up twice
    If KeyCode = 38 Then
        txtText.Text = txtText.Text & LastCommand
        TextBuffer = Len(LastCommand)
    End If
End Sub

Private Sub tmrStartMsgQueue_Timer()
    Call SendQueuedData
End Sub

Private Sub txtText_KeyDown(KeyCode As Integer, Shift As Integer)
'Prevent other keycodes to be allowed
    KeyCode = 0
End Sub

Private Sub txtText_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = 8 Then GoTo 1 Else GoTo 3
1         If TextBuffer = 0 Then GoTo 2
    'Backspace: User is allowed to delete 1 extra character
        TextBuffer = TextBuffer - 1
        Exit Sub
2         KeyAscii = 0
        Exit Sub
3         If TextBuffer > 64 Then
            KeyAscii = 0
            Exit Sub
            End If
    If KeyAscii = 13 Then GoTo 4
        TextBuffer = TextBuffer + 1
        Exit Sub
4 'User enterd a command
    UserInput = Right(txtText.Text, TextBuffer)
    LastCommand = UserInput
    If Right(UserInput, 1) = ":" Then GoTo ChgDrive
        TextBuffer = 0
        Select Case LCase$(UserInput)
            Case "": Addtext CurrentP & "Endieko>"
            Case Else:
                Select Case LCase$(Left$(UserInput, 3))
                    Case "cls": txtText.Text = CurrentP & "Endieko>"
                    TextBuffer = 0
                    Case "tim": ShowTime
                    Case "dat": ShowDate
                    Case "hel": ShowHelp
                    Case "exi": ServerTerminate
                    Case "shu": ServerTerminate
                    Case "res": ServerRestart
                    Case Else:
                    
                    If Right(CurrentP, 1) = "\" Then
                        EXEFile CurrentP & UserInput
                    Else
                        EXEFile CurrentP & UserInput
                    End If
                End Select
        End Select
        
KeyAscii = 0
Exit Sub
ChgDrive:
KeyAscii = 0
End Sub

Sub ServerRestart()
Addtext "Server Restarted"
Addtext CurrentP & "Endieko>"
TextBuffer = 0
Call mnuResetServer_Click
End Sub
Sub ServerTerminate()
' Closes the server
Addtext "Server shutdown in "
Addtext CurrentP & "Endieko>"
TextBuffer = 0
Call Shutdown_Server
End Sub

Sub ShowTime()
'=> Show the time
Addtext "The time is " & Time
Addtext CurrentP & "Endieko>"
TextBuffer = 0
End Sub

Sub ShowDate()
'=> Show the date
Addtext "The date is " & Date
Addtext CurrentP & "Endieko>"
TextBuffer = 0
End Sub

Sub ShowHelp()
'Show all known commands
Addtext "DOS Commands: cls, /ban Playername, /kick Playername, /shutdown"
Addtext CurrentP & "Endieko>"
TextBuffer = 0
End Sub

Sub EXEFile(FileName As String)
'Execute a file
'Auto-add .exe, .com or .bat
On Error GoTo 255
If FileExists(FileName) = False Then GoTo 1
Shell FileName, vbNormalFocus
Addtext CurrentP & "Endieko>"
Exit Sub
1 If FileExists(FileName & ".exe") = False Then GoTo 2
Shell FileName & ".exe", vbNormalFocus
Addtext CurrentP & "Endieko>"
Exit Sub
2 If FileExists(FileName & ".com") = False Then GoTo 3
Shell FileName & ".com", vbNormalFocus
Addtext CurrentP & "Endieko>"
Exit Sub
3 If FileExists(FileName & ".bat") = False Then GoTo 4
Shell FileName & ".bat", vbNormalFocus
Addtext CurrentP & "Endieko>"
Exit Sub
4 If FileExists(FileName & ".ini") = False Then GoTo 5
Shell FileName & ".ini", vbNormalFocus
Addtext CurrentP & "Endieko>"
Exit Sub
5 Shell FileName & ".txt", vbNormalFocus
Addtext CurrentP & "Endieko>"
Exit Sub
255 Addtext "'" & FileName & "' is not recognizable as an internal or external command, operable program or batch file."
Addtext CurrentP & "Endieko>"
End Sub

Function FileExists(FileName As String)
'Check of a file exists
'Returns: True if exists, false otherwise
File1.Refresh
On Error GoTo 1
Open FileName For Input As #1
FileExists = True
Close #1
Exit Function
1 FileExists = False
Close #1
End Function
