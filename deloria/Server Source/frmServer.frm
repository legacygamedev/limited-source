VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmServer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Konfuze Server"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   10575
   Icon            =   "frmServer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   10575
   StartUpPosition =   2  'CenterScreen
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
      Interval        =   500
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
   Begin MSWinsockLib.Winsock Socket 
      Index           =   0
      Left            =   2760
      Top             =   4800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   7858
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   370
      TabMaxWidth     =   1764
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Server"
      TabPicture(0)   =   "frmServer.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblRainIntensity"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label5"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame3"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "scrlRainIntensity"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "chkMod"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Command1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "Error Log"
      TabPicture(1)   =   "frmServer.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtErrorLog"
      Tab(1).ControlCount=   1
      Begin VB.CommandButton Command1 
         Caption         =   "Mass Kick"
         Height          =   255
         Left            =   9120
         TabIndex        =   14
         Top             =   4080
         Width           =   975
      End
      Begin VB.CheckBox chkMod 
         Caption         =   "Moderators Only"
         Height          =   255
         Left            =   7560
         TabIndex        =   13
         Top             =   4080
         Width           =   1455
      End
      Begin VB.HScrollBar scrlRainIntensity 
         Height          =   255
         Left            =   960
         Max             =   50
         Min             =   1
         TabIndex        =   9
         Top             =   4080
         Value           =   25
         Width           =   2895
      End
      Begin VB.TextBox txtErrorLog 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3570
         Left            =   -74880
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   360
         Width           =   10035
      End
      Begin VB.Frame Frame3 
         Caption         =   "Accounts In Use"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   7560
         TabIndex        =   3
         Top             =   2400
         Width           =   2535
         Begin VB.ListBox LstAccounts 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1215
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Width           =   2220
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Current Online Players"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   7560
         TabIndex        =   2
         Top             =   480
         Width           =   2535
         Begin VB.ListBox LstPlayers 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1380
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   2250
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Active Server Window / Broadcast"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3495
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   7215
         Begin VB.TextBox txtChat 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   3120
            Width           =   6915
         End
         Begin VB.TextBox txtText 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2850
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   4
            Top             =   240
            Width           =   6915
         End
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Current Weather: None"
         Height          =   195
         Left            =   5040
         TabIndex        =   12
         Top             =   4080
         Width           =   1710
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Weather"
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   4080
         Width           =   615
      End
      Begin VB.Label lblRainIntensity 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Intensity: 25"
         Height          =   195
         Left            =   3960
         TabIndex        =   10
         Top             =   4080
         Width           =   930
      End
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
   End
   Begin VB.Menu mnuLog 
      Caption         =   "&Log"
      Begin VB.Menu mnuServerLog 
         Caption         =   "Server Log"
      End
   End
   Begin VB.Menu mnuScript 
      Caption         =   "&Script"
      Begin VB.Menu mnuReloadScript 
         Caption         =   "Reload"
      End
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim i As Long
    For i = 1 To MAX_PLAYERS
        If GetPlayerAccess(i) <= 0 Then
            Call AlertMsg(i, "You have been kicked by the server!")
        End If
    Next i
End Sub

Private Sub scrlRainIntensity_Change()
    lblRainIntensity.Caption = "Intensity: " & Val(scrlRainIntensity.Value)
    RainIntensity = scrlRainIntensity.Value
    Call SendWeatherToAll
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo ErrorHandler
Dim lmsg As Long
    
    lmsg = x / Screen.TwipsPerPixelX
    Select Case lmsg
        Case WM_LBUTTONDBLCLK
            frmServer.WindowState = vbNormal
            frmServer.Show
    End Select
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "frmServer.frm", "Form_MouseMove", Err.Number, Err.Description
End Sub

Private Sub Form_Resize()
On Error GoTo ErrorHandler
    If frmServer.WindowState = vbMinimized Then
        frmServer.Hide
    End If
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
End Sub

Private Sub Form_Terminate()
On Error GoTo ErrorHandler
    Call DestroyServer
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "frmServer.frm", "Form_Terminate", Err.Number, Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrorHandler
    Call DestroyServer
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "frmServer.frm", "Form_Unload", Err.Number, Err.Description
End Sub

Private Sub mnuReloadScript_Click()
On Error GoTo ErrorHandler
If Scripting = 1 Then
    Set MyScript = Nothing
    Set clsScriptCommands = Nothing
    
    Set MyScript = New clsSadScript
    Set clsScriptCommands = New clsCommands
    MyScript.ReadInCode App.Path & "\Scripts\Main.txt", "Scripts\Main.txt", MyScript.SControl, False
    MyScript.SControl.AddObject "ScriptHardCode", clsScriptCommands, True
    Call TextAdd(frmServer.txtText, "Scripts reloaded.", True)
End If
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "frmServer.frm", "mnuReloadScript_Click", Err.Number, Err.Description
End Sub

Private Sub mnuServerLog_Click()
On Error GoTo ErrorHandler
    If mnuServerLog.Checked = True Then
        mnuServerLog.Checked = False
        ServerLog = False
    Else
        mnuServerLog.Checked = True
        ServerLog = True
    End If
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "frmServer.frm", "mnuServerLog_Click", Err.Number, Err.Description
End Sub


Private Sub PlayerTimer_Timer()
On Error GoTo ErrorHandler
Dim i As Long

If PlayerI <= MAX_PLAYERS Then
    If IsPlaying(PlayerI) Then
        Call SavePlayer(PlayerI, Player(PlayerI).CharNum)
        Call PlayerMsg(PlayerI, GetPlayerName(PlayerI) & " is now saved.", Yellow)
    End If
    PlayerI = PlayerI + 1
End If
If PlayerI >= MAX_PLAYERS Then
    PlayerI = 1
    PlayerTimer.Enabled = False
    tmrPlayerSave.Enabled = True
End If

ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "frmServer.frm", "PlayerTimer_Timer", Err.Number, Err.Description
End Sub

Private Sub tmrGameAI_Timer()
On Error GoTo ErrorHandler
    Call ServerLogic
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "frmServer.frm", "tmrGameAI_Timer", Err.Number, Err.Description
End Sub

Private Sub tmrPlayerSave_Timer()
On Error GoTo ErrorHandler
    Call PlayerSaveTimer
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "frmServer.frm", "tmrPlayerSave_Timer", Err.Number, Err.Description
End Sub

Private Sub tmrSpawnMapItems_Timer()
On Error GoTo ErrorHandler
    Call CheckSpawnMapItems
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "frmServer.frm", "tmrSpawnMapItems_Timer", Err.Number, Err.Description
End Sub

Private Sub txtText_GotFocus()
On Error GoTo ErrorHandler
    txtChat.SetFocus
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "frmServer.frm", "txtText_GotFocus", Err.Number, Err.Description
End Sub

Private Sub txtChat_KeyPress(KeyAscii As Integer)
On Error GoTo ErrorHandler
    If KeyAscii = vbKeyReturn And Trim(txtChat.Text) <> "" Then
        Call GlobalMsg("Server-- " & txtChat.Text, White)
        Call TextAdd(frmServer.txtText, "Server: " & txtChat.Text, True)
        txtChat.Text = ""
    End If
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "frmServer.frm", "txtChat_KeyPress", Err.Number, Err.Description
End Sub

Private Sub tmrShutdown_Timer()
On Error GoTo ErrorHandler
Static Secs As Long

    If Secs <= 0 Then Secs = 30
    Call GlobalMsg("Server Shutdown in " & Secs & " seconds.", BrightBlue)
    Call TextAdd(frmServer.txtText, "Automated Server Shutdown in " & Secs & " seconds.", True)
    Secs = Secs - 2
    If Secs <= 0 Then
        tmrShutdown.Enabled = False
        Call DestroyServer
    End If
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "frmServer.frm", "tmrShutdown_Timer", Err.Number, Err.Description
End Sub

Private Sub mnuShutdown_Click()
On Error GoTo ErrorHandler
    tmrShutdown.Enabled = True
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "frmServer.frm", "mnuShutdown_Click", Err.Number, Err.Description
End Sub

Private Sub mnuExit_Click()
On Error GoTo ErrorHandler
    Call DestroyServer
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "frmServer.frm", "mnuExit_Click", Err.Number, Err.Description
End Sub

Private Sub mnuReloadClasses_Click()
On Error GoTo ErrorHandler
    Call LoadClasses
    Call TextAdd(frmServer.txtText, "All classes reloaded.", True)
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "frmServer.frm", "mnuReloadClasses_Click", Err.Number, Err.Description
End Sub

Private Sub Socket_ConnectionRequest(index As Integer, ByVal requestID As Long)
On Error GoTo ErrorHandler
    Call AcceptConnection(index, requestID)
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "frmServer.frm", "Socket_ConnectionRequest", Err.Number, Err.Description
End Sub

Private Sub Socket_Accept(index As Integer, SocketId As Integer)
On Error GoTo ErrorHandler
    Call AcceptConnection(index, SocketId)
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "frmServer.frm", "Socket_Accept", Err.Number, Err.Description
End Sub

Private Sub Socket_DataArrival(index As Integer, ByVal bytesTotal As Long)
On Error GoTo ErrorHandler
    If IsConnected(index) Then
        Call IncomingData(index, bytesTotal)
    End If
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "frmServer.frm", "Socket_DataArrival", Err.Number, Err.Description
End Sub

Private Sub Socket_Close(index As Integer)
On Error GoTo ErrorHandler
    Call CloseSocket(index)
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "frmServer.frm", "Socket_Close", Err.Number, Err.Description
End Sub


