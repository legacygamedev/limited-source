VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmServer 
   Caption         =   "Dual Solace"
   ClientHeight    =   3225
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   10890
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmServer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   10890
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtText 
      Height          =   2295
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   120
      Width           =   7455
   End
   Begin VB.TextBox txtChat 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2520
      Width           =   7455
   End
   Begin VB.Timer tmrPlayerSave 
      Interval        =   60000
      Left            =   240
      Top             =   240
   End
   Begin VB.Timer tmrSpawnMapItems 
      Interval        =   1000
      Left            =   720
      Top             =   240
   End
   Begin VB.Timer tmrGameAI 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   1200
      Top             =   240
   End
   Begin VB.Timer tmrShutdown 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   1680
      Top             =   240
   End
   Begin VB.ListBox lstPlayers 
      Height          =   2370
      Left            =   7800
      TabIndex        =   1
      Top             =   720
      Width           =   3015
   End
   Begin VB.TextBox txtTotal 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   9510
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "Total"
      Top             =   420
      Width           =   615
   End
   Begin VB.Timer tmrTime 
      Interval        =   1000
      Left            =   2160
      Top             =   240
   End
   Begin MSWinsockLib.Winsock Socket 
      Index           =   0
      Left            =   5295
      Top             =   1905
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog cdFile 
      Left            =   6540
      Top             =   1845
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet Inet 
      Left            =   5842
      Top             =   1755
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      Protocol        =   4
   End
   Begin VB.Label lblIP 
      Alignment       =   2  'Center
      Caption         =   "IP"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   9
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label lblPort 
      Alignment       =   2  'Center
      Caption         =   "Port"
      Height          =   255
      Left            =   2580
      TabIndex        =   8
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "IP:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "Port:"
      Height          =   255
      Left            =   2100
      TabIndex        =   6
      Top             =   3000
      Width           =   375
   End
   Begin VB.Line Line1 
      X1              =   1980
      X2              =   1980
      Y1              =   3000
      Y2              =   3240
   End
   Begin VB.Line Line2 
      X1              =   3660
      X2              =   3660
      Y1              =   3000
      Y2              =   3240
   End
   Begin VB.Line Line3 
      X1              =   7680
      X2              =   7680
      Y1              =   0
      Y2              =   3240
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Player List"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7800
      TabIndex        =   5
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label4 
      Caption         =   "Total Players:"
      Height          =   255
      Left            =   8340
      TabIndex        =   4
      Top             =   420
      Width           =   1350
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000004&
      BackStyle       =   1  'Opaque
      Height          =   3240
      Left            =   0
      Top             =   0
      Width           =   10905
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&Server"
      Begin VB.Menu mnuShutdown 
         Caption         =   "Shutdown"
      End
      Begin VB.Menu mnuLock 
         Caption         =   "Lock"
      End
      Begin VB.Menu mnuUnlock 
         Caption         =   "Unlock"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuFileEditing 
      Caption         =   "&File Editing"
      Begin VB.Menu mnuServerMessage 
         Caption         =   "Edit Server Message"
      End
      Begin VB.Menu mnuEditScript 
         Caption         =   "Edit Scripts"
      End
      Begin VB.Menu mnuEditLibrary 
         Caption         =   "Edit Library"
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
      Begin VB.Menu mnuReloadLibrary 
         Caption         =   "Reload Library"
      End
   End
   Begin VB.Menu mnuLog 
      Caption         =   "&Log"
      Begin VB.Menu mnuServerLog 
         Caption         =   "Server Log"
      End
   End
   Begin VB.Menu mnuListOption 
      Caption         =   "&Player Options"
      Begin VB.Menu mnuKick 
         Caption         =   "Kick Player"
      End
      Begin VB.Menu mnuBan 
         Caption         =   "Ban Player"
      End
   End
   Begin VB.Menu mPopupSys 
      Caption         =   "&SysTray"
      Visible         =   0   'False
      Begin VB.Menu mPopRestore 
         Caption         =   "Restore"
      End
      Begin VB.Menu mnuSysLock 
         Caption         =   "Lock"
      End
      Begin VB.Menu mnuSysUnlock 
         Caption         =   "Unlock"
         Visible         =   0   'False
      End
      Begin VB.Menu mPopExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
'On Error GoTo errorhandler:
 With nid
  .cbSize = Len(nid)
  .hWnd = Me.hWnd
  .uID = vbNull
  .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
  .uCallbackMessage = WM_MOUSEMOVE
  .hIcon = Me.Icon
  .dwState = NIS_SHAREDICON
  .szTip = "RealFeel Server" & vbNullChar
 End With
 mnuListOption.Enabled = False
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("frmServer.frm", "Form_Load", Err.Number, Err.Description)
End Sub

Private Sub Form_Resize()
'On Error GoTo errorhandler:
    If frmServer.WindowState = vbMinimized Then
        frmServer.Hide
        Shell_NotifyIcon NIM_ADD, nid
    End If
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("frmServer.frm", "Form_Resize", Err.Number, Err.Description)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'On Error GoTo errorhandler:
'this procedure receives the callbacks from the System Tray icon.
Dim Result As Long
Dim Msg As Long
 'the value of X will vary depending upon the scalemode setting
 If Me.ScaleMode = vbPixels Then
  Msg = x
 Else
  Msg = x / Screen.TwipsPerPixelX
 End If
 Select Case Msg
  Case WM_LBUTTONDBLCLK    '515 restore form window
   Shell_NotifyIcon NIM_DELETE, nid
   Me.WindowState = vbNormal
   Me.Show
 Case WM_RBUTTONUP         '517 display popup menu
    ' This check keeps the popup menu from showing while maximized at this one random area.
    If Me.WindowState = vbMinimized Then
        Result = SetForegroundWindow(Me.hWnd)
        Me.PopupMenu Me.mPopupSys
    End If
 End Select
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("frmServer.frm", "Form_MouseMove", Err.Number, Err.Description)
End Sub

Private Sub Form_Terminate()
'On Error GoTo errorhandler:
    Shell_NotifyIcon NIM_DELETE, nid
    Call DestroyServer
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("frmServer.frm", "Form_Terminate", Err.Number, Err.Description)
End Sub

Private Sub Form_Unload(Cancel As Integer)
'On Error GoTo errorhandler:
    Shell_NotifyIcon NIM_DELETE, nid
    Call DestroyServer
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("frmServer.frm", "Form_Unload", Err.Number, Err.Description)
End Sub

Private Sub lstPlayers_click()
    If lstPlayers.List(lstPlayers.ListIndex) <> vbNullString Then
        mnuListOption.Enabled = True
    Else
        mnuListOption.Enabled = False
    End If
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("frmServer.frm", "lstPlayers_Click", Err.Number, Err.Description)
End Sub

Private Sub mnuBan_Click()
Dim AccountName As String

    If lstPlayers.List(lstPlayers.ListIndex) <> vbNullString Then
        mnuListOption.Enabled = True
    Else
        mnuListOption.Enabled = False
        Exit Sub
    End If

    AccountName = Left(lstPlayers.List(lstPlayers.ListIndex), InStr(1, lstPlayers.List(lstPlayers.ListIndex), "/") - 1)
    Call BanIndex(FindPlayer(AccountName), 0)
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("frmServer.frm", "mnuBan_Click", Err.Number, Err.Description)
End Sub

Private Sub mnuEditLibrary_Click()
'On Error GoTo errorhandler:
    frmLibrary.Visible = True
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("frmServer.frm", "mnuEditLibrary_Click", Err.Number, Err.Description)
End Sub

Private Sub mnuEditScript_Click()
Dim sFilename As String
'On Error GoTo errorhandler:
    frmEdit.Visible = True
    EditType = EDIT_OTHER
    
    '  Initialize the Dir$ function (and get the first filename if it exists).
    sFilename = Dir$(App.Path & "\Library\*.txt")
    
    '  Loop on the text files in the directory.
    Do While sFilename <> ""
        '  Add the current filename to the listbox
        frmEdit.lstLibrary.AddItem sFilename
        '  Advance to the next filename in the directory.
        sFilename = Dir$
    Loop
    
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("frmServer.frm", "mnuEditScript_Click", Err.Number, Err.Description)
End Sub

Private Sub mnuKick_Click()
Dim PlayerName As String
Dim n As Long

    If lstPlayers.List(lstPlayers.ListIndex) <> vbNullString Then
        mnuListOption.Enabled = True
    Else
        mnuListOption.Enabled = False
        Exit Sub
    End If

    PlayerName = Right(lstPlayers.List(lstPlayers.ListIndex), InStr(1, lstPlayers.List(lstPlayers.ListIndex), "/") - 1)
    n = FindPlayer(PlayerName)
    Call GlobalMsg(GetPlayerName(n) & " has been kicked from " & GAME_NAME & " by " & GetPlayerName(0) & "!", White)
    Call AddLog(GetPlayerName(0) & " has kicked " & GetPlayerName(n) & ".", ADMIN_LOG)
    Call AlertMsg(n, "You have been kicked by " & GetPlayerName(0) & "!")
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("frmServer.frm", "mnuBan_Click", Err.Number, Err.Description)
End Sub

Private Sub mnusysLock_Click()
Dim i As Long
    mnuSysUnlock.Visible = True
    mnuSysLock.Visible = False
    
    txtTotal.ForeColor = &HFF&
    txtTotal.Text = "Locked"
    
    ServerState = SERVER_LOCKED
    
    ' Loop through everyone online and kick anyone without access
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) And GetPlayerAccess(i) < 2 Then
            Call AlertMsg(i, "Server Closed for Maintenance.")
        End If
    Next i
      
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("frmServer.frm", "mnuSysLock_Click", Err.Number, Err.Description)
End Sub

Private Sub mnuSysUnLock_Click()
    mnuSysLock.Visible = True
    mnuSysUnlock.Visible = False
    
    txtTotal.ForeColor = &H0&
    txtTotal.Text = TotalOnlinePlayers
    
    ServerState = SERVER_UNLOCKED
    
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("frmServer.frm", "mnuSysUnLock_Click", Err.Number, Err.Description)
End Sub

Private Sub mnuLock_Click()
Dim i As Long
    mnuUnlock.Visible = True
    mnuLock.Visible = False
    
    txtTotal.ForeColor = &HFF&
    txtTotal.Text = "Locked"
    
    ServerState = SERVER_LOCKED
    
    ' Loop through everyone online and kick anyone without access
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) And GetPlayerAccess(i) < 2 Then
            Call AlertMsg(i, "Server Closed for Maintenance.")
        End If
    Next i
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("frmServer.frm", "mnuLock_Click", Err.Number, Err.Description)
End Sub

Private Sub mnuUnlock_Click()
    mnuLock.Visible = True
    mnuUnlock.Visible = False
    
    txtTotal.ForeColor = &H0&
    txtTotal.Text = TotalOnlinePlayers
    
    ServerState = SERVER_UNLOCKED
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("frmServer.frm", "mnuUnlock_Click", Err.Number, Err.Description)
End Sub

Private Sub mnuReloadLibrary_Click()
    frmLibrary.lstLibrary.Clear
    LoadLibrary
End Sub

Private Sub mnuReloadScript_Click()
'On Error GoTo errorhandler:
    Call LoadScripts
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("frmServer.frm", "mnuReloadScript_Click", Err.Number, Err.Description)
End Sub

Private Sub mnuServerLog_Click()
'On Error GoTo errorhandler:
    If mnuServerLog.Checked = True Then
        mnuServerLog.Checked = False
        ServerLog = False
    Else
        mnuServerLog.Checked = True
        ServerLog = True
    End If
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("frmServer.frm", "mnuServerLog_Click", Err.Number, Err.Description)
End Sub

Private Sub mnuServerMessage_Click()
'On Error GoTo errorhandler:
    frmEdit.Visible = True
    EditType = EDIT_SERVERMESSAGE
    frmEdit.rtfEdit.Text = GetVar(App.Path & "\Data\data.ini", "Strings", "Msg")
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("frmServer.frm", "mnuServerMessage_Click", Err.Number, Err.Description)
End Sub

Private Sub tmrGameAI_Timer()
'On Error GoTo errorhandler:
    Call ServerLogic
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("frmServer.frm", "tmrGameAI_Timer", Err.Number, Err.Description)
End Sub

Private Sub tmrPlayerSave_Timer()
'On Error GoTo errorhandler:
    Call PlayerSaveTimer
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("frmServer.frm", "tmrPlayerSave_Timer", Err.Number, Err.Description)
End Sub

Private Sub tmrSpawnMapItems_Timer()
'On Error GoTo errorhandler:
    Call CheckSpawnMapItems
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("frmServer.frm", "tmrSpawnMapItems_Timer", Err.Number, Err.Description)
End Sub

Private Sub tmrTime_Timer()
'On Error GoTo errorhandler:
Dim n As Long
Server_Second = Server_Second + 1
If Server_Second > 60 Then
    Server_Minute = Server_Minute + 1
            For n = 1 To MAX_SHOPS
                If Server_Minute * TIME_MINUTE >= Shop(n).Restock Then
                    Call ResetShopStock(n)
                End If
            Next n
    If Server_Minute > 60 Then
        Server_Hour = Server_Hour + 1
        If Server_Hour > 24 Then
            For n = 1 To MAX_SHOPS
                Call ResetShopStock(n)
            Next n
            Server_Hour = 1
        End If
        Server_Minute = 1
    End If
Server_Second = 1
End If
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("frmServer.frm", "tmrTime_Timer", Err.Number, Err.Description)
End Sub

Private Sub txtText_GotFocus()
'On Error GoTo errorhandler:
    txtChat.SetFocus
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("frmServer.frm", "txtText_GotFocus", Err.Number, Err.Description)
End Sub

Private Sub txtChat_KeyPress(KeyAscii As Integer)
'On Error GoTo errorhandler:
    If KeyAscii = vbKeyReturn And Trim$(txtChat.Text) <> "" Then
        Call GlobalMsg(txtChat.Text, White)
        Call TextAdd(frmServer.txtText, "Server: " & txtChat.Text, True)
        txtChat.Text = ""
    End If
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("frmServer.frm", "txtChat_KeyPress", Err.Number, Err.Description)
End Sub

Private Sub tmrShutdown_Timer()
'On Error GoTo errorhandler:
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
errorhandler:
  Call ReportError("frmServer.frm", "tmrShutdown_Timer", Err.Number, Err.Description)
End Sub

Private Sub mnuShutdown_Click()
'On Error GoTo errorhandler:
    tmrShutdown.Enabled = True
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("frmServer.frm", "mnuShutdown_Click", Err.Number, Err.Description)
End Sub

Private Sub mnuExit_Click()
'On Error GoTo errorhandler:
    Call DestroyServer
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("frmServer.frm", "mnuExit_Click", Err.Number, Err.Description)
End Sub

Private Sub mnuReloadClasses_Click()
'On Error GoTo errorhandler:
    Call LoadClasses
    Call TextAdd(frmServer.txtText, "All classes reloaded.", True)
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("frmServer.frm", "mnuReloadClasses_Click", Err.Number, Err.Description)
End Sub

Private Sub Socket_ConnectionRequest(Index As Integer, ByVal requestID As Long)
'On Error GoTo errorhandler:
    Call AcceptConnection(Index, requestID)
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("frmServer.frm", "Socket_ConnectionRequest", Err.Number, Err.Description)
End Sub

Private Sub Socket_Accept(Index As Integer, SocketId As Integer)
'On Error GoTo errorhandler:
    Call AcceptConnection(Index, SocketId)
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("frmServer.frm", "Socket_Accept", Err.Number, Err.Description)
End Sub

Private Sub Socket_DataArrival(Index As Integer, ByVal bytesTotal As Long)
'On Error GoTo errorhandler:
    If IsConnected(Index) Then
        Call IncomingData(Index, bytesTotal)
    End If
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("frmServer.frm", "Socket_DataArrival", Err.Number, Err.Description)
End Sub

Private Sub Socket_Close(Index As Integer)
'On Error GoTo errorhandler:
    Call CloseSocket(Index)
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("frmServer.frm", "Socket_Close", Err.Number, Err.Description)
End Sub

Private Sub mPopExit_Click()
    ' Called when user clicks the popup menu Exit command
    Call DestroyServer
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("frmServer.frm", "mPopExit_Click", Err.Number, Err.Description)
End Sub

Private Sub mPopRestore_Click()
    ' Called when the user clicks the popup menu Restore command
    Dim Result As Long
    Me.WindowState = vbNormal
    Result = SetForegroundWindow(Me.hWnd)
    Me.Show
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("frmServer.frm", "mPopRestore_Click", Err.Number, Err.Description)
End Sub

