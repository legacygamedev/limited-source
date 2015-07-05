VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "Mswinsck.ocx"
Begin VB.Form frmServer 
   Caption         =   "M:RPGe Game Server"
   ClientHeight    =   4425
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   8745
   Icon            =   "frmServer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   8745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrSP 
      Interval        =   1000
      Left            =   2160
      Top             =   240
   End
   Begin VB.Timer tmrShutdown 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   1680
      Top             =   240
   End
   Begin VB.Timer tmrGameAI 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   1200
      Top             =   240
   End
   Begin VB.Timer tmrSpawnMapItems 
      Interval        =   1000
      Left            =   720
      Top             =   240
   End
   Begin VB.Timer tmrPlayerSave 
      Interval        =   60000
      Left            =   240
      Top             =   240
   End
   Begin VB.TextBox txtChat 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   3960
      Width           =   8535
   End
   Begin VB.TextBox txtText 
      Height          =   3735
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   8535
   End
   Begin MSWinsockLib.Winsock Socket 
      Index           =   0
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
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
      Begin VB.Menu saveItems 
         Caption         =   "saveItems"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuReloadClasses 
         Caption         =   "Reload Classes"
      End
   End
   Begin VB.Menu mnuLog 
      Caption         =   "&Log"
      Begin VB.Menu mnuServerLog 
         Caption         =   "Server Log"
      End
      Begin VB.Menu viewLogs 
         Caption         =   "View Logs"
      End
   End
   Begin VB.Menu tim 
      Caption         =   "Time"
      Begin VB.Menu mnuNight 
         Caption         =   "Night?"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu map 
      Caption         =   "Maps"
      Begin VB.Menu map_Renew 
         Caption         =   "Renew Maps"
      End
   End
   Begin VB.Menu debug 
      Caption         =   "Debug"
      Begin VB.Menu showDebug 
         Caption         =   "Show Debug Window"
      End
   End
   Begin VB.Menu set 
      Caption         =   "Settings"
      Begin VB.Menu chkBroadcast 
         Caption         =   "Can broadcast?"
      End
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private SPcounter As Long

Private Sub chkBroadcast_Click()
If chkBroadcast.Checked = False Then
    blnBroadcast = True
    chkBroadcast.Checked = True
Else
    chkBroadcast.Checked = False
    blnBroadcast = False
End If
End Sub

Private Sub Form_Load()
    ServerLog = True
    blnNight = False
    mnuNight.Checked = False
    mnuServerLog.Checked = True
End Sub

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

Private Sub map_Renew_Click()
If MsgBox("are you sure? YOU PROBABLY DON'T WANT TO DO THIS!" & vbCrLf & "IT MAY KILL ALL YOUR MAPS AND IS NOT UNDOABLE", vbOKCancel, "!") = vbOK Then
    Call ConvertOldMapsToNew
End If
End Sub

Private Sub mnuNight_Click()
    If mnuNight.Checked = True Then
        mnuNight.Checked = False
        blnNight = False
        Call SendNight(False)
    Else
        mnuNight.Checked = True
        blnNight = True
        Call SendNight(True)
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



Private Sub saveItems_Click()
Dim i As Long
For i = 1 To MAX_ITEMS
Call SaveItem(i)
Next i
End Sub

Private Sub showDebug_Click()
showDebug = True
frmDebugWindow.Show
End Sub

Private Sub tmrGameAI_Timer()
    Call ServerLogic
End Sub

Private Sub tmrPlayerSave_Timer()
    Call PlayerSaveTimer
End Sub

Private Sub tmrSP_Timer()
Dim i As Long
Dim newSP As Long
    For i = 1 To MAX_PLAYERS
        If (IsPlaying(i)) Then
            If player(i).Char(player(i).CharNum).lastSentSP <> GetPlayerSP(i) Then
                SendSP (i)
                player(i).Char(player(i).CharNum).lastSentSP = GetPlayerSP(i)
            End If
        End If
    Next i
    'If SPcounter >= 2 Then
        For i = 1 To MAX_PLAYERS
            If (IsPlaying(i)) Then
                newSP = GetPlayerSP(i) + 10
                Call SetPlayerSP(i, newSP)
            End If
        Next i
        SPcounter = 0
   ' End If
    SPcounter = SPcounter + 1
End Sub

Private Sub tmrSpawnMapItems_Timer()
    Call CheckSpawnMapItems
End Sub

Private Sub txtText_GotFocus()
    txtChat.SetFocus
End Sub

Private Sub txtChat_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And Trim(txtChat.Text) <> "" Then
        Call ServerMsg(txtChat.Text)
        Call TextAdd(frmServer.txtText, "Server: " & txtChat.Text, True)
        txtChat.Text = ""
    End If
End Sub

Private Sub tmrShutdown_Timer()
Static Secs As Long

    If Secs <= 0 Then Secs = 30
    Call GlobalMsg("Server Shutdown in " & Secs & " seconds.", RGB_GlobalColor)
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

Private Sub Socket_ConnectionRequest(index As Integer, ByVal requestID As Long)
    Call AcceptConnection(index, requestID)
End Sub

Private Sub Socket_Accept(index As Integer, SocketId As Integer)
    Call AcceptConnection(index, SocketId)
End Sub

Private Sub Socket_DataArrival(index As Integer, ByVal bytesTotal As Long)
    If IsConnected(index) Then
        Call IncomingData(index, bytesTotal)
    End If
End Sub

Private Sub Socket_Close(index As Integer)
    Call CloseSocket(index)
End Sub











Private Sub viewLogs_Click()
    frmLogs.Show
    
End Sub
