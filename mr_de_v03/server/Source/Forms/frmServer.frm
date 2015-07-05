VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmServer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Server"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   11460
   Icon            =   "frmServer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   11460
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmServerInfo 
      Caption         =   "Server Info"
      Height          =   3015
      Left            =   7680
      TabIndex        =   2
      Top             =   0
      Width           =   3615
      Begin VB.CommandButton cmdSetExpMod 
         Caption         =   "Set"
         Height          =   255
         Left            =   2400
         TabIndex        =   5
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtExpMod 
         Height          =   285
         Left            =   1080
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lblExpMod 
         Caption         =   "Exp Mod (%):"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   375
         Width           =   975
      End
   End
   Begin VB.Timer tmrShutdown 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   240
      Top             =   240
   End
   Begin VB.TextBox txtChat 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   2520
      Width           =   7455
   End
   Begin VB.TextBox txtText 
      Height          =   2295
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   7455
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
      Begin VB.Menu mnuRestore 
         Caption         =   "Restore"
         Visible         =   0   'False
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
      Begin VB.Menu mnuSetAccess 
         Caption         =   "Set Access"
      End
   End
   Begin VB.Menu mnuLog 
      Caption         =   "&Log"
      Begin VB.Menu mnuServerLog 
         Caption         =   "Server Log"
      End
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub txtExpMod_Change()
    If Not IsNumeric(txtExpMod.Text) Then
        txtExpMod.Text = ExpMod
    End If
End Sub

Private Sub cmdSetExpMod_Click()
    SetExpMod frmServer.txtExpMod.Text
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Sys As Long
    Sys = X / Screen.TwipsPerPixelX
    Select Case Sys
        Case WM_LBUTTONDOWN:
            Me.PopupMenu mnuFile
    End Select
End Sub

Private Sub Form_Resize()
    If WindowState = vbMinimized Then
        Me.Hide
        Me.Refresh
        mnuRestore.Visible = True
           With nid
                .cbSize = Len(nid)
                .hWnd = Me.hWnd
                .uId = vbNull
                .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
                .uCallBackMessage = WM_MOUSEMOVE
                .hIcon = Me.Icon
                .szTip = Me.Caption & vbNullChar
           End With
        Shell_NotifyIcon NIM_ADD, nid
    Else
        Shell_NotifyIcon NIM_DELETE, nid
    End If
End Sub
Private Sub Form_Load()
   ShutOn = False
End Sub
Private Sub Form_Terminate()
    DestroyServer
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DestroyServer
End Sub

Private Sub mnuServerLog_Click()
    mnuServerLog.Checked = Not mnuServerLog.Checked
    ServerLog = Not mnuServerLog.Checked
End Sub

Private Sub mnuSetAccess_Click()
Dim Name As String
Dim i As Long, PlayerAccess As Byte

    Name = InputBox("What is the player name?", "Give Access to?", "")
    i = FindPlayer(Name)
    If i <= 0 Then Exit Sub
    
    PlayerAccess = Val(InputBox("What access level?", "Access:", "1"))
    If PlayerAccess < 0 Then Exit Sub
    If PlayerAccess > ADMIN_CREATOR Then Exit Sub
    
    If IsConnected(i) Then
        ' sloppy... but whatever
        Update_Access i, PlayerAccess
        
        'sendplayermsg(i, "Your access has been changed.", BrightRed)
        SendActionMsg Current_Map(i), "Your access has been changed.", AlertColor, ACTIONMSG_SCREEN, 0, 0, i
        SendPlayerData (i)
    Else
        AddText frmServer.txtChat, "Player is not online."
    End If

End Sub

Private Sub txtText_GotFocus()
    txtChat.SetFocus
End Sub

Private Sub txtChat_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Trim$(txtChat.Text) <> vbNullString Then
            SendGlobalMsg txtChat.Text, White
            AddText frmServer.txtText, "Server: " & txtChat.Text
            txtChat.Text = vbNullString
        End If
    End If
End Sub

Private Sub tmrShutdown_Timer()
Static Secs As Long
   
   If ShutOn = False Then
       Secs = 30
       AddText frmServer.txtText, "[Realm Event] Automated Server Shutdown Canceled!"
       SendGlobalMsg "[Realm Event] Server Shutdown Canceled!", BrightCyan
       tmrShutdown.Enabled = False
       Exit Sub
   End If
   
   If Secs <= 0 Then Secs = 30
   
   Secs = Secs - 1
   SendGlobalMsg "[Realm Event] Server Shutdown in " & Secs & " seconds.", BrightCyan
   AddText frmServer.txtText, "Automated Server Shutdown in " & Secs & " seconds."
   
   If Secs <= 0 Then
       tmrShutdown.Enabled = False
       DestroyServer
   End If
End Sub

Private Sub mnuShutdown_Click()
    If ShutOn = False Then
        tmrShutdown.Enabled = True
        mnuShutdown.Caption = "Cancel Shutdown"
        ShutOn = True
    ElseIf ShutOn = True Then
        mnuShutdown.Caption = "Shutdown"
        ShutOn = False
    End If
End Sub

Private Sub mnuExit_Click()
    DestroyServer
End Sub

Private Sub mnuReloadClasses_Click()
    LoadClasses
    AddText frmServer.txtText, "All classes reloaded."
End Sub
Private Sub mnuRestore_Click()
    WindowState = vbNormal
    Me.Show
    mnuRestore.Visible = False
End Sub

Private Sub Socket_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    AcceptConnection Index, requestID
End Sub

'Private Sub Socket_Accept(Index As Integer, SocketId As Integer)
'    AcceptConnection Index, SocketId
'End Sub

Private Sub Socket_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    If IsConnected(Index) Then
        IncomingData Index, bytesTotal
    End If
End Sub

Private Sub Socket_Close(Index As Integer)
    CloseSocket Index
End Sub


