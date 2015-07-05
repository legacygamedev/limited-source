VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL32.OCX"
Begin VB.Form frmServer 
   Caption         =   "Mirage Source Engine Server"
   ClientHeight    =   3600
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   6735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   5953
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Console"
      TabPicture(0)   =   "frmServer.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "txtText"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "txtChat"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Players"
      TabPicture(1)   =   "frmServer.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "lvwInfo"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.TextBox txtChat 
         Height          =   375
         Left            =   -74880
         TabIndex        =   2
         Top             =   2880
         Width           =   6255
      End
      Begin VB.TextBox txtText 
         Height          =   2295
         Left            =   -74880
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   480
         Width           =   6255
      End
      Begin MSComctlLib.ListView lvwInfo 
         Height          =   2775
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   4895
         View            =   3
         Arrange         =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Index"
            Object.Width           =   1147
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "IP Address"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Account"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Character"
            Object.Width           =   2999
         EndProperty
      End
   End
   Begin VB.Timer tmrShutdown 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   3240
      Top             =   0
   End
   Begin VB.Timer tmrGameAI 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2760
      Top             =   0
   End
   Begin MSWinsockLib.Winsock Socket 
      Index           =   0
      Left            =   2280
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
      Begin VB.Menu mnuReloadClasses 
         Caption         =   "Reload Classes"
      End
      Begin VB.Menu mnuRMaps 
         Caption         =   "Reload Maps"
      End
      Begin VB.Menu mnuRSpells 
         Caption         =   "Reload Spells"
      End
      Begin VB.Menu mnuRShop 
         Caption         =   "Reload Shops"
      End
      Begin VB.Menu mnuRNPC 
         Caption         =   "Reload NPC"
      End
      Begin VB.Menu mnuRItems 
         Caption         =   "Reload Items"
      End
   End
   Begin VB.Menu mnuLog 
      Caption         =   "&Log"
      Begin VB.Menu mnuServerLog 
         Caption         =   "Server Log"
      End
   End
   Begin VB.Menu mnuKick 
      Caption         =   "Kick"
      Visible         =   0   'False
      Begin VB.Menu mnuKickPlayer 
         Caption         =   "Kick"
      End
      Begin VB.Menu mnuDisconnect 
         Caption         =   "Disconnect"
      End
      Begin VB.Menu mnuBan 
         Caption         =   "Ban"
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
    Call UsersOnline_Start
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

Private Sub lvwInfo_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    'When a ColumnHeader object is clicked, the ListView control is sorted by the subitems of that column.
    'Set the SortKey to the Index of the ColumnHeader - 1
    'Set Sorted to True to sort the list.
    
    If lvwInfo.SortOrder = lvwAscending Then
        lvwInfo.SortOrder = lvwDescending
    Else
        lvwInfo.SortOrder = lvwAscending
    End If
        
    
    lvwInfo.SortKey = ColumnHeader.Index - 1
    lvwInfo.Sorted = True
    Debug.Print ColumnHeader.Index & " " & ColumnHeader.Text & " " & ColumnHeader.Width
End Sub

Private Sub mnuRItems_Click()
    Call LoadItems
    Call TextAdd(frmServer.txtText, "All items reloaded.", True)
End Sub

Private Sub mnuRMaps_Click()
    Call LoadMaps
    Call TextAdd(frmServer.txtText, "All maps reloaded.", True)
End Sub

Private Sub mnuRNPC_Click()
    Call LoadNpcs
    Call TextAdd(frmServer.txtText, "All npcs reloaded.", True)
End Sub

Private Sub mnuRShop_Click()
    Call LoadShops
    Call TextAdd(frmServer.txtText, "All shops reloaded.", True)
End Sub

Private Sub mnuRSpells_Click()
    Call LoadSpells
    Call TextAdd(frmServer.txtText, "All spells reloaded.", True)
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

Private Sub tmrGameAI_Timer()
    Call ServerLogic
End Sub

Private Sub txtText_GotFocus()
    txtChat.SetFocus
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

Private Sub Socket_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    Call AcceptConnection(Index, requestID)
End Sub

Private Sub Socket_Accept(Index As Integer, SocketId As Integer)
    Call AcceptConnection(Index, SocketId)
End Sub

Private Sub Socket_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    If IsConnected(Index) Then
        Call IncomingData(Index, bytesTotal)
    End If
End Sub

Private Sub Socket_Close(Index As Integer)
    Call CloseSocket(Index)
End Sub

Sub UsersOnline_Start()
Dim i As Integer
    For i = 1 To MAX_PLAYERS
        Dim Add() As ListItem
        frmServer.lvwInfo.ListItems.Add (i)
        If i < 10 Then
            frmServer.lvwInfo.ListItems(i).Text = "00" & i
        ElseIf i < 100 Then
            frmServer.lvwInfo.ListItems(i).Text = "0" & i
        Else
            frmServer.lvwInfo.ListItems(i).Text = i
        End If
        frmServer.lvwInfo.ListItems(i).SubItems(1) = vbNullString
        frmServer.lvwInfo.ListItems(i).SubItems(2) = vbNullString
        frmServer.lvwInfo.ListItems(i).SubItems(3) = vbNullString
    Next i
End Sub

Private Sub lvwInfo_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If Button = vbRightButton Then
        PopupMenu mnuKick
    End If
    
End Sub

Private Sub mnuKickPlayer_Click()
Dim Index As Long
Dim Name As String

    Name = frmServer.lvwInfo.SelectedItem.SubItems(3)
    If Not Name = "Not Playing" Then
        Call AlertMsg(FindPlayer(Name), "You have been kicked by the server owner!")
        CloseSocket (FindPlayer(Name))
    End If
End Sub

Sub mnuDisconnectPlayer_Click()
Dim Name As String
    
    Name = frmServer.lvwInfo.SelectedItem.SubItems(3)
    If Not Name = "Not Playing" Then
        CloseSocket (FindPlayer(Name))
    End If
End Sub

Sub mnuBanPlayer_click()
Dim Name As String

    Name = frmServer.lvwInfo.SelectedItem.SubItems(3)
    If Not Name = "Not Playing" Then
        Call ServerBanIndex(FindPlayer(Name))
    End If
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
