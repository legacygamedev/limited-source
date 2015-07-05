VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmServer 
   Caption         =   "Mirage Source Server"
   ClientHeight    =   3555
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6720
   Icon            =   "frmServer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   6720
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
      TabHeight       =   503
      TabCaption(0)   =   "Console"
      TabPicture(0)   =   "frmServer.frx":1708A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "txtText"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "txtChat"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Players"
      TabPicture(1)   =   "frmServer.frx":170A6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lvwInfo"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Control "
      TabPicture(2)   =   "frmServer.frx":170C2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraDatabase"
      Tab(2).Control(1)=   "fraServer"
      Tab(2).ControlCount=   2
      Begin VB.Frame fraDatabase 
         Caption         =   "Database"
         Height          =   1935
         Left            =   -73320
         TabIndex        =   8
         Top             =   480
         Width           =   1935
         Begin VB.CommandButton cmdReloadItems 
            Caption         =   "Reload Items"
            Height          =   255
            Left            =   240
            TabIndex        =   14
            Top             =   1440
            Width           =   1455
         End
         Begin VB.CommandButton cmdReloadNPCs 
            Caption         =   "Reload NPCs"
            Height          =   255
            Left            =   240
            TabIndex        =   13
            Top             =   1200
            Width           =   1455
         End
         Begin VB.CommandButton cmdReloadShops 
            Caption         =   "Reload Shops"
            Height          =   255
            Left            =   240
            TabIndex        =   12
            Top             =   960
            Width           =   1455
         End
         Begin VB.CommandButton CmdReloadSpells 
            Caption         =   "Reload Spells"
            Height          =   255
            Left            =   240
            TabIndex        =   11
            Top             =   720
            Width           =   1455
         End
         Begin VB.CommandButton cmdReloadMaps 
            Caption         =   "Reload Maps"
            Height          =   255
            Left            =   240
            TabIndex        =   10
            Top             =   480
            Width           =   1455
         End
         Begin VB.CommandButton cmdReloadClasses 
            Caption         =   "Reload Classes"
            Height          =   255
            Left            =   240
            TabIndex        =   9
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.Frame fraServer 
         Caption         =   "Server"
         Height          =   1215
         Left            =   -74880
         TabIndex        =   1
         Top             =   480
         Width           =   1455
         Begin VB.CheckBox chkServerLog 
            Caption         =   "Server Log"
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   840
            Width           =   1215
         End
         Begin VB.CommandButton cmdExit 
            Caption         =   "Exit"
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   480
            Width           =   975
         End
         Begin VB.CommandButton cmdShutDown 
            Caption         =   "Shut Down"
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.TextBox txtChat 
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   2880
         Width           =   6255
      End
      Begin VB.TextBox txtText 
         Height          =   2295
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   480
         Width           =   6255
      End
      Begin MSComctlLib.ListView lvwInfo 
         Height          =   2775
         Left            =   -74880
         TabIndex        =   4
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
   Begin MSWinsockLib.Winsock Socket 
      Index           =   0
      Left            =   6360
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Menu mnuKick 
      Caption         =   "&Kick"
      Visible         =   0   'False
      Begin VB.Menu mnuKickPlayer 
         Caption         =   "Kick"
      End
      Begin VB.Menu mnuDisconnectPlayer 
         Caption         =   "Disconnect"
      End
      Begin VB.Menu mnuBanPlayer 
         Caption         =   "Ban"
      End
      Begin VB.Menu mnuAdminPlayer 
         Caption         =   "Make Admin"
      End
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ******************************************
' **            Mirage Source 4           **
' ******************************************

' ********************
' ** Winsock object **
' ********************
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

' ********************

Private Sub chkServerLog_Click()
    ' if its not 0, then its true
    If Not chkServerLog.Value Then
        ServerLog = True
    End If
End Sub

Private Sub cmdExit_Click()
    Call DestroyServer
End Sub

Private Sub cmdReloadClasses_Click()
    Call LoadClasses
    Call TextAdd("All classes reloaded.")
End Sub

Private Sub cmdReloadItems_Click()
    Call LoadItems
    Call TextAdd("All items reloaded.")
End Sub

Private Sub cmdReloadMaps_Click()
    Call LoadMaps
    Call TextAdd("All maps reloaded.")
End Sub

Private Sub cmdReloadNPCs_Click()
    Call LoadNpcs
    Call TextAdd("All npcs reloaded.")
End Sub

Private Sub cmdReloadShops_Click()
    Call LoadShops
    Call TextAdd("All shops reloaded.")
End Sub

Private Sub CmdReloadSpells_Click()
    Call LoadSpells
    Call TextAdd("All spells reloaded.")
End Sub

Private Sub cmdShutDown_Click()
    isShuttingDown = True
    cmdShutDown.Enabled = False
End Sub

Private Sub Form_Load()
    Call UsersOnline_Start
End Sub

Private Sub Form_Resize()
    If frmServer.WindowState = vbMinimized Then
        frmServer.Hide
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = True
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

Private Sub txtText_GotFocus()
    txtChat.SetFocus
End Sub

Private Sub txtChat_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If LenB(Trim$(txtChat.Text)) > 0 Then
            Call GlobalMsg(txtChat.Text, White)
            Call TextAdd("Server: " & txtChat.Text)
            txtChat.Text = vbNullString
        End If
        KeyAscii = 0
    End If
End Sub

Sub UsersOnline_Start()
Dim i As Integer
    For i = 1 To MAX_PLAYERS
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
    Next
End Sub

Private Sub lvwInfo_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If Button = vbRightButton Then
        PopupMenu mnuKick
    End If
    
End Sub

Private Sub mnuKickPlayer_Click()
Dim Name As String

    Name = frmServer.lvwInfo.SelectedItem.SubItems(3)
    If Not Name = "Not Playing" Then
        Call AlertMsg(FindPlayer(Name), "You have been kicked by the server owner!")
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

Sub mnuAdminPlayer_click()
Dim Name As String

    Name = frmServer.lvwInfo.SelectedItem.SubItems(3)
    If Not Name = "Not Playing" Then
        Call SetPlayerAccess(FindPlayer(Name), 4)
        Call SendPlayerData(FindPlayer(Name))
        Call PlayerMsg(FindPlayer(Name), "You have been granted administrator access.", Pink)
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim lmsg As Long
   
    lmsg = x / Screen.TwipsPerPixelX
    Select Case lmsg
        Case WM_LBUTTONDBLCLK
            frmServer.WindowState = vbNormal
            frmServer.Show
            txtText.SelStart = Len(txtText.Text)
    End Select
End Sub
