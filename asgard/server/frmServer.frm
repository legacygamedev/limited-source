VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmServer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asgard Server"
   ClientHeight    =   5505
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   10665
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmServer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   367
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   711
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command13 
      Caption         =   "Kick"
      Height          =   255
      Left            =   240
      TabIndex        =   32
      Top             =   4440
      Width           =   1575
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Ban"
      Height          =   255
      Left            =   1920
      TabIndex        =   31
      Top             =   4440
      Width           =   1575
   End
   Begin VB.CommandButton Command22 
      Caption         =   "Mute"
      Height          =   255
      Left            =   3600
      TabIndex        =   30
      Top             =   4440
      Width           =   1575
   End
   Begin VB.CommandButton Command23 
      Caption         =   "UnMute"
      Height          =   255
      Left            =   5280
      TabIndex        =   29
      Top             =   4440
      Width           =   1575
   End
   Begin VB.CommandButton Command66 
      Caption         =   "Refresh"
      Height          =   255
      Left            =   9000
      TabIndex        =   27
      Top             =   4440
      Width           =   1575
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Gridlines"
      Height          =   255
      Left            =   7920
      TabIndex        =   26
      Top             =   4440
      Value           =   1  'Checked
      Width           =   975
   End
   Begin VB.Frame Frame5 
      Caption         =   "Chat Options"
      Height          =   2415
      Left            =   9360
      TabIndex        =   17
      Top             =   120
      Width           =   1335
      Begin VB.Timer tmrChatLogs 
         Interval        =   1000
         Left            =   840
         Top             =   1560
      End
      Begin VB.CheckBox chkBC 
         Caption         =   "Broadcast"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox chkE 
         Caption         =   "Emote"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   480
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox chkM 
         Caption         =   "Map"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   720
         Value           =   1  'Checked
         Width           =   615
      End
      Begin VB.CheckBox chkP 
         Caption         =   "Private"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   960
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox chkG 
         Caption         =   "Global"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   1200
         Value           =   1  'Checked
         Width           =   735
      End
      Begin VB.CheckBox chkA 
         Caption         =   "Admin"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   1440
         Value           =   1  'Checked
         Width           =   735
      End
      Begin VB.CommandButton Command60 
         Caption         =   "Save Logs"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   2040
         Width           =   975
      End
   End
   Begin VB.Timer PlayerTimer 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   8280
      Top             =   2520
   End
   Begin VB.Timer tmrShutdown 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   8760
      Top             =   2520
   End
   Begin VB.Timer tmrGameAI 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   9240
      Top             =   2520
   End
   Begin VB.Timer tmrSpawnMapItems 
      Interval        =   1000
      Left            =   9720
      Top             =   2520
   End
   Begin VB.Timer tmrPlayerSave 
      Interval        =   60000
      Left            =   7800
      Top             =   2520
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   4800
      Width           =   10335
      Begin VB.CommandButton Command1 
         Caption         =   "Shutdown"
         Height          =   255
         Left            =   8760
         TabIndex        =   6
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Exit"
         Height          =   255
         Left            =   7200
         TabIndex        =   5
         Top             =   240
         Width           =   1455
      End
      Begin VB.CheckBox GMOnly 
         Caption         =   "GMs Only"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   975
      End
      Begin VB.CheckBox Closed 
         Caption         =   "Closed"
         Height          =   255
         Left            =   1320
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox mnuServerLog 
         Caption         =   "Server Log"
         Height          =   255
         Left            =   3720
         TabIndex        =   2
         Top             =   240
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox chkChat 
         Caption         =   "Save Logs"
         Height          =   255
         Left            =   2400
         TabIndex        =   1
         Top             =   240
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.Label ShutdownTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Shutdown: Not Active"
         Height          =   195
         Left            =   5520
         TabIndex        =   7
         Top             =   240
         Width           =   1575
      End
   End
   Begin MSWinsockLib.Winsock Socket 
      Index           =   0
      Left            =   10200
      Top             =   2520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin TabDlg.SSTab SSTab2 
      Height          =   2400
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   4233
      _Version        =   393216
      Style           =   1
      Tabs            =   7
      TabsPerRow      =   7
      TabHeight       =   353
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Main"
      TabPicture(0)   =   "frmServer.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "txtChat"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "txtText(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Broadcast"
      TabPicture(1)   =   "frmServer.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtText(1)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Global"
      TabPicture(2)   =   "frmServer.frx":0902
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtText(2)"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Map"
      TabPicture(3)   =   "frmServer.frx":091E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "txtText(3)"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Private"
      TabPicture(4)   =   "frmServer.frx":093A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "txtText(4)"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Admin"
      TabPicture(5)   =   "frmServer.frx":0956
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "txtText(5)"
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "Emote"
      TabPicture(6)   =   "frmServer.frx":0972
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "txtText(6)"
      Tab(6).ControlCount=   1
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
         Height          =   1770
         Index           =   0
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   16
         Top             =   240
         Width           =   9075
      End
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
         TabIndex        =   15
         Top             =   2040
         Width           =   9075
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
         Height          =   2010
         Index           =   1
         Left            =   -74880
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   14
         Top             =   240
         Width           =   9075
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
         Height          =   2010
         Index           =   2
         Left            =   -74880
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Top             =   240
         Width           =   9075
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
         Height          =   2010
         Index           =   3
         Left            =   -74880
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Top             =   240
         Width           =   9075
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
         Height          =   2010
         Index           =   4
         Left            =   -74880
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   240
         Width           =   9075
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
         Height          =   2010
         Index           =   5
         Left            =   -74880
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   240
         Width           =   9075
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
         Height          =   2010
         Index           =   6
         Left            =   -74880
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   240
         Width           =   9075
      End
   End
   Begin MSComctlLib.ListView lvUsers 
      Height          =   1575
      Left            =   240
      TabIndex        =   33
      Top             =   2760
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   2778
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Index"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Account"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Character"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Level"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Sprite"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Access"
         Object.Width           =   1235
      EndProperty
   End
   Begin VB.Label TPO 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Players Online:"
      Height          =   195
      Left            =   240
      TabIndex        =   28
      Top             =   2520
      Width           =   1485
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Chat Log Save In"
      Height          =   195
      Left            =   120
      TabIndex        =   25
      Top             =   2520
      Visible         =   0   'False
      Width           =   10485
   End
   Begin VB.Menu mnuScripts 
      Caption         =   "Scripts"
      Begin VB.Menu mnuEScripts 
         Caption         =   "Enable Scripting"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuRScripts 
         Caption         =   "Reload"
      End
      Begin VB.Menu mnuEditScript 
         Caption         =   "Edit..."
      End
   End
   Begin VB.Menu mnuClasses 
      Caption         =   "Classes"
      Begin VB.Menu mnuRClasses 
         Caption         =   "Reload"
      End
      Begin VB.Menu mnuEClasses 
         Caption         =   "Edit..."
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnuEAdmin 
         Caption         =   "Admin Log"
      End
      Begin VB.Menu mnuEBanlist 
         Caption         =   "Ban List"
      End
      Begin VB.Menu mnuEPlayer 
         Caption         =   "Player Log"
      End
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CM As Long
Dim num As Long

Private Sub Check1_Click()
    If Check1.Value = Checked Then
        lvUsers.GridLines = True
    Else
        lvUsers.GridLines = False
    End If
End Sub

Private Sub Command1_Click()
If tmrShutdown.Enabled = False Then
    tmrShutdown.Enabled = True
End If
End Sub

Private Sub Command13_Click()
Dim index As Long
index = lvUsers.ListItems(lvUsers.SelectedItem.index).Text

If index > 0 Then
    If IsPlaying(index) Then
        Call GlobalMsg(GetPlayerName(index) & " has been kicked by the server!", White)
        Call AlertMsg(index, "You have been kicked by the server!")
    End If
End If
End Sub

Private Sub Command15_Click()
If lvUsers.ListItems(lvUsers.SelectedItem.index).Text > 0 Then
    If IsPlaying(lvUsers.ListItems(lvUsers.SelectedItem.index).Text) Then
        Call BanByServer(lvUsers.ListItems(lvUsers.SelectedItem.index).Text, "")
    End If
End If
End Sub

Private Sub Command2_Click()
    Call DestroyServer
End Sub

Private Sub Command20_Click()

End Sub

Private Sub Command22_Click()
Dim index As Long
index = lvUsers.ListItems(lvUsers.SelectedItem.index).Text

    Call PlayerMsg(index, "You have been muted!", White)
    Call TextAdd(frmServer.txtText(0), GetPlayerName(index) & " has been muted!", True)
    Player(index).Mute = True
End Sub

Private Sub Command23_Click()
Dim index As Long
index = lvUsers.ListItems(lvUsers.SelectedItem.index).Text

    Call PlayerMsg(index, "You have been unmuted!", White)
    Call TextAdd(frmServer.txtText(0), GetPlayerName(index) & " has been unmuted!", True)
    Player(index).Mute = False
End Sub

Private Sub Command28_Click()
    AFileName = "Scripts/Main.txt"
    Unload frmEditor
    frmEditor.Show
End Sub

Private Sub Command3_Click()
End Sub

Private Sub Command32_Click()

End Sub

Private Sub Command33_Click()

End Sub

Private Sub Command34_Click()
End Sub

Private Sub Command35_Click()
End Sub

Private Sub Command36_Click()
End Sub

Private Sub Command37_Click()
End Sub

Private Sub Command38_Click()
End Sub

Private Sub Command39_Click()
End Sub

Private Sub Command4_Click()
End Sub

Private Sub Command40_Click()
End Sub

Private Sub Command41_Click()
End Sub

Private Sub Command42_Click()
End Sub

Private Sub Command43_Click()
End Sub

Private Sub Command44_Click()
    AFileName = "player.txt"
    Unload frmEditor
    frmEditor.Show
End Sub

Private Sub Command45_Click()
End Sub

Private Sub Command5_Click()
End Sub

Private Sub Command46_Click()

End Sub

Private Sub Command47_Click()

End Sub

Private Sub Command58_Click()
    If GameTime = TIME_DAY Then
        GameTime = TIME_NIGHT
    ElseIf GameTime = TIME_NIGHT Then
        GameTime = TIME_DAY
    End If
    Call SendTimeToAll
End Sub

Private Sub Command59_Click()
End Sub

Private Sub Command6_Click()
End Sub

Private Sub Command60_Click()
    Call SaveLogs
End Sub

Private Sub Command61_Click()
End Sub

Private Sub Command62_Click()
    GameWeather = WEATHER_NONE
    Call SendWeatherToAll
End Sub

Private Sub Command63_Click()
    GameWeather = WEATHER_THUNDER
    Call SendWeatherToAll
End Sub

Private Sub Command64_Click()
    GameWeather = WEATHER_RAINING
    Call SendWeatherToAll
End Sub

Private Sub Command65_Click()
    GameWeather = WEATHER_SNOWING
    Call SendWeatherToAll
End Sub

Private Sub Command66_Click()
Dim i As Long

    Call RemovePLR
    
    For i = 1 To MAX_PLAYERS
        Call ShowPLR(i)
    Next i
End Sub

Private Sub Command7_Click()
Dim index As Long

End Sub

Private Sub Command8_Click()
End Sub

Private Sub CustomMsg_Click(index As Integer)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim lmsg As Long
    
    lmsg = x
    Select Case lmsg
        Case WM_LBUTTONDBLCLK
            frmServer.WindowState = vbNormal
            frmServer.Show
    End Select
End Sub
Private Sub mnuRestore_Click()
WindowState = vbNormal
Me.Show
End Sub
Private Sub Form_Resize()
   If frmServer.WindowState = vbMinimized Then
       frmServer.Hide
       Me.Refresh
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

Private Sub Form_Terminate()
    Call DestroyServer
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call DestroyServer
End Sub

Private Sub Label7_Click()
    Shell ("http://www.ipchicken.com"), vbNormalNoFocus
End Sub

Private Sub lstTopics_Click()
Dim FileName As String
Dim hFile As Long
End Sub

Private Sub mnuCond_Click()
    'frmEditCondition.Show
End Sub

Private Sub mnuEAdmin_Click()
    AFileName = "admin.txt"
    Unload frmEditor
    frmEditor.Show
End Sub

Private Sub mnuEBanlist_Click()
    AFileName = "banlist.txt"
    Unload frmEditor
    frmEditor.Show
End Sub

Private Sub mnuEClasses_Click()
    AFileName = "Classes\Info.ini"
    Unload frmEditor
    frmEditor.Show
End Sub

Private Sub mnuEditScript_Click()
    AFileName = "Scripts/Main.txt"
    Unload frmEditor
    frmEditor.Show
End Sub

Private Sub mnuEPlayer_Click()
    AFileName = "banlist.txt"
    Unload frmEditor
    frmEditor.Show
End Sub

Private Sub mnuEScripts_Click()

'Turn ON
If Scripting = 0 Then
    Scripting = 1
    PutVar App.Path & "\Data.ini", "CONFIG", "Scripting", 1
    
    If Scripting = 1 Then
        Set MyScript = New clsSadScript
        Set clsScriptCommands = New clsCommands
        MyScript.ReadInCode App.Path & "\Scripts\Main.txt", "Scripts\Main.txt", MyScript.SControl, False
        MyScript.SControl.AddObject "ScriptHardCode", clsScriptCommands, True
    End If
Else

'Turn OFF

    Scripting = 0
    PutVar App.Path & "\Data.ini", "CONFIG", "Scripting", 0
    
    If Scripting = 0 Then
        Set MyScript = Nothing
        Set clsScriptCommands = Nothing
    End If
End If

mnuEScripts.Checked = Scripting

End Sub

Private Sub mnuRClasses_Click()
    Call LoadClasses
    Call TextAdd(frmServer.txtText(0), "All classes reloaded.", True)
End Sub

Private Sub mnuRScripts_Click()
If Scripting = 1 Then
    Set MyScript = Nothing
    Set clsScriptCommands = Nothing
    
    Set MyScript = New clsSadScript
    Set clsScriptCommands = New clsCommands
    MyScript.ReadInCode App.Path & "\Scripts\Main.txt", "Scripts\Main.txt", MyScript.SControl, False
    MyScript.SControl.AddObject "ScriptHardCode", clsScriptCommands, True
    Call TextAdd(frmServer.txtText(0), "Scripts reloaded.", True)
End If
End Sub

Private Sub mnuServerLog_Click()
    If mnuServerLog.Value = Checked Then
        ServerLog = False
    Else
        ServerLog = True
    End If
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

Private Sub Say_Click(index As Integer)
    Call GlobalMsg(Trim(CMessages(index + 1).Message), White)
    Call TextAdd(frmServer.txtText(0), "Quick Msg: " & Trim(CMessages(index + 1).Message), True)
End Sub

Private Sub scrlMap_Change()
End Sub

Private Sub scrlMM_Change()
End Sub

Private Sub scrlMX_Change()
End Sub

Private Sub scrlMY_Change()
End Sub

Private Sub scrlRainIntensity_Change()
End Sub

Private Sub scrlX_Change()
End Sub

Private Sub scrlY_Change()
End Sub

Private Sub tmrChatLogs_Timer()
Static ChatSecs As Long
Dim SaveTime As Long

SaveTime = 3600

    If frmServer.chkChat.Value = Unchecked Then
        ChatSecs = SaveTime
        Label6.Caption = "Chat Log Save Disabled!"
        Exit Sub
    End If
    
    If ChatSecs <= 0 Then ChatSecs = SaveTime
    If ChatSecs > 60 Then
        Label6.Caption = "Chat Log Save In " & Int(ChatSecs / 60) & " Minute(s)"
    Else
        Label6.Caption = "Chat Log Save In " & Int(ChatSecs) & " Second(s)"
    End If
    
    ChatSecs = ChatSecs - 1
    
    If ChatSecs <= 0 Then
        Call TextAdd(txtText(0), "Chat Logs Have Been Saved!", True)
        Call SaveLogs
        ChatSecs = 0
    End If
End Sub

Private Sub tmrCycle_Timer()
End Sub

Private Sub tmrGameAI_Timer()
    Call ServerLogic
End Sub

Private Sub tmrPlayerSave_Timer()
    Call PlayerSaveTimer
End Sub

Private Sub tmrSpawnMapItems_Timer()
    Call CheckSpawnMapItems
End Sub

Private Sub txtChat_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And Trim(txtChat.Text) <> "" Then
        Call GlobalMsg(txtChat.Text, White)
        Call TextAdd(frmServer.txtText(0), "Server: " & txtChat.Text, True)
        txtChat.Text = ""
    End If
    If KeyAscii = vbKeyReturn Then KeyAscii = 0
End Sub

Private Sub tmrShutdown_Timer()
Static Secs As Long

    If Secs <= 0 Then Secs = 30
    ShutdownTime.Caption = "Shutdown: " & Secs & " Seconds"
    If Secs = 30 Then Call TextAdd(frmServer.txtText(0), "Automated Server Shutdown in " & Secs & " seconds.", True)
    If Secs = 30 Then Call GlobalMsg("Server Shutdown in " & Secs & " seconds.", BrightBlue)
    If Secs = 25 Then Call GlobalMsg("Server Shutdown in " & Secs & " seconds.", BrightBlue)
    If Secs = 20 Then Call GlobalMsg("Server Shutdown in " & Secs & " seconds.", BrightBlue)
    If Secs = 15 Then Call GlobalMsg("Server Shutdown in " & Secs & " seconds.", BrightBlue)
    If Secs = 10 Then Call GlobalMsg("Server Shutdown in " & Secs & " seconds.", BrightBlue)
    If Secs < 6 Then
        Call GlobalMsg("Server Shutdown in " & Secs & " seconds.", BrightBlue)
    End If
    
    Secs = Secs - 1
    If Secs <= 0 Then
        tmrShutdown.Enabled = False
        Call DestroyServer
    End If
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

Private Sub txtText_GotFocus(index As Integer)
    txtChat.SetFocus
End Sub
