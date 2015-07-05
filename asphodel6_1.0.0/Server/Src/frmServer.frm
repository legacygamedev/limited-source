VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "msinet.ocx"
Begin VB.Form frmServer 
   AutoRedraw      =   -1  'True
   Caption         =   "Asphodel Server <Loading...>"
   ClientHeight    =   4095
   ClientLeft      =   60
   ClientTop       =   675
   ClientWidth     =   6735
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Fixedsys"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmServer.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4095
   ScaleWidth      =   6735
   StartUpPosition =   2  'CenterScreen
   Begin InetCtlsObjects.Inet INet 
      Index           =   0
      Left            =   5760
      Top             =   3480
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Socket 
      Index           =   0
      Left            =   6240
      Top             =   3600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3615
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   6376
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   503
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Console"
      TabPicture(0)   =   "frmServer.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "txtText"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "txtChat"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Players"
      TabPicture(1)   =   "frmServer.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lstPlayers"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Control"
      TabPicture(2)   =   "frmServer.frx":0902
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "chkStaffOnly"
      Tab(2).Control(1)=   "chkEmoteChat"
      Tab(2).Control(2)=   "chkPrivateChat"
      Tab(2).Control(3)=   "chkGlobalChat"
      Tab(2).Control(4)=   "chkMapChat"
      Tab(2).Control(5)=   "scrlSellBack"
      Tab(2).Control(6)=   "scrlLevelLimit"
      Tab(2).Control(7)=   "chkPVPLevel"
      Tab(2).Control(8)=   "chkAdminSafety"
      Tab(2).Control(9)=   "lblSellBack"
      Tab(2).Control(10)=   "Label1"
      Tab(2).Control(11)=   "lblLimit"
      Tab(2).ControlCount=   12
      TabCaption(3)   =   "News Box"
      TabPicture(3)   =   "frmServer.frx":091E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "cmdLoadOld"
      Tab(3).Control(1)=   "cmdSaveNews"
      Tab(3).Control(2)=   "txtNews"
      Tab(3).ControlCount=   3
      TabCaption(4)   =   "Ban List"
      TabPicture(4)   =   "frmServer.frx":093A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "cmdRemoveAccount"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "cmdRemoveIP"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "cmdAddAccount"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "txtAccountBan"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).Control(4)=   "lstAccountBans"
      Tab(4).Control(4).Enabled=   0   'False
      Tab(4).Control(5)=   "cmdAddIP"
      Tab(4).Control(5).Enabled=   0   'False
      Tab(4).Control(6)=   "txtIPBan"
      Tab(4).Control(6).Enabled=   0   'False
      Tab(4).Control(7)=   "lstIPBans"
      Tab(4).Control(7).Enabled=   0   'False
      Tab(4).Control(8)=   "Label5"
      Tab(4).Control(8).Enabled=   0   'False
      Tab(4).Control(9)=   "lblAccountBans"
      Tab(4).Control(9).Enabled=   0   'False
      Tab(4).Control(10)=   "Line1"
      Tab(4).Control(10).Enabled=   0   'False
      Tab(4).Control(11)=   "Label3"
      Tab(4).Control(11).Enabled=   0   'False
      Tab(4).Control(12)=   "lblIPBans"
      Tab(4).Control(12).Enabled=   0   'False
      Tab(4).ControlCount=   13
      Begin VB.CommandButton cmdRemoveAccount 
         Caption         =   "Remove Account Ban"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -71280
         TabIndex        =   31
         Top             =   2160
         Width           =   2655
      End
      Begin VB.CommandButton cmdRemoveIP 
         Caption         =   "Remove IP Ban"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   30
         Top             =   2160
         Width           =   2655
      End
      Begin VB.CommandButton cmdAddAccount 
         Caption         =   "Add Account Ban"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -71280
         TabIndex        =   29
         Top             =   2880
         Width           =   2655
      End
      Begin VB.TextBox txtAccountBan 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -70800
         TabIndex        =   27
         Top             =   2520
         Width           =   2175
      End
      Begin VB.ListBox lstAccountBans 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1500
         Left            =   -71280
         TabIndex        =   26
         Top             =   600
         Width           =   2655
      End
      Begin VB.CommandButton cmdAddIP 
         Caption         =   "Add IP Ban"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74760
         TabIndex        =   24
         Top             =   2880
         Width           =   2655
      End
      Begin VB.TextBox txtIPBan 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -74400
         TabIndex        =   22
         Top             =   2520
         Width           =   2295
      End
      Begin VB.ListBox lstIPBans 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1500
         Left            =   -74760
         TabIndex        =   21
         Top             =   600
         Width           =   2655
      End
      Begin VB.ListBox lstPlayers 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2970
         Left            =   -74880
         TabIndex        =   19
         Top             =   480
         Width           =   6495
      End
      Begin VB.CheckBox chkStaffOnly 
         Caption         =   "Staff only allowed to login"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   18
         Top             =   1920
         Width           =   3615
      End
      Begin VB.CheckBox chkEmoteChat 
         Caption         =   "Emote Chat"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -70080
         TabIndex        =   17
         Top             =   1560
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.CheckBox chkPrivateChat 
         Caption         =   "Private Chat"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -71880
         TabIndex        =   16
         Top             =   1560
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.CheckBox chkGlobalChat 
         Caption         =   "Global Chat"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -73560
         TabIndex        =   15
         Top             =   1560
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.CheckBox chkMapChat 
         Caption         =   "Map Chat"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   14
         Top             =   1560
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.CommandButton cmdLoadOld 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -73080
         TabIndex        =   13
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton cmdSaveNews 
         Caption         =   "Save News"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   12
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txtNews 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2775
         Left            =   -74880
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   720
         Width           =   6495
      End
      Begin VB.HScrollBar scrlSellBack 
         Height          =   255
         Left            =   -73680
         Max             =   100
         Min             =   1
         TabIndex        =   8
         Top             =   1200
         Value           =   50
         Width           =   3255
      End
      Begin VB.HScrollBar scrlLevelLimit 
         Height          =   255
         Left            =   -73200
         Max             =   100
         Min             =   1
         TabIndex        =   6
         Top             =   840
         Value           =   10
         Width           =   3255
      End
      Begin VB.CheckBox chkPVPLevel 
         Caption         =   "PVP Level Limit:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   5
         Top             =   840
         Value           =   1  'Checked
         Width           =   2295
      End
      Begin VB.CheckBox chkAdminSafety 
         Caption         =   "Admin Safety"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   4
         Top             =   480
         Value           =   1  'Checked
         Width           =   2295
      End
      Begin VB.TextBox txtChat 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         TabIndex        =   2
         Top             =   3120
         Width           =   6495
      End
      Begin VB.TextBox txtText 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2775
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   360
         Width           =   6495
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Acc:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -71280
         TabIndex        =   28
         Top             =   2570
         Width           =   375
      End
      Begin VB.Label lblAccountBans 
         Alignment       =   2  'Center
         Caption         =   "Account Bans"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -71280
         TabIndex        =   25
         Top             =   360
         Width           =   2655
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         X1              =   -71640
         X2              =   -71640
         Y1              =   360
         Y2              =   3360
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "IP:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   23
         Top             =   2570
         Width           =   375
      End
      Begin VB.Label lblIPBans 
         Alignment       =   2  'Center
         Caption         =   "IP Bans"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   20
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label lblSellBack 
         AutoSize        =   -1  'True
         Caption         =   "50% of worth"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   -70320
         TabIndex        =   10
         Top             =   1200
         Width           =   1170
      End
      Begin VB.Label Label1 
         Caption         =   "Sell Item For:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   9
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label lblLimit 
         AutoSize        =   -1  'True
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   -69720
         TabIndex        =   7
         Top             =   840
         Width           =   210
      End
   End
   Begin VB.Label lblServer 
      Alignment       =   2  'Center
      Caption         =   "Loading..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   120
      Width           =   6735
   End
   Begin VB.Menu mnuServer 
      Caption         =   "&Server"
      Enabled         =   0   'False
      Begin VB.Menu mnuMinimize 
         Caption         =   "Minimize"
      End
      Begin VB.Menu mnuShutDown 
         Caption         =   "Shut Down"
         Begin VB.Menu mnuShutDownAutomated 
            Caption         =   "Automated"
         End
         Begin VB.Menu mnuShutDownForced 
            Caption         =   "Forced"
         End
      End
      Begin VB.Menu mnuConsoleLog 
         Caption         =   "Console Log"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuDatabase 
      Caption         =   "&Database"
      Enabled         =   0   'False
      Begin VB.Menu mnuReload 
         Caption         =   "Reload"
         Begin VB.Menu mnuReloadAll 
            Caption         =   "All"
         End
         Begin VB.Menu mnuReloadSpells 
            Caption         =   "Spells"
         End
         Begin VB.Menu mnuReloadItems 
            Caption         =   "Items"
         End
         Begin VB.Menu mnuReloadNPCs 
            Caption         =   "NPCs"
         End
         Begin VB.Menu mnuReloadShops 
            Caption         =   "Shops"
         End
         Begin VB.Menu mnuReloadMaps 
            Caption         =   "Maps"
         End
         Begin VB.Menu mnuReloadClasses 
            Caption         =   "Classes"
         End
         Begin VB.Menu mnuReloadAnims 
            Caption         =   "Animations"
         End
      End
   End
   Begin VB.Menu mnuKick 
      Caption         =   "&Kick"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnuMute 
         Caption         =   "Mute"
      End
      Begin VB.Menu mnuKickPlayer 
         Caption         =   "Kick"
      End
      Begin VB.Menu mnuDisconnectPlayer 
         Caption         =   "Disconnect"
      End
      Begin VB.Menu mnuBanPlayer 
         Caption         =   "Ban"
      End
      Begin VB.Menu mnuRemoveAdmin 
         Caption         =   "Remove Admin"
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

' ------------------------------------------
' --              Asphodel 6              --
' ------------------------------------------

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
' ** Handling Form  **
' ********************

' this will auto adjust all the controls
Private Sub Form_Resize()

    If Me.WindowState = vbMinimized Then Exit Sub
    
    If Me.Height <= 3000 Then Me.Height = 2700
    
    Me.lblServer.Width = Me.Width
    
    SSTab1.Width = Me.Width - 240
    SSTab1.Height = Me.Height - 1350
    
    txtChat.Width = SSTab1.Width - 240
    txtText.Width = txtChat.Width
    txtNews.Width = txtChat.Width
    
    txtText.Height = SSTab1.Height - 840
    txtChat.Top = (txtText.Top + txtText.Height) + 30
    txtNews.Height = SSTab1.Height - 840
    
    lstPlayers.Height = SSTab1.Height - 480
    lstPlayers.Width = SSTab1.Width - 240
    
    Line1.Y2 = SSTab1.Height - 255
    
    If SSTab1.Height - 2115 > 150 Then
        lstAccountBans.Height = SSTab1.Height - 2115
        lstIPBans.Height = lstAccountBans.Height
        cmdRemoveAccount.Top = (lstAccountBans.Top + lstAccountBans.Height) + 60
        cmdRemoveIP.Top = cmdRemoveAccount.Top
        txtAccountBan.Top = cmdRemoveAccount.Top + 360
        txtIPBan.Top = txtAccountBan.Top
        Label5.Top = txtAccountBan.Top + 50
        Label3.Top = txtIPBan.Top + 50
        cmdAddAccount.Top = txtAccountBan.Top + 360
        cmdAddIP.Top = cmdAddAccount.Top
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mnuShutDownForced_Click
End Sub

Private Sub Form_Terminate()
    mnuShutDownForced_Click
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim lmsg As Long
   
    lmsg = X / Screen.TwipsPerPixelX
    Select Case lmsg
    
        Case WM_LBUTTONDBLCLK
        
            Me.Show
            txtText.SelStart = Len(txtText.Text)
            
            Me.Refresh
            lstPlayers.Refresh
            txtChat.Refresh
            txtText.Refresh
            txtNews.Refresh
            
    End Select
    
End Sub

Private Sub lblServer_Click()
    If lblServer.Caption <> "Loading..." Then
        Clipboard.Clear
        Clipboard.SetText ACTUAL_IP
        Call TextAdd(txtText, "IP Address copied to clipboard.")
    End If
End Sub

' ********************
' ** Handle Listbox **
' ********************

Private Sub lstPlayers_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then PopupMenu mnuKick
End Sub

' ********************
' *  Handle Ban List *
' ********************

Private Sub cmdAddAccount_Click()
Dim FileName As String
Dim BIndex As Long
Dim F As Long

    FileName = App.Path & "\data\bans.ini"
    
    ' Make sure the file exists
    If Not FileExist("data\bans.ini") Then
        F = FreeFile
        Open FileName For Output As #F
        Close #F
    End If
    
    If LenB(Trim$(txtAccountBan.Text)) > 0 Then
        BIndex = Val(GetVar(FileName, "ACCOUNT", "Total")) + 1
        PutVar FileName, "ACCOUNT", "Total", CStr(BIndex)
        PutVar FileName, "ACCOUNT", "Account" & BIndex, txtAccountBan.Text
        txtAccountBan.Text = vbNullString
        Load_BanTable
    End If
    
End Sub

Private Sub cmdAddIP_Click()
Dim FileName As String
Dim BIndex As Long
Dim F As Long

    FileName = App.Path & "\data\bans.ini"
    
    ' Make sure the file exists
    If Not FileExist("data\bans.ini") Then
        F = FreeFile
        Open FileName For Output As #F
        Close #F
    End If
    
    If LenB(Trim$(txtIPBan.Text)) > 0 Then
        If Not IsIP(txtIPBan.Text) Then
            MsgBox "Not a valid IP!", vbOKOnly + vbCritical, GAME_NAME
            Exit Sub
        End If
        BIndex = Val(GetVar(FileName, "IP", "Total")) + 1
        PutVar FileName, "IP", "Total", CStr(BIndex)
        PutVar FileName, "IP", "IP" & BIndex, txtIPBan.Text
        txtIPBan.Text = vbNullString
        Load_BanTable
    End If
    
End Sub

Private Sub cmdRemoveIP_Click()

    If lstIPBans.ListIndex >= 0 Then
        If LenB(Trim$(lstIPBans.List(lstIPBans.ListIndex))) < 1 Then Exit Sub
        lstIPBans.List(lstIPBans.ListIndex) = vbNullString
        Dim FileName As String
        FileName = App.Path & "\data\bans.ini"
        If lstIPBans.ListIndex + 1 = Val(GetVar(FileName, "IP", "Total")) Then
            PutVar FileName, "IP", "Total", Val(GetVar(FileName, "IP", "Total")) - 1
        End If
        PutVar FileName, "IP", "IP" & lstIPBans.ListIndex + 1, vbNullString
        lstIPBans.ListIndex = -1
        Load_BanTable
    End If
    
End Sub

Private Sub cmdRemoveAccount_Click()

    If lstAccountBans.ListIndex >= 0 Then
        If LenB(Trim$(lstAccountBans.List(lstAccountBans.ListIndex))) < 1 Then Exit Sub
        lstAccountBans.List(lstAccountBans.ListIndex) = vbNullString
        Dim FileName As String
        FileName = App.Path & "\data\bans.ini"
        If lstAccountBans.ListIndex + 1 = Val(GetVar(FileName, "ACCOUNT", "Total")) Then
            PutVar FileName, "ACCOUNT", "Total", Val(GetVar(FileName, "ACCOUNT", "Total")) - 1
        End If
        PutVar FileName, "ACCOUNT", "Account" & lstAccountBans.ListIndex + 1, vbNullString
        lstAccountBans.ListIndex = -1
        Load_BanTable
    End If
    
End Sub

' ********************
' **    News box    **
' ********************

Private Sub cmdLoadOld_Click()
    txtNews.Text = GAME_NEWS
End Sub

Private Sub cmdSaveNews_Click()
    GAME_NEWS = txtNews.Text
    PutVar App.Path & "\data\news.ini", "CONTENT", "News", GAME_NEWS
End Sub

' ********************
' **    Chat box    **
' ********************

Private Sub txtText_GotFocus()
On Error Resume Next

    txtChat.SetFocus
    
End Sub

Private Sub txtChat_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If LenB(Trim$(txtChat.Text)) > 0 Then
            Call GlobalMsg(txtChat.Text, Color.White)
            Call TextAdd(txtText, "Server: " & txtChat.Text)
            txtChat.Text = vbNullString
        End If
        KeyAscii = 0
    End If
End Sub

' ********************
' ** Server Control **
' ********************

Private Sub scrlLevelLimit_Change()
    lblLimit.Caption = scrlLevelLimit.Value
    PutVar App.Path & "\data\config.ini", "SETUP", "PVP_Level", CStr(scrlLevelLimit.Value)
End Sub

Private Sub chkEmoteChat_Click()
    PutVar App.Path & "\data\config.ini", "SETUP", "EmoteChat", CStr(chkEmoteChat.Value)
End Sub

Private Sub chkGlobalChat_Click()
    PutVar App.Path & "\data\config.ini", "SETUP", "GlobalChat", CStr(chkGlobalChat.Value)
End Sub

Private Sub chkMapChat_Click()
    PutVar App.Path & "\data\config.ini", "SETUP", "MapChat", CStr(chkMapChat.Value)
End Sub

Private Sub chkPrivateChat_Click()
    PutVar App.Path & "\data\config.ini", "SETUP", "PrivateChat", CStr(chkPrivateChat.Value)
End Sub

Private Sub chkStaffOnly_Click()
    PutVar App.Path & "\data\config.ini", "SETUP", "StaffOnly", CStr(chkStaffOnly.Value)
End Sub

Private Sub chkPVPLevel_Click()

    If chkPVPLevel.Value = 0 Then
        scrlLevelLimit.Enabled = False
        lblLimit.Enabled = False
    Else
        scrlLevelLimit.Enabled = True
        lblLimit.Enabled = True
    End If
    
    PutVar App.Path & "\data\config.ini", "SETUP", "PVP_LevelOn", CStr(chkPVPLevel.Value)
    
End Sub

Private Sub chkAdminSafety_Click()
    PutVar App.Path & "\data\config.ini", "SETUP", "Staff_Safe", CStr(chkAdminSafety.Value)
End Sub

Private Sub scrlSellBack_Change()
    lblSellBack.Caption = scrlSellBack.Value & "% of worth"
    PutVar App.Path & "\data\config.ini", "SETUP", "SellBack", CStr(scrlSellBack.Value)
End Sub

Private Sub scrlSellBack_Scroll()
    scrlSellBack_Change
End Sub

Private Sub scrlLevelLimit_Scroll()
    scrlLevelLimit_Change
End Sub

' ********************
' **  Handle Menus  **
' ********************

Private Sub mnuReloadAnims_Click()
Dim i As Long
Dim SplitString() As String
Dim FileName As String

    FileName = App.Path & "\data\config.ini"
    
    TOTAL_ANIMFRAMES = Val(GetVar(FileName, "ANIMATION", "Total_AnimFrames"))
    CONFIG_STANDFRAME = Val(GetVar(FileName, "ANIMATION", "StandFrame"))
    
    If GetVar(FileName, "ANIMATION", "WalkFrames") = vbNullString Then
        MsgBox "You need to make sure you have walk frames specified in the Data\config.ini!", vbOKOnly + vbCritical, "Error"
        Exit Sub
    End If
    
    If Not InStr(1, GetVar(FileName, "ANIMATION", "WalkFrames"), ",", vbTextCompare) > 0 Then
        ReDim WalkFrame(1 To 1)
        TOTAL_WALKFRAMES = 1
        WalkFrame(1) = Val(GetVar(FileName, "ANIMATION", "WalkFrames"))
        GoTo Skip1
    End If
    
    SplitString = Split(GetVar(FileName, "ANIMATION", "WalkFrames"), ",", , vbTextCompare)
    
    ReDim WalkFrame(1 To UBound(SplitString) + 1)
    TOTAL_WALKFRAMES = UBound(WalkFrame)
    
    For i = 0 To UBound(SplitString)
        WalkFrame(i + 1) = Val(SplitString(i)) - 1
    Next
    
Skip1:
    
    If GetVar(FileName, "ANIMATION", "AttackFrames") = vbNullString Then GoTo Skip2
    
    If Not InStr(1, GetVar(FileName, "ANIMATION", "AttackFrames"), ",", vbTextCompare) > 0 Then
        ReDim AttackFrame(1 To 1)
        TOTAL_ATTACKFRAMES = 1
        AttackFrame(1) = Val(GetVar(FileName, "ANIMATION", "AttackFrames"))
        GoTo Skip2
    End If
    
    SplitString = Split(GetVar(FileName, "ANIMATION", "AttackFrames"), ",", , vbTextCompare)
    
    ReDim AttackFrame(1 To UBound(SplitString) + 1)
    TOTAL_ATTACKFRAMES = UBound(AttackFrame)
    
    For i = 0 To UBound(SplitString)
        AttackFrame(i + 1) = Val(SplitString(i)) - 1
    Next
    
Skip2:
    
    Direction_Anim(E_Direction.Up_) = Val(GetVar(FileName, "ANIMATION", "Anim_Up"))
    Direction_Anim(E_Direction.Down_) = Val(GetVar(FileName, "ANIMATION", "Anim_Down"))
    Direction_Anim(E_Direction.Left_) = Val(GetVar(FileName, "ANIMATION", "Anim_Left"))
    Direction_Anim(E_Direction.Right_) = Val(GetVar(FileName, "ANIMATION", "Anim_Right"))
    
    WALKANIM_SPEED = Val(GetVar(FileName, "ANIMATION", "WalkAnim_Speed"))
    
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            SendGameOptions i
            SendPlayerData i
        End If
    Next
    
End Sub

Private Sub mnuMute_Click()
Dim UseIndex As Long

    UseIndex = lstPlayers.ListIndex + 1
    
    If Not Player(UseIndex).Char(TempPlayer(UseIndex).CharNum).Muted Then
        Player(UseIndex).Char(TempPlayer(UseIndex).CharNum).MuteTime = GetTickCountNew + 300000
        Player(UseIndex).Char(TempPlayer(UseIndex).CharNum).Muted = True
        
        PlayerMsg UseIndex, "You have been muted for 5 minutes by the server!", Color.BrightRed
        TextAdd txtText, "You have muted " & GetPlayerName(UseIndex) & " for 5 minutes."
    Else
        Player(UseIndex).Char(TempPlayer(UseIndex).CharNum).Muted = False
        Player(UseIndex).Char(TempPlayer(UseIndex).CharNum).MuteTime = 0
        
        PlayerMsg UseIndex, "You have been unmuted by the server!", Color.BrightRed
        TextAdd txtText, "You have unmuted " & GetPlayerName(UseIndex) & "."
    End If
    
    UpdatePlayerTable lstPlayers.ListIndex + 1
    
End Sub

Private Sub mnuKickPlayer_Click()

    If lstPlayers.List(lstPlayers.ListIndex) <> lstPlayers.ListIndex + 1 & ") None" Then
        Call AlertMsg(lstPlayers.ListIndex + 1, "You have been kicked by the server!")
    End If
    
End Sub

Private Sub mnuDisconnectPlayer_Click()

    If lstPlayers.List(lstPlayers.ListIndex) <> lstPlayers.ListIndex + 1 & ") None" Then
        CloseSocket lstPlayers.ListIndex + 1
    End If
    
End Sub

Private Sub mnuBanPlayer_click()

    If lstPlayers.List(lstPlayers.ListIndex) <> lstPlayers.ListIndex + 1 & ") None" Then Call ServerBanIndex(lstPlayers.ListIndex + 1)
    
End Sub

Private Sub mnuAdminPlayer_click()

    If lstPlayers.List(lstPlayers.ListIndex) <> lstPlayers.ListIndex + 1 & ") None" Then
        Call SetPlayerAccess(lstPlayers.ListIndex + 1, 4)
        Call SendPlayerData(lstPlayers.ListIndex + 1)
        Call PlayerMsg(lstPlayers.ListIndex + 1, "You have been granted administrator access.", Color.Pink)
        UpdatePlayerTable lstPlayers.ListIndex + 1
    End If
    
End Sub

Private Sub mnuRemoveAdmin_click()

    If lstPlayers.List(lstPlayers.ListIndex) <> lstPlayers.ListIndex + 1 & ") None" Then
        Call SetPlayerAccess(lstPlayers.ListIndex + 1, 0)
        Call SendPlayerData(lstPlayers.ListIndex + 1)
        Call PlayerMsg(lstPlayers.ListIndex + 1, "You have lost administrator access.", Color.Pink)
        UpdatePlayerTable lstPlayers.ListIndex + 1
    End If
    
End Sub

Private Sub mnuConsoleLog_Click()

    mnuConsoleLog.Checked = Not mnuConsoleLog.Checked
    ServerLog = mnuConsoleLog.Checked
    Call TextAdd(txtText, "Server Logging: " & ServerLog)
    
    If mnuConsoleLog.Checked Then PutVar App.Path & "\data\config.ini", "SETUP", "Logging", CStr(1) Else PutVar App.Path & "\data\config.ini", "SETUP", "Logging", CStr(0)
    
End Sub

Private Sub mnuShutDownAutomated_Click()
    
    mnuShutDownAutomated.Checked = Not mnuShutDownAutomated.Checked
    isShuttingDown = mnuShutDownAutomated.Checked
    
    If Not isShuttingDown Then
        Secs = -1
        Call TextAdd(txtText, "Automated shut down canceled.")
        Call GlobalMsg("Server shut down has been cancelled.", Color.BrightBlue)
    End If
    
End Sub

Private Sub mnuShutDownForced_Click()
    Call GlobalMsg("Server shut down! Good bye.", Color.BrightRed)
    ServerOnline = False
End Sub

Private Sub mnuReloadClasses_Click()
    Call LoadClasses
    Call TextAdd(txtText, "All classes reloaded.")
End Sub

Private Sub mnuReloadMaps_Click()
    Call LoadMaps
    Call TextAdd(txtText, "All maps reloaded.")
End Sub

Private Sub mnuReloadItems_Click()
    Call LoadItems
    Call TextAdd(txtText, "All items reloaded.")
End Sub

Private Sub mnuReloadNPCs_Click()
    Call LoadNpcs
    Call TextAdd(txtText, "All npcs reloaded.")
End Sub

Private Sub mnuReloadShops_Click()
    Call LoadShops
    Call TextAdd(txtText, "All shops reloaded.")
End Sub

Private Sub mnuReloadSpells_Click()
    Call LoadSpells
    Call TextAdd(txtText, "All spells reloaded.")
End Sub

Private Sub mnuReloadAll_Click()
    Call mnuReloadClasses_Click
    Call mnuReloadMaps_Click
    Call mnuReloadItems_Click
    Call mnuReloadNPCs_Click
    Call mnuReloadShops_Click
    Call mnuReloadSpells_Click
    Call TextAdd(txtText, "All database items fully reloaded.")
End Sub

Private Sub mnuMinimize_Click()
    Me.Hide
End Sub
