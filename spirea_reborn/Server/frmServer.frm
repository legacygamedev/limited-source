VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmServer 
   Caption         =   "Mirage Server"
   ClientHeight    =   2775
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   6495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picMap 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   600
      ScaleHeight     =   143
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   351
      TabIndex        =   17
      Top             =   240
      Visible         =   0   'False
      Width           =   5295
      Begin VB.CommandButton Command41 
         Caption         =   "Cancel"
         Height          =   315
         Left            =   120
         TabIndex        =   19
         Top             =   1800
         Width           =   735
      End
      Begin VB.ListBox lstNPC 
         Height          =   1815
         Left            =   3480
         TabIndex        =   18
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label MapInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Map"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   33
         Top             =   0
         Width           =   300
      End
      Begin VB.Label MapInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Revision:"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   32
         Top             =   240
         Width           =   660
      End
      Begin VB.Label MapInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Moral:"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   31
         Top             =   480
         Width           =   450
      End
      Begin VB.Label MapInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Up:"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   30
         Top             =   720
         Width           =   255
      End
      Begin VB.Label MapInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Down:"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   29
         Top             =   960
         Width           =   465
      End
      Begin VB.Label MapInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Left:"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   28
         Top             =   1200
         Width           =   345
      End
      Begin VB.Label MapInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Right:"
         Height          =   195
         Index           =   6
         Left            =   1800
         TabIndex        =   27
         Top             =   240
         Width           =   435
      End
      Begin VB.Label MapInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Music:"
         Height          =   195
         Index           =   7
         Left            =   1800
         TabIndex        =   26
         Top             =   480
         Width           =   450
      End
      Begin VB.Label MapInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BootMap:"
         Height          =   195
         Index           =   8
         Left            =   1800
         TabIndex        =   25
         Top             =   720
         Width           =   690
      End
      Begin VB.Label MapInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BootX:"
         Height          =   195
         Index           =   9
         Left            =   1800
         TabIndex        =   24
         Top             =   960
         Width           =   480
      End
      Begin VB.Label MapInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BootY:"
         Height          =   195
         Index           =   10
         Left            =   1800
         TabIndex        =   23
         Top             =   1200
         Width           =   480
      End
      Begin VB.Label MapInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Shop:"
         Height          =   195
         Index           =   11
         Left            =   120
         TabIndex        =   22
         Top             =   2640
         Width           =   420
      End
      Begin VB.Label MapInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Indoors:"
         Height          =   195
         Index           =   12
         Left            =   120
         TabIndex        =   21
         Top             =   2880
         Width           =   615
      End
      Begin VB.Label MapInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NPCs"
         Height          =   195
         Index           =   13
         Left            =   3000
         TabIndex        =   20
         Top             =   120
         Width           =   375
      End
   End
   Begin VB.PictureBox picChangeInfo 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   2055
      Left            =   3720
      ScaleHeight     =   2025
      ScaleWidth      =   2625
      TabIndex        =   48
      Top             =   360
      Visible         =   0   'False
      Width           =   2655
      Begin VB.CommandButton cmdsetMAGI 
         Caption         =   "Change Magic"
         Height          =   255
         Left            =   1200
         TabIndex        =   61
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox txtMagi 
         Height          =   285
         Left            =   120
         TabIndex        =   60
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CommandButton cmdSSPEED 
         Caption         =   "Change Speed"
         Height          =   255
         Left            =   1200
         TabIndex        =   59
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox txtSpeed 
         Height          =   285
         Left            =   120
         TabIndex        =   58
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CommandButton cmdSDEF 
         Caption         =   "Change Defense"
         Height          =   255
         Left            =   1200
         TabIndex        =   57
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox txtDef 
         Height          =   285
         Left            =   120
         TabIndex        =   56
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton cmdSSSTR 
         Caption         =   "Change Strength"
         Height          =   255
         Left            =   1200
         TabIndex        =   55
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txtStr 
         Height          =   285
         Left            =   120
         TabIndex        =   54
         Top             =   600
         Width           =   1095
      End
      Begin VB.CommandButton cmdSA 
         Caption         =   "Change Access"
         Height          =   255
         Left            =   1200
         TabIndex        =   53
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox txtAccess 
         Height          =   285
         Left            =   120
         TabIndex        =   52
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdCL 
         Caption         =   "Change Level"
         Height          =   255
         Left            =   1200
         TabIndex        =   51
         Top             =   120
         Width           =   1335
      End
      Begin VB.TextBox txtlevel 
         Height          =   285
         Left            =   120
         TabIndex        =   50
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdClosepic 
         Caption         =   "Close"
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   1680
         Width           =   975
      End
   End
   Begin VB.PictureBox picExp 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   3600
      ScaleHeight     =   55
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   191
      TabIndex        =   43
      Top             =   1920
      Visible         =   0   'False
      Width           =   2895
      Begin VB.TextBox txtExp 
         Appearance      =   0  'Flat
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
         Left            =   960
         TabIndex        =   45
         Top             =   120
         Width           =   1875
      End
      Begin VB.CommandButton Command39 
         Caption         =   "Cancel"
         Height          =   315
         Left            =   2040
         TabIndex        =   46
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton Command40 
         Caption         =   "Execute"
         Height          =   315
         Left            =   1200
         TabIndex        =   44
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Experience:"
         Height          =   255
         Left            =   0
         TabIndex        =   47
         Top             =   120
         Width           =   855
      End
   End
   Begin VB.PictureBox picWarp 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   63
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   215
      TabIndex        =   34
      Top             =   1800
      Visible         =   0   'False
      Width           =   3255
      Begin VB.CommandButton Command37 
         Caption         =   "Warp"
         Height          =   255
         Left            =   2520
         TabIndex        =   39
         Top             =   120
         Width           =   615
      End
      Begin VB.HScrollBar scrlMM 
         Height          =   255
         Left            =   960
         Min             =   1
         TabIndex        =   38
         Top             =   120
         Value           =   1
         Width           =   1455
      End
      Begin VB.HScrollBar scrlMX 
         Height          =   255
         Left            =   960
         TabIndex        =   37
         Top             =   360
         Width           =   1455
      End
      Begin VB.HScrollBar scrlMY 
         Height          =   255
         Left            =   960
         TabIndex        =   36
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton Command38 
         Caption         =   "Cancel"
         Height          =   255
         Left            =   2520
         TabIndex        =   35
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblMM 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Map: 1"
         Height          =   195
         Left            =   120
         TabIndex        =   42
         Top             =   120
         Width           =   495
      End
      Begin VB.Label lblMX 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "X: 0"
         Height          =   195
         Left            =   120
         TabIndex        =   41
         Top             =   360
         Width           =   285
      End
      Begin VB.Label lblMY 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Y: 0"
         Height          =   195
         Left            =   120
         TabIndex        =   40
         Top             =   600
         Width           =   285
      End
   End
   Begin VB.Timer tmrShutdown 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   0
      Top             =   2760
   End
   Begin VB.Timer tmrGameAI 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   1080
      Top             =   2760
   End
   Begin VB.Timer tmrSpawnMapItems 
      Interval        =   1000
      Left            =   720
      Top             =   2760
   End
   Begin VB.Timer tmrPlayerSave 
      Interval        =   60000
      Left            =   360
      Top             =   2760
   End
   Begin MSWinsockLib.Winsock Socket 
      Index           =   0
      Left            =   1440
      Top             =   2760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2775
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   4895
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Console"
      TabPicture(0)   =   "frmServer.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "txtChat"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "txtText"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "PlayerTimer"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Players"
      TabPicture(1)   =   "frmServer.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Command5"
      Tab(1).Control(1)=   "lvwInfo"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Guide"
      TabPicture(2)   =   "frmServer.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lstTopics"
      Tab(2).Control(1)=   "TopicTitle"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Commands"
      TabPicture(3)   =   "frmServer.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame6"
      Tab(3).Control(1)=   "Frame9"
      Tab(3).ControlCount=   2
      Begin VB.Timer PlayerTimer 
         Enabled         =   0   'False
         Left            =   10000
         Top             =   0
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Change Info"
         Height          =   255
         Left            =   -74880
         TabIndex        =   62
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox txtText 
         Height          =   1455
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Top             =   720
         Width           =   6255
      End
      Begin VB.TextBox txtChat 
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   2280
         Width           =   6255
      End
      Begin VB.ListBox lstTopics 
         Height          =   2010
         Left            =   -74880
         TabIndex        =   13
         Top             =   600
         Width           =   2535
      End
      Begin VB.Frame TopicTitle 
         Caption         =   "Topic Title"
         Height          =   2295
         Left            =   -72240
         TabIndex        =   11
         Top             =   300
         Width           =   3615
         Begin VB.TextBox txtTopic 
            Height          =   1935
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   12
            Top             =   240
            Width           =   3375
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Commands"
         Height          =   2055
         Left            =   -74880
         TabIndex        =   3
         Top             =   480
         Width           =   1695
         Begin VB.CommandButton Command9 
            Caption         =   "Mass Kick"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Width           =   1455
         End
         Begin VB.CommandButton Command31 
            Caption         =   "Mass Kill"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   480
            Width           =   1455
         End
         Begin VB.CommandButton Command12 
            Caption         =   "Mass Heal"
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   720
            Width           =   1455
         End
         Begin VB.CommandButton Command36 
            Caption         =   "Map Info"
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   960
            Width           =   1455
         End
         Begin VB.CommandButton Command33 
            Caption         =   "Mass Experience"
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   1200
            Width           =   1455
         End
         Begin VB.CommandButton Command35 
            Caption         =   "Map List"
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   1440
            Width           =   1455
         End
         Begin VB.CommandButton Command32 
            Caption         =   "Mass Warp"
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   1680
            Width           =   1455
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Map List"
         Height          =   2055
         Left            =   -73080
         TabIndex        =   1
         Top             =   480
         Width           =   2415
         Begin VB.ListBox MapList 
            Height          =   1620
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   2175
         End
      End
      Begin MSComctlLib.ListView lvwInfo 
         Height          =   2055
         Left            =   -74880
         TabIndex        =   16
         Top             =   360
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   3625
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
      Begin VB.Menu mnuKick 
         Caption         =   "Kick Player"
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
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCL_Click()
Dim index As Long
index = lvwInfo.ListItems(lvwInfo.SelectedItem.index).Text
Call SetPlayerLevel(index, txtlevel.Text)
Call SendPlayerData(index)
Call SendStats(index)
End Sub

Private Sub cmdClosepic_Click()
picChangeInfo.Visible = False
End Sub

Private Sub cmdSA_Click()
Dim index As Long
index = lvwInfo.ListItems(lvwInfo.SelectedItem.index).Text
Call SetPlayerAccess(index, txtAccess.Text)
Call SendPlayerData(index)
End Sub

Private Sub cmdSDEF_Click()
Dim index As Long
index = lvwInfo.ListItems(lvwInfo.SelectedItem.index).Text
Call SetPlayerDEF(index, txtDef.Text)
Call SendPlayerData(index)
Call SendStats(index)
End Sub

Private Sub cmdsetMAGI_Click()
Dim index As Long
index = lvwInfo.ListItems(lvwInfo.SelectedItem.index).Text
Call SetPlayerMAGI(index, txtMagi.Text)
Call SendPlayerData(index)
Call SendStats(index)
End Sub

Private Sub cmdSSPEED_Click()
Dim index As Long
index = lvwInfo.ListItems(lvwInfo.SelectedItem.index).Text
Call SetPlayerSPEED(index, txtSpeed.Text)
Call SendPlayerData(index)
Call SendStats(index)
End Sub

Private Sub cmdSSSTR_Click()
Dim index As Long
index = lvwInfo.ListItems(lvwInfo.SelectedItem.index).Text
Call SetPlayerSTR(index, txtStr.Text)
Call SendPlayerData(index)
Call SendStats(index)
End Sub


Private Sub Command12_Click()
Dim index As Long

    For index = 1 To MAX_PLAYERS

        If IsPlaying(index) = True Then
            Call SetPlayerHP(index, GetPlayerMaxHP(index))
            Call SendHP(index)
            Call PlayerMsg(index, "You have been healed by the server!", BrightGreen)
        End If
    Next
End Sub



Private Sub Command31_Click()
Dim index As Long

    For index = 1 To MAX_PLAYERS

        If IsPlaying(index) = True Then
            If GetPlayerAccess(index) <= 0 Then
                Call SetPlayerHP(index, 0)
                Call PlayerMsg(index, "You have been killed by the server!", BrightRed)

                ' Warp player away
            
                    Call PlayerWarp(index, START_MAP, START_X, START_Y)
                End If
                Call SetPlayerHP(index, GetPlayerMaxHP(index))
                Call SetPlayerMP(index, GetPlayerMaxMP(index))
                Call SetPlayerSP(index, GetPlayerMaxSP(index))
                Call SendHP(index)
                Call SendMP(index)
                Call SendSP(index)
           
    End If
    Next
End Sub

Private Sub Command32_Click()
    scrlMM.Max = MAX_MAPS
    scrlMX.Max = MAX_MAPX
    scrlMY.Max = MAX_MAPY
    picWarp.Visible = True
End Sub

Private Sub Command33_Click()
   picExp.Visible = True
End Sub

Private Sub Command35_Click()
Dim i As Long

    MapList.Clear
    For i = 1 To MAX_MAPS
        MapList.AddItem i & ": " & Map(i).Name
    Next
    frmServer.MapList.Selected(0) = True

End Sub

Private Sub Command36_Click()
Dim index As Long
Dim i As Long

    index = MapList.ListIndex + 1
    MapInfo(0).Caption = "Map " & index & " - " & Map(index).Name
    MapInfo(1).Caption = "Revision: " & Map(index).Revision
    MapInfo(2).Caption = "Moral: " & Map(index).Moral
    MapInfo(3).Caption = "Up: " & Map(index).Up
    MapInfo(4).Caption = "Down: " & Map(index).Down
    MapInfo(5).Caption = "Left: " & Map(index).Left
    MapInfo(6).Caption = "Right: " & Map(index).Right
    MapInfo(7).Caption = "Music: " & Map(index).Music
    MapInfo(8).Caption = "BootMap: " & Map(index).BootMap
    MapInfo(9).Caption = "BootX: " & Map(index).BootX
    MapInfo(10).Caption = "BootY: " & Map(index).BootY
    MapInfo(11).Caption = "Shop: " & Map(index).Shop
    MapInfo(12).Caption = "Indoors: " & Map(index).Indoors
    lstNPC.Clear
    For i = 1 To MAX_MAP_NPCS
        lstNPC.AddItem i & ": " & Npc(Map(index).Npc(i)).Name
    Next
    picMap.Visible = True
End Sub

Private Sub Command37_Click()
Dim i As Long

    Call GlobalMsg("The server has warped everyone to Map:" & scrlMM.Value & " X:" & scrlMX.Value & " Y:" & scrlMY.Value, Yellow)
    For i = 1 To MAX_PLAYERS

        If IsPlaying(i) = True Then
            If GetPlayerAccess(i) <= 1 Then
               Call PlayerWarp(i, scrlMM.Value, scrlMX.Value, scrlMY.Value)
               End If
   End If
   Next
End Sub

Private Sub Command38_Click()
picWarp.Visible = False
End Sub

Private Sub Command39_Click()
picExp.Visible = False
End Sub



Private Sub Command40_Click()
Dim index As Long

    If IsNumeric(txtExp.Text) = False Then
        MsgBox "Enter a numerical value!"
        Exit Sub
    End If

    If txtExp.Text >= 0 Then
        Call GlobalMsg("The server gave everyone " & txtExp.Text & " experience!", BrightGreen)
        For index = 1 To MAX_PLAYERS

            If IsPlaying(index) = True Then
                Call SetPlayerExp(index, GetPlayerExp(index) + txtExp.Text)
                Call CheckPlayerLevelUp(index)
            End If
        Next
    End If
    picExp.Visible = False
End Sub

Private Sub Command41_Click()
picMap.Visible = False
End Sub

Private Sub Command5_Click()
picChangeInfo.Visible = True
End Sub

Private Sub Command9_Click()
Dim index As Long

    For index = 1 To MAX_PLAYERS

        If IsPlaying(index) = True Then
            If GetPlayerAccess(index) <= 0 Then
                Call GlobalMsg(GetPlayerName(index) & " has been kicked by the server!", White)
                Call AlertMsg(index, "You have been kicked by the server!")
            End If
        End If
    Next
End Sub

Private Sub Form_Load()
Dim i As Long
For i = 1 To MAX_PLAYERS
    Call UsersOnline_Start(i)
Next
    MapList.Clear
    For i = 1 To MAX_MAPS
        MapList.AddItem i & ": " & Map(i).Name
Next
    frmServer.MapList.Selected(0) = True
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

Private Sub lstTopics_Click()
Dim FileName As String, inputdata As String
Dim hFile As Long
Dim x As Long

    txtTopic.Text = ""
    TopicTitle.Caption = lstTopics.List(lstTopics.ListIndex)
    FileName = lstTopics.ListIndex + 1 & ".txt"

    x = 0
    
    If FileExist("Guide\" & FileName) = True And FileName <> "" Then
        hFile = FreeFile
        Open App.Path & "\Guide\" & FileName For Input As #hFile
            Do Until EOF(1)
                Line Input #1, inputdata
                If x = 0 Then
                    x = 1
                Else
                    txtTopic.Text = txtTopic.Text & inputdata & vbCrLf
                End If
            Loop
            
        Close #hFile
    End If
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
        
    
    lvwInfo.SortKey = ColumnHeader.index - 1
    lvwInfo.Sorted = True
    Debug.Print ColumnHeader.index & " " & ColumnHeader.Text & " " & ColumnHeader.Width
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

Private Sub scrlMM_Change()
    lblMM.Caption = "Map: " & scrlMM.Value
End Sub

Private Sub scrlMX_Change()
    lblMX.Caption = "X: " & scrlMX.Value
End Sub

Private Sub scrlMY_Change()
    lblMY.Caption = "Y: " & scrlMY.Value
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

Private Sub txtText_GotFocus()
    txtChat.SetFocus
End Sub

Private Sub txtChat_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And Trim(txtChat.Text) <> "" Then
        Call GlobalMsg(txtChat.Text, White)
        Call TextAdd(frmServer.txtText, "Server: " & txtChat.Text, True)
        txtChat.Text = ""
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


Private Sub txtTopic_Change()
Dim FileName As String, inputdata As String
Dim hFile As Long
Dim x As Long

    txtTopic.Text = ""
    TopicTitle.Caption = lstTopics.List(lstTopics.ListIndex)
    FileName = lstTopics.ListIndex + 1 & ".txt"

    x = 0
    
    If FileExist("Guide\" & FileName) = True And FileName <> "" Then
        hFile = FreeFile
        Open App.Path & "\Guide\" & FileName For Input As #hFile
            Do Until EOF(1)
                Line Input #1, inputdata
                If x = 0 Then
                    x = 1
                Else
                    txtTopic.Text = txtTopic.Text & inputdata & vbCrLf
                End If
            Loop
            
        Close #hFile
    End If
End Sub

Private Sub lvwInfo_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If Button = vbRightButton Then
        PopupMenu mnuKick
    End If
    
End Sub


Private Sub mnuKickPlayer_Click()
Dim index As Long
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
Private Sub PlayerTimer_Timer()
Dim i As Long

    If PlayerI <= MAX_PLAYERS Then
        If IsPlaying(PlayerI) Then
            Call SavePlayer(PlayerI)
            Call PlayerMsg(PlayerI, GetPlayerName(PlayerI) & ", you have been saved.", BrightGreen)
        End If
        PlayerI = PlayerI + 1
    End If
   
    If PlayerI >= MAX_PLAYERS Then
        PlayerI = 1
        PlayerTimer.Enabled = False
        tmrPlayerSave.Enabled = True
    End If
   
End Sub
