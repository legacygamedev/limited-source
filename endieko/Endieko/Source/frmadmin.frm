VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmAdmin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Admin Panel"
   ClientHeight    =   1455
   ClientLeft      =   150
   ClientTop       =   795
   ClientWidth     =   4215
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "frmadmin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   4215
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtSprite 
      Height          =   255
      Left            =   2280
      TabIndex        =   12
      Top             =   1080
      Width           =   1695
   End
   Begin VB.TextBox txtMap 
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   1080
      Width           =   1695
   End
   Begin VB.TextBox txtAccess 
      Height          =   255
      Left            =   2280
      TabIndex        =   7
      Top             =   360
      Width           =   1695
   End
   Begin VB.TextBox txtPlayer 
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   360
      Width           =   1695
   End
   Begin TabDlg.SSTab sstMain 
      Height          =   3855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   6800
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Help"
      TabPicture(0)   =   "frmadmin.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      Begin VB.Frame Frame4 
         Caption         =   "Possible Commands"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   4335
         Begin VB.TextBox Text9 
            Height          =   2055
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   5
            Text            =   "frmadmin.frx":0028
            Top             =   240
            Width           =   4095
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Credits"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   240
         TabIndex        =   1
         Top             =   2880
         Width           =   4335
         Begin VB.CommandButton cmdCloseHelp 
            Caption         =   "Close"
            Height          =   495
            Left            =   2760
            TabIndex        =   14
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label7 
            Caption         =   "Website: www.key2heaven.net/phoenix"
            Height          =   255
            Left            =   120
            TabIndex        =   3
            Top             =   480
            Width           =   2775
         End
         Begin VB.Label Label6 
            Caption         =   "Programmers: William, Obsidian, Relikk"
            Height          =   255
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   2535
         End
      End
   End
   Begin VB.Label lblSpriteNumber 
      Caption         =   "Sprite Number:"
      Height          =   255
      Left            =   2280
      TabIndex        =   13
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label lblMapNumber 
      Caption         =   "Map Number:"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label lblPlayerName 
      Caption         =   "Player Name:"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label lblAccessLevel 
      Caption         =   "Access Level:"
      Height          =   255
      Left            =   2280
      TabIndex        =   8
      Top             =   120
      Width           =   1695
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuPlayer 
      Caption         =   "Player"
      Begin VB.Menu mnuBanPlayer 
         Caption         =   "Ban"
      End
      Begin VB.Menu mnuKickPlayer 
         Caption         =   "Kick"
      End
      Begin VB.Menu mnuMyLocation 
         Caption         =   "My Location"
      End
      Begin VB.Menu mnuPlayerLocation 
         Caption         =   "Location"
      End
      Begin VB.Menu mnuSetAccess 
         Caption         =   "Set Access"
      End
      Begin VB.Menu mnuSetMySprite 
         Caption         =   "Set My Sprite"
      End
      Begin VB.Menu mnuSetPlayerSprite 
         Caption         =   "Set Player Sprite"
      End
      Begin VB.Menu mnuWarpMeTo 
         Caption         =   "Warp Me To"
      End
      Begin VB.Menu mnuWarpToMe 
         Caption         =   "Warp To Me"
      End
   End
   Begin VB.Menu mnuMap 
      Caption         =   "Map"
      Begin VB.Menu mnuRespawnMap 
         Caption         =   "Respawn Map"
      End
      Begin VB.Menu mnuWarpTo 
         Caption         =   "Warp To"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnuChangeMOTD 
         Caption         =   "Edit MOTD"
      End
      Begin VB.Menu mnuMapEditor 
         Caption         =   "Edit Map"
      End
      Begin VB.Menu mnuEditItem 
         Caption         =   "Edit Items"
      End
      Begin VB.Menu mnuEditNPC 
         Caption         =   "Edit NPCs"
      End
      Begin VB.Menu mnuEditShop 
         Caption         =   "Edit Shops"
      End
      Begin VB.Menu mnuEditSpell 
         Caption         =   "Edit Spells"
      End
      Begin VB.Menu mnuEditPet 
         Caption         =   "Edit Pets"
      End
      Begin VB.Menu mnuEditArrow 
         Caption         =   "Edit Arrows"
      End
   End
   Begin VB.Menu mnuOther 
      Caption         =   "Other"
      Begin VB.Menu mnuBanList 
         Caption         =   "Ban List"
      End
      Begin VB.Menu mnuMapList 
         Caption         =   "Map List"
      End
      Begin VB.Menu mnuDayNight 
         Caption         =   "Day/Night"
      End
      Begin VB.Menu mnuChangeWeather 
         Caption         =   "Change Weather"
         Begin VB.Menu mnuWeatherNone 
            Caption         =   "None"
         End
         Begin VB.Menu mnuWeatherRain 
            Caption         =   "Rain"
         End
         Begin VB.Menu mnuThunder 
            Caption         =   "Thunder"
         End
         Begin VB.Menu mnuWeatherSnow 
            Caption         =   "Snow"
         End
      End
      Begin VB.Menu mnuDisableChat 
         Caption         =   "Disable Chat"
         Begin VB.Menu mnuNChat 
            Caption         =   "Normal Chat"
         End
         Begin VB.Menu mnuBChat 
            Caption         =   "Broadcast Chat"
         End
         Begin VB.Menu mnuGChat 
            Caption         =   "Global Chat"
         End
         Begin VB.Menu mnuAChat 
            Caption         =   "Admin Chat"
         End
         Begin VB.Menu mnuEChat 
            Caption         =   "Emote Chat"
         End
         Begin VB.Menu mnuPChat 
            Caption         =   "Private Chat"
         End
         Begin VB.Menu mnuGuildChat 
            Caption         =   "Guild Chat"
         End
         Begin VB.Menu mnuPartyChat 
            Caption         =   "Party Chat"
         End
      End
      Begin VB.Menu mnuDUChat 
         Caption         =   "Disable User Chat"
         Begin VB.Menu mnuDNChat 
            Caption         =   "Normal Chat"
         End
         Begin VB.Menu mnuDBChat 
            Caption         =   "Broadcast Chat"
         End
         Begin VB.Menu mnuDGChat 
            Caption         =   "Global Chat"
         End
         Begin VB.Menu mnuDAChat 
            Caption         =   "Admin Chat"
         End
         Begin VB.Menu mnuDEChat 
            Caption         =   "Emote Chat"
         End
         Begin VB.Menu mnuDPrivChat 
            Caption         =   "Private Chat"
         End
         Begin VB.Menu mnuDGuildChat 
            Caption         =   "Guild Chat"
         End
         Begin VB.Menu mnuDPChat 
            Caption         =   "Party Chat"
         End
      End
      Begin VB.Menu mnuAdminClear 
         Caption         =   "Clear Chat Box"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuTopics 
         Caption         =   "Topics"
      End
   End
End
Attribute VB_Name = "frmAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCloseHelp_Click()
    frmAdmin.Height = 2190
    frmAdmin.Width = 4305
    sstMain.Visible = False
    lblMapNumber.Visible = True
    lblPlayerName.Visible = True
    lblAccessLevel.Visible = True
    lblSpriteNumber.Visible = True
    txtMap.Visible = True
    txtSprite.Visible = True
    txtPlayer.Visible = True
    txtAccess.Visible = True
End Sub

Private Sub Form_Load()
    If GetPlayerAccess(MyIndex) = 1 Then
        mnuPlayer.Visible = True
        mnuKickPlayer.Visible = True
        
        mnuMap.Visible = False
        mnuEdit.Visible = False
        mnuOther.Visible = False
        mnuMyLocation.Visible = False
        mnuEdit.Visible = False
        mnuSetMySprite.Visible = False
        mnuSetPlayerSprite.Visible = False
        mnuBanPlayer.Visible = False
        mnuPlayerLocation.Visible = False
        mnuWarpMeTo.Visible = False
        mnuWarpToMe.Visible = False
    ElseIf GetPlayerAccess(MyIndex) = 2 Then
        mnuPlayer.Visible = True
        mnuKickPlayer.Visible = True
        mnuMyLocation.Visible = True
        mnuEdit.Visible = True
        mnuMapEditor.Visible = True
        mnuSetMySprite.Visible = True
        mnuSetPlayerSprite.Visible = True
        mnuRespawnMap.Visible = True
        mnuChangeMOTD.Visible = True
        mnuBanPlayer.Visible = True
        
        mnuBanList.Visible = False
        mnuAdminClear.Visible = False
        mnuDayNight.Visible = False
        mnuMapList.Visible = False
        mnuPlayerLocation.Visible = False
        mnuWarpMeTo.Visible = False
        mnuWarpTo.Visible = False
        mnuWarpToMe.Visible = False
        mnuEditArrow.Visible = False
        mnuEditItem.Visible = False
        mnuEditNPC.Visible = False
        mnuEditShop.Visible = False
        mnuEditSpell.Visible = False
    ElseIf GetPlayerAccess(MyIndex) > 3 Then
    Else
        Call PlayerMsg("You are not authorized to view the Admin Panel.", MyIndex)
        Unload Me
    End If
End Sub

Private Sub mnuAdminClear_Click()
    If GetPlayerAccess(MyIndex) >= ADMIN_MAPPER Then
        Call SendData("adminclear" & SEP_CHAR & END_CHAR)
    Else
        Call AddText("You are not authorized to carry out that action", BrightRed)
    End If
End Sub

Private Sub mnuBanList_Click()
    If GetPlayerAccess(MyIndex) >= ADMIN_DEVELOPER Then
        Call SendBanList
    Else
        Call AddText("You are not authorized to carry out that action", BrightRed)
    End If
End Sub

Private Sub mnuBanPlayer_Click()
If txtPlayer.Text = vbNullString Then
    Exit Sub
Else
    If GetPlayerAccess(MyIndex) >= ADMIN_DEVELOPER Then
        Call SendBan(Trim(txtPlayer.Text))
    Else
        Call AddText("You are not authorized to carry out that action", BrightRed)
    End If
End If
End Sub

Private Sub mnuDAChat_Click()
    If GetPlayerAccess(MyIndex) >= ADMIN_DEVELOPER Then
        If txtPlayer.Text <> vbNullString Then
            Call PlayerAdmin(txtPlayer.Text)
        Else
            Call MsgBox("Please Enter Player Name.")
        End If
    Else
        Call AddText("You are not authorized to carry out that action", BrightRed)
    End If
End Sub

Private Sub mnuDayNight_Click()
    If GetPlayerAccess(MyIndex) >= ADMIN_DEVELOPER Then
        Call SendGameTime
        Call SendData("daynight" & SEP_CHAR & END_CHAR)
    Else
        Call AddText("You are not authorized to carry out that action", BrightRed)
    End If
End Sub

Private Sub mnuDBChat_Click()
    If GetPlayerAccess(MyIndex) >= ADMIN_DEVELOPER Then
        If txtPlayer.Text <> vbNullString Then
            Call PlayerBroadcast(txtPlayer.Text)
        Else
            Call MsgBox("Please Enter Player Name.")
        End If
    Else
        Call AddText("You are not authorized to carry out that action", BrightRed)
    End If
End Sub

Private Sub mnuDEChat_Click()
    If GetPlayerAccess(MyIndex) >= ADMIN_DEVELOPER Then
        If txtPlayer.Text <> vbNullString Then
            Call PlayerEmote(txtPlayer.Text)
        Else
            Call MsgBox("Please Enter Player Name.")
        End If
    Else
        Call AddText("You are not authorized to carry out that action", BrightRed)
    End If
End Sub

Private Sub mnuDGChat_Click()
    If GetPlayerAccess(MyIndex) >= ADMIN_DEVELOPER Then
        If txtPlayer.Text <> vbNullString Then
            Call PlayerGlobal(txtPlayer.Text)
        Else
            Call MsgBox("Please Enter Player Name.")
        End If
    Else
        Call AddText("You are not authorized to carry out that action", BrightRed)
    End If
End Sub

Private Sub mnuDGuildChat_Click()
    If GetPlayerAccess(MyIndex) >= ADMIN_DEVELOPER Then
        If txtPlayer.Text <> vbNullString Then
            Call PlayerGuild(txtPlayer.Text)
        Else
            Call MsgBox("Please Enter Player Name.")
        End If
    Else
        Call AddText("You are not authorized to carry out that action", BrightRed)
    End If
End Sub

Private Sub mnuDNChat_Click()
    If GetPlayerAccess(MyIndex) >= ADMIN_DEVELOPER Then
        If txtPlayer.Text <> vbNullString Then
            Call PlayerMap(txtPlayer.Text)
        Else
            Call MsgBox("Please Enter Player Name.")
        End If
    Else
        Call AddText("You are not authorized to carry out that action", BrightRed)
    End If
End Sub

Private Sub mnuDPChat_Click()
    If GetPlayerAccess(MyIndex) >= ADMIN_DEVELOPER Then
        If txtPlayer.Text <> vbNullString Then
            Call PlayerParty(txtPlayer.Text)
        Else
            Call MsgBox("Please Enter Player Name.")
        End If
    Else
        Call AddText("You are not authorized to carry out that action", BrightRed)
    End If
End Sub

Private Sub mnuEditArrow_Click()
    If GetPlayerAccess(MyIndex) >= ADMIN_DEVELOPER Then
        Call SendRequestEditArrow
    Else
        Call AddText("You are not authorized to carry out that action", BrightRed)
    End If
End Sub

Private Sub mnuEditItem_Click()
    If GetPlayerAccess(MyIndex) >= ADMIN_DEVELOPER Then
        Call SendRequestEditItem
    Else
        Call AddText("You are not authorized to carry out that action", BrightRed)
    End If
End Sub

Private Sub mnuEditNPC_Click()
    If GetPlayerAccess(MyIndex) >= ADMIN_DEVELOPER Then
        Call SendRequestEditNpc
    Else
        Call AddText("You are not authorized to carry out that action", BrightRed)
    End If
End Sub

Private Sub mnuEditShop_Click()
    If GetPlayerAccess(MyIndex) >= ADMIN_DEVELOPER Then
        Call SendRequestEditShop
    Else
        Call AddText("You are not authorized to carry out that action", BrightRed)
    End If
End Sub

Private Sub mnuEditSpell_Click()
    If GetPlayerAccess(MyIndex) >= ADMIN_DEVELOPER Then
        Call SendRequestEditSpell
    Else
        Call AddText("You are not authorized to carry out that action", BrightRed)
    End If
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuKickPlayer_Click()
If txtPlayer.Text = vbNullString Then
    Exit Sub
Else
    If GetPlayerAccess(MyIndex) >= ADMIN_MONITER Then
        Call SendKick(Trim(txtPlayer.Text))
    Else
        Call AddText("You are not authorized to carry out that action", BrightRed)
    End If
End If
End Sub

Private Sub mnuMapEditor_Click()
    If GetPlayerAccess(MyIndex) >= ADMIN_MAPPER Then
        Call SendRequestEditMap
    Else
        Call AddText("You are not authorized to carry out that action", BrightRed)
    End If
End Sub

Private Sub mnuMapList_Click()
    If GetPlayerAccess(MyIndex) >= ADMIN_MAPPER Then
        Call SendData("mapreport" & SEP_CHAR & END_CHAR)
    Else
        Call AddText("You are not authorized to carry out that action", BrightRed)
    End If
End Sub

Private Sub mnuMyLocation_Click()
    If GetPlayerAccess(MyIndex) >= ADMIN_MAPPER Then
        Call SendRequestLocation
    Else
        Call AddText("You are not authorized to carry out that action", BrightRed)
    End If
End Sub

Private Sub mnuPChat_Click()
    If GetPlayerAccess(MyIndex) >= ADMIN_DEVELOPER Then
        If txtPlayer.Text <> vbNullString Then
            Call PlayerPrivate(txtPlayer.Text)
        Else
            Call MsgBox("Please Enter Player Name.")
        End If
    Else
        Call AddText("You are not authorized to carry out that action", BrightRed)
    End If
End Sub

Private Sub mnuRespawnMap_Click()
    If GetPlayerAccess(MyIndex) >= ADMIN_MAPPER Then
        Call SendMapRespawn
    Else
        Call AddText("You are not authorized to carry out that action", BrightRed)
    End If
End Sub

Private Sub mnuSetAccess_Click()
If txtPlayer.Text = vbNullString Or txtAccess.Text = vbNullString Then
    Exit Sub
Else
    If GetPlayerAccess(MyIndex) >= ADMIN_DEVELOPER Then
        Call SetPlayerAccess(txtPlayer.Text, txtAccess.Text)
    Else
        Call AddText("You are not authorized to carry out that action", BrightRed)
    End If
End If
End Sub

Private Sub mnuSetMySprite_Click()
If txtSprite.Text = vbNullString Then
    Exit Sub
Else
    If GetPlayerAccess(MyIndex) >= ADMIN_MAPPER Then
        Call SendSetSprite(Val(txtSprite.Text))
    Else
        Call AddText("You are not authorized to carry out that action", BrightRed)
    End If
End If
End Sub

Private Sub mnuSetPlayerSprite_Click()
If txtPlayer.Text = vbNullString Or txtSprite.Text = vbNullString Then
    Exit Sub
Else
    If GetPlayerAccess(MyIndex) >= ADMIN_MAPPER Then
        If Trim(txtPlayer.Text) <> "" Then
            If Trim(txtSprite.Text) <> "" Then
                Call SendSetPlayerSprite(Trim(txtPlayer.Text), Trim(txtSprite.Text))
            End If
        End If
    Else
        Call AddText("You are not authorized to carry out that action", BrightRed)
    End If
End If
End Sub

'Private Sub mnuThunder_Click()
'    If GetPlayerAccess(MyIndex) >= ADMIN_MAPPER Then
'        Call SendData("weather" & SEP_CHAR & WEATHER_THUNDER & SEP_CHAR & END_CHAR)
'    Else
'        Call AddText("You are not authorized to carry out that action", BrightRed)
'    End If
'End Sub

Private Sub mnuTopics_Click()
    If sstMain.Visible = False Then
        frmAdmin.Height = 4590
        frmAdmin.Width = 4830
        lblMapNumber.Visible = False
        lblPlayerName.Visible = False
        lblAccessLevel.Visible = False
        lblSpriteNumber.Visible = False
        txtMap.Visible = False
        txtSprite.Visible = False
        txtPlayer.Visible = False
        txtAccess.Visible = False
        sstMain.Visible = True
    Else
        sstMain.Visible = False
    End If
End Sub

Private Sub mnuWarpMeTo_Click()
If txtPlayer.Text = vbNullString Then
    Exit Sub
Else
    If GetPlayerAccess(MyIndex) >= ADMIN_MAPPER Then
        Call WarpMeTo(Trim(txtPlayer.Text))
    Else
        Call AddText("You are not authorized to carry out that action", BrightRed)
    End If
End If
End Sub

Private Sub mnuWarpTo_Click()
If txtMap.Text = vbNullString Then
    Exit Sub
Else
    If GetPlayerAccess(MyIndex) >= ADMIN_MAPPER Then
        Call WarpTo(Val(txtMap.Text))
    Else
        Call AddText("You are not authorized to carry out that action", BrightRed)
    End If
End If
End Sub

Private Sub mnuWeatherNone_Click()
    If GetPlayerAccess(MyIndex) >= ADMIN_MAPPER Then
        Call SendData("weather" & SEP_CHAR & WEATHER_NONE & SEP_CHAR & END_CHAR)
    Else
        Call AddText("You are not authorized to carry out that action", BrightRed)
    End If
End Sub

Private Sub mnuWeatherRain_Click()
    If GetPlayerAccess(MyIndex) >= ADMIN_MAPPER Then
        Call SendData("weather" & SEP_CHAR & WEATHER_RAINING & SEP_CHAR & END_CHAR)
    Else
        Call AddText("You are not authorized to carry out that action", BrightRed)
    End If
End Sub

Private Sub mnuWeatherSnow_Click()
    If GetPlayerAccess(MyIndex) >= ADMIN_MAPPER Then
        Call SendData("weather" & SEP_CHAR & WEATHER_SNOWING & SEP_CHAR & END_CHAR)
    Else
        Call AddText("You are not authorized to carry out that action", BrightRed)
    End If
End Sub
