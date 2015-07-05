VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmadmin 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "  Administration Panel"
   ClientHeight    =   4560
   ClientLeft      =   135
   ClientTop       =   495
   ClientWidth     =   2400
   Icon            =   "frmadmin.frx":0000
   LinkTopic       =   "frmadmin"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmadmin.frx":0FC2
   ScaleHeight     =   4560
   ScaleWidth      =   2400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnclose 
      Caption         =   "Close Admin Panel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   4080
      Width           =   1695
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4335
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   7646
      _Version        =   393216
      TabHeight       =   353
      TabMaxWidth     =   1235
      BackColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Players"
      TabPicture(0)   =   "frmadmin.frx":3A90
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "btnSetAccess"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "btnKick"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "btnWarpMeTo"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "btnBan"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "btnPlayerSprite"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "btnSprite"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtSprite"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtAccess"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtPlayer"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "World"
      TabPicture(1)   =   "frmadmin.frx":3AAC
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "btneditmap"
      Tab(1).Control(1)=   "Command65"
      Tab(1).Control(2)=   "Command64"
      Tab(1).Control(3)=   "Command63"
      Tab(1).Control(4)=   "Command62"
      Tab(1).Control(5)=   "btnWarpto"
      Tab(1).Control(6)=   "btnRespawn"
      Tab(1).Control(7)=   "btnLOC"
      Tab(1).Control(8)=   "txtMap"
      Tab(1).Control(9)=   "Label2"
      Tab(1).ControlCount=   10
      TabCaption(2)   =   "Editors"
      TabPicture(2)   =   "frmadmin.frx":3AC8
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Command3"
      Tab(2).Control(1)=   "Command2"
      Tab(2).Control(2)=   "Command1"
      Tab(2).Control(3)=   "tnEditElement"
      Tab(2).Control(4)=   "tnEditEmoticon"
      Tab(2).Control(5)=   "tnEditArrow"
      Tab(2).Control(6)=   "tnEditNPC"
      Tab(2).Control(7)=   "btnEditShops"
      Tab(2).Control(8)=   "btnedititem"
      Tab(2).Control(9)=   "btneditspell"
      Tab(2).ControlCount=   10
      Begin VB.CommandButton Command3 
         Caption         =   "Edit Scripts"
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
         Left            =   -74760
         TabIndex        =   33
         Top             =   3600
         Width           =   1695
      End
      Begin VB.CommandButton btneditmap 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         Caption         =   "Edit Map"
         Default         =   -1  'True
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
         Left            =   -74760
         MaskColor       =   &H8000000D&
         TabIndex        =   32
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Edit Quests"
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
         Left            =   -74760
         TabIndex        =   31
         Top             =   3240
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Edit Skills"
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
         Left            =   -74760
         TabIndex        =   30
         Top             =   2880
         Width           =   1695
      End
      Begin VB.CommandButton tnEditElement 
         Caption         =   "Edit Element"
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
         Left            =   -74760
         TabIndex        =   29
         Top             =   2520
         Width           =   1695
      End
      Begin VB.CommandButton tnEditEmoticon 
         Caption         =   "Edit Emotion"
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
         Left            =   -74760
         TabIndex        =   28
         Top             =   2160
         Width           =   1695
      End
      Begin VB.CommandButton tnEditArrow 
         Caption         =   "Edit Arrows"
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
         Left            =   -74760
         TabIndex        =   27
         Top             =   1800
         Width           =   1695
      End
      Begin VB.CommandButton tnEditNPC 
         Caption         =   "Edit NPC's"
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
         Left            =   -74760
         TabIndex        =   26
         Top             =   1440
         Width           =   1695
      End
      Begin VB.CommandButton btnEditShops 
         Caption         =   "Edit Shops"
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
         Left            =   -74760
         TabIndex        =   25
         Top             =   1080
         Width           =   1695
      End
      Begin VB.CommandButton btnedititem 
         Caption         =   "Edit Items"
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
         Left            =   -74760
         TabIndex        =   24
         Top             =   720
         Width           =   1695
      End
      Begin VB.CommandButton btneditspell 
         Caption         =   "Edit Spells"
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
         Left            =   -74760
         TabIndex        =   23
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton Command65 
         Caption         =   "Snow"
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
         Left            =   -74760
         TabIndex        =   22
         Top             =   3240
         Width           =   1695
      End
      Begin VB.CommandButton Command64 
         Caption         =   "Rain"
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
         Left            =   -74760
         TabIndex        =   21
         Top             =   2520
         Width           =   1695
      End
      Begin VB.CommandButton Command63 
         Caption         =   "Thunder"
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
         Left            =   -74760
         TabIndex        =   20
         Top             =   2880
         Width           =   1695
      End
      Begin VB.CommandButton Command62 
         Caption         =   "None"
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
         Left            =   -74760
         TabIndex        =   19
         Top             =   3600
         Width           =   1695
      End
      Begin VB.CommandButton btnWarpto 
         Caption         =   "Warp To"
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
         Left            =   -74760
         TabIndex        =   17
         Top             =   1320
         Width           =   1695
      End
      Begin VB.CommandButton btnRespawn 
         Caption         =   "Respawn"
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
         Left            =   -74760
         TabIndex        =   16
         Top             =   2040
         Width           =   1695
      End
      Begin VB.CommandButton btnLOC 
         Caption         =   "Location"
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
         Left            =   -74760
         TabIndex        =   15
         Top             =   1680
         Width           =   1695
      End
      Begin VB.TextBox txtMap 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -74760
         TabIndex        =   14
         Text            =   "Enter Map Number"
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox txtPlayer 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   12
         Text            =   "Enter Player Name"
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox txtAccess 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   11
         Text            =   "Enter Access Level"
         Top             =   3360
         Width           =   1695
      End
      Begin VB.TextBox txtSprite 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   10
         Text            =   "Enter Sprite Number"
         Top             =   2640
         Width           =   1695
      End
      Begin VB.CommandButton btnSprite 
         Caption         =   "Set Sprite"
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
         Left            =   240
         TabIndex        =   9
         Top             =   2160
         Width           =   1695
      End
      Begin VB.CommandButton btnPlayerSprite 
         Caption         =   "Set Player Sprite"
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
         Left            =   240
         TabIndex        =   8
         Top             =   2400
         Width           =   1695
      End
      Begin VB.CommandButton btnBan 
         Caption         =   "Ban Player"
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
         Left            =   240
         TabIndex        =   7
         Top             =   1800
         Width           =   1695
      End
      Begin VB.CommandButton btnWarpMeTo 
         Caption         =   "Warp Me To"
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
         Left            =   240
         TabIndex        =   6
         Top             =   1080
         Width           =   1695
      End
      Begin VB.CommandButton btnKick 
         Caption         =   "Kick"
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
         Left            =   240
         TabIndex        =   5
         Top             =   1440
         Width           =   1695
      End
      Begin VB.CommandButton btnSetAccess 
         Caption         =   "Set Access"
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
         Left            =   240
         TabIndex        =   4
         Top             =   3120
         Width           =   1695
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Map Number :"
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
         Left            =   -74760
         TabIndex        =   18
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Player Name :"
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
         Left            =   240
         TabIndex        =   13
         Top             =   480
         Width           =   1695
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Player Name :"
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
      Left            =   5520
      TabIndex        =   3
      Top             =   3600
      Width           =   1695
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Sprite Number:"
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
      Left            =   5520
      TabIndex        =   2
      Top             =   3960
      Width           =   1695
   End
End
Attribute VB_Name = "frmadmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnBan_Click()

    If GetPlayerAccess(MyIndex) >= ADMIN_DEVELOPER Then
        Call SendBan(Trim$(txtPlayer.Text))
         Else: Call AddText("You are not authorized to carry out that action", BrightRed)
        End If

    End Sub

Private Sub btnclose_Click()

    frmadmin.Visible = False

End Sub

Private Sub btnedititem_Click()

    If GetPlayerAccess(MyIndex) >= ADMIN_DEVELOPER Then
        Call SendRequestEditItem
         Else: Call AddText("You are not authorized to carry out that action", BrightRed)
        End If

    End Sub

Private Sub btneditmap_Click()

    If GetPlayerAccess(MyIndex) >= ADMIN_MAPPER Then
        Call SendRequestEditMap
        Call WriteINI("CONFIG", "Res", 1, (App.Path & "\config.ini"))
        Screen_RESIZED = 0
         Else: Call AddText("You are not authorized to carry out that action", BrightRed)
        End If

    End Sub

Private Sub btnEditNPC_Click()

    If GetPlayerAccess(MyIndex) >= ADMIN_DEVELOPER Then
        Call SendRequestEditNpc
         Else: Call AddText("You are not authorized to carry out that action", BrightRed)
        End If

    End Sub

Private Sub btnEditShops_Click()

    If GetPlayerAccess(MyIndex) >= ADMIN_DEVELOPER Then
        Call SendRequestEditShop
         Else: Call AddText("You are not authorized to carry out that action", BrightRed)
        End If

    End Sub

Private Sub btneditspell_Click()

    If GetPlayerAccess(MyIndex) >= ADMIN_DEVELOPER Then
        Call SendRequestEditSpell
         Else: Call AddText("You are not authorized to carry out that action", BrightRed)
        End If

    End Sub

Private Sub btnkick_Click()

    If GetPlayerAccess(MyIndex) >= ADMIN_MONITER Then
        Call SendKick(Trim$(txtPlayer.Text))
         Else: Call AddText("You are not authorized to carry out that action", BrightRed)
        End If

    End Sub

Private Sub btnLOC_Click()

    If GetPlayerAccess(MyIndex) >= ADMIN_MAPPER Then
        Call SendRequestLocation
         Else: Call AddText("You are not authorized to carry out that action", BrightRed)
        End If

    End Sub

Private Sub btnPlayerSprite_Click()

    If GetPlayerAccess(MyIndex) >= ADMIN_MAPPER Then
        If Trim$(txtPlayer.Text) <> vbNullString Then
            If Trim$(txtSprite.Text) <> vbNullString And IsNumeric(txtSprite.Text) Then
                Call SendSetPlayerSprite(Trim$(txtPlayer.Text), Trim$(txtSprite.Text))
            End If

        End If
         Else: Call AddText("You are not authorized to carry out that action", BrightRed)
        End If

    End Sub

Private Sub btnRespawn_Click()

    If GetPlayerAccess(MyIndex) >= ADMIN_MAPPER Then
        Call SendMapRespawn
         Else: Call AddText("You are not authorized to carry out that action", BrightRed)
        End If

    End Sub

Private Sub btnSetAccess_Click()

    On Error Resume Next

    If GetPlayerAccess(MyIndex) >= ADMIN_CREATOR Then
        Call SendSetAccess(Trim$(txtPlayer.Text), Trim$(txtAccess.Text))
         Else: Call AddText("You are not authorized to carry out that action", BrightRed)
        End If

    End Sub

Private Sub btnSprite_Click()

    If GetPlayerAccess(MyIndex) >= ADMIN_MAPPER Then
        Call SendSetSprite(val(txtSprite.Text))
         Else: Call AddText("You are not authorized to carry out that action", BrightRed)
        End If

    End Sub

Private Sub btnWarpmeTo_Click()

    If GetPlayerAccess(MyIndex) >= ADMIN_MAPPER Then
        Call WarpMeTo(Trim$(txtPlayer.Text))
         Else: Call AddText("You are not authorized to carry out that action", BrightRed)
        End If

    End Sub

Private Sub btnWarptome_Click()

    If GetPlayerAccess(MyIndex) >= ADMIN_MAPPER Then
        Call WarpToMe(Trim$(txtPlayer.Text))
         Else: Call AddText("You are not authorized to carry out that action", BrightRed)
        End If

    End Sub

Private Sub btnWarpto_Click()

    If GetPlayerAccess(MyIndex) >= ADMIN_MAPPER Then
        Call WarpTo(val(txtMap.Text))
         Else: Call AddText("You are not authorized to carry out that action", BrightRed)
        End If

    End Sub

Private Sub Command1_Click()

    If GetPlayerAccess(MyIndex) >= ADMIN_DEVELOPER Then
        Call SendRequestEditSkill
         Else: Call AddText("You are not authorized to carry out that action", BrightRed)
        End If

    End Sub

Private Sub Command2_Click()

    If GetPlayerAccess(MyIndex) >= ADMIN_DEVELOPER Then
        Call SendRequestEditQuest
         Else: Call AddText("You are not authorized to carry out that action", BrightRed)
        End If

    End Sub

Private Sub Command3_Click()

    Call EditMain

End Sub

Private Sub Command62_Click()
Dim i As Long

    i = 0
    Call SendData(PacketID.Weather & SEP_CHAR & i & SEP_CHAR & END_CHAR)

End Sub

Private Sub Command63_Click()
Dim i As Long

    i = 3
    Call SendData(PacketID.Weather & SEP_CHAR & i & SEP_CHAR & END_CHAR)

End Sub

Private Sub Command64_Click()
Dim i As Long

    i = 1
    Call SendData(PacketID.Weather & SEP_CHAR & i & SEP_CHAR & END_CHAR)

End Sub

Private Sub Command65_Click()
Dim i As Long

    i = 2
    Call SendData(PacketID.Weather & SEP_CHAR & i & SEP_CHAR & END_CHAR)

End Sub

Private Sub fpscheck_Click()

    Call AddText("FPS: " & GameFPS, Pink)

End Sub

Private Sub MapReport_Click()

    Call SendData(PacketID.MapReport & SEP_CHAR & END_CHAR)

End Sub

Private Sub PlayerInfo_Click()

    If txtPlayer = vbNullString Then Call MsgBox("You need to enter a player Name", vbCritical, GAME_NAME)
    Call SendData(PacketID.PlayerInfoRequest & SEP_CHAR & txtPlayer & SEP_CHAR & END_CHAR)

End Sub

Private Sub tnEditArrow_Click()

    If GetPlayerAccess(MyIndex) >= ADMIN_DEVELOPER Then
        Call SendRequestEditArrow
         Else: Call AddText("You are not authorized to carry out that action", BrightRed)
        End If

    End Sub

Private Sub tnEditElement_Click()

    If GetPlayerAccess(MyIndex) >= ADMIN_DEVELOPER Then
        Call SendRequestEditElement
         Else: Call AddText("You are not authorized to carry out that action", BrightRed)
        End If

    End Sub

Private Sub tnEditEmoticon_Click()

    If GetPlayerAccess(MyIndex) >= ADMIN_DEVELOPER Then
        Call SendRequestEditEmoticon
         Else: Call AddText("You are not authorized to carry out that action", BrightRed)
        End If

    End Sub

Private Sub tnEditNPC_Click()

    If GetPlayerAccess(MyIndex) >= ADMIN_DEVELOPER Then
        Call SendRequestEditNpc
         Else: Call AddText("You are not authorized to carry out that action", BrightRed)
        End If

    End Sub

