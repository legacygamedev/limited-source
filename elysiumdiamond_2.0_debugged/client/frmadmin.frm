VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL3N.OCX"
Begin VB.Form frmadmin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Administration Panel"
   ClientHeight    =   5505
   ClientLeft      =   75
   ClientTop       =   525
   ClientWidth     =   4830
   Icon            =   "frmadmin.frx":0000
   LinkTopic       =   "frmadmin"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   4830
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
      Left            =   2760
      TabIndex        =   18
      Top             =   5040
      Width           =   1695
   End
   Begin VB.Frame Frame4 
      Caption         =   "Developer Commands"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   2640
      TabIndex        =   11
      Top             =   2400
      Width           =   1935
      Begin VB.CommandButton tnEditEmoticon 
         Caption         =   "Edit Emotion"
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
         TabIndex        =   30
         Top             =   1680
         Width           =   1695
      End
      Begin VB.CommandButton tnEditArrow 
         Caption         =   "Edit Arrows"
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
         TabIndex        =   29
         Top             =   1440
         Width           =   1695
      End
      Begin VB.CommandButton tnEditNPC 
         Caption         =   "Edit NPC's"
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
         TabIndex        =   16
         Top             =   1200
         Width           =   1695
      End
      Begin VB.CommandButton btnEditShops 
         Caption         =   "Edit Shops"
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
         Top             =   960
         Width           =   1695
      End
      Begin VB.CommandButton btnedititem 
         Caption         =   "Edit Items"
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
         TabIndex        =   14
         Top             =   720
         Width           =   1695
      End
      Begin VB.CommandButton btneditspell 
         Caption         =   "Edit Spells"
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
         TabIndex        =   13
         Top             =   480
         Width           =   1695
      End
      Begin VB.CommandButton btnMapeditor 
         Caption         =   "Map Editor"
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
         TabIndex        =   12
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Sprite Commands"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   2640
      TabIndex        =   6
      Top             =   840
      Width           =   1935
      Begin VB.CommandButton btnPlayerSprite 
         Caption         =   "Set Player Sprite"
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
         TabIndex        =   22
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox txtSprite 
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
         TabIndex        =   8
         Top             =   1080
         Width           =   1695
      End
      Begin VB.CommandButton btnSprite 
         Caption         =   "Set Sprite"
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
         TabIndex        =   7
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label3 
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
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Map Commands"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   240
      TabIndex        =   2
      Top             =   3600
      Width           =   1935
      Begin VB.CommandButton btnWarpto 
         Caption         =   "Warp To"
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
         TabIndex        =   3
         Top             =   720
         Width           =   1695
      End
      Begin VB.CommandButton btnRespawn 
         Caption         =   "Respawn"
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
         TabIndex        =   17
         Top             =   480
         Width           =   1695
      End
      Begin VB.CommandButton btnLOC 
         Caption         =   "Location"
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
         TabIndex        =   10
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox txtMap 
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
         TabIndex        =   4
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Map Number:"
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
         TabIndex        =   5
         Top             =   1080
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Player Commands"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   1935
      Begin VB.CommandButton btnSetAccess 
         Caption         =   "Set Access"
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
         TabIndex        =   19
         Top             =   960
         Width           =   1695
      End
      Begin VB.CommandButton btnKick 
         Caption         =   "Kick"
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
         TabIndex        =   24
         Top             =   720
         Width           =   1695
      End
      Begin VB.CommandButton btnWarpMeTo 
         Caption         =   "Warp Me To"
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
         TabIndex        =   25
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox txtPlayer 
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
         TabIndex        =   27
         Top             =   2280
         Width           =   1695
      End
      Begin VB.CommandButton btnBan 
         Caption         =   "Banish"
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
         TabIndex        =   26
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox txtAccess 
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
         TabIndex        =   20
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000010&
         X1              =   120
         X2              =   1800
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Label Label4 
         Caption         =   "Access Level:"
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
         TabIndex        =   21
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Player Name:"
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
         TabIndex        =   1
         Top             =   2040
         Width           =   1575
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5295
      Left            =   120
      TabIndex        =   23
      Top             =   120
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   9340
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   353
      TabMaxWidth     =   1764
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Menu"
      TabPicture(0)   =   "frmadmin.frx":0FC2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Line1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         X1              =   2280
         X2              =   2280
         Y1              =   720
         Y2              =   5160
      End
      Begin VB.Label Label5 
         Caption         =   "Administrative panel version 0.3"
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
         TabIndex        =   28
         Top             =   360
         Width           =   2175
      End
   End
End
Attribute VB_Name = "frmadmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnPlayerSprite_Click()
    If GetPlayerAccess(MyIndex) >= ADMIN_MAPPER Then
        If Trim(txtPlayer.Text) <> "" Then
            If Trim(txtSprite.Text) <> "" Then
                Call SendSetPlayerSprite(Trim(txtPlayer.Text), Trim(txtSprite.Text))
            End If
        End If
    Else: Call AddText("You are not authorized to carry out that action", BrightRed)
    End If
End Sub
Private Sub btnBan_Click()
If GetPlayerAccess(MyIndex) >= ADMIN_DEVELOPER Then
Call SendBan(Trim(txtPlayer.Text))
Else: Call AddText("You are not authorized to carry out that action", BrightRed)
End If
End Sub




Private Sub btnedititem_Click()
If GetPlayerAccess(MyIndex) >= ADMIN_DEVELOPER Then
Call SendRequestEditItem
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
Call SendKick(Trim(txtPlayer.Text))
Else: Call AddText("You are not authorized to carry out that action", BrightRed)
End If
End Sub

Private Sub btnLOC_Click()
If GetPlayerAccess(MyIndex) >= ADMIN_MAPPER Then
Call SendRequestLocation
Else: Call AddText("You are not authorized to carry out that action", BrightRed)
End If
End Sub

Private Sub btnMapeditor_Click()
If GetPlayerAccess(MyIndex) >= ADMIN_MAPPER Then
Call SendRequestEditMap
Else: Call AddText("You are not authorized to carry out that action", BrightRed)
End If
End Sub
Private Sub btnRespawn_Click()
If GetPlayerAccess(MyIndex) >= ADMIN_MAPPER Then
Call SendMapRespawn
Else: Call AddText("You are not authorized to carry out that action", BrightRed)
End If
End Sub
Private Sub btnWarpmeTo_Click()
If GetPlayerAccess(MyIndex) >= ADMIN_MAPPER Then
Call WarpMeTo(Trim(txtPlayer.Text))
Else: Call AddText("You are not authorized to carry out that action", BrightRed)
End If
End Sub

Private Sub btnWarpto_Click()
If GetPlayerAccess(MyIndex) >= ADMIN_MAPPER Then
Call WarpTo(Val(txtMap.Text))
Else: Call AddText("You are not authorized to carry out that action", BrightRed)
End If
End Sub

Private Sub btnWarptome_Click()
If GetPlayerAccess(MyIndex) >= ADMIN_MAPPER Then
Call WarpToMe(Trim(txtPlayer.Text))
Else: Call AddText("You are not authorized to carry out that action", BrightRed)
End If
End Sub

Private Sub btnclose_Click()
frmadmin.Visible = False
End Sub

Private Sub btnSprite_Click()
If GetPlayerAccess(MyIndex) >= ADMIN_MAPPER Then
Call SendSetSprite(Val(txtSprite.Text))
Else: Call AddText("You are not authorized to carry out that action", BrightRed)
End If
End Sub

Private Sub btnSetAccess_Click()
   If GetPlayerAccess(MyIndex) >= ADMIN_CREATOR Then
      Call SendSetAccess(Trim(txtPlayer.Text), Trim(txtAccess.Text))
   Else: Call AddText("You are not authorized to carry out that action", BrightRed)
   End If
End Sub

Private Sub tnEditArrow_Click()
If GetPlayerAccess(MyIndex) >= ADMIN_DEVELOPER Then
Call SendRequestEditArrow
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

