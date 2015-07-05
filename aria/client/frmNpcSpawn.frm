VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmNpcSpawn 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NPC Spawn Tile"
   ClientHeight    =   1530
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   5025
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1530
   ScaleWidth      =   5025
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   2355
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
      TabCaption(0)   =   "NPC Spawn"
      TabPicture(0)   =   "frmNpcSpawn.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblNPC"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "scrNPC"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdOk"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdCancel"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
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
         Left            =   2640
         TabIndex        =   3
         Top             =   960
         Width           =   1935
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "Ok"
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
         TabIndex        =   2
         Top             =   960
         Width           =   1935
      End
      Begin VB.HScrollBar scrNPC 
         Height          =   255
         Left            =   360
         Max             =   15
         Min             =   1
         TabIndex        =   1
         Top             =   600
         Value           =   1
         Width           =   4095
      End
      Begin VB.Label lblNPC 
         Caption         =   "0"
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
         Left            =   840
         TabIndex        =   5
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "NPC:"
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
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmNpcSpawn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    NPCSpawnNum = scrNPC.Value
    Unload Me
End Sub

Private Sub Form_Load()
    scrNPC.Max = MAX_MAP_NPCS
    If NPCSpawnNum < scrNPC.Min Then NPCSpawnNum = scrNPC.Min
    If NPCSpawnNum > scrNPC.Max Then NPCSpawnNum = scrNPC.Max
    scrNPC.Value = NPCSpawnNum
End Sub

Private Sub scrNPC_Change()
    If Map(GetPlayerMap(MyIndex)).Npc(scrNPC.Value) <> 0 Then
        lblNPC.Caption = scrNPC.Value & " - " & Npc(Map(GetPlayerMap(MyIndex)).Npc(scrNPC.Value)).Name
    Else
        lblNPC.Caption = scrNPC.Value & " - Currently None"
    End If
End Sub
