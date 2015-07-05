VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmNPCSpawn 
   Caption         =   "NPC Spawn"
   ClientHeight    =   1920
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4230
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   1920
   ScaleWidth      =   4230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab SSTab1 
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   2990
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   353
      TabMaxWidth     =   2117
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Attributes"
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "scrMapSlot"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Command1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Command2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      Begin VB.CommandButton Command2 
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
         Left            =   2400
         TabIndex        =   5
         Top             =   1320
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
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
         TabIndex        =   4
         Top             =   1320
         Width           =   1575
      End
      Begin VB.HScrollBar scrMapSlot 
         Height          =   255
         Left            =   240
         Max             =   15
         Min             =   1
         TabIndex        =   1
         Top             =   600
         Value           =   1
         Width           =   2895
      End
      Begin VB.Label Label3 
         Caption         =   "Map NPCs: Go to Map Editor -> Properties -> NPCs"
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
         TabIndex        =   6
         Top             =   960
         Width           =   3375
      End
      Begin VB.Label Label2 
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
         Left            =   3240
         TabIndex        =   3
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "NPC Map Slot"
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
         TabIndex        =   2
         Top             =   360
         Width           =   2415
      End
   End
End
Attribute VB_Name = "frmNPCSpawn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

    NPCSpawnNum = scrMapSlot.Value
    Unload Me

End Sub

Private Sub Command2_Click()

    Unload Me

End Sub

Private Sub Form_Load()

    'Prevent some nasty errors

    If NPCSpawnNum > 0 And NPCSpawnNum < 16 Then
        scrMapSlot.Value = NPCSpawnNum
    End If

    Label2.Caption = scrMapSlot.Value

End Sub

Private Sub scrMapSlot_Change()

    Label2.Caption = scrMapSlot.Value

End Sub

