VERSION 5.00
Begin VB.Form frmNPCSpawn 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NPC Spawn"
   ClientHeight    =   2790
   ClientLeft      =   90
   ClientTop       =   570
   ClientWidth     =   2790
   Icon            =   "frmNPCSpawn.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   2790
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
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
      Left            =   600
      TabIndex        =   7
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
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
      Left            =   1680
      TabIndex        =   6
      Top             =   2400
      Width           =   975
   End
   Begin VB.HScrollBar scrlNum2 
      Height          =   255
      Left            =   240
      Max             =   50
      Min             =   1
      TabIndex        =   4
      Top             =   840
      Value           =   1
      Width           =   2295
   End
   Begin VB.HScrollBar scrlNum3 
      Height          =   255
      Left            =   240
      Max             =   30
      TabIndex        =   2
      Top             =   1560
      Width           =   2295
   End
   Begin VB.HScrollBar scrlNum 
      Height          =   255
      Left            =   240
      Min             =   1
      TabIndex        =   0
      Top             =   240
      Value           =   1
      Width           =   2295
   End
   Begin VB.Label lblNum2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NPC Amount: 1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label lblNum3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Spawn Range By Tile: 0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label lblNum 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NPC Number: 1 -"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   1200
   End
End
Attribute VB_Name = "frmNPCSpawn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    NPCSpawnNum = scrlNum.Value
    NPCSpawnAmount = scrlNum2.Value
    NPCSpawnRange = scrlNum3.Value
    Unload Me
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    scrlNum.Max = MAX_NPCS
    scrlNum2.Max = MAX_ATTRIBUTE_NPCS
    
    If NPCSpawnNum < scrlNum.Min Then NPCSpawnNum = scrlNum.Min
    scrlNum.Value = NPCSpawnNum
    If NPCSpawnAmount < scrlNum2.Min Then NPCSpawnAmount = scrlNum2.Min
    scrlNum2.Value = NPCSpawnAmount
    If NPCSpawnRange < scrlNum2.Min Then NPCSpawnRange = scrlNum3.Min
    scrlNum3.Value = NPCSpawnRange
End Sub

Private Sub scrlNum_Change()
    lblNum.Caption = "NPC Number: " & scrlNum.Value & " - " & Trim(Npc(scrlNum.Value).Name)
End Sub

Private Sub scrlNum2_Change()
    lblNum2.Caption = "NPC Amount: " & scrlNum2.Value
End Sub

Private Sub scrlNum3_Change()
    lblNum3.Caption = "Spawn Range By Tile: " & scrlNum3.Value
End Sub
