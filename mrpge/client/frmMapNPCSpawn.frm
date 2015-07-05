VERSION 5.00
Begin VB.Form frmMapNPCSpawn 
   Caption         =   "Form1"
   ClientHeight    =   1500
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5055
   LinkTopic       =   "Form1"
   ScaleHeight     =   1500
   ScaleWidth      =   5055
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar scrlItem 
      Height          =   300
      Left            =   840
      Max             =   14
      Min             =   1
      TabIndex        =   2
      Top             =   600
      Value           =   1
      Width           =   3255
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   2295
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   2640
      TabIndex        =   0
      Top             =   960
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "NPC"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   615
   End
   Begin VB.Label lblItem 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   255
      Left            =   4080
      TabIndex        =   5
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "NPC"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   840
      TabIndex        =   3
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "frmMapNPCSpawn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
On Error Resume Next
    lblName.Caption = Trim(Npc(map.Npc(scrlItem.value)).Name)
End Sub

Private Sub cmdOK_Click()
    If lblName.Caption <> "NA" Then
        EditorNPC_Num = scrlItem.value
    Else
        MsgBox "please select an npc next time"
    End If
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub



Private Sub scrlItem_Change()
    
    If map.Npc(scrlItem.value) > 0 And map.Npc(scrlItem.value) < MAX_NPCS Then
        lblItem.Caption = str(scrlItem.value)
        lblName.Caption = Trim(Npc(map.Npc(scrlItem.value)).Name)
    Else
        lblName.Caption = "NA"
    End If
End Sub

