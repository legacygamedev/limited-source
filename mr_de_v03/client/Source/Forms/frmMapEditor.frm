VERSION 5.00
Begin VB.Form frmMapEditor 
   Caption         =   "Map Editor"
   ClientHeight    =   9060
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3855
   LinkTopic       =   "Form1"
   ScaleHeight     =   604
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   257
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraMobs 
      Caption         =   "Mobs"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Left            =   120
      TabIndex        =   1
      Top             =   3960
      Visible         =   0   'False
      Width           =   1815
      Begin VB.ComboBox cmbDir 
         Height          =   315
         ItemData        =   "frmMapEditor.frx":0000
         Left            =   120
         List            =   "frmMapEditor.frx":0014
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   4560
         Width           =   975
      End
      Begin VB.HScrollBar scrlMob 
         Height          =   255
         Left            =   120
         Max             =   15
         Min             =   1
         TabIndex        =   7
         Top             =   480
         Value           =   1
         Width           =   1575
      End
      Begin VB.ListBox lstMobs 
         Height          =   2205
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   1575
      End
      Begin VB.CommandButton cmbAdd 
         Caption         =   "Add"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   3480
         Width           =   1575
      End
      Begin VB.CommandButton cmbRemove 
         Caption         =   "Remove"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   3720
         Width           =   1575
      End
      Begin VB.ComboBox cmbNpcs 
         Height          =   315
         ItemData        =   "frmMapEditor.frx":0035
         Left            =   120
         List            =   "frmMapEditor.frx":0048
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   3120
         Width           =   1575
      End
      Begin VB.CommandButton cmdReCalc 
         Caption         =   "Trim"
         Height          =   255
         Left            =   1200
         TabIndex        =   2
         Top             =   4560
         Width           =   495
      End
      Begin VB.Label lblMobNum 
         Alignment       =   1  'Right Justify
         Caption         =   "1"
         Height          =   255
         Left            =   1200
         TabIndex        =   11
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Mob Num"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblNpcCount 
         Caption         =   "Npc Count:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   4080
         Width           =   1575
      End
      Begin VB.Label lblNpcLimit 
         Caption         =   "Npc Limit: "
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   4320
         Width           =   1455
      End
   End
   Begin VB.Frame fraLayers 
      Caption         =   "Layers"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   120
      TabIndex        =   34
      Top             =   3960
      Width           =   1815
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   120
         TabIndex        =   45
         Top             =   3960
         Width           =   1575
      End
      Begin VB.OptionButton optFringe 
         Caption         =   "Fringe"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   1575
         Width           =   1215
      End
      Begin VB.OptionButton optAnim 
         Caption         =   "Animation"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   840
         Width           =   1215
      End
      Begin VB.OptionButton optMask 
         Caption         =   "Mask"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton optGround 
         Caption         =   "Ground"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton optM2anim 
         Caption         =   "Animation"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   1335
         Width           =   1470
      End
      Begin VB.OptionButton optMask2 
         Caption         =   "Coat"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   1095
         Width           =   1590
      End
      Begin VB.OptionButton optF2anim 
         Caption         =   "Animation"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   2295
         Width           =   1215
      End
      Begin VB.OptionButton optFringe2 
         Caption         =   "Cover"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   2055
         Width           =   1470
      End
      Begin VB.OptionButton optFAnim 
         Caption         =   "Animation"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   1815
         Width           =   1590
      End
      Begin VB.CommandButton cmdFill 
         Caption         =   "Fill"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   3720
         Width           =   1575
      End
   End
   Begin VB.PictureBox picBack 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   3360
      Left            =   120
      ScaleHeight     =   224
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   224
      TabIndex        =   33
      Top             =   240
      Width           =   3360
      Begin VB.Shape shpSelection 
         BorderColor     =   &H000000FF&
         Height          =   480
         Left            =   0
         Top             =   0
         Width           =   480
      End
   End
   Begin VB.CommandButton cmdProperties 
      Caption         =   "Properties"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   31
      Top             =   8160
      Width           =   1695
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   30
      Top             =   8760
      Width           =   1695
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   29
      Top             =   8520
      Width           =   1695
   End
   Begin VB.Frame fraRandom 
      Caption         =   "Random"
      Height          =   1695
      Left            =   2040
      TabIndex        =   23
      Top             =   4800
      Width           =   1695
      Begin VB.CheckBox chkRandomTile 
         Caption         =   "Random"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   1320
         Width           =   1215
      End
      Begin VB.PictureBox picRandomTile 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   480
         Index           =   3
         Left            =   720
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   27
         Top             =   795
         Width           =   480
      End
      Begin VB.PictureBox picRandomTile 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   480
         Index           =   2
         Left            =   120
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   26
         Top             =   795
         Width           =   480
      End
      Begin VB.PictureBox picRandomTile 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   480
         Index           =   1
         Left            =   720
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   25
         Top             =   240
         Width           =   480
      End
      Begin VB.PictureBox picRandomTile 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   480
         Index           =   0
         Left            =   120
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   24
         Top             =   240
         Width           =   480
      End
   End
   Begin VB.OptionButton optNpcs 
      Caption         =   "Npcs"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   22
      Top             =   4440
      Width           =   1575
   End
   Begin VB.OptionButton optLayers 
      Caption         =   "Layers"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   21
      Top             =   3960
      Value           =   -1  'True
      Width           =   1575
   End
   Begin VB.OptionButton optAttribs 
      Caption         =   "Attributes"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   20
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Frame fraAttribs 
      Caption         =   "Attributes"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   120
      TabIndex        =   12
      Top             =   3960
      Visible         =   0   'False
      Width           =   1815
      Begin VB.OptionButton optKeyOpen 
         Caption         =   "Switch"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   19
         Top             =   2415
         Width           =   1215
      End
      Begin VB.OptionButton optBlocked 
         Caption         =   "Blocked"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1200
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton optWarp 
         Caption         =   "Warp"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CommandButton cmdClear2 
         Caption         =   "Clear"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   120
         TabIndex        =   16
         Top             =   3960
         Width           =   1575
      End
      Begin VB.OptionButton optItem 
         Caption         =   "Item"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   120
         TabIndex        =   15
         Top             =   1665
         Width           =   1215
      End
      Begin VB.OptionButton optNpcAvoid 
         Caption         =   "Npc Avoid"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   120
         TabIndex        =   14
         Top             =   1920
         Width           =   1215
      End
      Begin VB.OptionButton optKey 
         Caption         =   "Locked Door"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   120
         TabIndex        =   13
         Top             =   2160
         Width           =   1575
      End
   End
   Begin VB.VScrollBar scrlPicture 
      Height          =   3375
      Left            =   3480
      Max             =   550
      TabIndex        =   0
      Top             =   240
      Width           =   255
   End
End
Attribute VB_Name = "frmMapEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdFill_Click()
    MapEditorFillLayer
End Sub

Private Sub Form_Load()
    scrlMob.Max = MAX_MOBS
    cmbDir.ListIndex = 0
End Sub

Private Sub Form_Activate()
Dim i As Long
    
    cmbNpcs.Clear
    For i = 1 To MAX_NPCS
        cmbNpcs.AddItem i & ": " & Trim$(Npc(i).Name)
    Next
    cmbNpcs.ListIndex = 0
    
    LoadMob scrlMob.Value
    lblNpcLimit.Caption = "Npc Limit: " & MapNpcLimit()
End Sub

Private Sub picRandomTile_Click(Index As Integer)
    RandomTileSelected = Index
End Sub

Private Sub optLayers_Click()
    If optLayers.Value Then
        fraLayers.Visible = True
        fraAttribs.Visible = False
        fraMobs.Visible = False
    End If
End Sub

Private Sub optAttribs_Click()
    If optAttribs.Value Then
        fraLayers.Visible = False
        fraAttribs.Visible = True
        fraMobs.Visible = False
    End If
End Sub

Private Sub optNpcs_Click()
    If optNpcs.Value Then
        fraLayers.Visible = False
        fraAttribs.Visible = False
        fraMobs.Visible = True
        LoadMob scrlMob.Value
    End If
End Sub

Private Sub picBack_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MapEditorChooseTile Button, Shift, X, Y
End Sub

Private Sub picBack_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MapEditorChooseTile Button, Shift, X, Y
End Sub

Private Sub cmdSend_Click()
    Call MapEditorSend
    frmMainGame.picScreen.SetFocus
End Sub

Private Sub cmdCancel_Click()
    Call MapEditorCancel
End Sub

Private Sub cmdProperties_Click()
    frmMapProperties.Show vbModal
End Sub

Private Sub optWarp_Click()
    frmMapWarp.Show vbModal
End Sub

Private Sub optItem_Click()
    frmMapItem.Show vbModal
End Sub

Private Sub optKey_Click()
    frmMapKey.Show vbModal
End Sub

Private Sub optKeyOpen_Click()
    frmKeyOpen.Show vbModal
End Sub

Private Sub scrlPicture_Change()
    Call MapEditorTileScroll
End Sub

Private Sub cmdClear_Click()
    Call MapEditorClearLayer
End Sub

Private Sub cmdClear2_Click()
    Call MapEditorClearAttribs
End Sub

'***************
'** Mob Stuff **
'***************

Private Sub scrlMob_Change()
    lblMobNum = scrlMob.Value
    LoadMob scrlMob.Value
End Sub

Private Sub cmbAdd_Click()
Dim MobNum As Long
Dim tempArr() As Long

    MobNum = scrlMob.Value
        
    ' Check limit
    If NpcCount >= MapNpcLimit Then Exit Sub
    
    With Map.Mobs(MobNum)
        
        .NpcCount = .NpcCount + 1
        ReDim Preserve .Npc(.NpcCount)
        
        .Npc(.NpcCount) = cmbNpcs.ListIndex + 1
    End With
    
    LoadMob MobNum
End Sub

Private Sub cmbRemove_Click()
Dim MobNum As Long
Dim tempArr() As Long
Dim i As Long, X As Long

    MobNum = scrlMob.Value
    
    ' Check if we have a npc selected
    If lstMobs.ListIndex < 0 Then Exit Sub
    
    With Map.Mobs(MobNum)
        .Npc(lstMobs.ListIndex + 1) = 0
        
        ReDim tempArr(1 To .NpcCount)
        For i = 1 To .NpcCount
            If .Npc(i) > 0 Then
                X = X + 1
                tempArr(X) = .Npc(i)
            End If
        Next
        
        .NpcCount = .NpcCount - 1
        ReDim .Npc(.NpcCount)
        
        For i = 1 To .NpcCount
            .Npc(i) = tempArr(i)
        Next
    End With
    
    LoadMob MobNum
End Sub

Private Sub cmdReCalc_Click()
Dim i As Long
Dim n As Long

    For i = 1 To MAX_MOBS
        If Map.Mobs(i).NpcCount > 0 Then
            For n = 1 To Map.Mobs(i).NpcCount
                If Map.Mobs(i).Npc(n) <= 0 Then Map.Mobs(i).NpcCount = Map.Mobs(i).NpcCount - 1
            Next
        End If
    Next
    LoadMob scrlMob.Value
    lblNpcLimit.Caption = "Npc Limit: " & MapNpcLimit()
End Sub

Private Sub LoadMob(ByVal MobNum As Long)
Dim i As Long
    
    lstMobs.Clear
    With Map.Mobs(MobNum)
        For i = 1 To .NpcCount
            If .Npc(i) > 0 Then
                lstMobs.AddItem i & ": " & Trim$(Npc(.Npc(i)).Name)
            End If
        Next
        If lstMobs.ListCount > 0 Then lstMobs.ListIndex = 0
    End With
    
    lblNpcCount.Caption = "Npc Count: " & NpcCount
End Sub

Private Function NpcCount() As Long
Dim i As Long
    
    NpcCount = 0
    For i = 1 To MAX_MOBS
        NpcCount = NpcCount + Map.Mobs(i).NpcCount
    Next
End Function

