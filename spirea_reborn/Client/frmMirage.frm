VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMirage 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Spirea Reborn"
   ClientHeight    =   8895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10800
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   593
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   720
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.PictureBox picInv 
      Appearance      =   0  'Flat
      BackColor       =   &H00BD8C64&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2685
      Left            =   8760
      ScaleHeight     =   177
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   135
      TabIndex        =   2
      Top             =   6120
      Visible         =   0   'False
      Width           =   2055
      Begin VB.ListBox lstInv 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1500
         ItemData        =   "frmMirage.frx":0000
         Left            =   120
         List            =   "frmMirage.frx":0002
         TabIndex        =   4
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label lblCancel 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   720
         TabIndex        =   7
         Top             =   2400
         Width           =   675
      End
      Begin VB.Label lblDropItem 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Drop Item"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   600
         TabIndex        =   6
         Top             =   2160
         Width           =   945
      End
      Begin VB.Label lblUseItem 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Use Item"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   600
         TabIndex        =   5
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Inventory"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   600
         TabIndex        =   3
         Top             =   0
         Width           =   915
      End
   End
   Begin VB.PictureBox picPlayerSpells 
      Appearance      =   0  'Flat
      BackColor       =   &H00BD8C64&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2685
      Left            =   8760
      ScaleHeight     =   177
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   135
      TabIndex        =   8
      Top             =   6120
      Visible         =   0   'False
      Width           =   2055
      Begin VB.ListBox lstSpells 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1500
         ItemData        =   "frmMirage.frx":0004
         Left            =   120
         List            =   "frmMirage.frx":0006
         TabIndex        =   9
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Spells"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   720
         TabIndex        =   12
         Top             =   0
         Width           =   615
      End
      Begin VB.Label lblCast 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cast"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   840
         TabIndex        =   11
         Top             =   1920
         Width           =   435
      End
      Begin VB.Label lblSpellsCancel 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   720
         TabIndex        =   10
         Top             =   2160
         Width           =   675
      End
   End
   Begin RichTextLib.RichTextBox txtChat 
      Height          =   1815
      Left            =   120
      TabIndex        =   1
      Top             =   6960
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   3201
      _Version        =   393217
      BackColor       =   0
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmMirage.frx":0008
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picScreen 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   6720
      Left            =   120
      ScaleHeight     =   446
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   574
      TabIndex        =   0
      Top             =   240
      Width           =   8640
   End
   Begin MSWinsockLib.Winsock Socket 
      Left            =   120
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5835
      Left            =   8760
      Picture         =   "frmMirage.frx":007F
      ScaleHeight     =   5805
      ScaleWidth      =   1980
      TabIndex        =   16
      Top             =   240
      Width           =   2010
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Height          =   495
         Left            =   0
         TabIndex        =   22
         Top             =   4080
         Width           =   1935
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Height          =   495
         Left            =   0
         TabIndex        =   21
         Top             =   3720
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Height          =   495
         Left            =   0
         TabIndex        =   20
         Top             =   3240
         Width           =   1215
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Height          =   495
         Left            =   120
         TabIndex        =   19
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Height          =   495
         Left            =   0
         TabIndex        =   18
         Top             =   1920
         Width           =   1815
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Height          =   495
         Left            =   0
         TabIndex        =   17
         Top             =   2880
         Width           =   1215
      End
   End
   Begin VB.Label lblSP 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0%"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3480
      TabIndex        =   15
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label lblMP 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0%"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   255
      Left            =   1800
      TabIndex        =   14
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label lblHP 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0%"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   0
      Width           =   1575
   End
End
Attribute VB_Name = "frmMirage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim Y As Long
Dim X As Long

X = MsgBox("Are you sure you want to fill the map?", vbYesNo)
If X = vbNo Then
    Exit Sub
End If

If frmMapEditor.optAttribs.Value = False Then
    For Y = 0 To MAX_MAPY
        For X = 0 To MAX_MAPX
            With Map.Tile(X, Y)
                If frmMapEditor.optGround.Value = True Then .Ground = EditorTileY * 14 + EditorTileX
                If frmMapEditor.optMask.Value = True Then .Mask = EditorTileY * 14 + EditorTileX
                If frmMapEditor.optAnim.Value = True Then .Anim = EditorTileY * 14 + EditorTileX
              
                If frmMapEditor.optFringe.Value = True Then .Fringe = EditorTileY * 14 + EditorTileX
               
            End With
        Next X
    Next Y
Else
        For Y = 0 To MAX_MAPY
            For X = 0 To MAX_MAPX
                With Map.Tile(X, Y)
                  If frmMapEditor.optBlocked.Value = True Then .Type = TILE_TYPE_BLOCKED
                  If frmMapEditor.optWarp.Value = True Then .Type = TILE_TYPE_WARP
                 
                  If frmMapEditor.optItem.Value = True Then .Type = TILE_TYPE_ITEM

                  If frmMapEditor.optNpcAvoid.Value = True Then .Type = TILE_TYPE_NPCAVOID
                  If frmMapEditor.optKey.Value = True Then .Type = TILE_TYPE_KEY
                  If frmMapEditor.optKeyOpen.Value = True Then .Type = TILE_TYPE_KEYOPEN
            End With
        Next X
    Next Y
End If
End Sub

Private Sub Form_Resize()
    Call ResizeGUI
End Sub

Private Sub Form_Terminate()
    Call GameDestroy
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call GameDestroy
End Sub

Private Sub fraLayers_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub Label2_Click()
    Call SendData("getstats" & END_CHAR)
End Sub

Private Sub Label3_Click()
    Call UpdateInventory
    picInv.Visible = True
End Sub

Private Sub Label4_Click()
    frmTraining.Show vbModal
End Sub

Private Sub Label6_Click()
    Call SendData("trade" & END_CHAR)
End Sub

Private Sub Label7_Click()
    Call GameDestroy
End Sub

Private Sub Label8_Click()
    Call SendData("spells" & END_CHAR)
End Sub

Private Sub Label9_Click()

End Sub

Private Sub picGUI_Click()

End Sub

Private Sub picInv_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

     SOffsetX = X
    SOffsetY = Y
End Sub

Private Sub picInv_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Call MovePicture(frmMirage.picInv, Button, Shift, X, Y)
End Sub

Private Sub picPlayerSpells_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SOffsetX = X
    SOffsetY = Y
End Sub

Private Sub picPlayerSpells_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Call MovePicture(frmMirage.picPlayerSpells, Button, Shift, X, Y)
End Sub

Private Sub picScreen_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call EditorMouseDown(Button, Shift, X, Y)
    Call PlayerSearch(Button, Shift, X, Y)
End Sub

Private Sub picScreen_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call EditorMouseDown(Button, Shift, X, Y)
    If Button = vbRightButton Then
    If Not InEditor Then
        Call CharMove(X, Y)
    End If
    End If
End Sub

Private Sub Socket_DataArrival(ByVal bytesTotal As Long)
    If IsConnected Then
        Call IncomingData(bytesTotal)
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Call HandleKeypresses(KeyAscii)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Call CheckInput(1, KeyCode, Shift)
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Call CheckInput(0, KeyCode, Shift)
        If KeyCode = vbKeyF1 Then
        If Player(MyIndex).Access > 0 Then
            frmadmin.Visible = False
            frmadmin.Visible = True
        End If
    End If
End Sub

Private Sub txtChat_GotFocus()
    frmMirage.picScreen.SetFocus
End Sub

Private Sub picInventory_Click()

End Sub

Private Sub lblUseItem_Click()
    Call SendUseItem(frmMirage.lstInv.ListIndex + 1)
End Sub

Private Sub lblDropItem_Click()
Dim Value As Long
Dim InvNum As Long

    InvNum = frmMirage.lstInv.ListIndex + 1

    If GetPlayerInvItemNum(MyIndex, InvNum) > 0 And GetPlayerInvItemNum(MyIndex, InvNum) <= MAX_ITEMS Then
        If Item(GetPlayerInvItemNum(MyIndex, InvNum)).Type = ITEM_TYPE_CURRENCY Then
            ' Show them the drop dialog
            frmDrop.Show vbModal
        Else
            Call SendDropItem(frmMirage.lstInv.ListIndex + 1, 0)
        End If
    End If
End Sub

Private Sub lblCast_Click()
    If Player(MyIndex).Spell(lstSpells.ListIndex + 1) > 0 Then
        If GetTickCount > Player(MyIndex).AttackTimer + 1000 Then
            If Player(MyIndex).Moving = 0 Then
                Call SendData("cast" & SEP_CHAR & lstSpells.ListIndex + 1 & END_CHAR)
                Player(MyIndex).Attacking = 1
                Player(MyIndex).AttackTimer = GetTickCount
                Player(MyIndex).CastedSpell = YES
            Else
                Call AddText("Cannot cast while walking!", BrightRed)
            End If
        End If
    Else
        Call AddText("No spell here.", BrightRed)
    End If
End Sub

Private Sub lblCancel_Click()
    picInv.Visible = False
End Sub

Private Sub lblSpellsCancel_Click()
    picPlayerSpells.Visible = False
End Sub

Private Sub picSpells_Click()

End Sub

Private Sub picStats_Click()

End Sub

Private Sub picTrain_Click()

End Sub

Private Sub picTrade_Click()

End Sub

Private Sub picQuit_Click()

End Sub

