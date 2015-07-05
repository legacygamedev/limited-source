VERSION 5.00
Begin VB.Form frmEditNpc 
   BorderStyle     =   0  'None
   Caption         =   "Npc Editor"
   ClientHeight    =   8430
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8535
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   562
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   569
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkBuildingNpc 
      Caption         =   "Building"
      Height          =   255
      Left            =   7200
      TabIndex        =   66
      Top             =   2640
      Width           =   975
   End
   Begin VB.CheckBox chkTreeNpc 
      Caption         =   "Tree"
      Height          =   255
      Left            =   7200
      TabIndex        =   65
      Top             =   2040
      Width           =   975
   End
   Begin VB.HScrollBar scrlNpcQuestNum 
      Height          =   255
      Left            =   5400
      Max             =   50
      TabIndex        =   63
      Top             =   7320
      Width           =   2415
   End
   Begin VB.HScrollBar scrlNpcQuest 
      Height          =   255
      Left            =   5400
      Max             =   50
      Min             =   1
      TabIndex        =   58
      Top             =   6600
      Value           =   1
      Width           =   2415
   End
   Begin VB.PictureBox picSprites 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   240
      ScaleHeight     =   375
      ScaleWidth      =   735
      TabIndex        =   56
      Top             =   8040
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Timer tmrNpcSprite 
      Interval        =   50
      Left            =   7560
      Top             =   7920
   End
   Begin VB.ComboBox cmbShopLink 
      Height          =   315
      Left            =   5160
      Style           =   2  'Dropdown List
      TabIndex        =   55
      Top             =   6000
      Width           =   3015
   End
   Begin VB.CheckBox chkLinkWithShop 
      Caption         =   "Link With Shop"
      Height          =   255
      Left            =   4680
      TabIndex        =   54
      Top             =   5640
      Width           =   1455
   End
   Begin VB.ComboBox cmbHitOnlyWith 
      Height          =   315
      Left            =   5160
      Style           =   2  'Dropdown List
      TabIndex        =   53
      Top             =   5160
      Width           =   3015
   End
   Begin VB.CheckBox chkHitOnlyWith 
      Caption         =   "Hit Only With"
      Height          =   255
      Left            =   4680
      TabIndex        =   52
      Top             =   4800
      Width           =   1455
   End
   Begin VB.ComboBox cmbNpcRespawn 
      Height          =   315
      Left            =   5160
      Style           =   2  'Dropdown List
      TabIndex        =   51
      Top             =   4320
      Width           =   3015
   End
   Begin VB.CheckBox chkNpcRespawn 
      Caption         =   "Respawn"
      Height          =   255
      Left            =   4680
      TabIndex        =   50
      Top             =   3960
      Width           =   975
   End
   Begin VB.ComboBox cmbNpcBehavior 
      Height          =   315
      ItemData        =   "frmEditNpc.frx":0000
      Left            =   5400
      List            =   "frmEditNpc.frx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   49
      Top             =   3480
      Width           =   2775
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2070
      Left            =   5160
      ScaleHeight     =   136
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   104
      TabIndex        =   46
      Top             =   1200
      Width           =   1590
      Begin VB.PictureBox picSprite 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   960
         Left            =   555
         ScaleHeight     =   930
         ScaleWidth      =   450
         TabIndex        =   47
         Top             =   555
         Width           =   480
      End
   End
   Begin VB.CheckBox chkBigNpc 
      Caption         =   "Big Npc"
      Height          =   255
      Left            =   7200
      TabIndex        =   45
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton cmdEditNpcCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4560
      TabIndex        =   44
      Top             =   7920
      Width           =   2055
   End
   Begin VB.CommandButton cmdEditNpcOk 
      Caption         =   "OK"
      Height          =   375
      Left            =   1920
      TabIndex        =   43
      Top             =   7920
      Width           =   2175
   End
   Begin VB.TextBox txtDropChance 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2760
      TabIndex        =   42
      Text            =   "0"
      Top             =   7200
      Width           =   1215
   End
   Begin VB.HScrollBar scrlDropValue 
      Height          =   255
      Left            =   1320
      Max             =   1000
      TabIndex        =   40
      Top             =   6840
      Width           =   2655
   End
   Begin VB.HScrollBar scrlDropNum 
      Height          =   255
      Left            =   1320
      Max             =   255
      TabIndex        =   38
      Top             =   6480
      Width           =   2655
   End
   Begin VB.HScrollBar scrlDropItem 
      Height          =   255
      Left            =   1320
      Max             =   255
      Min             =   1
      TabIndex        =   35
      Top             =   5760
      Value           =   1
      Width           =   2655
   End
   Begin VB.TextBox txtNpcSpawnSecs 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2760
      TabIndex        =   29
      Text            =   "0"
      Top             =   5040
      Width           =   1215
   End
   Begin VB.HScrollBar scrlNpcExpGiven 
      Height          =   255
      Left            =   1200
      TabIndex        =   26
      Top             =   4560
      Width           =   2775
   End
   Begin VB.ComboBox cmbExpType 
      Height          =   315
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   24
      Top             =   4080
      Width           =   2775
   End
   Begin VB.HScrollBar scrlNpcStartHP 
      Height          =   255
      Left            =   1200
      Max             =   1000
      TabIndex        =   21
      Top             =   3360
      Width           =   2775
   End
   Begin VB.HScrollBar scrlNpcRange 
      Height          =   255
      Left            =   1200
      Max             =   255
      TabIndex        =   18
      Top             =   2880
      Width           =   2775
   End
   Begin VB.HScrollBar scrlNpcSprite 
      Height          =   255
      Left            =   5280
      Max             =   100
      TabIndex        =   15
      Top             =   720
      Width           =   2535
   End
   Begin VB.HScrollBar scrlNpcMAGI 
      Height          =   255
      Left            =   1200
      Max             =   255
      TabIndex        =   9
      Top             =   2160
      Width           =   2775
   End
   Begin VB.HScrollBar scrlNpcSPEED 
      Height          =   255
      Left            =   1200
      Max             =   255
      TabIndex        =   8
      Top             =   1680
      Width           =   2775
   End
   Begin VB.HScrollBar scrlNpcDEF 
      Height          =   255
      Left            =   1200
      Max             =   255
      TabIndex        =   7
      Top             =   1200
      Width           =   2775
   End
   Begin VB.HScrollBar scrlNpcSTR 
      Height          =   255
      Left            =   1200
      Max             =   255
      TabIndex        =   2
      Top             =   720
      Width           =   2775
   End
   Begin VB.TextBox txtNpcName 
      Height          =   285
      Left            =   2640
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   3615
   End
   Begin VB.Label lblNpcQuestNum 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   195
      Left            =   7920
      TabIndex        =   64
      Top             =   7320
      Width           =   90
   End
   Begin VB.Label lblNpcQuestName 
      Height          =   255
      Left            =   5400
      TabIndex        =   62
      Top             =   6960
      Width           =   2775
   End
   Begin VB.Label Label20 
      Caption         =   "Number"
      Height          =   255
      Left            =   4680
      TabIndex        =   61
      Top             =   7320
      Width           =   615
   End
   Begin VB.Label Label19 
      Caption         =   "Quest"
      Height          =   255
      Left            =   4680
      TabIndex        =   60
      Top             =   6960
      Width           =   615
   End
   Begin VB.Label lblNpcQuest 
      AutoSize        =   -1  'True
      Caption         =   "1"
      Height          =   195
      Left            =   7920
      TabIndex        =   59
      Top             =   6600
      Width           =   90
   End
   Begin VB.Label Label18 
      Caption         =   "Quests"
      Height          =   255
      Left            =   4680
      TabIndex        =   57
      Top             =   6600
      Width           =   495
   End
   Begin VB.Label Label17 
      Caption         =   "Behavior"
      Height          =   255
      Left            =   4560
      TabIndex        =   48
      Top             =   3480
      Width           =   735
   End
   Begin VB.Label lblDropValue 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   195
      Left            =   4080
      TabIndex        =   41
      Top             =   6840
      Width           =   90
   End
   Begin VB.Label lblDropNum 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   195
      Left            =   4080
      TabIndex        =   39
      Top             =   6480
      Width           =   90
   End
   Begin VB.Label lblDropItemName 
      Height          =   255
      Left            =   1320
      TabIndex        =   37
      Top             =   6120
      Width           =   2895
   End
   Begin VB.Label lblDropItem 
      AutoSize        =   -1  'True
      Caption         =   "1"
      Height          =   195
      Left            =   4080
      TabIndex        =   36
      Top             =   5760
      Width           =   90
   End
   Begin VB.Label Label16 
      Caption         =   "Drop Item Chance 1 out of "
      Height          =   255
      Left            =   600
      TabIndex        =   34
      Top             =   7200
      Width           =   1935
   End
   Begin VB.Label Label15 
      Caption         =   "Value"
      Height          =   255
      Left            =   360
      TabIndex        =   33
      Top             =   6840
      Width           =   855
   End
   Begin VB.Label Label14 
      Caption         =   "Number"
      Height          =   255
      Left            =   360
      TabIndex        =   32
      Top             =   6480
      Width           =   855
   End
   Begin VB.Label Label13 
      Caption         =   "Item"
      Height          =   255
      Left            =   360
      TabIndex        =   31
      Top             =   6120
      Width           =   855
   End
   Begin VB.Label Label12 
      Caption         =   "Dropping"
      Height          =   255
      Left            =   360
      TabIndex        =   30
      Top             =   5760
      Width           =   855
   End
   Begin VB.Label Label11 
      Caption         =   "Spawn Rate (in Seconds)"
      Height          =   255
      Left            =   720
      TabIndex        =   28
      Top             =   5040
      Width           =   1935
   End
   Begin VB.Label lblNpcExpGiven 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   195
      Left            =   4080
      TabIndex        =   27
      Top             =   4560
      Width           =   90
   End
   Begin VB.Label Label10 
      Caption         =   "Exp Given"
      Height          =   255
      Left            =   360
      TabIndex        =   25
      Top             =   4560
      Width           =   855
   End
   Begin VB.Label Label9 
      Caption         =   "Exp Type"
      Height          =   255
      Left            =   360
      TabIndex        =   23
      Top             =   4080
      Width           =   855
   End
   Begin VB.Label lblNpcStartHP 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   195
      Left            =   4080
      TabIndex        =   22
      Top             =   3360
      Width           =   90
   End
   Begin VB.Label Label8 
      Caption         =   "Start HP"
      Height          =   255
      Left            =   360
      TabIndex        =   20
      Top             =   3360
      Width           =   855
   End
   Begin VB.Label lblNpcRange 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   195
      Left            =   4080
      TabIndex        =   19
      Top             =   2880
      Width           =   90
   End
   Begin VB.Label Label7 
      Caption         =   "Range"
      Height          =   255
      Left            =   360
      TabIndex        =   17
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label lblNpcSprite 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   195
      Left            =   7920
      TabIndex        =   16
      Top             =   720
      Width           =   90
   End
   Begin VB.Label Label6 
      Caption         =   "Sprite"
      Height          =   255
      Left            =   4560
      TabIndex        =   14
      Top             =   720
      Width           =   615
   End
   Begin VB.Label lblNpcMAGI 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   195
      Left            =   4080
      TabIndex        =   13
      Top             =   2160
      Width           =   90
   End
   Begin VB.Label lblNpcSPEED 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   195
      Left            =   4080
      TabIndex        =   12
      Top             =   1680
      Width           =   90
   End
   Begin VB.Label lblNpcDEF 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   195
      Left            =   4080
      TabIndex        =   11
      Top             =   1200
      Width           =   90
   End
   Begin VB.Label lblNpcSTR 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   195
      Left            =   4080
      TabIndex        =   10
      Top             =   720
      Width           =   90
   End
   Begin VB.Label Label5 
      Caption         =   "Magic"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Speed"
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Defence"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Strength"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Name"
      Height          =   255
      Left            =   1920
      TabIndex        =   1
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "frmEditNpc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'   This file is part of the Cerberus Engine 2nd Edition.
'
'    The Cerberus Engine 2nd Edition is free software; you can redistribute it
'    and/or modify it under the terms of the GNU General Public License as
'    published by the Free Software Foundation; either version 2 of the License,
'    or (at your option) any later version.
'
'    Cerberus 2nd Edition is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License
'    along with Cerberus 2nd Edition; if not, write to the Free Software
'    Foundation, Inc., 51 Franklin St, Fifth Floor, Boston, MA  02110-1301  USA

Option Explicit

Private Sub scrlNpcSTR_Change()
    lblNpcSTR.Caption = STR(scrlNpcSTR.Value)
End Sub

Private Sub scrlNpcDEF_Change()
    lblNpcDEF.Caption = STR(scrlNpcDEF.Value)
End Sub

Private Sub scrlNpcSPEED_Change()
    lblNpcSPEED.Caption = STR(scrlNpcSPEED.Value)
End Sub

Private Sub scrlNpcMAGI_Change()
    lblNpcMAGI.Caption = STR(scrlNpcMAGI.Value)
End Sub

Private Sub scrlNpcRange_Change()
    lblNpcRange.Caption = STR(scrlNpcRange.Value)
End Sub

Private Sub scrlNpcStartHP_Change()
    lblNpcStartHP.Caption = STR(scrlNpcStartHP.Value)
End Sub

Private Sub scrlNpcExpGiven_Change()
    lblNpcExpGiven.Caption = STR(scrlNpcExpGiven.Value)
End Sub

Private Sub scrlNpcQuest_Change()
    lblNpcQuest.Caption = STR(scrlNpcQuest.Value)
    scrlNpcQuestNum.Value = Npc(EditorIndex).QuestNPC(scrlNpcQuest.Value)
End Sub

Private Sub scrlNpcQuestNum_Change()
    lblNpcQuestNum.Caption = STR(scrlNpcQuestNum.Value)
    If scrlNpcQuestNum.Value > 0 Then
        lblNpcQuestName.Caption = Trim(Quest(scrlNpcQuestNum.Value).Name)
    Else
        lblNpcQuestName.Caption = ""
    End If
    Npc(EditorIndex).QuestNPC(scrlNpcQuest.Value) = scrlNpcQuestNum.Value
End Sub

Private Sub scrlDropItem_Change()
    txtDropChance.Text = Npc(EditorIndex).ItemNPC(scrlDropItem.Value).Chance
    scrlDropNum.Value = Npc(EditorIndex).ItemNPC(scrlDropItem.Value).ItemNum
    scrlDropValue.Value = Npc(EditorIndex).ItemNPC(scrlDropItem.Value).ItemValue
    lblDropItem.Caption = scrlDropItem.Value
End Sub

Private Sub scrlDropNum_Change()
    lblDropNum.Caption = STR(scrlDropNum.Value)
    lblDropItemName.Caption = ""
    If scrlDropNum.Value > 0 Then
        lblDropItemName.Caption = Trim(Item(scrlDropNum.Value).Name)
    End If
    Npc(EditorIndex).ItemNPC(scrlDropItem.Value).ItemNum = scrlDropNum.Value
End Sub

Private Sub scrlDropValue_Change()
    Npc(EditorIndex).ItemNPC(scrlDropItem.Value).ItemValue = scrlDropValue.Value
    lblDropValue.Caption = STR(scrlDropValue.Value)
End Sub

Private Sub txtDropChance_Change()
    Npc(EditorIndex).ItemNPC(scrlDropItem.Value).Chance = Val(txtDropChance.Text)
End Sub

Private Sub scrlNpcSprite_Change()
    lblNpcSprite.Caption = STR(scrlNpcSprite.Value)
End Sub

Private Sub chkBigNpc_Click()
frmCClient.ScaleMode = 3
    If chkBigNpc.Value = Checked Then
        chkTreeNpc.Value = Unchecked
        chkBuildingNpc.Value = Unchecked
        frmEditNpc.picSprites.Picture = LoadPicture(App.Path & "\GFX\bigsprites.bmp")
        picSprite.Width = 64
        picSprite.Height = 64
        picSprite.Top = 36
        picSprite.Left = 20
        frmEditNpc.scrlNpcSprite.Value = 0
    Else
        frmEditNpc.picSprites.Picture = LoadPicture(App.Path & "\GFX\sprites.bmp")
        picSprite.Width = 32
        picSprite.Height = 64
        picSprite.Top = 37
        picSprite.Left = 37
        frmEditNpc.scrlNpcSprite.Value = 0
    End If
End Sub

Private Sub chkTreeNpc_Click()
frmCClient.ScaleMode = 3
    If chkTreeNpc.Value = Checked Then
        chkBigNpc.Value = Unchecked
        chkBuildingNpc.Value = Unchecked
        frmEditNpc.picSprites.Picture = LoadPicture(App.Path & "\GFX\treesprites.bmp")
        picSprite.Width = 96
        picSprite.Height = 128
        picSprite.Top = 4
        picSprite.Left = 4
        frmEditNpc.scrlNpcSprite.Value = 0
    Else
        frmEditNpc.picSprites.Picture = LoadPicture(App.Path & "\GFX\sprites.bmp")
        picSprite.Width = 32
        picSprite.Height = 64
        picSprite.Top = 37
        picSprite.Left = 37
        frmEditNpc.scrlNpcSprite.Value = 0
    End If
End Sub

Private Sub cmbNpcBehavior_Click()
Dim i As Long
    If cmbNpcBehavior.ListIndex = NPC_BEHAVIOR_RESOURCE Then
        chkNpcRespawn.Value = Checked
        chkNpcRespawn.Enabled = True
        cmbNpcRespawn.Enabled = True
        cmbNpcRespawn.Clear
        cmbNpcRespawn.AddItem "None", 0
        cmbNpcRespawn.ListIndex = 0
        For i = 1 To MAX_NPCS
            cmbNpcRespawn.AddItem i & ": " & Trim(Npc(i).Name)
        Next i
        chkHitOnlyWith.Value = Checked
        chkHitOnlyWith.Enabled = True
        cmbHitOnlyWith.Enabled = True
        cmbHitOnlyWith.Clear
        cmbHitOnlyWith.AddItem "None", 0
        cmbHitOnlyWith.ListIndex = 0
        For i = 1 To MAX_ITEMS
            If Item(i).Type = ITEM_TYPE_TOOL Then
                cmbHitOnlyWith.AddItem i & ": " & Trim(Item(i).Name)
            Else
                cmbHitOnlyWith.AddItem i & ": " & "Not Tool Type"
            End If
        Next i
        chkLinkWithShop.Value = Unchecked
        chkLinkWithShop.Enabled = False
        cmbShopLink.Enabled = False
        cmbShopLink.Clear
        cmbShopLink.AddItem "None", 0
        cmbShopLink.ListIndex = 0
        scrlNpcQuest.Value = 1
        scrlNpcQuest.Enabled = False
        scrlNpcQuestNum.Value = 0
        scrlNpcQuestNum.Enabled = False
        lblNpcQuest.Caption = 0
        lblNpcQuest.Enabled = False
        lblNpcQuestNum.Caption = 0
        lblNpcQuestNum.Enabled = False
        scrlNpcRange.Enabled = False
        scrlNpcSTR.Enabled = False
        scrlNpcDEF.Enabled = True
        scrlNpcSPEED.Enabled = False
        scrlNpcMAGI.Enabled = False
        chkTreeNpc.Enabled = True
        chkTreeNpc.Value = Unchecked
        chkBuildingNpc.Enabled = True
        chkBuildingNpc.Value = Unchecked
    ElseIf cmbNpcBehavior.ListIndex = NPC_BEHAVIOR_SHOPKEEPER Then
        chkNpcRespawn.Value = Unchecked
        chkNpcRespawn.Enabled = False
        cmbNpcRespawn.Enabled = False
        cmbNpcRespawn.Clear
        cmbNpcRespawn.AddItem "None", 0
        cmbNpcRespawn.ListIndex = 0
        chkHitOnlyWith.Value = Unchecked
        chkHitOnlyWith.Enabled = False
        cmbHitOnlyWith.Clear
        cmbHitOnlyWith.AddItem "None", 0
        cmbHitOnlyWith.ListIndex = 0
        cmbHitOnlyWith.Enabled = False
        chkLinkWithShop.Value = Checked
        chkLinkWithShop.Enabled = True
        cmbShopLink.Enabled = True
        cmbShopLink.Clear
        cmbShopLink.AddItem "None", 0
        cmbShopLink.ListIndex = 0
        For i = 1 To MAX_SHOPS
            cmbShopLink.AddItem i & ": " & Trim(Shop(i).Name)
        Next i
        scrlNpcQuest.Value = 1
        scrlNpcQuest.Enabled = False
        scrlNpcQuestNum.Value = 0
        scrlNpcQuestNum.Enabled = False
        lblNpcQuest.Caption = 0
        lblNpcQuest.Enabled = False
        lblNpcQuestNum.Caption = 0
        lblNpcQuestNum.Enabled = False
        scrlNpcRange.Enabled = False
        scrlNpcSTR.Enabled = False
        scrlNpcDEF.Enabled = False
        scrlNpcSPEED.Enabled = False
        scrlNpcMAGI.Enabled = False
        scrlNpcMAGI.Enabled = False
        chkTreeNpc.Enabled = False
        chkTreeNpc.Value = Unchecked
        chkBuildingNpc.Enabled = False
        chkBuildingNpc.Value = Unchecked
     ElseIf cmbNpcBehavior.ListIndex = NPC_BEHAVIOR_FRIENDLY Then
        chkNpcRespawn.Value = Unchecked
        chkNpcRespawn.Enabled = False
        cmbNpcRespawn.Enabled = False
        cmbNpcRespawn.Clear
        cmbNpcRespawn.AddItem "None", 0
        cmbNpcRespawn.ListIndex = 0
        chkHitOnlyWith.Value = Unchecked
        chkHitOnlyWith.Enabled = False
        cmbHitOnlyWith.Clear
        cmbHitOnlyWith.AddItem "None", 0
        cmbHitOnlyWith.ListIndex = 0
        cmbHitOnlyWith.Enabled = False
        chkLinkWithShop.Value = Checked
        chkLinkWithShop.Enabled = False
        cmbShopLink.Enabled = False
        cmbShopLink.Clear
        cmbShopLink.AddItem "None", 0
        cmbShopLink.ListIndex = 0
        scrlNpcQuest.Max = MAX_NPC_QUESTS
        scrlNpcQuestNum.Max = MAX_QUESTS
        scrlNpcQuest.Value = 1
        scrlNpcQuest.Enabled = True
        'scrlNpcQuestNum.Value = 0
        scrlNpcQuestNum.Enabled = True
        lblNpcQuest.Caption = 1
        lblNpcQuest.Enabled = True
        lblNpcQuestNum.Caption = 0
        lblNpcQuestNum.Enabled = True
        scrlNpcRange.Enabled = False
        scrlNpcSTR.Enabled = False
        scrlNpcDEF.Enabled = False
        scrlNpcSPEED.Enabled = False
        scrlNpcMAGI.Enabled = False
        scrlNpcMAGI.Enabled = False
        chkTreeNpc.Enabled = False
        chkTreeNpc.Value = Unchecked
        chkBuildingNpc.Enabled = False
        chkBuildingNpc.Value = Unchecked
     Else
        chkNpcRespawn.Value = Checked
        chkNpcRespawn.Enabled = True
        cmbNpcRespawn.Enabled = True
        cmbNpcRespawn.Clear
        cmbNpcRespawn.AddItem "None", 0
        cmbNpcRespawn.ListIndex = 0
        For i = 1 To MAX_NPCS
            cmbNpcRespawn.AddItem i & ": " & Trim(Npc(i).Name)
        Next i
        chkHitOnlyWith.Value = Checked
        chkHitOnlyWith.Enabled = True
        cmbHitOnlyWith.Enabled = True
        cmbHitOnlyWith.Clear
        cmbHitOnlyWith.AddItem "None", 0
        cmbHitOnlyWith.ListIndex = 0
        For i = 1 To MAX_ITEMS
            If Item(i).Type = ITEM_TYPE_WEAPON Then
                cmbHitOnlyWith.AddItem i & ": " & Trim(Item(i).Name)
            Else
                cmbHitOnlyWith.AddItem i & ": " & "Not Weapon Type"
            End If
        Next i
        chkLinkWithShop.Value = Unchecked
        chkLinkWithShop.Enabled = False
        cmbShopLink.Enabled = False
        cmbShopLink.Clear
        cmbShopLink.AddItem "None", 0
        cmbShopLink.ListIndex = 0
        scrlNpcQuest.Value = 1
        scrlNpcQuest.Enabled = False
        scrlNpcQuestNum.Value = 0
        scrlNpcQuestNum.Enabled = False
        lblNpcQuest.Caption = 0
        lblNpcQuest.Enabled = False
        lblNpcQuestNum.Caption = 0
        lblNpcQuestNum.Enabled = False
        scrlNpcRange.Enabled = True
        scrlNpcSTR.Enabled = True
        scrlNpcDEF.Enabled = True
        scrlNpcSPEED.Enabled = True
        scrlNpcMAGI.Enabled = True
        scrlNpcMAGI.Enabled = True
        chkTreeNpc.Enabled = False
        chkTreeNpc.Value = Unchecked
        chkBuildingNpc.Enabled = False
        chkBuildingNpc.Value = Unchecked
    End If
End Sub

Private Sub cmdEditNpcOK_Click()
    Call NpcEditorOk
End Sub

Private Sub cmdEditNpcCancel_Click()
    Call NpcEditorCancel
End Sub

Private Sub tmrNpcSprite_Timer()
    Call NpcEditorBltSprite
End Sub
