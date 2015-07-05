VERSION 5.00
Begin VB.Form frmQuests 
   Caption         =   "Form1"
   ClientHeight    =   6390
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7275
   LinkTopic       =   "Form1"
   ScaleHeight     =   426
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   485
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picItemDesc 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4725
      Left            =   5400
      ScaleHeight     =   313
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   120
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   1830
      Begin VB.Label lblItemDescName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ItemDescName"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H003E8CA6&
         Height          =   240
         Left            =   60
         TabIndex        =   15
         Top             =   75
         Width           =   1695
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblRequirement 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Requirement"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H001D2B34&
         Height          =   180
         Left            =   435
         TabIndex        =   14
         Top             =   480
         Width           =   960
      End
      Begin VB.Label lblItemDescReq 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "DescReq"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H001D2B34&
         Height          =   180
         Left            =   600
         TabIndex        =   13
         Top             =   840
         Width           =   630
      End
      Begin VB.Label lblItemName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Type"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H001D2B34&
         Height          =   255
         Left            =   75
         TabIndex        =   12
         Top             =   240
         Width           =   1680
      End
   End
   Begin VB.Frame frmNpcQuest 
      Height          =   6015
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   6735
      Begin VB.CommandButton cmdDrop 
         Caption         =   "Drop"
         Height          =   375
         Left            =   4560
         TabIndex        =   17
         Top             =   5520
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.PictureBox picRewards 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   1575
         Left            =   120
         ScaleHeight     =   105
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   433
         TabIndex        =   10
         Top             =   3840
         Width           =   6495
         Begin VB.Label lblSelected 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Selected"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   0
            TabIndex        =   16
            Top             =   0
            Visible         =   0   'False
            Width           =   765
         End
      End
      Begin VB.CommandButton cmdComplete 
         Caption         =   "Complete"
         Height          =   375
         Left            =   4560
         TabIndex        =   8
         Top             =   5520
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdAccept 
         Caption         =   "Accept"
         Height          =   375
         Left            =   4560
         TabIndex        =   6
         Top             =   5520
         Width           =   975
      End
      Begin VB.CommandButton cmdBack 
         Caption         =   "Back"
         Height          =   375
         Left            =   5640
         TabIndex        =   5
         Top             =   5520
         Width           =   975
      End
      Begin VB.Label lblNeeds 
         AutoSize        =   -1  'True
         Caption         =   "Quest Needs"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   2760
         Width           =   930
      End
      Begin VB.Label lblDescription 
         Caption         =   "Description"
         Height          =   2175
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   6495
      End
      Begin VB.Label lblName 
         Caption         =   "Name"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   4695
      End
   End
   Begin VB.ListBox lstNpcQuests 
      Height          =   2205
      Left            =   240
      TabIndex        =   7
      Top             =   600
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.ListBox lstPlayerQuests 
      Height          =   2205
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.Label lblQuestType 
      Caption         =   "Label1"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   240
      Width           =   6855
   End
End
Attribute VB_Name = "frmQuests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private QuestType As Long
Private SelectReward As Long


Private Sub cmdBack_Click()
    If QuestType = 1 Then
        ' Go back to the npc lst
        frmNpcQuest.Visible = False
        lstNpcQuests.Visible = True
    ElseIf QuestType = 2 Then
        ' Go back to the quest lst
        frmNpcQuest.Visible = False
        lstPlayerQuests.Visible = True
    End If
End Sub




'
' NPC Quests
'
Private Sub lstNpcQuests_Click()
Dim i As Long
Dim RewardCount As Long
Dim SelectRewardCount As Long
Dim QuestNum As Long
Dim QuestProgressNum As Long
Dim rec As RECT
Dim rec_pos As RECT

    QuestType = 1
    
    ' Load up the quest info
    lstNpcQuests.Visible = False
    frmNpcQuest.Visible = True
    picRewards.Cls
    lblSelected.Visible = False
    cmdDrop.Visible = False
    
    ' Load in the data
    QuestNum = lstNpcQuests.ItemData(lstNpcQuests.ListIndex)
    lblName.Caption = Trim$(Quest(QuestNum).Name)
    lblDescription.Caption = QuestDescription(QuestNum) 'Trim$(Quest(QuestNum).Description)
        
    ' Check if you're on quest
    QuestProgressNum = OnQuest(QuestNum)
    If QuestProgressNum > 0 Then
        ' Quest Needs
        lblNeeds.Caption = vbNullString
        For i = 1 To MAX_QUEST_NEEDS
            Select Case Quest(QuestNum).QuestNeeds(i).QuestType
                Case QuestTypes.KillNpc
                    lblNeeds.Caption = lblNeeds.Caption & "You've killed " & Player(MyIndex).QuestProgress(QuestProgressNum).Progress(i) & " / " & Quest(QuestNum).QuestNeeds(i).Data2 & " " & Trim$(Npc(Quest(QuestNum).QuestNeeds(i).Data1).Name) & vbNewLine
                Case QuestTypes.ItemCollection
                    lblNeeds.Caption = lblNeeds.Caption & "You have " & Player(MyIndex).QuestProgress(QuestProgressNum).Progress(i) & " / " & Quest(QuestNum).QuestNeeds(i).Data2 & " " & Trim$(Item(Quest(QuestNum).QuestNeeds(i).Data1).Name) & vbNewLine
                Case QuestTypes.ExploreMap
                    If Player(MyIndex).QuestProgress(QuestProgressNum).Progress(i) Then
                        lblNeeds.Caption = lblNeeds.Caption & "You have explored the area."
                    Else
                        lblNeeds.Caption = lblNeeds.Caption & "You have not explored the area yet."
                    End If
            End Select
        Next
        
        ' Hide the accept button
        cmdAccept.Visible = False
        
        ' If this is the turn in NPC, show the "Complete" button
        If Quest(QuestNum).EndNPC = MapNpc(QuestMapNpcNum).Num Then
            cmdComplete.Visible = True
            cmdComplete.Enabled = True
        
             ' Check if you're done with the quest
            For i = 1 To MAX_QUEST_NEEDS
                ' Check for any progresses that aren't complete
                ' Make sure it's an actual quest need
                If Quest(QuestNum).QuestNeeds(i).QuestType <> QuestTypes.None Then
                    If Player(MyIndex).QuestProgress(QuestProgressNum).Progress(i) < Quest(QuestNum).QuestNeeds(i).Data2 Then
                        cmdComplete.Enabled = False
                        Exit For
                    End If
                End If
            Next
        End If
    Else
        ' Just reset the buttons
        cmdAccept.Visible = True
        cmdComplete.Visible = False
        
        ' Quest Needs
        lblNeeds.Caption = vbNullString
        For i = 1 To MAX_QUEST_NEEDS
            Select Case Quest(QuestNum).QuestNeeds(i).QuestType
                Case QuestTypes.KillNpc
                    lblNeeds.Caption = lblNeeds.Caption & "Kill " & Quest(QuestNum).QuestNeeds(i).Data2 & " " & Trim$(Npc(Quest(QuestNum).QuestNeeds(i).Data1).Name) & vbNewLine
                Case QuestTypes.ItemCollection
                    lblNeeds.Caption = lblNeeds.Caption & "Collect " & Quest(QuestNum).QuestNeeds(i).Data2 & " " & Trim$(Item(Quest(QuestNum).QuestNeeds(i).Data1).Name) & vbNewLine
                Case QuestTypes.ExploreMap
                    lblNeeds.Caption = lblNeeds.Caption & "Explore the correct area."
            End Select
        Next
    End If
    
    ' Draw the rewards
    For i = 1 To MAX_QUEST_REWARDS
        If Quest(QuestNum).Rewards(i).ItemNum Then
            With rec
                .Top = Item(Quest(QuestNum).Rewards(i).ItemNum).Pic * PIC_Y
                .Bottom = .Top + PIC_Y
                .Left = 0
                .Right = .Left + PIC_X
            End With
            
            If Not Quest(QuestNum).Rewards(i).SelectionOnly Then
                RewardCount = RewardCount + 1
                With rec_pos
                    .Top = 16
                    .Bottom = .Top + PIC_Y
                    .Left = 16 + ((16 + 32) * (RewardCount - 1))
                    .Right = .Left + PIC_X
                End With
            Else
                SelectRewardCount = SelectRewardCount + 1
                With rec_pos
                    .Top = (PIC_Y * 2) + 8
                    .Bottom = .Top + PIC_Y
                    .Left = 16 + ((16 + 32) * (SelectRewardCount - 1))
                    .Right = .Left + PIC_X
                End With
            End If
            
            DD_ItemSurf.BltToDC frmQuests.picRewards.hdc, rec, rec_pos
            
            DrawText frmQuests.picRewards.hdc, rec_pos.Left, rec_pos.Top - 10, Quest(QuestNum).Rewards(i).ItemValue, QBColor(White)
        End If
    Next
    If SelectRewardCount Then DrawText frmQuests.picRewards.hdc, 2, (PIC_Y * 2) - 16, "Select one:", vbWhite
    frmQuests.picRewards.Refresh
End Sub

Private Sub cmdAccept_Click()
Dim QuestNum As Long

    'Get the quest num
    QuestNum = lstNpcQuests.ItemData(lstNpcQuests.ListIndex)
    SendAcceptQuest QuestNum
    QuestMapNpcNum = 0
    SelectReward = 0
    Unload frmQuests
End Sub

Private Sub cmdComplete_Click()
Dim QuestNum As Long
Dim QuestProgressNum As Long
    
    QuestNum = lstNpcQuests.ItemData(lstNpcQuests.ListIndex)
    QuestProgressNum = OnQuest(QuestNum)
    SendCompleteQuest QuestProgressNum, SelectReward
    QuestMapNpcNum = 0
    SelectReward = 0
    Unload frmQuests
End Sub


'
' Player Quests
'
Private Sub lstPlayerQuests_Click()
Dim i As Long
Dim RewardCount As Long
Dim SelectRewardCount As Long
Dim QuestNum As Long
Dim QuestProgressNum As Long
Dim rec As RECT
Dim rec_pos As RECT

    QuestType = 2
    
    ' Load up the quest info
    lstPlayerQuests.Visible = False
    frmNpcQuest.Visible = True
    picRewards.Cls
    lblSelected.Visible = False
    
    ' Load in the data
    QuestNum = lstPlayerQuests.ItemData(lstPlayerQuests.ListIndex)
    lblName.Caption = Trim$(Quest(QuestNum).Name)
    lblDescription.Caption = QuestDescription(QuestNum) 'Trim$(Quest(QuestNum).Description)
        
    ' Check if you're on quest
    QuestProgressNum = OnQuest(QuestNum)
    If QuestProgressNum > 0 Then
        ' Quest Needs
        lblNeeds.Caption = vbNullString
        For i = 1 To MAX_QUEST_NEEDS
            Select Case Quest(QuestNum).QuestNeeds(i).QuestType
                Case QuestTypes.KillNpc
                    lblNeeds.Caption = lblNeeds.Caption & "You've killed " & Player(MyIndex).QuestProgress(QuestProgressNum).Progress(i) & " / " & Quest(QuestNum).QuestNeeds(i).Data2 & " " & Trim$(Npc(Quest(QuestNum).QuestNeeds(i).Data1).Name) & vbNewLine
                Case QuestTypes.ItemCollection
                    lblNeeds.Caption = lblNeeds.Caption & "You have " & Player(MyIndex).QuestProgress(QuestProgressNum).Progress(i) & " / " & Quest(QuestNum).QuestNeeds(i).Data2 & " " & Trim$(Item(Quest(QuestNum).QuestNeeds(i).Data1).Name) & vbNewLine
                Case QuestTypes.ExploreMap
                    If Player(MyIndex).QuestProgress(QuestProgressNum).Progress(i) Then
                        lblNeeds.Caption = lblNeeds.Caption & "You have explored the area."
                    Else
                        lblNeeds.Caption = lblNeeds.Caption & "You have not explored the area yet."
                    End If
            End Select
        Next
        
        ' Draw the rewards
        For i = 1 To MAX_QUEST_REWARDS
            If Quest(QuestNum).Rewards(i).ItemNum Then
                With rec
                    .Top = Item(Quest(QuestNum).Rewards(i).ItemNum).Pic * PIC_Y
                    .Bottom = .Top + PIC_Y
                    .Left = 0
                    .Right = .Left + PIC_X
                End With
                
                If Not Quest(QuestNum).Rewards(i).SelectionOnly Then
                    RewardCount = RewardCount + 1
                    With rec_pos
                        .Top = 16
                        .Bottom = .Top + PIC_Y
                        .Left = 16 + ((16 + 32) * (RewardCount - 1))
                        .Right = .Left + PIC_X
                    End With
                Else
                    SelectRewardCount = SelectRewardCount + 1
                    With rec_pos
                        .Top = (PIC_Y * 2) + 8
                        .Bottom = .Top + PIC_Y
                        .Left = 16 + ((16 + 32) * (SelectRewardCount - 1))
                        .Right = .Left + PIC_X
                    End With
                End If
                
                DD_ItemSurf.BltToDC frmQuests.picRewards.hdc, rec, rec_pos
                
                DrawText frmQuests.picRewards.hdc, rec_pos.Left, rec_pos.Top - 10, Quest(QuestNum).Rewards(i).ItemValue, QBColor(White)
            End If
        Next
        If SelectRewardCount Then DrawText frmQuests.picRewards.hdc, 2, (PIC_Y * 2) - 16, "Select one:", vbWhite
        frmQuests.picRewards.Refresh
        
        ' Hide the buttons
        cmdAccept.Visible = False
        cmdComplete.Visible = False
        cmdDrop.Visible = True
    End If

End Sub

Private Sub cmdDrop_Click()
Dim QuestNum As Long
Dim QuestProgressNum As Long
    
    QuestNum = lstPlayerQuests.ItemData(lstPlayerQuests.ListIndex)
    QuestProgressNum = OnQuest(QuestNum)
    SendDropQuest QuestProgressNum
    QuestMapNpcNum = 0
    SelectReward = 0
    Unload frmQuests
End Sub

Private Sub picRewards_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim RewardNum As Long
Dim QuestNum As Long
Dim i As Long
Dim SelectRewardCount As Long
    
    ' Exit if it's player lst
    If QuestType = 2 Then Exit Sub
    
    If Button = vbLeftButton Then
        If QuestType = 1 Then
            QuestNum = lstNpcQuests.ItemData(lstNpcQuests.ListIndex)
        ElseIf QuestType = 2 Then
            QuestNum = lstPlayerQuests.ItemData(lstPlayerQuests.ListIndex)
        End If
                
        ' If this is the turn in NPC, allow clicking
        If Quest(QuestNum).EndNPC <> MapNpc(QuestMapNpcNum).Num Then Exit Sub
        
        RewardNum = IsItem(QuestNum, X, Y)
    
        If RewardNum <> 0 Then
            SelectReward = RewardNum
            ' tell tehm it's selected
            For i = 1 To MAX_QUEST_REWARDS
                If Quest(QuestNum).Rewards(i).ItemNum Then
                    If Quest(QuestNum).Rewards(i).SelectionOnly Then
                        SelectRewardCount = SelectRewardCount + 1
                        If i = RewardNum Then
                            lblSelected.Top = (PIC_Y * 2) + 28
                            lblSelected.Left = ((16 + 32) * (SelectRewardCount - 1))
                            lblSelected.Visible = True
                            Exit Sub
                        End If
                    End If
                End If
            Next
        End If
        
        SelectReward = 0
        lblSelected.Visible = False
    End If
End Sub

Private Sub picRewards_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim RewardNum As Long, ItemNum As Long, ItemType As Long
Dim X2 As Long, Y2 As Long
Dim QuestNum As Long
    
    If QuestType = 1 Then
        QuestNum = lstNpcQuests.ItemData(lstNpcQuests.ListIndex)
    ElseIf QuestType = 2 Then
        QuestNum = lstPlayerQuests.ItemData(lstPlayerQuests.ListIndex)
    End If
    
    RewardNum = IsItem(QuestNum, X, Y)

    If RewardNum <> 0 Then
        ItemNum = Quest(QuestNum).Rewards(RewardNum).ItemNum
        ItemType = Item(ItemNum).Type

        lblItemDescName.Caption = Trim$(Item(ItemNum).Name)
        Select Case ItemType
            Case ITEM_TYPE_NONE
                lblItemName.Caption = "Item"
                lblRequirement.Caption = ItemReq(ItemNum)
                lblItemDescReq.Caption = "Value: " & Quest(QuestNum).Rewards(RewardNum).ItemValue & vbNewLine

            Case ITEM_TYPE_EQUIPMENT
                lblItemName.Caption = EquipmentName(Item(ItemNum).Data1)
                lblRequirement.Caption = ItemReq(ItemNum)
                lblItemDescReq.Caption = ItemDesc(ItemNum)

            Case ITEM_TYPE_POTION
                lblItemName.Caption = "Potion"
                lblRequirement.Caption = ItemReq(ItemNum)
                lblItemDescReq.Caption = ItemDesc(ItemNum)

            Case ITEM_TYPE_KEY
                lblItemName.Caption = "Key"
                lblRequirement.Caption = ItemReq(ItemNum)
                lblItemDescReq.Caption = "Amount: " & Quest(QuestNum).Rewards(RewardNum).ItemValue & vbNewLine

            Case ITEM_TYPE_SPELL
                lblItemName.Caption = "Spell"
                lblRequirement.Caption = ItemReq(ItemNum)
                lblItemDescReq.Caption = Trim$(Spell(Item(ItemNum).Data1).Name) & vbNewLine

        End Select

        lblItemDescReq.Top = lblRequirement.Top + lblRequirement.Height
        picItemDesc.Height = lblItemDescReq.Top + lblItemDescReq.Height

        X2 = (X - picItemDesc.Width) + picRewards.Left
        Y2 = Y '(Y + (picRewards.Top / picRewards.ScaleHeight)) + 20
        
        picItemDesc.Top = Y2
        picItemDesc.Left = X2

        picItemDesc.Visible = True
        Exit Sub
    End If

    picItemDesc.Visible = False
End Sub

Private Function IsItem(ByVal QuestNum As Long, ByVal X As Single, ByVal Y As Single) As Long
Dim tempRec As RECT
Dim i As Long
Dim RewardCount As Long
Dim SelectRewardCount As Long

    For i = 1 To MAX_QUEST_REWARDS
        If Quest(QuestNum).Rewards(i).ItemNum And Quest(QuestNum).Rewards(i).ItemNum <= MAX_ITEMS Then
            If Not Quest(QuestNum).Rewards(i).SelectionOnly Then
                RewardCount = RewardCount + 1
                With tempRec
                    .Top = 16
                    .Bottom = .Top + PIC_Y
                    .Left = 16 + ((16 + 32) * (RewardCount - 1))
                    .Right = .Left + PIC_X
                End With
            Else
                SelectRewardCount = SelectRewardCount + 1
                With tempRec
                    .Top = (PIC_Y * 2) + 8
                    .Bottom = .Top + PIC_Y
                    .Left = 16 + ((16 + 32) * (SelectRewardCount - 1))
                    .Right = .Left + PIC_X
                End With
            End If
            
            If X >= tempRec.Left And X <= tempRec.Right Then
                If Y >= tempRec.Top And Y <= tempRec.Bottom Then
                    IsItem = i
                    Exit Function
                End If
            End If
        End If
    Next
End Function


