Attribute VB_Name = "modQuest"
Option Explicit

Public Const MAX_QUESTS As Long = 1000      ' The max amount of quests in the game
Public Const MAX_QUEST_NEEDS As Long = 3    ' The max amount of 'steps' in 1 quest
Public Const MAX_QUEST_REWARDS As Long = 3  ' The max amount of rewards for 1 quest
Public Const MAX_PLAYER_QUESTS As Long = 10 ' The max amount of quests a player can have at once

Public Const QUEST_LENGTH As Long = 512     ' The max amount of characters

Public Const QUEST_STATUS_INCOMPLETE As Byte = 0    ' Quest status - Used for Player().Char.CompletedQuest(QuestNum)
Public Const QUEST_STATUS_COMPLETE As Byte = 1      ' Quest status - Used for Player().Char.CompletedQuest(QuestNum)

Public Const QUEST_PERCENT As Byte = 15     ' The % of damage you do to a npc to get credit for a kill

Public QuestMapNpcNum As Long

Public Enum QuestTypes
    None
    KillNpc
    ItemCollection
    ExploreMap
End Enum

Public Type QuestNeedsUDT
    QuestType As Long
    Data1 As Long       ' KillNPC: Npcnum, ItemCollection: ItemNum
    Data2 As Long       ' How many needed of above
End Type

Public Type RewardsUDT
    ItemNum As Long     ' The ItemNum of the reward
    ItemValue As Long   ' How many you get
    SelectionOnly As Boolean    ' If true, player can select 1 of MAX_QUEST_REWARDS. If false, will get no matter what.
End Type

Public Type QuestUDT
    Name As String * NAME_LENGTH
    Description As String * QUEST_LENGTH
    
    AcceptMessage As String * QUEST_LENGTH          ' Message sent to player when quest is accepted
    DeniedMessage As String * QUEST_LENGTH          ' If the player is denied for whatever reason
    IncompletedMessage As String * QUEST_LENGTH    ' When the player talks to the NPC while on the quest, but not competed
    CompletedMessage As String * QUEST_LENGTH       ' When the player finishes the quest
    
    Repeatable As Boolean       ' If repeatable, the player can do this quest as many times as they'd like
    
    QuestReq As Long
    LevelReq As Long
    ClassReq As Long
    
    QuestNeeds(1 To MAX_QUEST_NEEDS) As QuestNeedsUDT
    Rewards(1 To MAX_QUEST_REWARDS) As RewardsUDT   ' Rewards for completing quest
    RewardExp As Long                               ' How much exp you get for completing quest
    
    StartNPC As Long            ' The NpcNum that can start the quest
    EndNPC As Long              ' The NpcNum you turn the quest into
    
    GiveItemNum As Long         ' THe NPC will give this item to the player on quest Accept.
    GiveItemValue As Long       ' How many of the above item the NPC will give the player on accept
End Type
Public Quest(1 To MAX_QUESTS) As QuestUDT

Public LastQuestClicked As Long ' For Quest GUI
Public CurrentSelectedQuest As Long ' For quest info GUI
Public SelectReward As Long ' for quest selection on turn in

' Used to 'cache' quests
Public Type NpcQuestsRec
    QuestCount As Long
    QuestList() As Long
End Type
Public NpcQuests(1 To MAX_NPCS) As NpcQuestsRec

'
' Quests
'
Public Function Get_QuestData(ByRef QuestNum As Long) As Byte()
Dim QuestData() As Byte
    ReDim QuestData(0 To QuestSize - 1)
    CopyMemory QuestData(0), ByVal VarPtr(Quest(QuestNum)), QuestSize
    Get_QuestData = QuestData
End Function

Public Sub Set_QuestData(ByRef QuestNum As Long, ByRef QuestData() As Byte)
    CopyMemory ByVal VarPtr(Quest(QuestNum)), ByVal VarPtr(QuestData(0)), QuestSize
End Sub

Public Sub QuestEditor()
Dim i As Long
    
    InQuestEditor = True
        
    frmIndex.Show
    frmIndex.lstIndex.Clear
    
    ' Add the names
    For i = 1 To MAX_QUESTS
        frmIndex.lstIndex.AddItem i & ": " & Trim$(Quest(i).Name)
    Next
    
    frmIndex.lstIndex.ListIndex = 0
End Sub

Public Sub QuestEditorInit()
Dim i As Long
Dim n As Long

    With frmQuestEditor
        .scrlStartNpc.Max = MAX_NPCS
        .scrlEndNpc.Max = MAX_NPCS
        .scrlPart.Max = MAX_QUEST_NEEDS
        .scrlRewardIndex.Max = MAX_QUEST_REWARDS
        .scrlPrvQuest.Max = MAX_QUESTS
        .scrlGiveItem.Max = MAX_ITEMS
        
        .txtName.Text = Trim$(Quest(EditorIndex).Name)
        .txtDescription.Text = Trim$(Quest(EditorIndex).Description)
        .scrlStartNpc.Value = Quest(EditorIndex).StartNPC
        .scrlEndNpc.Value = Quest(EditorIndex).EndNPC
        If Quest(EditorIndex).Repeatable Then
            .chkRepeatable.Value = 1
        Else
            .chkRepeatable.Value = 0
        End If
        
        .txtAccept.Text = Trim$(Quest(EditorIndex).AcceptMessage)
        .txtDenied.Text = Trim$(Quest(EditorIndex).DeniedMessage)
        .txtComplete.Text = Trim$(Quest(EditorIndex).CompletedMessage)
        .txtIncomplete.Text = Trim$(Quest(EditorIndex).IncompletedMessage)
        
        .scrlLevel.Value = Quest(EditorIndex).LevelReq
        .scrlPrvQuest.Value = Quest(EditorIndex).QuestReq
        
        '
        ' CLASS REQ
        '
        For i = 0 To MAX_CLASSES
            If i > 0 Then Load .chkClass(i)
            .chkClass(i).Visible = True
            .chkClass(i).Top = .chkClass(i).Top + (.chkClass(i).Height * i)
            
            .chkClass(i).Caption = Trim$(Class(i).Name)
            ' If the flag is true, set the checkbox
            If Quest(EditorIndex).ClassReq And (2 ^ i) Then
                .chkClass(i).Value = 1
            Else
                .chkClass(i).Value = 0
            End If
        Next
    
        ' Class Req Frame
        .frameClasses.Height = .frameClasses.Height + (MAX_CLASSES * .chkClass(0).Height)
        
        '
        ' QUEST NEEDS
        '
        For i = 1 To MAX_QUEST_NEEDS
            Load .cmbQuestType(i)
            Load .cmbNeedsIndex(i)
            Load .txtRequired(i)
            
            ' add the quest types
            .cmbQuestType(i).AddItem "None"
            .cmbQuestType(i).AddItem "Kill NPC"
            .cmbQuestType(i).AddItem "Item Collection"
            .cmbQuestType(i).AddItem "Explore Map"
            
            .cmbQuestType(i).ListIndex = Quest(EditorIndex).QuestNeeds(i).QuestType
            
            .txtRequired(i).Text = Quest(EditorIndex).QuestNeeds(i).Data2
            
            .cmbNeedsIndex(i).AddItem "None"
            Select Case Quest(EditorIndex).QuestNeeds(i).QuestType
                Case QuestTypes.KillNpc
                    For n = 1 To MAX_NPCS
                        .cmbNeedsIndex(i).AddItem n & ": " & Trim$(Npc(n).Name)
                    Next
                Case QuestTypes.ItemCollection
                    For n = 1 To MAX_ITEMS
                        .cmbNeedsIndex(i).AddItem n & ": " & Trim$(Item(n).Name)
                    Next
                Case QuestTypes.ExploreMap
                    For n = 1 To MAX_MAPS
                        .cmbNeedsIndex(i).AddItem n ' TODO: Add mapnames & ": " & Trim$(Map(n).Name)
                    Next
                    .txtRequired(i).Visible = False
            End Select
            .cmbNeedsIndex(i).ListIndex = Quest(EditorIndex).QuestNeeds(i).Data1
        Next
        ' Now show only Needs(1)
        .cmbQuestType(1).Visible = True
        .cmbNeedsIndex(1).Visible = True
        '.txtRequired(1).Visible = True
        
        '
        ' QUEST REWARDS
        '
        For i = 1 To MAX_QUEST_NEEDS
            Load .cmbItemIndex(i)
            Load .txtItemValue(i)
            Load .chkSelectionOnly(i)
            .cmbItemIndex(i).AddItem "None"
            For n = 1 To MAX_ITEMS
                .cmbItemIndex(i).AddItem n & ": " & Trim$(Item(n).Name)
            Next
            .cmbItemIndex(i).ListIndex = Quest(EditorIndex).Rewards(i).ItemNum
            .txtItemValue(i).Text = Quest(EditorIndex).Rewards(i).ItemValue
            If Quest(EditorIndex).Rewards(i).SelectionOnly Then
                .chkSelectionOnly(i).Value = 1
            Else
                .chkSelectionOnly(i).Value = 0
            End If
        Next
        ' Now show only Rewards(1)
        .cmbItemIndex(1).Visible = True
        .txtItemValue(1).Visible = True
        .chkSelectionOnly(1).Visible = True
        
        .txtExp.Text = Quest(EditorIndex).RewardExp
        
        ' GiveItems
        .scrlGiveItem.Value = Quest(EditorIndex).GiveItemNum
        .txtGiveItemValue.Text = Quest(EditorIndex).GiveItemValue
    End With

    frmQuestEditor.Show vbModal
End Sub

Public Sub QuestEditorOk()
Dim i As Long
Dim n As Long

    With frmQuestEditor
        Quest(EditorIndex).Name = .txtName.Text
        Quest(EditorIndex).Description = .txtDescription.Text
        Quest(EditorIndex).StartNPC = .scrlStartNpc.Value
        Quest(EditorIndex).EndNPC = .scrlEndNpc.Value
        Quest(EditorIndex).Repeatable = .chkRepeatable.Value
        
        Quest(EditorIndex).AcceptMessage = .txtAccept.Text
        Quest(EditorIndex).DeniedMessage = .txtDenied.Text
        Quest(EditorIndex).CompletedMessage = .txtComplete.Text
        Quest(EditorIndex).IncompletedMessage = .txtIncomplete.Text
        
        Quest(EditorIndex).LevelReq = .scrlLevel.Value
        Quest(EditorIndex).QuestReq = .scrlPrvQuest.Value
        
        For i = 0 To MAX_CLASSES
            If .chkClass(i).Value Then
                Quest(EditorIndex).ClassReq = Quest(EditorIndex).ClassReq Or (2 ^ i)
            Else
                Quest(EditorIndex).ClassReq = Quest(EditorIndex).ClassReq And Not (2 ^ i)
            End If
        Next
            
        '
        ' QUEST NEEDS
        '
        For i = 1 To MAX_QUEST_NEEDS
            Quest(EditorIndex).QuestNeeds(i).QuestType = .cmbQuestType(i).ListIndex
            Quest(EditorIndex).QuestNeeds(i).Data1 = .cmbNeedsIndex(i).ListIndex
            Quest(EditorIndex).QuestNeeds(i).Data2 = .txtRequired(i).Text
        Next
        
        '
        ' QUEST REWARDS
        '
        For i = 1 To MAX_QUEST_REWARDS
            Quest(EditorIndex).Rewards(i).ItemNum = .cmbItemIndex(i).ListIndex
            Quest(EditorIndex).Rewards(i).ItemValue = .txtItemValue(i).Text
            Quest(EditorIndex).Rewards(i).SelectionOnly = .chkSelectionOnly(i).Value
        Next
        
        Quest(EditorIndex).RewardExp = .txtExp.Text
        
        ' GiveItems
        Quest(EditorIndex).GiveItemNum = .scrlGiveItem.Value
        Quest(EditorIndex).GiveItemValue = .txtGiveItemValue.Text
    End With
    
    SendSaveQuest EditorIndex
    InQuestEditor = False
    Unload frmQuestEditor

    QuestEditor
End Sub

Public Sub QuestEditorCancel()
    InQuestEditor = False
    Unload frmQuestEditor

    QuestEditor
End Sub

'
' Quests
'
Sub SendRequestEditQuest()
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 4
    Buffer.WriteLong SMsgRequestEditQuest
    
    SendData Buffer.ToArray()
End Sub

Sub SendEditQuest()
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 8
    Buffer.WriteLong SMsgEditQuest
    Buffer.WriteLong EditorIndex
    
    SendData Buffer.ToArray()
End Sub

Sub SendSaveQuest(ByVal QuestNum As Long)
Dim Buffer As clsBuffer
Dim QuestData() As Byte
Dim QuestSize As Long

    Set Buffer = New clsBuffer
    
    QuestSize = LenB(Quest(QuestNum))
    ReDim QuestData(0 To QuestSize - 1)

    Buffer.PreAllocate QuestSize + 8
    Buffer.WriteLong SMsgSaveQuest
    Buffer.WriteLong QuestNum
    CopyMemory QuestData(0), ByVal VarPtr(Quest(QuestNum)), QuestSize
    Buffer.WriteBytes QuestData
    
    SendData Buffer.ToArray()
End Sub

Public Function CanAcceptQuest(ByVal Index As Long, ByVal QuestNum As Long) As Boolean
Dim i As Long
        
    ' Check player quests
    If Player(Index).ActiveQuestCount > MAX_PLAYER_QUESTS Then
        ' Error Message
        Exit Function
    End If
    
    ' Check for quest requirement
    If Quest(QuestNum).QuestReq > 0 Then
        If Player(Index).CompletedQuests(Quest(QuestNum).QuestReq) = QUEST_STATUS_INCOMPLETE Then
            ' Error message
            Exit Function
        End If
    End If
    
    ' Check for class requirement
    ' Will check your current class to the quest
    ' Checks the binary flag is set for your class
    If Not Quest(QuestNum).ClassReq And (2 ^ Current_Class(Index)) Then
        Exit Function
    End If
    
    ' Check for level requirement
    If Quest(QuestNum).LevelReq > 0 Then
        ' If there's a level requirement then check if you can use it
        ' Checks if your level is below the req and if so - will exit
        If Current_Level(Index) < Quest(QuestNum).LevelReq Then
            Exit Function
        End If
    End If
    
    CanAcceptQuest = True
End Function

Sub SendAcceptQuest(ByVal QuestNum As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SMsgAcceptQuest
    Buffer.WriteLong QuestNum
    Buffer.WriteLong QuestMapNpcNum
    
    SendData Buffer.ToArray()
End Sub

Public Function OnQuest(ByVal QuestNum As Long) As Long
Dim i As Long
    For i = 1 To MAX_PLAYER_QUESTS
        If Player(MyIndex).QuestProgress(i).QuestNum = QuestNum Then
            OnQuest = i
            Exit Function
        End If
    Next
End Function

Sub SendCompleteQuest(ByVal QuestProgressNum As Long, ByVal SelectReward As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SMsgCompleteQuest
    Buffer.WriteLong QuestProgressNum
    Buffer.WriteLong QuestMapNpcNum
    Buffer.WriteLong SelectReward
    
    SendData Buffer.ToArray()
End Sub

Public Function QuestDescription(ByVal QuestNum As Long) As String
    QuestDescription = Trim$(Quest(QuestNum).Description)
    QuestDescription = Replace$(QuestDescription, "%PLAYERNAME%", Current_Name(MyIndex))
    QuestDescription = Replace$(QuestDescription, "%PLAYERCLASS%", Trim$(Class(Current_Class(MyIndex)).Name))
End Function

Public Sub Update_Npc_Quests()
Dim i As Long
    For i = 1 To MAX_NPCS
        Update_Npcs_Quest i
    Next
End Sub

Public Sub Update_Npcs_Quest(ByVal NpcNum As Long)
Dim i As Long

    ' Make sure it's a valid NPcNum
    If NpcNum <= 0 Then Exit Sub
    
    ' Clear out any old quest
    NpcQuests(NpcNum).QuestCount = 0
    ReDim NpcQuests(NpcNum).QuestList(NpcQuests(NpcNum).QuestCount)
    
    ' If there's a NpcNum then get the quests
    If NpcNum Then
        ' Loop quests and check if this npc is either Start or End
        For i = 1 To MAX_QUESTS
            '
            If Quest(i).StartNPC = NpcNum Or Quest(i).EndNPC = NpcNum Then
                ' Set the count and update the array
                NpcQuests(NpcNum).QuestCount = NpcQuests(NpcNum).QuestCount + 1
                ReDim Preserve NpcQuests(NpcNum).QuestList(NpcQuests(NpcNum).QuestCount)
                ' Set the quest number
                NpcQuests(NpcNum).QuestList(NpcQuests(NpcNum).QuestCount) = i
            End If
        Next
    End If
End Sub

Public Sub SendDropQuest(ByVal QuestProgressNum As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SMsgDropQuest
    Buffer.WriteLong QuestProgressNum
    
    SendData Buffer.ToArray()
End Sub

Public Sub UpdateQuestList()
Dim i As Long
Dim ii As Long

    frmMainGame.lstQuests.Clear
    For i = 1 To MAX_PLAYER_QUESTS
        If Player(MyIndex).QuestProgress(i).QuestNum Then
            frmMainGame.lstQuests.AddItem Trim$(Quest(Player(MyIndex).QuestProgress(i).QuestNum).Name)
            frmMainGame.lstQuests.ItemData(ii) = Player(MyIndex).QuestProgress(i).QuestNum
            ii = ii + 1
        End If
    Next
    
    frmMainGame.lblQuestProgress.Caption = vbNullString
    
    If ii Then
        ' Make sure LastQuest is valid
        If LastQuestClicked < 0 Then Exit Sub
        If LastQuestClicked > frmMainGame.lstQuests.ListCount - 1 Then Exit Sub
        
        frmMainGame.lstQuests.ListIndex = LastQuestClicked
    End If
End Sub
