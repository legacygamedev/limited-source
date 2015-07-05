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

Public Enum QuestTypes
    None
    KillNpc
    ItemCollection
    ExploreMap
End Enum

Public Type QuestNeedsUDT
    QuestType As Long
    Data1 As Long       ' KillNPC: Npcnum, ItemCollection: ItemNum, ExploreMap: MapNum
    Data2 As Long       ' How many needed of above, ExploreMap: always 1
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

Public Type NpcQuestsRec
    QuestCount As Long
    QuestList() As Long
End Type
Public NpcQuests(1 To MAX_NPCS) As NpcQuestsRec

Public Function QuestCount() As Long
Dim i As Long
Dim n As Long

    For i = 1 To MAX_QUESTS
        If Trim$(Quest(i).Name) <> vbNullString Then
            n = n + 1
        End If
    Next
    QuestCount = n
End Function
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

Public Sub CacheQuests()
Dim Buffer As clsBuffer
Dim i As Long
    
    Set Buffer = New clsBuffer

    Buffer.PreAllocate ((QuestSize + 4) * QuestCount) + 4
    Buffer.WriteLong QuestCount
    
    For i = 1 To MAX_QUESTS
        If Trim$(Quest(i).Name) <> vbNullString Then
            Buffer.WriteLong i
            Buffer.WriteBytes Get_QuestData(i)
        End If
    Next
    
    Buffer.CompressBuffer
    
    QuestsCache() = Buffer.ToArray()
End Sub

Sub CheckQuests()
Dim i As Long
    
    For i = 1 To MAX_QUESTS
        If Not FileExist(QuestPath & "\quest" & i & ".mir", True) Then
            SaveQuest i
        End If
    Next
End Sub

Sub SaveQuest(ByVal QuestNum As Long)
Dim FileName As String
Dim f  As Long

    FileName = QuestPath & "\quest" & QuestNum & ".mir"
        
    f = FreeFile
    Open FileName For Binary As #f
        Put #f, , Quest(QuestNum)
    Close #f
End Sub

Sub LoadQuests()
Dim FileName As String
Dim i As Long
Dim f As Long

    CheckQuests
    
    For i = 1 To MAX_QUESTS
        FileName = QuestPath & "\quest" & i & ".mir"
        f = FreeFile
        Open FileName For Binary As #f
            Get #f, , Quest(i)
        Close #f
    Next
End Sub

Sub ClearQuests()
Dim i As Long
    For i = 1 To MAX_QUESTS
        ClearQuest i
    Next
End Sub

Public Sub ClearQuest(ByVal QuestNum As Long)
    ZeroMemory ByVal VarPtr(Quest(QuestNum)), LenB(Quest(QuestNum))
    Quest(QuestNum).Name = vbNullString
    Quest(QuestNum).Description = vbNullString
    Quest(QuestNum).AcceptMessage = vbNullString
    Quest(QuestNum).DeniedMessage = vbNullString
    Quest(QuestNum).IncompletedMessage = vbNullString
    Quest(QuestNum).CompletedMessage = vbNullString
End Sub

Public Function CanAcceptQuest(ByVal Index As Long, ByVal QuestNum As Long) As Boolean
Dim i As Long
        
    ' Check if the player is already on the quest
    If OnQuest(Index, QuestNum) Then
        SendPlayerMsg Index, "You are already on this quest.", BrightRed
        Exit Function
    End If
     
    ' Check if quest is repeatable - If it's not, you can only do it once
    If Not Quest(QuestNum).Repeatable Then
        ' Check if the player already completed this quest
        If Player(Index).Char.CompletedQuests(QuestNum) = QUEST_STATUS_COMPLETE Then
            Exit Function
        End If
    End If
    
    ' Check player quests
    If Player(Index).Char.ActiveQuestCount > MAX_PLAYER_QUESTS Then
        ' Error Message
        SendPlayerMsg Index, "You can not accept anymore quests at this time.", BrightRed
        Exit Function
    End If
    
    ' Check for quest requirement
    If Quest(QuestNum).QuestReq > 0 Then
        If Player(Index).Char.CompletedQuests(Quest(QuestNum).QuestReq) = QUEST_STATUS_INCOMPLETE Then
            ' Error message
            'SendPlayerMsg Index, Quest(QuestNum).DeniedMessage, BrightRed
            SendPlayerQuestMsg Index, QuestNum, Quest(QuestNum).DeniedMessage, True, BrightRed
            Exit Function
        End If
    End If
    
    ' Check for class requirement
    ' Will check your current class to the quest
    ' Checks the binary flag is set for your class
    If Not Quest(QuestNum).ClassReq And (2 ^ Current_Class(Index)) Then
        SendActionMsg Current_Map(Index), "Your class can not accept this quest.", ActionColor, ACTIONMSG_SCREEN, 0, 0, Index
        'SendPlayerMsg Index, Quest(QuestNum).DeniedMessage, BrightRed
        SendPlayerQuestMsg Index, QuestNum, Quest(QuestNum).DeniedMessage, True, BrightRed
        Exit Function
    End If
    
    ' Check for level requirement
    If Quest(QuestNum).LevelReq > 0 Then
        ' If there's a level requirement then check if you can use it
        ' Checks if your level is below the req and if so - will exit
        If Current_Level(Index) < Quest(QuestNum).LevelReq Then
            SendActionMsg Current_Map(Index), "Level Required: " & Quest(QuestNum).LevelReq, ActionColor, ACTIONMSG_SCREEN, 0, 0, Index
            'SendPlayerMsg Index, Quest(QuestNum).DeniedMessage, BrightRed
            SendPlayerQuestMsg Index, QuestNum, Quest(QuestNum).DeniedMessage, True, BrightRed
            Exit Function
        End If
    End If
    
    ' Checks for GiveItem
    If Quest(QuestNum).GiveItemNum Then
        If FindOpenInvSlot(Index, Quest(QuestNum).GiveItemNum) <= 0 Then
            SendPlayerMsg Index, "You do not have enough inventory space to accept this quest.", BrightRed
            Exit Function
        End If
    End If
    
    CanAcceptQuest = True
End Function

Public Sub OnQuestAccept(ByVal Index As Long, ByVal QuestNum As Long)
Dim i As Long
Dim n As Long

    ' Check if the player can accept the quest
    If CanAcceptQuest(Index, QuestNum) Then
        ' Find the first open player quest slot
        For i = 1 To MAX_PLAYER_QUESTS
            If Player(Index).Char.QuestProgress(i).QuestNum = 0 Then
                ' Set the quest
                Player(Index).Char.QuestProgress(i).QuestNum = QuestNum
                 ' Update active quest count
                Player(Index).Char.ActiveQuestCount = Player(Index).Char.ActiveQuestCount + 1
                
                ' Tell them they accepted it
                SendActionMsg Current_Map(Index), "You have accepted a quest!", AlertColor, ACTIONMSG_SCREEN, 0, 0, Index
                'SendPlayerMsg Index, Quest(QuestNum).AcceptMessage, AlertColor
                SendPlayerQuestMsg Index, QuestNum, Quest(QuestNum).AcceptMessage, True, AlertColor
                
                ' Gives the player the GiveITem
                If Quest(QuestNum).GiveItemNum Then
                    GiveItem Index, Quest(QuestNum).GiveItemNum, Quest(QuestNum).GiveItemValue
                End If
                
                ' Used only for ItemCollection - We need to count how many of the item we already have
                For n = 1 To MAX_QUEST_NEEDS
                    If Quest(QuestNum).QuestNeeds(n).QuestType = QuestTypes.ItemCollection Then
                        Player(Index).Char.QuestProgress(i).Progress(n) = Clamp(Current_InvItemCount(Index, Quest(QuestNum).QuestNeeds(n).Data1), 0, Quest(QuestNum).QuestNeeds(n).Data2)
                        Exit For
                    End If
                Next
                SendPlayerQuest Index, i
                Exit Sub
            End If
        Next
        ' If we get here there was an error somehow
        SendPlayerMsg Index, "You can not accept the quest at this time.", AlertColor
    End If
End Sub

Public Function Current_InvItemCount(ByVal Index As Long, ByVal ItemNum As Long) As Long
Dim i As Long
    For i = 1 To MAX_INV
        If Current_InvItemNum(Index, i) = ItemNum Then
            Current_InvItemCount = Current_InvItemCount + Current_InvItemValue(Index, i)
        End If
    Next
End Function

' Num is used for both NpcNum and ItemNum
Public Sub OnUpdateQuestProgress(ByVal Index As Long, ByVal Num As Long, ByVal Value As Long, ByVal AddValue As Boolean, ByVal QuestType As QuestTypes)
Dim i As Long
Dim QuestNum As Long
Dim QuestNeedsNum As Long

    ' Check if this is needed for the quest
    For i = 1 To MAX_PLAYER_QUESTS
        QuestNum = Player(Index).Char.QuestProgress(i).QuestNum
        If QuestNum > 0 Then
            For QuestNeedsNum = 1 To MAX_QUEST_NEEDS
                ' Check if quest is the right type
                If Quest(QuestNum).QuestNeeds(QuestNeedsNum).QuestType = QuestType Then
                    ' Checking the NUM to the data1
                    If Quest(QuestNum).QuestNeeds(QuestNeedsNum).Data1 = Num Then
                        ' Now update the player's progress. (Gets clamped it to the max needed)
                        If AddValue Then
                            Player(Index).Char.QuestProgress(i).Progress(QuestNeedsNum) = Clamp(Player(Index).Char.QuestProgress(i).Progress(QuestNeedsNum) + Value, 0, Quest(QuestNum).QuestNeeds(QuestNeedsNum).Data2)
                        Else
                            Player(Index).Char.QuestProgress(i).Progress(QuestNeedsNum) = Clamp(Value, 0, Quest(QuestNum).QuestNeeds(QuestNeedsNum).Data2)
                        End If
                        'SendPlayerMsg Index, Player(Index).Char.QuestProgress(i).Progress(QuestNeedsNum), AlertColor
                        SendPlayerQuest Index, i
                    End If
                End If
            Next
        End If
    Next
End Sub

Public Sub OnQuestTurnIn(ByVal Index As Long, ByVal QuestProgressNum As Long, ByVal SelectReward As Long)
Dim i As Long
Dim n As Long
Dim QuestNum As Long
Dim NeededInvSpace As Long

    ' Check to make sure it's a valid QuestProgressNum
    If QuestProgressNum <= 0 Then Exit Sub
    If QuestProgressNum > MAX_PLAYER_QUESTS Then Exit Sub
    If SelectReward < 0 Then Exit Sub
    If SelectReward > MAX_QUEST_REWARDS Then Exit Sub
    
    QuestNum = Player(Index).Char.QuestProgress(QuestProgressNum).QuestNum
    
    ' Make sure there's a quest here
    If QuestNum = 0 Then Exit Sub
    
    ' Check if you're done with the quest
    For i = 1 To MAX_QUEST_NEEDS
        ' Check for any progresses that aren't complete
        ' Make sure it's an actual quest need
        If Quest(QuestNum).QuestNeeds(i).QuestType <> QuestTypes.None Then
            If Player(Index).Char.QuestProgress(QuestProgressNum).Progress(i) < Quest(QuestNum).QuestNeeds(i).Data2 Then
                'SendPlayerMsg Index, Quest(QuestNum).IncompletedMessage, AlertColor
                SendPlayerQuestMsg Index, QuestNum, Quest(QuestNum).IncompletedMessage, False, AlertColor
                Exit Sub
            End If
        End If
    Next
    
    ' First Check if we have enough space for the rewards
      
    ' Check if there are any selection needs if SelectReward = 0
    ' This is needed incase you have to pick a reward
    If SelectReward = 0 Then
        For i = 1 To MAX_QUEST_REWARDS
            If Quest(QuestNum).Rewards(i).ItemNum > 0 Then
                If Quest(QuestNum).Rewards(i).SelectionOnly Then
                    SendPlayerMsg Index, "You must select 1 reward.", AlertColor
                    Exit Sub
                End If
            End If
        Next
    Else
        ' Make sure the selection is valid
        If Quest(QuestNum).Rewards(SelectReward).ItemNum Then
            ' Check if we have enough room for the selection item
            If FindOpenInvSlot(Index, Quest(QuestNum).Rewards(SelectReward).ItemNum) <= 0 Then
                SendPlayerMsg Index, "You do not have enough inventory space to turn this quest in.", BrightRed
                Exit Sub
            End If
        Else
            SendPlayerMsg Index, "Invalid selection", BrightRed
            Exit Sub
        End If
    End If
    
    ' Now we check if we have enough room for non selection items
    For i = 1 To MAX_QUEST_REWARDS
        If Quest(QuestNum).Rewards(i).ItemNum > 0 Then
            If Not Quest(QuestNum).Rewards(i).SelectionOnly Then
                If FindOpenInvSlot(Index, Quest(QuestNum).Rewards(i).ItemNum) <= 0 Then
                    SendPlayerMsg Index, "You do not have enough inventory space to turn this quest in.", BrightRed
                    Exit Sub
                End If
            End If
        End If
    Next

    '
    ' Now that we passed all requirements we can turn it in
    '
    ' Take the ItemCollection items, if we can't take it for whatever reason then exit
    For i = 1 To MAX_QUEST_NEEDS
        If Quest(QuestNum).QuestNeeds(i).QuestType = QuestTypes.ItemCollection Then
            If Not CanTakeItem(Index, Quest(QuestNum).QuestNeeds(i).Data1, Quest(QuestNum).QuestNeeds(i).Data2) Then
                SendPlayerMsg Index, "You don't have the right items to turn in.", BrightRed
                Exit Sub
            End If
        End If
    Next
    
    ' Give the SelectReward items
    If SelectReward Then
        GiveItem Index, Quest(QuestNum).Rewards(SelectReward).ItemNum, Quest(QuestNum).Rewards(SelectReward).ItemValue
    End If
    
    ' Give the rest of the rewards
    For i = 1 To MAX_QUEST_REWARDS
        If Quest(QuestNum).Rewards(i).ItemNum Then
            If Not Quest(QuestNum).Rewards(i).SelectionOnly Then
                GiveItem Index, Quest(QuestNum).Rewards(i).ItemNum, Quest(QuestNum).Rewards(i).ItemValue
            End If
        End If
    Next
    
    ' Give Exp
    Update_Exp Index, Current_Exp(Index) + Quest(QuestNum).RewardExp
    SendActionMsg Current_Map(Index), "+" & Quest(QuestNum).RewardExp & " EXP!", Yellow, ACTIONMSG_SCROLL, Current_X(Index), Current_Y(Index), Index
                    
    ' Update the completed quest flag
    Player(Index).Char.CompletedQuests(QuestNum) = QUEST_STATUS_COMPLETE
    
    ' Clear the quest
    ClearPlayerQuest Index, QuestProgressNum
    
    ' Send the completed message
    'SendPlayerMsg Index, Quest(QuestNum).CompletedMessage, AlertColor
    SendPlayerQuestMsg Index, QuestNum, Quest(QuestNum).CompletedMessage, False, AlertColor
End Sub

Public Sub ClearPlayerQuest(ByVal Index As Long, ByVal QuestProgressNum As Long)
Dim i As Long

    ' Check to make sure it's a valid QuestProgressNum
    If QuestProgressNum <= 0 Then Exit Sub
    If QuestProgressNum > MAX_PLAYER_QUESTS Then Exit Sub
    
    ' Make sure there's a quest here
    If Player(Index).Char.QuestProgress(QuestProgressNum).QuestNum = 0 Then Exit Sub
    
    ' Clear out this quest now
'    Player(Index).Char.QuestProgress(QuestProgressNum).QuestNum = 0
'    For i = 1 To MAX_QUEST_NEEDS
'        Player(Index).Char.QuestProgress(QuestProgressNum).Progress(i) = 0
'    Next
    ZeroMemory ByVal VarPtr(Player(Index).Char.QuestProgress(QuestProgressNum)), LenB(Player(Index).Char.QuestProgress(QuestProgressNum))
    
    ' Update the ActiveQuestCount
    Player(Index).Char.ActiveQuestCount = Player(Index).Char.ActiveQuestCount - 1
    
    ' Send the quest
    SendPlayerQuest Index, QuestProgressNum
End Sub


Public Function IsQuestAvailable(ByVal Index As Long, ByVal QuestNum As Long) As Boolean
Dim i As Long
    
    ' Check if quest is repeatable - If it's not, you can only do it once
    If Not Quest(QuestNum).Repeatable Then
        ' Check if the player already completed this quest
        If Player(Index).Char.CompletedQuests(QuestNum) = QUEST_STATUS_COMPLETE Then
            Exit Function
        End If
    End If
    
'    ' Check if player is on quest
'    For i = 1 To MAX_PLAYER_QUESTS
'        If Player(Index).Char.QuestProgress(i).QuestNum = QuestNum Then
'            Exit Function
'        End If
'    Next
    
    ' Check for quest requirement
    If Quest(QuestNum).QuestReq > 0 Then
        If Player(Index).Char.CompletedQuests(Quest(QuestNum).QuestReq) = QUEST_STATUS_INCOMPLETE Then
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
    
    IsQuestAvailable = True
End Function

Public Function CanTurnInQuest(ByVal Index As Long, ByVal QuestProgressNum As Long) As Boolean
Dim i As Long
Dim QuestNum As Long

    QuestNum = Player(Index).Char.QuestProgress(QuestProgressNum).QuestNum
    
    ' Check if it's a real quest
    If QuestNum = 0 Then Exit Function
    
    ' Check if you're done with the quest
    For i = 1 To MAX_QUEST_NEEDS
        ' Check for any progresses that aren't complete
        If Player(Index).Char.QuestProgress(QuestProgressNum).Progress(i) < Quest(QuestNum).QuestNeeds(i).Data2 Then
            'SendPlayerMsg Index, Quest(QuestNum).IncompletedMessage, AlertColor
            SendPlayerQuestMsg Index, QuestNum, Quest(QuestNum).IncompletedMessage, False, AlertColor
            Exit Function
        End If
    Next
End Function

Public Sub SendAvailableQuests(ByVal Index As Long, ByVal NpcNum As Long, ByVal MapNpcNum As Long)
Dim Buffer As clsBuffer
Dim i As Long
Dim n As Long
Dim QuestNum As Long

    Set Buffer = New clsBuffer
    
    Buffer.WriteLong CMsgAvailableQuests
    Buffer.WriteLong MapNpcNum
    For i = 1 To NpcQuests(NpcNum).QuestCount
        QuestNum = NpcQuests(NpcNum).QuestList(i)
        If Quest(QuestNum).StartNPC = NpcNum Then
            ' Need to send if quest is available
            If IsQuestAvailable(Index, QuestNum) Then
                Buffer.WriteLong QuestNum  ' Quest Number
            End If
        ElseIf Quest(QuestNum).EndNPC = NpcNum Then
            ' Need to send if player is on quest
            For n = 1 To MAX_PLAYER_QUESTS
                If Player(Index).Char.QuestProgress(n).QuestNum = QuestNum Then
                    Buffer.WriteLong QuestNum
                End If
            Next
        End If
    Next
    
    SendDataTo Index, Buffer.ToArray()
End Sub

Public Sub SendPlayerQuests(ByVal Index As Long)
Dim Buffer As clsBuffer
Dim i As Long
Dim n As Long

    Set Buffer = New clsBuffer
    
    Buffer.WriteLong CMsgPlayerQuests
    For i = 1 To MAX_PLAYER_QUESTS
        Buffer.WriteLong Player(Index).Char.QuestProgress(i).QuestNum
        For n = 1 To MAX_QUEST_NEEDS
            Buffer.WriteLong Player(Index).Char.QuestProgress(i).Progress(n)
        Next
    Next
    
    SendDataTo Index, Buffer.ToArray()
End Sub

Public Sub SendPlayerQuest(ByVal Index As Long, ByVal QuestProgressNum As Long)
Dim Buffer As clsBuffer
Dim i As Long

    Set Buffer = New clsBuffer
    
    Buffer.WriteLong CMsgPlayerQuest
    Buffer.WriteLong QuestProgressNum
    Buffer.WriteLong Player(Index).Char.QuestProgress(QuestProgressNum).QuestNum
    For i = 1 To MAX_QUEST_NEEDS
        Buffer.WriteLong Player(Index).Char.QuestProgress(QuestProgressNum).Progress(i)
    Next
    
    SendDataTo Index, Buffer.ToArray()
End Sub

Public Function OnQuest(ByVal Index As Long, ByVal QuestNum As Long) As Long
Dim i As Long
    For i = 1 To MAX_PLAYER_QUESTS
        If Player(Index).Char.QuestProgress(i).QuestNum = QuestNum Then
            OnQuest = i
            Exit Function
        End If
    Next
End Function

Public Sub CheckPlayerQuests(ByVal Index As Long)
Dim i As Long
    ' Loop through the quests and see if the quest was deleted
    For i = 1 To MAX_PLAYER_QUESTS
        If Player(Index).Char.QuestProgress(i).QuestNum Then
            ' Check if the quest exists
            If LenB(Quest(Player(Index).Char.QuestProgress(i).QuestNum).Name) = 0 Then
                ZeroMemory ByVal VarPtr(Player(Index).Char.QuestProgress(i)), LenB(Player(Index).Char.QuestProgress(i))
            End If
        End If
    Next
End Sub

Public Sub SendPlayerQuestMsg(ByVal Index As Long, ByVal QuestNum As Long, ByVal QuestMsg As String, ByVal Start As Boolean, ByVal Color As Byte)
Dim Buffer As clsBuffer
    
    ' Check if the message is null
    QuestMsg = Trim$(QuestMsg)
    If LenB(QuestMsg) = 0 Then Exit Sub
    
    ' Start = StartNpc
    If Start Then
        ' Make sure there's a NPC
        If Quest(QuestNum).StartNPC Then
            QuestMsg = Trim$(Npc(Quest(QuestNum).StartNPC).Name) & ": " & QuestMsg
        End If
    Else
        ' Make sure there's a NPC
        If Quest(QuestNum).StartNPC Then
            QuestMsg = Trim$(Npc(Quest(QuestNum).EndNPC).Name) & ": " & QuestMsg
        End If
    End If
    
    QuestMsg = Replace$(QuestMsg, "%PLAYERNAME%", Current_Name(Index))
    QuestMsg = Replace$(QuestMsg, "%PLAYERCLASS%", GetClassName(Current_Class(Index)))
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate Len(QuestMsg) + 7
    Buffer.WriteLong CMsgChatMsg
    Buffer.WriteString QuestMsg
    Buffer.WriteByte Color
    
    SendDataTo Index, Buffer.ToArray()
End Sub


Public Sub Update_Npc_Quests()
Dim i As Long
    
    For i = 1 To MAX_NPCS
        Update_Npcs_Quest i
    Next
End Sub

Public Sub Update_Npcs_Quest(ByVal NpcNum As Long)
Dim i As Long

    ' Make sure it's a valid NpcNum
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

Public Sub OnQuestDrop(ByVal Index As Long, ByVal QuestProgressNum As Long)
Dim QuestNum As Long

    ' Make sure it's a valid quest
    QuestNum = Player(Index).Char.QuestProgress(QuestProgressNum).QuestNum
    If QuestNum Then
        ' Let's clear it
        ZeroMemory ByVal VarPtr(Player(Index).Char.QuestProgress(QuestProgressNum)), LenB(Player(Index).Char.QuestProgress(QuestProgressNum))
        ' Send it
        SendPlayerQuest Index, QuestProgressNum
    End If
End Sub
