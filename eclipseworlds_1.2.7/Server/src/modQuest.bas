Attribute VB_Name = "modQuest"
Option Explicit

Public Const QUESTICON_LENGTH = 5

Public Const QUESTNAME_LENGTH = 40

Public Const QUESTDESC_LENGTH = 300

Public Quest()                         As QuestRec

Public PlayerQuest(1 To MAX_PLAYERS)   As PlayerQuestRec

'Constants to use for tasks
Public Const TASK_KILL                 As Byte = 1

Public Const TASK_GATHER               As Byte = 2

Public Const TASK_VARIABLE             As Byte = 3

Public Const TASK_GETSKILL             As Byte = 4

Public Const ACTION_GIVE_ITEM          As Byte = 5

Public Const ACTION_TAKE_ITEM          As Byte = 6

Public Const ACTION_SHOWMSG            As Byte = 7

Public Const ACTION_ADJUST_LVL         As Byte = 8

Public Const ACTION_ADJUST_EXP         As Byte = 9

Public Const ACTION_WARP               As Byte = 10

Public Const ACTION_ADJUST_STAT_LVL    As Byte = 11

Public Const ACTION_ADJUST_SKILL_LVL   As Byte = 12

Public Const ACTION_ADJUST_SKILL_EXP   As Byte = 13

Public Const ACTION_ADJUST_STAT_POINTS As Byte = 14

Public Const ACTION_SETVARIABLE        As Byte = 15

Public Const ACTION_PLAYSOUND          As Byte = 16

Public Enum Questing

    GivingItem = 1
    TakingItem
    QuestKill

End Enum

Private Type QuestAmountRec

    ID() As Integer

End Type

Public Type PlayerQuestRec

    QuestCompleted() As Byte
    QuestCLI() As Integer
    QuestTask() As Integer
    QuestAmount() As QuestAmountRec

End Type

Private Type RequirementsRec

    AccessReq As Long
    LevelReq As Long
    GenderReq As Long
    ClassReq As Long
    SkillReq As Long
    SkillLevelReq As Long
    Stat_Req(1 To Stats.Stat_count - 1) As Long

End Type

Private Type ActionRec

    TextHolder As String * QUESTDESC_LENGTH
    ActionID As Byte
    Amount As Long
    MainData As Long
    SecondaryData As Long
    TertiaryData As Long
    QuadData As Long

End Type

Private Type CLIRec

    ItemIndex As Long
    isNPC As Long
    Max_Actions As Long
    Action() As ActionRec

End Type

Private Type QuestRec

    Name As String * QUESTNAME_LENGTH
    Description As String * QUESTDESC_LENGTH
    CanBeRetaken As Byte
    Rank As String * QUESTICON_LENGTH
    OutOfOrder As Byte

    'Maxes
    Max_CLI As Long

    'Main data
    CLI() As CLIRec
    Requirements As RequirementsRec

End Type

Private Type TempQuestRec

    CLI() As CLIRec

End Type

Public Type FindQuestRec

    QuestIndex As Long
    CLIIndex As Long
    ActionIndex As Long

End Type


'/////////////////////////////////////////////////////////
'/////////////////QUEST SUBS AND FUNCTIONS////////////////
'/////////////////////////////////////////////////////////

Function GetPlayerQuestCLI(ByVal index As Long, ByVal QuestID As Long)
    If index < 1 Or index > Player_HighIndex Then Exit Function
    ReDim Preserve Account(index).Chars(GetPlayerChar(index)).QuestCLI(MAX_QUESTS)
    GetPlayerQuestCLI = Account(index).Chars(GetPlayerChar(index)).QuestCLI(QuestID)
End Function

Function GetPlayerQuestTask(ByVal index As Long, ByVal QuestID As Long)
    ReDim Preserve Account(index).Chars(GetPlayerChar(index)).QuestTask(MAX_QUESTS)
    GetPlayerQuestTask = Account(index).Chars(GetPlayerChar(index)).QuestTask(QuestID)
End Function

Function GetPlayerQuestAmount(ByVal index As Long, ByVal QuestID As Long, ByVal NPCNum As Long)
    If index < 1 Or index > MAX_PLAYERS Then Exit Function

    Exit Function
    ReDim Preserve Account(index).Chars(GetPlayerChar(index)).QuestAmount(MAX_QUESTS).ID(MAX_NPCS)
    GetPlayerQuestAmount = Account(index).Chars(GetPlayerChar(index)).QuestAmount(QuestID).ID(NPCNum)
End Function

Sub SetPlayerQuestCLI(ByVal index As Long, ByVal QuestID As Long, Value As Long)
    ReDim Preserve Account(index).Chars(GetPlayerChar(index)).QuestCLI(MAX_QUESTS)
    Account(index).Chars(GetPlayerChar(index)).QuestCLI(QuestID) = Value
End Sub

Sub SetPlayerQuestTask(ByVal index As Long, ByVal QuestID As Long, Value As Long)
    ReDim Preserve Account(index).Chars(GetPlayerChar(index)).QuestTask(MAX_QUESTS)
    Account(index).Chars(GetPlayerChar(index)).QuestTask(QuestID) = Value
End Sub
    
Sub SetPlayerQuestAmount(ByVal index As Long, _
                         ByVal QuestID As Long, _
                         Value As Long, _
                         Optional ByVal NPCNum As Long = 0, _
                         Optional ByVal PlusVal As Boolean = False)

    Dim i As Long

    If index < 1 Or index > MAX_PLAYERS Then Exit Sub
    If QuestID < 1 Or QuestID > MAX_QUESTS Then Exit Sub
    
    ReDim Preserve Account(index).Chars(GetPlayerChar(index)).QuestAmount(MAX_QUESTS)
    ReDim Preserve Account(index).Chars(GetPlayerChar(index)).QuestAmount(MAX_NPCS)

    If PlusVal Then
        Account(index).Chars(GetPlayerChar(index)).QuestAmount(QuestID).ID(NPCNum) = Account(index).Chars(GetPlayerChar(index)).QuestAmount(QuestID).ID(NPCNum) + Value
    Else

        If Not NPCNum > 0 Then

            For i = 1 To MAX_NPCS
                Account(index).Chars(GetPlayerChar(index)).QuestAmount(QuestID).ID(i) = Value
            Next i

        Else
            Account(index).Chars(GetPlayerChar(index)).QuestAmount(QuestID).ID(NPCNum) = Value
        End If
    End If

End Sub

Function GetPlayerQuestCompleted(ByVal index As Long, ByVal QuestID As Long) As Boolean

    If index < 1 Or index > MAX_PLAYERS Then Exit Function

    ReDim Preserve Account(index).Chars(GetPlayerChar(index)).QuestCompleted(MAX_QUESTS)
    GetPlayerQuestCompleted = Account(index).Chars(GetPlayerChar(index)).QuestCompleted(QuestID)
End Function

Sub SetPlayerQuestCompleted(ByVal index As Long, ByVal QuestID As Long, Value As Byte)

    If index < 1 Or index > MAX_PLAYERS Then Exit Sub

    ReDim Preserve Account(index).Chars(GetPlayerChar(index)).QuestCompleted(MAX_QUESTS)
    Account(index).Chars(GetPlayerChar(index)).QuestCompleted(QuestID) = Value
End Sub

Sub SetPlayerTask(ByVal index As Long, ByVal QuestID As Long, Value As Long)

    If index < 1 Or index > MAX_PLAYERS Then Exit Sub

    Account(index).Chars(GetPlayerChar(index)).QuestTask(QuestID) = Value
End Sub

Function GetPlayerTotalQuestAmount(ByVal index As Long, ByVal QuestID As Long)

    Dim i As Long

    If index < 1 Or index > MAX_PLAYERS Then Exit Function

    For i = 1 To MAX_NPCS
        GetPlayerTotalQuestAmount = GetPlayerTotalQuestAmount + Account(index).Chars(GetPlayerChar(index)).QuestAmount(QuestID).ID(i)
    Next i

End Function

Public Function IsQuestCLI(ByVal index As Long, ByVal NPCIndex As Long) As FindQuestRec

    Dim i    As Long, ii As Long, III As Long

    Dim temp As FindQuestRec

    If index < 1 Or index > MAX_PLAYERS Then Exit Function

    'Dynamically find the correct quest item
    For i = 1 To MAX_QUESTS

        With Quest(i)

            For ii = 1 To .Max_CLI

                'See if this npc is within a started quest first.
                If .CLI(ii).ItemIndex = NPCIndex Then    'found a matching quest cli item, this npc is part of a quest
                    If IsInQuest(index, i) Then
                        temp.QuestIndex = i
                        temp.CLIIndex = ii
                        IsQuestCLI = temp

                        Exit Function

                    End If
                End If

            Next ii

        End With

    Next i

    For i = 1 To MAX_QUESTS

        With Quest(i)

            For ii = 1 To .Max_CLI

                'It's not within a started quest, so see if it's a start to a new quest
                If .CLI(ii).ItemIndex = NPCIndex Then    'found a matching quest cli item, this npc is part of a quest
                    If ii = 1 Then
                        temp.QuestIndex = i
                        temp.CLIIndex = ii
                        IsQuestCLI = temp

                        Exit Function

                    End If
                End If

            Next ii

        End With

    Next i

End Function

Public Sub CheckQuest(ByVal index As Long, _
                      QuestIndex As Long, _
                      CLIIndex As Long, _
                      TaskIndex As Long)

    Dim i As Long, ii As Long
    
    Exit Sub

    'Is the PlayerQuest on this quest?  If not, cancel out.
    If IsInQuest(index, QuestIndex) Then

        'Is the PlayerQuest currently on this Chronological list item?
        If GetPlayerQuestCLI(index, QuestIndex) = CLIIndex Then
            Call HandleQuestTask(index, QuestIndex, CLIIndex, GetPlayerQuestTask(index, QuestIndex))
        Else

            'Dynamically show message from last known cli
            If GetPlayerQuestCLI(index, QuestIndex) - 1 > 0 Then

                For i = Quest(QuestIndex).CLI(GetPlayerQuestCLI(index, QuestIndex) - 1).Max_Actions To 1 Step -1

                    With Quest(QuestIndex).CLI(GetPlayerQuestCLI(index, QuestIndex) - 1).Action(i)

                        'quit early if we run into a task.  Means we don't have a msg to display
                        If .ActionID > 0 And .ActionID < 4 Then Exit For

                        If .ActionID = ACTION_SHOWMSG Then
                            Call PlayerMsg(index, Trim$(.TextHolder), .TertiaryData, True, QuestIndex, Trim$(NPC(Quest(QuestIndex).CLI(GetPlayerQuestCLI(index, QuestIndex)).ItemIndex).Name))

                            Exit For

                        End If

                    End With

                Next i

            End If
        End If

    Else

        'lets start this quest if the CLI is the first greeter
        If CLIIndex = 1 Then

            ' see if the Player has taken it all ready and if it can be retaken
            If IsQuestCompleted(index, QuestIndex) Then
                If Quest(QuestIndex).CanBeRetaken = False Then

                    'See if we have a retort message for a quest that cannot be retaken
                    For i = 1 To Quest(QuestIndex).Max_CLI
                        For ii = 1 To Quest(QuestIndex).CLI(i).Max_Actions

                            If Quest(QuestIndex).CLI(i).Action(ii).ActionID = ACTION_SHOWMSG Then
                                If Quest(QuestIndex).CLI(i).Action(ii).QuadData = vbChecked Then
                                    Call PlayerMsg(index, Trim$(Quest(QuestIndex).CLI(i).Action(ii).TextHolder), Quest(QuestIndex).CLI(i).Action(ii).TertiaryData, True, QuestIndex, Trim$(NPC(Quest(QuestIndex).CLI(CLIIndex).ItemIndex).Name))

                                    Exit Sub

                                End If
                            End If

                        Next ii
                    Next i
                    
                    Exit Sub

                End If
            End If

            'not in a quest, check the requirements
            With Quest(QuestIndex).Requirements

                'check level
                If .LevelReq > 0 Then
                    If Not GetPlayerLevel(index) >= .LevelReq Then
                        Call PlayerMsg(index, "Your level does not meet the requirements to start this mission.", BrightRed, True, QuestIndex, Trim$(NPC(Quest(QuestIndex).CLI(CLIIndex).ItemIndex).Name))

                        Exit Sub

                    End If
                End If

                'check class
                If .ClassReq > 0 Then
                    If Not GetPlayerClass(index) = .ClassReq Then
                        Call PlayerMsg(index, "Your class does not meet the requirements to start this mission.", BrightRed, True, QuestIndex, Trim$(NPC(Quest(QuestIndex).CLI(CLIIndex).ItemIndex).Name))

                        Exit Sub

                    End If
                End If

                'check gender - DOES NOT APPEAR TO EXIST IN NIN
                If .GenderReq > 0 Then
                'TODO
                 '   If Not GetPlayerQuestGender(Index) = .GenderReq Then
                 '       Call PlayerMsg(Index, "Your gender does not meet the requirements to start this quest.", BrightRed, True, QuestIndex)
                 '       Exit Sub
                 '   End If
                End If
                'check access
                If Not GetPlayerAccess(index) >= .AccessReq Then
                    Call PlayerMsg(index, "Your administrative access level does not meet the requirements to start this mission.", BrightRed, True, QuestIndex, Trim$(NPC(Quest(QuestIndex).CLI(CLIIndex).ItemIndex).Name))

                    Exit Sub

                End If

                'check skill - DOES NOT APPEAR TO EXIST IN NIN
                'If .SkillReq > 0 Then
                '    If Not GetPlayerQuestSkill(index, .SkillReq) >= .SkillLevelReq Then
                '        Call PlayerMsg(index, "Your " & GetSkillName(.SkillLevelReq) & " level does not meet the requirements to start this quest.", BrightRed, True, QuestIndex)
                '        Exit Sub
                '    End If
                'End If
                'check stats
                For i = 1 To Stats.Stat_count - 1

                    If Not GetPlayerStat(index, i) >= .Stat_Req(i) Then
                        Call PlayerMsg(index, "Your stats do not meet the requirements to start this mission.", BrightRed, True, QuestIndex, Trim$(NPC(Quest(QuestIndex).CLI(CLIIndex).ItemIndex).Name))

                        Exit Sub

                    End If

                Next i

            End With

            'send the request to the PlayerQuest
            Call SendPlayerQuestRequest(index, QuestIndex)
        End If
    End If

End Sub

Public Sub HandleQuestTask(ByVal index As Long, _
                           ByVal QuestID As Long, _
                           ByVal CLIID As Long, _
                           ByVal TaskID As Long, _
                           Optional ByVal ShowRebuttal As Boolean = True)

    Dim i      As Long, GaveItem As Boolean

    Dim NPCNum As Long

    'Manage the current task the PlayerQuest is on and move PlayerQuest forward through the quest.
    If QuestID < 1 Or QuestID > MAX_QUESTS Then Exit Sub
    If CLIID < 1 Or CLIID > Quest(QuestID).Max_CLI Then Exit Sub
    If index < 1 Or index > MAX_PLAYERS Then Exit Sub
    
    'If Not ValidArray1(Quest(QuestID).CLI) Then Exit Sub
    'If Not ValidArray2(Quest(QuestID).CLI(II).Action) Then Exit Sub
    
    TaskID = TaskID - 1
    
    With Quest(QuestID).CLI(CLIID)
        NPCNum = .ItemIndex

        ' Figure out what we need to do.
        If Quest(QuestID).Max_CLI > 0 Then
            If .Max_Actions > 0 And .Max_Actions >= TaskID Then
                
                Select Case .Action(TaskID).ActionID
    
                    Case TASK_GATHER
    
                        'check if the PlayerQuest gathered enough of the item
                        If HasItem(index, .Action(TaskID).MainData) >= .Action(TaskID).Amount Then
                            'PlayerQuest has the right amount.  move forward.
                            Call PlayerMsg(index, "Mission Task Completed: Gather " & .Action(TaskID).Amount & " " & Trim$(Item(.Action(TaskID).MainData).Name) & "('s)", BrightGreen)
                            Call CheckNextTask(index, QuestID, CLIID, TaskID)
    
                            If .Action(TaskID).SecondaryData = vbChecked Then    'take the item
                                Call TakeInvItem(index, .Action(TaskID).MainData, .Action(TaskID).Amount, True)
                            End If
    
                        Else
    
                            'we don't have the required amount, see if we need to say a rebuttal msg
                            If ShowRebuttal Then Call CheckRebuttal(index, QuestID, CLIID, TaskID)
                        End If
    
                        Exit Sub
    
                    Case TASK_KILL
                
                        If GetPlayerQuestAmount(index, QuestID, .Action(TaskID).MainData) >= .Action(TaskID).Amount Then
                            'PlayerQuest has the right amount.  move forward.
                            Call CheckNextTask(index, QuestID, CLIID, TaskID)
                        Else
    
                            'we don't have the required amount, see if we need to say a rebuttal msg
                            If ShowRebuttal Then Call CheckRebuttal(index, QuestID, CLIID, TaskID)
                        End If
    
                    Case TASK_GETSKILL
                        'check if the PlayerQuest gained the right skill amount
                        'If getPlayerQuestskill(index, .Action(TaskID).MainData) >= .Action(TaskID).Amount Then
                        'Call SetPlayerQuestAmount(index, QuestID, 0)
                        Call CheckNextTask(index, QuestID, CLIID, TaskID)
    
                        'Else
                        'we don't have the required amount, see if we need to say a rebuttal msg
                        'If ShowRebuttal Then Call CheckRebuttal(index, QuestID, CLIID, TaskID)
                        'End If
                    Case TASK_VARIABLE
    
                        If CBool(.Action(TaskID).MainData) = True Then 'Variable, not switch
                            If Account(index).Chars(GetPlayerChar(index)).Variables(.Action(TaskID).SecondaryData) >= .Action(TaskID).Amount Then
                                Call CheckNextTask(index, QuestID, CLIID, TaskID)
                            Else
    
                                If ShowRebuttal Then Call CheckRebuttal(index, QuestID, CLIID, TaskID)
                            End If
    
                        Else 'now its a switch
    
                            If Account(index).Chars(GetPlayerChar(index)).Switches(.Action(TaskID).SecondaryData) = .Action(TaskID).Amount Then
                                Call CheckNextTask(index, QuestID, CLIID, TaskID)
                            Else
    
                                If ShowRebuttal Then Call CheckRebuttal(index, QuestID, CLIID, TaskID)
                            End If
                        End If
                
                    Case ACTION_SETVARIABLE
    
                        If CBool(.Action(TaskID).MainData) = True Then 'Variable, not switch
                            If .Action(TaskID).TertiaryData = vbChecked Then 'Set the value instead of adding
                               Account(index).Chars(GetPlayerChar(index)).Variables(.Action(TaskID).SecondaryData) = .Action(TaskID).Amount
                            Else
                               Account(index).Chars(GetPlayerChar(index)).Variables(.Action(TaskID).SecondaryData) = Account(index).Chars(GetPlayerChar(index)).Variables(.Action(TaskID).SecondaryData) + .Action(TaskID).Amount
                            End If
    
                            Call SendPlayerData(index)
                        Else 'now its a switch
                            Account(index).Chars(GetPlayerChar(index)).Switches(.Action(TaskID).SecondaryData) = .Action(TaskID).Amount
                            Call SendPlayerData(index)
                        End If
    
                    Case ACTION_GIVE_ITEM
    
                        'give the PlayerQuest so many of a certain item
                        If Item(.Action(TaskID).MainData).stackable > 0 Then
                            If .Action(TaskID).MainData > 1 Then
                                GaveItem = GiveInvItem(index, .Action(TaskID).MainData, .Action(TaskID).Amount, True)
                            Else
                                GaveItem = True
                                Call GiveInvItem(index, 1, .Action(TaskID).Amount)
                            End If
    
                        Else
    
                            For i = 1 To .Action(TaskID).Amount
                                GaveItem = GiveInvItem(index, .Action(TaskID).MainData, 1, True)
                            Next i
    
                        End If
    
                        If Not GaveItem Then
                            Call PlayerMsg(index, "Not enough space in your inventory.  Please come back when you can hold everything I have to give you.", BrightRed, True, QuestID, Trim$(NPC(.ItemIndex).Name))
    
                            Exit Sub
    
                        End If
    
                    Case ACTION_TAKE_ITEM
                        'take the PlayerQuest's item
                        Call TakeInvItem(index, .Action(TaskID).MainData, .Action(TaskID).Amount, True)
    
                    Case ACTION_SHOWMSG

                        'show the Player a message
                        If .ItemIndex = 0 Then Exit Sub
                        Call PlayerMsg(index, ModifyTxt(index, QuestID, Trim$(.Action(TaskID).TextHolder)), .Action(TaskID).TertiaryData, True, QuestID, Trim$(NPC(.ItemIndex).Name))
    
                    Case ACTION_ADJUST_LVL
                        Call SetPlayerLevel(index, .Action(TaskID).Amount, .Action(TaskID).MainData)
                        Call SendPlayerLevel(index)
    
                    Case ACTION_ADJUST_EXP
                        Call SetPlayerExp(index, .Action(TaskID).Amount, .Action(TaskID).MainData)
                        Call SendPlayerExp(index)
    
                    Case ACTION_ADJUST_STAT_LVL
                        Call SetPlayerStat(index, .Action(TaskID).SecondaryData, .Action(TaskID).Amount, .Action(TaskID).MainData)
                        Call SendPlayerStats(index)
    
                    Case ACTION_ADJUST_SKILL_LVL
                        Call SetPlayerSkill(index, .Action(TaskID).Amount, .Action(TaskID).SecondaryData, .Action(TaskID).MainData)
                        Call SendPlayerSkills(index)
    
                    Case ACTION_ADJUST_SKILL_LVL
                        Call SetPlayerSkill(index, .Action(TaskID).Amount, .Action(TaskID).SecondaryData, .Action(TaskID).MainData)
                        Call SendPlayerSkills(index)
    
                    Case ACTION_ADJUST_SKILL_EXP
                        Call SetPlayerSkillExp(index, .Action(TaskID).Amount, .Action(TaskID).SecondaryData, .Action(TaskID).MainData)
                        Call SendPlayerSkills(index)
    
                    Case ACTION_WARP
                        Call PlayerWarp(index, .Action(TaskID).Amount, .Action(TaskID).MainData, .Action(TaskID).SecondaryData, , DIR_DOWN)
                    
                    Case ACTION_PLAYSOUND
                        Call SendQuestSound(index, .Action(TaskID).MainData, .Action(TaskID).SecondaryData)
    
                    Case Else
                        'continue on in case we missed something.  This will make it harder to find bugs, but will run smoother for the user
                End Select
            
                'Continue if we processed an action.
                If .Action(TaskID).ActionID > 4 Then
                    Call CheckNextTask(index, QuestID, CLIID, TaskID)
    
                    For i = 1 To Quest(QuestID).Max_CLI
                        Call SendShowTaskCompleteOnNPC(index, Quest(QuestID).CLI(i).ItemIndex, False)
                    Next i
    
                    Call SendPlayerQuest(index)
                End If
            
            End If
        End If

    End With

End Sub

Public Sub CheckNextTask(ByVal index As Long, _
                         QuestID As Long, _
                         CLIID As Long, _
                         TaskID As Long)

    If index < 1 Or index > MAX_PLAYERS Then Exit Sub

    With Quest(QuestID).CLI(CLIID)

        ' move on to next task if there is one
        If TaskID > .Max_Actions Then GoTo NextCLI

        'check if the next task is a rebuttal, if so, skip it
        If .Action(TaskID + 1).ActionID = ACTION_SHOWMSG Then
            If .Action(TaskID + 1).SecondaryData = vbChecked Then
                If Not TaskID + 2 > .Max_Actions Then
                    'skip this rebuttal task
                    Call SetPlayerTask(index, QuestID, TaskID + 2)
                Else
                    GoTo NextCLI
                End If

            Else
                Call SetPlayerTask(index, QuestID, TaskID + 1)
            End If

        Else
            Call SetPlayerTask(index, QuestID, TaskID + 1)
        End If

        Call SendPlayerQuest(index)
        Call HandleQuestTask(index, QuestID, GetPlayerQuestCLI(index, QuestID), GetPlayerQuestTask(index, QuestID), False)

        Exit Sub

NextCLI:

        'move on to next cli item if there is one
        If Not CLIID + 1 > Quest(QuestID).Max_CLI Then
            Call SetPlayerQuestCLI(index, QuestID, CLIID + 1)
            Call SetPlayerTask(index, QuestID, 1)
            Call SendPlayerQuest(index)
            'We don't want to move straight for the next task here.  The Player has to talk to them to start it.
        Else
            'quest completed
            Call MarkQuestCompleted(index, QuestID)
            Call SetPlayerQuestCLI(index, QuestID, 0)
            Call SetPlayerTask(index, QuestID, 0)
            Call SetPlayerQuestAmount(index, QuestID, 0)
            Call SendPlayerQuest(index)
        End If

    End With

End Sub

Public Sub CheckRebuttal(ByVal index As Long, _
                         QuestID As Long, _
                         CLIID As Long, _
                         TaskID As Long)

    Dim i As Long

    With Quest(QuestID).CLI(CLIID)

        For i = TaskID To .Max_Actions

            If .Action(i).ActionID = ACTION_SHOWMSG Then
                If .Action(i).SecondaryData = vbChecked Then
                    'send the msg
                    Call PlayerMsg(index, ModifyTxt(index, QuestID, Trim$(.Action(i).TextHolder)), .Action(i).TertiaryData, True, QuestID, Trim$(NPC(.ItemIndex).Name))

                    Exit Sub

                End If
            End If

        Next i

    End With

End Sub

Public Function ModifyTxt(ByVal index As Integer, _
                          ByVal QuestID As Long, _
                          ByVal Msg As String) As String

    Dim nMsg As String

    Dim i    As Long, ii As Long, ID As Long

    nMsg = Replace$(Msg, "<kills>", GetPlayerTotalQuestAmount(index, QuestID))    'replace with PlayerQuest kill amount
    ModifyTxt = nMsg
    
    i = GetPlayerQuestCLI(index, QuestID)
    ii = GetPlayerQuestTask(index, QuestID)
    ID = Quest(QuestID).CLI(i).Action(ii).SecondaryData

    If ID > 0 Then
        If Quest(QuestID).CLI(i).Action(ii).ActionID = TASK_VARIABLE Then

            'working with variable
            If CBool(Quest(QuestID).CLI(i).Action(ii).MainData) = True Then
                nMsg = Replace$(ModifyTxt, "<amount>", Account(index).Chars(GetPlayerChar(index)).Variables(ID))
            Else

                'working with switch
                If CBool(Account(index).Chars(GetPlayerChar(index)).Switches(ID)) = True Then
                    nMsg = Replace$(ModifyTxt, "<torf>", "True")
                Else
                    nMsg = Replace$(ModifyTxt, "<torf>", "False")
                End If
            End If
        End If
    End If

    ModifyTxt = nMsg
End Function

Public Function IsQuestCompleted(ByVal index As Long, ByVal QuestID As Long) As Boolean
Dim i As Long
    IsQuestCompleted = False
    If Not QuestID > 0 Then Exit Function
    
    ReDim Account(index).Chars(GetPlayerChar(index)).QuestCompleted(MAX_QUESTS)
    If GetPlayerQuestCompleted(index, QuestID) = True Then
        IsQuestCompleted = True
    End If
End Function

Public Sub MarkQuestCompleted(ByVal index As Long, ByVal QuestID As Long)

    Dim i As Long

    If Not QuestID > 0 Then Exit Sub

    Call SetPlayerQuestCompleted(index, QuestID, 1)
End Sub

Private Function IsInQuest(ByVal index As Long, ByVal QuestID As Long) As Boolean

    If Not QuestID > 0 Then Exit Function

    If GetPlayerQuestCLI(index, QuestID) > 0 Then IsInQuest = True
End Function

Public Sub SendPlayerQuestRequest(ByVal index As Long, ByVal QuestID As Long)

    Dim buffer As clsBuffer

    If index < 1 Or index > Player_HighIndex Then Exit Sub
    If QuestID < 1 Or QuestID > MAX_QUESTS Then Exit Sub
    
    'Call QuitQuest(Index, QuestID, False)

    Set buffer = New clsBuffer
    buffer.WriteLong SQuestRequest
    buffer.WriteLong QuestID
    Call SendDataTo(index, buffer.ToArray())
    Set buffer = Nothing
End Sub

Public Function HasQuestItems(ByVal index As Long, _
                              QuestID As Long, _
                              Optional ByVal ReturnIfNot As Boolean = False) As String

    Dim i As Long, CLIIndex As Long, TaskIndex As Long

    CLIIndex = GetPlayerQuestCLI(index, QuestID)
    TaskIndex = GetPlayerQuestTask(index, QuestID)

    HasQuestItems = 0

    If CLIIndex > 0 Then
        If TaskIndex > 0 Then

            For i = TaskIndex To 1 Step -1

                If Quest(QuestID).Max_CLI >= CLIIndex Then
                    If Quest(QuestID).CLI(CLIIndex).Max_Actions >= i Then
                        If Quest(QuestID).CLI(CLIIndex).Action(i).ActionID = TASK_GATHER Then
                            If HasItem(index, Quest(QuestID).CLI(CLIIndex).Action(i).MainData) >= Quest(QuestID).CLI(CLIIndex).Action(i).Amount Then
                                HasQuestItems = Quest(QuestID).CLI(CLIIndex).ItemIndex    'return the npc number
        
                                Exit Function
        
                            Else
        
                                If ReturnIfNot Then
                                    'return a value meant to be parsed so we can distinguish the returned value
                                    HasQuestItems = Quest(QuestID).CLI(CLIIndex).ItemIndex & "|" & "Why leave it empty... lol"
        
                                    Exit Function
        
                                End If
                            End If
                        End If
                    End If
                End If

            Next i

        End If
    End If

End Function

Public Sub QuestUpdate(ByVal PlayerID As Long, _
                       ByVal Todo As Questing, _
                       Optional ByVal Data1 As Long = 0, _
                       Optional ByVal Data2 As Long = 0)

    Dim i       As Long, ii As Long, III As Long, index As Long

    Dim Parse() As String

    Dim NPCNum  As Long

    Dim Kills   As Long, Needed As Long

    Dim ResetIt As Boolean

    Select Case Todo

        Case Questing.TakingItem

            'check quests
            For i = 1 To MAX_QUESTS
                Parse() = Split(HasQuestItems(PlayerID, i, True), "|")

                If UBound(Parse()) > 0 Then
                    NPCNum = Parse(0)

                    If NPCNum > 0 Then
                        Call SendShowTaskCompleteOnNPC(PlayerID, NPCNum, False)
                    End If
                End If

            Next i

        Case Questing.GivingItem

            'check quests
            For i = 1 To MAX_QUESTS
                NPCNum = HasQuestItems(PlayerID, i)

                If NPCNum > 0 Then
                    ii = GetPlayerQuestCLI(PlayerID, i)
                    III = GetPlayerQuestTask(PlayerID, i)
                
                    If Quest(i).CLI(ii).Action(III).ActionID = TASK_GATHER Then
                        If Quest(i).CLI(ii).Action(III).MainData = Data1 Then
                            'found a quest, let's see if we move on from it.
                            Data2 = Quest(i).CLI(ii).Action(III).TertiaryData
                        End If
                    End If
                
                    If Data2 = vbChecked Then
                        Call PlayerMsg(PlayerID, "Mission Task Completed!  You gathered the required items.", BrightGreen, True, i, Trim$(NPC(Quest(i).CLI(ii).ItemIndex).Name))
                        Call HandleQuestTask(PlayerID, i, ii, III, False)
                    Else
                        Call SendShowTaskCompleteOnNPC(PlayerID, NPCNum, True)
                    End If
                End If

            Next i

        Case Questing.QuestKill

            'Cycle through all the quests the PlayerQuest could be in
            For i = 1 To MAX_QUESTS
                index = PlayerID
                NPCNum = Data1
                ii = GetPlayerQuestCLI(index, i)
                III = GetPlayerQuestTask(index, i)
                
                If ii < 1 Then GoTo NextLoop
                If III < 1 Then GoTo NextLoop
                If Not Quest(i).Max_CLI > 0 Then Exit Sub
                If Not Quest(i).CLI(ii).Max_Actions > 0 Then Exit Sub
                
                If ii > 0 Then
                    If III > 0 Then
                        
                        'If out-of-order is selected, add a kill count for any npc the player attacks.
                        If CBool(Quest(i).OutOfOrder) = True Then
                            Call SetPlayerQuestAmount(index, i, 1, NPCNum, True)
                        End If
                            
                        'Make sure the PlayerQuest's current task for this quest is to kill enemies
                        If Quest(i).Max_CLI > 0 Then
                            ReDim Preserve Quest(i).CLI(1 To Quest(i).Max_CLI)
                            
                            If Quest(i).CLI(ii).Action(III).ActionID = TASK_KILL Then
    
                                'Make sure this is the NPC we're supposed to kill for this quest
                                If Quest(i).CLI(ii).Action(III).MainData = NPCNum Then
                                
                                    If Quest(i).CLI(ii).Action(III).QuadData <> 0 Then

                                        'reset the kill count for the selected NPC('s) | only once
                                        If ResetIt Then Call SetPlayerQuestAmount(index, i, 0, Quest(i).CLI(ii).Action(III).QuadData)
                                        ResetIt = True
                                    End If
                            
                                    If CBool(Quest(i).OutOfOrder) = False Then
                                        Call SetPlayerQuestAmount(index, i, 1, NPCNum, True)
                                    End If
                                
                                    Kills = GetPlayerQuestAmount(index, i, NPCNum)
                                    Needed = Quest(i).CLI(ii).Action(III).Amount
                                    
                                    'check if the player killed enough
                                    If Not Kills >= Needed Then
                                        Call PlayerMsg(index, "Mission Kills: " & Kills & " / " & Needed, White)
                                        'Call SendActionMsg(GetPlayerMap(index), Kills & "/" & Needed & " kills", Green, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32))
                                    Else
                                        ResetIt = False

                                        If Quest(i).CLI(ii).Action(III).TertiaryData = False Then
                                            Call PlayerMsg(index, "Mission Task Completed!  Kills: " & GetPlayerQuestAmount(index, i, NPCNum) & " / " & Quest(i).CLI(ii).Action(III).Amount & "  Go back and speak with " & Trim$(NPC(Quest(i).CLI(ii).ItemIndex).Name) & " to continue.", BrightGreen, True, i)
                                            Call SendShowTaskCompleteOnNPC(index, Quest(i).CLI(ii).ItemIndex, True)
                                        Else
                                            Call PlayerMsg(index, "Mission Task Completed!  Kills: " & GetPlayerQuestAmount(index, i, NPCNum) & " / " & Quest(i).CLI(ii).Action(III).Amount, BrightGreen, True, i)
                                            Call HandleQuestTask(index, i, ii, III, False)
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    
                        Call SendPlayerQuest(index)
                    End If
                End If
                
NextLoop:

            Next i

        Case Else: Exit Sub
    End Select

End Sub

Private Function ValidArray2(Var() As ActionRec) As Boolean

    On Error GoTo hell

    ValidArray2 = False

    If UBound(Var) > 0 Then ValidArray2 = True

    Exit Function

hell:
    ValidArray2 = False
End Function

Private Function ValidArray1(Var() As CLIRec) As Boolean

    On Error GoTo hell

    ValidArray1 = False

    If UBound(Var) > 0 Then ValidArray1 = True

    Exit Function

hell:
    ValidArray1 = False
End Function

Public Sub QuitQuest(ByVal index As Long, _
                     ByVal QuestNum As Long, _
                     Optional ByVal SendMsg As Boolean = True)

    Dim i As Long

    If QuestNum < 1 Or QuestNum > MAX_QUESTS Then Exit Sub
    
    Call SetPlayerQuestCLI(index, QuestNum, 0)
    Call SetPlayerTask(index, QuestNum, 0)
    Call SetPlayerQuestAmount(index, QuestNum, 0)
    Call SendPlayerQuest(index)

    For i = 1 To Quest(QuestNum).Max_CLI
        Call SendShowTaskCompleteOnNPC(index, i, False)
    Next i

    If SendMsg Then Call PlayerMsg(index, "You have abandoned the mission. (" & Trim$(Quest(QuestNum).Name) & ")", BrightGreen)
End Sub

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~DATABASING~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' ************
' ** Quests **
' ************
Public Sub SaveQuest(ByVal QuestNum As Long)

    Dim filename As String

    Dim F        As Long

    Dim i        As Long, ii As Long

    filename = App.path & "\data\quests\" & QuestNum & ".dat"
    F = FreeFile

    Open filename For Binary As #F
    Put #F, , Quest(QuestNum).Name
    Put #F, , Quest(QuestNum).Description
    Put #F, , Quest(QuestNum).CanBeRetaken
    Put #F, , Quest(QuestNum).Max_CLI
    Put #F, , Quest(QuestNum).Rank
    Put #F, , Quest(QuestNum).OutOfOrder
    
    Put #F, , Quest(QuestNum).Requirements.AccessReq
    Put #F, , Quest(QuestNum).Requirements.ClassReq
    Put #F, , Quest(QuestNum).Requirements.GenderReq
    Put #F, , Quest(QuestNum).Requirements.LevelReq
    Put #F, , Quest(QuestNum).Requirements.SkillLevelReq
    Put #F, , Quest(QuestNum).Requirements.SkillReq
    
    For i = 1 To Stats.Stat_count - 1
        Put #F, , Quest(QuestNum).Requirements.Stat_Req(i)
    Next i

    If Quest(QuestNum).Max_CLI > 0 Then

        For i = 1 To Quest(QuestNum).Max_CLI
            Put #F, , Quest(QuestNum).CLI(i).ItemIndex
            Put #F, , Quest(QuestNum).CLI(i).isNPC
            Put #F, , Quest(QuestNum).CLI(i).Max_Actions
    
            For ii = 1 To Quest(QuestNum).CLI(i).Max_Actions
                Put #F, , Quest(QuestNum).CLI(i).Action(ii).TextHolder
                Put #F, , Quest(QuestNum).CLI(i).Action(ii).ActionID
                Put #F, , Quest(QuestNum).CLI(i).Action(ii).Amount
                Put #F, , Quest(QuestNum).CLI(i).Action(ii).MainData
                Put #F, , Quest(QuestNum).CLI(i).Action(ii).QuadData
                Put #F, , Quest(QuestNum).CLI(i).Action(ii).SecondaryData
                Put #F, , Quest(QuestNum).CLI(i).Action(ii).TertiaryData
            Next ii
        Next i

    End If
    
    Close #F

    DoEvents
End Sub

Public Sub LoadQuests()

    Dim filename As String

    Dim F        As Long

    Dim i        As Long, ii As Long, III As Long
    
    Call CheckQuests

    For i = 1 To MAX_QUESTS

        filename = App.path & "\data\quests\" & i & ".dat"
        F = FreeFile

        Open filename For Binary As #F
        Get #F, , Quest(i).Name
        Get #F, , Quest(i).Description
        Get #F, , Quest(i).CanBeRetaken
        Get #F, , Quest(i).Max_CLI
        Get #F, , Quest(i).Rank
        Get #F, , Quest(i).OutOfOrder
        
        Get #F, , Quest(i).Requirements.AccessReq
        Get #F, , Quest(i).Requirements.ClassReq
        Get #F, , Quest(i).Requirements.GenderReq
        Get #F, , Quest(i).Requirements.LevelReq
        Get #F, , Quest(i).Requirements.SkillLevelReq
        Get #F, , Quest(i).Requirements.SkillReq
        
        For ii = 1 To Stats.Stat_count - 1
            Get #F, , Quest(i).Requirements.Stat_Req(ii)
        Next ii

        If Quest(i).Max_CLI > 0 Then
            ReDim Quest(i).CLI(1 To Quest(i).Max_CLI)
            
            For ii = 1 To Quest(i).Max_CLI
                Get #F, , Quest(i).CLI(ii).ItemIndex
                Get #F, , Quest(i).CLI(ii).isNPC
                Get #F, , Quest(i).CLI(ii).Max_Actions

                If Quest(i).CLI(ii).Max_Actions > 0 Then
                    ReDim Quest(i).CLI(ii).Action(Quest(i).CLI(ii).Max_Actions)

                    For III = 1 To Quest(i).CLI(ii).Max_Actions
                        Get #F, , Quest(i).CLI(ii).Action(III).TextHolder
                        Get #F, , Quest(i).CLI(ii).Action(III).ActionID
                        Get #F, , Quest(i).CLI(ii).Action(III).Amount
                        Get #F, , Quest(i).CLI(ii).Action(III).MainData
                        Get #F, , Quest(i).CLI(ii).Action(III).QuadData
                        Get #F, , Quest(i).CLI(ii).Action(III).SecondaryData
                        Get #F, , Quest(i).CLI(ii).Action(III).TertiaryData
                    Next III

                End If

            Next ii

        End If
        
        Close #F

        DoEvents
    Next i

End Sub

Public Sub ClearQuests()

    Dim i As Long

    For i = 1 To MAX_QUESTS
        Call ClearQuest(i)
    Next

End Sub

Public Sub ClearQuest(ByVal QuestNum As Long)
    Quest(QuestNum).Name = vbNullString
    Quest(QuestNum).Description = vbNullString
    Quest(QuestNum).Rank = vbNullString
    Call ZeroMemory(ByVal VarPtr(Quest(QuestNum)), LenB(Quest(QuestNum)))
    Quest(QuestNum).Requirements.ClassReq = 0
    Quest(QuestNum).Requirements.GenderReq = 0
    Quest(QuestNum).Requirements.SkillReq = 0
End Sub

Public Sub CheckQuests()

    Dim i As Long

    For i = 1 To MAX_QUESTS

        If Not FileExist("\data\quests\" & i & ".dat") Then
            Call SaveQuest(i)
        End If

    Next

End Sub

Sub ClearPlayerQuests()
    Dim i As Long
    
    For i = 1 To MAX_PLAYERS
        ClearPlayerQuest (i)
    Next
End Sub

Sub ClearPlayerQuest(ByVal index As Long)
    Dim i As Long

    Call ZeroMemory(ByVal VarPtr(PlayerQuest(index)), LenB(PlayerQuest(index)))

    ReDim PlayerQuest(index).QuestCompleted(1 To MAX_QUESTS)
    ReDim PlayerQuest(index).QuestTask(1 To MAX_QUESTS)
    ReDim PlayerQuest(index).QuestCLI(1 To MAX_QUESTS)
    ReDim PlayerQuest(index).QuestAmount(1 To MAX_QUESTS)
    
    For i = 1 To MAX_QUESTS
        ReDim PlayerQuest(index).QuestAmount(i).ID(1 To MAX_NPCS)
    Next
End Sub

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~QUEST TCP~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Sub SendShowTaskCompleteOnNPC(ByVal index As Long, _
                              ByVal NPCNum As Long, _
                              ShowIt As Boolean)

    If NPCNum < 1 Or NPCNum > MAX_NPCS Then Exit Sub
    If index < 1 Or index > Player_HighIndex Then Exit Sub

    NPC(NPCNum).ShowQuestCompleteIcon = Abs(ShowIt)
    Call SendNPCs(index)
    Call SaveNPCs
End Sub

Sub SendQuests(ByVal index As Long)

    Dim i As Long
    
    For i = 1 To MAX_QUESTS

        If Len(Trim$(Quest(i).Name)) > 0 Then
            Call SendUpdateQuestTo(index, i)
        End If

    Next

End Sub

Sub SendUpdateQuestTo23(ByVal index As Long, ByVal QuestNum As Integer)

    Dim buffer As clsBuffer

    Dim i      As Long, ii As Long

    Set buffer = New clsBuffer

    buffer.WriteLong SUpdateQuest

    With Quest(QuestNum)
        buffer.WriteLong QuestNum
        buffer.WriteString .Name
        buffer.WriteString .Description
        buffer.WriteLong .CanBeRetaken
        buffer.WriteLong .Max_CLI
        buffer.WriteString .Rank
        buffer.WriteByte .OutOfOrder
        
        buffer.WriteLong .Requirements.AccessReq
        buffer.WriteLong .Requirements.ClassReq
        buffer.WriteLong .Requirements.GenderReq
        buffer.WriteLong .Requirements.LevelReq
        buffer.WriteLong .Requirements.SkillLevelReq
        buffer.WriteLong .Requirements.SkillReq
        
        For i = 1 To Stats.Stat_count - 1
            buffer.WriteLong .Requirements.Stat_Req(i)
        Next i

        If .Max_CLI > 0 Then

            For i = 1 To .Max_CLI
                buffer.WriteLong .CLI(i).ItemIndex
                buffer.WriteLong .CLI(i).isNPC
                buffer.WriteLong .CLI(i).Max_Actions
    
                For ii = 1 To .CLI(i).Max_Actions
                    buffer.WriteString .CLI(i).Action(ii).TextHolder
                    buffer.WriteLong .CLI(i).Action(ii).ActionID
                    buffer.WriteLong .CLI(i).Action(ii).Amount
                    buffer.WriteLong .CLI(i).Action(ii).MainData
                    buffer.WriteLong .CLI(i).Action(ii).QuadData
                    buffer.WriteLong .CLI(i).Action(ii).SecondaryData
                    buffer.WriteLong .CLI(i).Action(ii).TertiaryData
                Next ii
            Next i

        End If

    End With

    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing

End Sub

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~DATA HANDLER~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Sub HandleQuestEditor(ByVal index As Long, _
                                   ByRef Data() As Byte, _
                                   ByVal StartAddr As Long, _
                                   ByVal ExtraVar As Long)

    Dim buffer   As clsBuffer

    Dim EventNum As Long

    ' Prevent hacking
    If GetPlayerAccess(index) < STAFF_MAPPER Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteLong SEditQuest
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub HandleSaveQuest2(ByVal index As Long, _
                           ByRef Data() As Byte, _
                           ByVal StartAddr As Long, _
                           ByVal ExtraVar As Long)

    Dim buffer   As clsBuffer

    Dim i        As Long, ii As Long

    Dim QuestNum As Long

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    QuestNum = buffer.ReadLong

    With Quest(QuestNum)

        .Name = buffer.ReadString
        .Description = buffer.ReadString
        .CanBeRetaken = buffer.ReadLong
        .Max_CLI = buffer.ReadLong
        .Rank = buffer.ReadString
        .OutOfOrder = buffer.ReadByte
        
        .Requirements.AccessReq = buffer.ReadLong
        .Requirements.ClassReq = buffer.ReadLong
        .Requirements.GenderReq = buffer.ReadLong
        .Requirements.LevelReq = buffer.ReadLong
        .Requirements.SkillLevelReq = buffer.ReadLong
        .Requirements.SkillReq = buffer.ReadLong

        For i = 1 To Stats.Stat_count - 1
            .Requirements.Stat_Req(i) = buffer.ReadLong
        Next i

        If .Max_CLI > 0 Then
            ReDim .CLI(1 To .Max_CLI)

            For i = 1 To .Max_CLI
                .CLI(i).ItemIndex = buffer.ReadLong
                .CLI(i).isNPC = buffer.ReadLong
                .CLI(i).Max_Actions = buffer.ReadLong

                If .CLI(i).Max_Actions > 0 Then
                    ReDim Preserve .CLI(i).Action(1 To .CLI(i).Max_Actions)

                    For ii = 1 To .CLI(i).Max_Actions
                        .CLI(i).Action(ii).TextHolder = buffer.ReadString
                        .CLI(i).Action(ii).ActionID = buffer.ReadLong
                        .CLI(i).Action(ii).Amount = buffer.ReadLong
                        .CLI(i).Action(ii).MainData = buffer.ReadLong
                        .CLI(i).Action(ii).QuadData = buffer.ReadLong
                        .CLI(i).Action(ii).SecondaryData = buffer.ReadLong
                        .CLI(i).Action(ii).TertiaryData = buffer.ReadLong
                    Next ii

                End If

            Next i

        End If

    End With

    Call SaveQuest(QuestNum)

    Set buffer = Nothing
End Sub

Public Sub HandleQuitQuest(ByVal index As Long, _
                           ByRef Data() As Byte, _
                           ByVal StartAddr As Long, _
                           ByVal ExtraVar As Long)

    Dim buffer   As clsBuffer

    Dim QuestNum As Long

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    QuestNum = buffer.ReadLong
    Set buffer = Nothing

    Call QuitQuest(index, QuestNum)
End Sub

Public Sub HandleAcceptQuest(ByVal index As Long, _
                             ByRef Data() As Byte, _
                             ByVal StartAddr As Long, _
                             ByVal ExtraVar As Long)

    Dim buffer  As clsBuffer

    Dim QuestID As Long

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    QuestID = buffer.ReadLong
    Set buffer = Nothing

    If QuestID < 1 Or QuestID > MAX_QUESTS Then Exit Sub

    'set the PlayerQuest questid to this quest, and set the cli/greeter to the first one in the quest
    Call SetPlayerQuestCLI(index, QuestID, 1)
    Call SetPlayerTask(index, QuestID, 2)
    Call SendPlayerQuest(index)

    'Start processing the tasks of the quest.
    Call HandleQuestTask(index, QuestID, GetPlayerQuestCLI(index, QuestID), GetPlayerQuestTask(index, QuestID), False)
End Sub

Public Sub HandleRequestQuests(ByVal index As Long, _
                               ByRef Data() As Byte, _
                               ByVal StartAddr As Long, _
                               ByVal ExtraVar As Long)
    SendQuests index
End Sub

Public Sub Var_Switch_Change(ByVal index As Long, _
                             ByVal Variable As Boolean, _
                             ByVal Value As Long)

    Dim i         As Long, ii As Long, CLIID As Long, TASK As Long

    Dim SkipConvo As Boolean

    Call SendPlayerData(index)
    
    For i = 1 To MAX_QUESTS
        CLIID = GetPlayerQuestCLI(index, i)
        TASK = GetPlayerQuestTask(index, i)
        
        If CLIID > 0 Then
            If Quest(i).Max_CLI >= CLIID Then

                If Quest(i).CLI(CLIID).Max_Actions >= TASK Then

                    With Quest(i).CLI(CLIID).Action(TASK)
    
                        If .ActionID = TASK_VARIABLE Then
                            If CBool(.MainData) = Variable Then
                                If Value = .Amount Then
                                    SkipConvo = CBool(.QuadData)
                                    
                                    If SkipConvo Then
                                        Call PlayerMsg(index, "Mission Task Completed.  (" & Trim$(.TextHolder) & ")", BrightGreen)
                                        Call HandleQuestTask(index, i, CLIID, TASK)

                                        Exit Sub

                                    Else
                                        Call PlayerMsg(index, "Mission Task Completed.  (" & Trim$(.TextHolder) & ")", BrightGreen)
                                        Call PlayerMsg(index, "Go back and speak with " & Trim$(NPC(Quest(i).CLI(CLIID).ItemIndex).Name) & " to continue.", BrightGreen)

                                        Exit Sub

                                    End If
                                End If
                            End If
                        End If
    
                    End With

                End If

            End If
        End If

    Next i

End Sub

Public Sub SendQuestSound(ByVal index As Long, _
                          ByVal SoundToPlay As Long, _
                          ByVal ToWho As Byte)

    Dim buffer As clsBuffer

    Dim Data() As Byte

    Dim i      As Long

    If SoundToPlay < 0 Then Exit Sub
    If ToWho < 0 Or ToWho > 2 Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteLong SPlaySound
    buffer.WriteLong SoundToPlay
    Data() = buffer.ToArray()
    Set buffer = Nothing
    
    Select Case ToWho
    
        Case 0 'Player
            SendDataTo index, Data()

        Case 1 'Map

            For i = 1 To Player_HighIndex

                If GetPlayerMap(i) = GetPlayerMap(index) Then
                    SendDataTo index, Data()
                End If

            Next i

        Case 2 'All

            For i = 1 To Player_HighIndex
                SendDataTo index, Data()
            Next i
    
    End Select

End Sub

Public Function HasQuestSkill(ByVal index As Long, QuestID As Long, Optional ByVal ReturnIfNot As Boolean = False) As Long
Dim i As Long, CLIIndex As Long, TaskIndex As Long
    CLIIndex = GetPlayerQuestCLI(index, QuestID)
    TaskIndex = GetPlayerQuestTask(index, QuestID)
    
    HasQuestSkill = 0
    
    If CLIIndex > 0 Then
        If TaskIndex > 0 Then
            For i = TaskIndex To 1 Step -1
                If Quest(QuestID).CLI(CLIIndex).Action(i).ActionID = TASK_GETSKILL Then
                    If GetPlayerSkill(index, Quest(QuestID).CLI(CLIIndex).Action(i).MainData) >= Quest(QuestID).CLI(CLIIndex).Action(i).Amount Then
                        HasQuestSkill = Quest(QuestID).CLI(CLIIndex).ItemIndex 'return the npc number
                        Exit Function
                    Else
                        If ReturnIfNot Then
                            'return a value meant to be parsed so we can distinguish the returned value
                            HasQuestSkill = Quest(QuestID).CLI(CLIIndex).ItemIndex & "|" & "Can't be empty... lol"
                            Exit Function
                        End If
                    End If
                End If
            Next i
        End If
    End If
End Function

Sub SendPlayerQuest(ByVal index As Long)
    Dim i      As Long, ii As Long
            
    For i = 1 To MAX_QUESTS

        With PlayerQuest(index)
                Account(index).Chars(GetPlayerChar(index)).QuestCompleted(i) = .QuestCompleted(i)
                Account(index).Chars(GetPlayerChar(index)).QuestCLI(i) = .QuestCLI(i)
                Account(index).Chars(GetPlayerChar(index)).QuestTask(i) = .QuestTask(i)
                    
            For ii = 1 To MAX_NPCS
                Account(index).Chars(GetPlayerChar(index)).QuestAmount(i).ID(ii) = .QuestAmount(i).ID(ii)
            Next ii
        End With
    Next i
    
    SendPlayerData (index)
            
End Sub

Sub SetPlayerQuestData(ByVal index As Long)
    Dim i As Long, ii As Long
    
    For i = 1 To MAX_QUESTS

        With PlayerQuest(index)
                .QuestCompleted(i) = GetPlayerQuestCompleted(index, i)
                .QuestCLI(i) = GetPlayerQuestCLI(index, i)
                .QuestTask(i) = GetPlayerQuestTask(index, i)
                    
            For ii = 1 To MAX_NPCS
                .QuestAmount(i).ID(ii) = GetPlayerQuestAmount(index, i, ii)
            Next ii
                
        End With

    Next i

End Sub
