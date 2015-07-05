Attribute VB_Name = "modQuests"
Option Explicit

Function PlayerHasQuest(ByVal index As Long) As Boolean
    Dim Qplayer As PlayerRec
    Qplayer = player(index).Char(player(index).CharNum)
    If Qplayer.CurrentQuest > 0 Then
        PlayerHasQuest = True
    Else
        PlayerHasQuest = False
    End If
End Function

Function SetCurrentQuest(ByVal index As Long, ByVal QuestID As Long, ByRef Qplayer As PlayerRec) As Boolean
    If QuestID > 0 And QuestID < MAX_QUESTS Then
        If Qplayer.level >= Quests(QuestID).requiredLevel Then
            If Qplayer.CurrentQuest = 0 Then
                Qplayer.CurrentQuest = QuestID
                Call SetCurrentQuestStatus(index, 1, Qplayer)
                SetCurrentQuest = True
                Exit Function
            Else
                Call PlayerMsg(index, "You currently have a quest.", RGB_AlertColor)
            End If
        Else
            Call PlayerMsg(index, "You are of too low level to start this quest.", RGB_AlertColor)
        End If
    End If
    SetCurrentQuest = False
End Function
Sub SetCurrentQuestStatus(ByVal index As Long, ByVal status As Long, ByRef Qplayer As PlayerRec)
    If status > 0 And status < 4 Then
        Qplayer.QuestStatus = status
        Exit Sub
    End If
End Sub

Sub StartQuest(ByVal index As Long, ByVal QuestID As Long)
    If player(index).Char(player(index).CharNum).CurrentQuest = QuestID And player(index).Char(player(index).CharNum).QuestStatus = 2 Then
        'player is has finished quest?
        FinishQuest (index)
        Exit Sub
    End If
    If (SetCurrentQuest(index, QuestID, player(index).Char(player(index).CharNum)) = True) Then
        Call SendQuestMessage(index, Quests(QuestID).StartQuestMsg)
    End If
End Sub

Public Sub checkQuestProgression(ByVal index As Long)
Dim QuestID As Long
QuestID = player(index).Char(player(index).CharNum).CurrentQuest
    If QuestID > 0 And QuestID < MAX_QUESTS And player(index).Char(player(index).CharNum).QuestStatus = 1 Then
        If HasItem(index, Quests(QuestID).ItemToObtain) Then
            Call SetCurrentQuestStatus(index, 2, player(index).Char(player(index).CharNum))
            'FinishQuest (index)
            Call SendQuestMessage(index, Quests(QuestID).GetItemQuestMsg)
        End If
    End If
End Sub

Sub FinishQuest(ByVal index As Long)
Dim QuestID As Long
Dim QuestStatus As Long
QuestID = player(index).Char(player(index).CharNum).CurrentQuest
QuestStatus = player(index).Char(player(index).CharNum).QuestStatus
    If QuestID > 0 And QuestID < MAX_QUESTS And QuestStatus = 2 Then
        Call TakeItem(index, Quests(QuestID).ItemToObtain, 1)
        Call GiveItem(index, Quests(QuestID).ItemGiven, Quests(QuestID).ItemValGiven)
        Call GiveItem(index, 2, Quests(QuestID).goldGiven)
        Call SetPlayerExp(index, (GetPlayerExp(index) + Quests(QuestID).ExpGiven))
        Call SendQuestMessage(index, Quests(QuestID).FinishQuestMessage)
        Call SendStatsInfo(index)
        player(index).Char(player(index).CharNum).CurrentQuest = 0
        player(index).Char(player(index).CharNum).QuestStatus = 0
    End If
End Sub

Sub SendQuestMessage(ByVal index As Long, ByVal msg As String)
Dim packet As String

    packet = "QUESTMSG" & SEP_CHAR & procQuestMsg(index, msg) & SEP_CHAR & END_CHAR
    Call SendDataTo(index, packet)
End Sub




Public Function procQuestMsg(ByVal index As Long, ByVal msg As String) As String
    Dim QuestID As Long
    
    Dim Qplayer As PlayerRec
    Qplayer = player(index).Char(player(index).CharNum)
    QuestID = Qplayer.CurrentQuest
    msg = Replace(msg, "#NAME#", Trim(Qplayer.name), , , vbTextCompare)
    msg = Replace(msg, "#LEVEL#", Trim(Qplayer.level), , , vbTextCompare)
    msg = Replace(msg, "#ITEM#", Trim(Item(Quests(QuestID).ItemToObtain).name), , , vbTextCompare)
    procQuestMsg = msg
End Function
