Attribute VB_Name = "modFuncs"
Option Explicit

Function DoQuest(ByVal questnum As Long, ByVal Index As Long, ByVal npcnum As Long)
Dim BoB

If ReadINI(Player(Index).Char(Player(Index).CharNum).Name, "QUEST" & Npc(npcnum).Quest, App.Path + "\qflag.ini") = 0 Then
If MeetReq(questnum, Index) Then
    If Quest(questnum).StartOn = 0 Then
        Call WriteINI(Player(Index).Char(Player(Index).CharNum).Name, "QUEST" & Npc(npcnum).Quest, "1", App.Path + "\qflag.ini")
        Call QuestMsg(Index, "----Quest Recieved----", BrightGreen)
        Call QuestMsg(Index, "" & Trim(Npc(npcnum).Name) & " says, '" & Trim(Quest(Npc(npcnum).Quest).Start) & "'", SayColor)
    ElseIf Quest(questnum).StartOn = 1 Then
        Call GiveQuestItem(Index, Quest(questnum).StartItem, Quest(questnum).Startval, npcnum)
    End If
Else
If ReadINI(Player(Index).Char(Player(Index).CharNum).Name, "QUEST" & Npc(npcnum).Quest, App.Path + "\qflag.ini") = 2 Then
Call QuestMsg(Index, "" & Trim(Npc(npcnum).Name) & " says, '" & Trim(Quest(Npc(npcnum).Quest).After) & "'", SayColor)
Exit Function
End If
    Call QuestMsg(Index, "" & Trim(Npc(npcnum).Name) & " says, '" & Trim(Quest(Npc(npcnum).Quest).Before) & "'", SayColor)
End If
Exit Function
End If

If ReadINI(Player(Index).Char(Player(Index).CharNum).Name, "QUEST" & Npc(npcnum).Quest, App.Path + "\qflag.ini") = 1 Then
    Call SendDataTo(Index, "questprompt" & SEP_CHAR & questnum & SEP_CHAR & npcnum & SEP_CHAR & END_CHAR)
End If

If ReadINI(Player(Index).Char(Player(Index).CharNum).Name, "QUEST" & Npc(npcnum).Quest, App.Path + "\qflag.ini") = 2 Then
Call QuestMsg(Index, "" & Trim(Npc(npcnum).Name) & " says, '" & Trim(Quest(Npc(npcnum).Quest).After) & "'", SayColor)
Exit Function
End If

End Function
Sub SaveLine(File As Integer, Header As String, Var As String, Value As String)
    Print #File, Var & "=" & Value
End Sub

Function MeetReq(questnum As Long, Index As Long) As Boolean
If Quest(questnum).ClassIsReq = 0 And Quest(questnum).LevelIsReq = 0 Then
    MeetReq = True
    Exit Function
ElseIf Quest(questnum).ClassIsReq = 1 And Quest(questnum).LevelIsReq = 0 Then
    If Player(Index).Char(Player(Index).CharNum).Class = Quest(questnum).ClassReq Then
        MeetReq = True
        Exit Function
    Else
        MeetReq = False
        Exit Function
    End If
ElseIf Quest(questnum).ClassIsReq = 0 And Quest(questnum).LevelIsReq = 1 Then
    If Player(Index).Char(Player(Index).CharNum).Level >= Quest(questnum).LevelReq Then
        MeetReq = True
        Exit Function
    Else
        MeetReq = False
        Exit Function
    End If
ElseIf Quest(questnum).ClassIsReq = 1 And Quest(questnum).LevelIsReq = 1 Then
    If Player(Index).Char(Player(Index).CharNum).Class = Quest(questnum).ClassReq And Player(Index).Char(Player(Index).CharNum).Level >= Quest(questnum).LevelReq Then
        MeetReq = True
        Exit Function
    Else
        MeetReq = False
        Exit Function
    End If
End If

End Function

Sub GiveQuestItem(ByVal Index As Long, ByVal ItemNum As Long, ByVal ItemVal As Long, ByVal npcnum As Long)
Dim i As Long
Dim Curr As Boolean
Dim Has As Boolean
    ' Check for subscript out of range
    If IsPlaying(Index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Sub
    End If
    
    If Item(ItemNum).Type = 12 Then Curr = True Else Curr = False
    
    For i = 1 To MAX_INV
        If Curr = True Then
            If GetPlayerInvItemNum(Index, i) = ItemNum Then
                Call SetPlayerInvItemValue(Index, i, GetPlayerInvItemValue(Index, i) + ItemVal)
                Call SendInventoryUpdate(Index, i)
                Has = True
                Exit For
            End If
        Else
            If GetPlayerInvItemNum(Index, i) = 0 Then
                Call SetPlayerInvItemNum(Index, i, ItemNum)
                Call SetPlayerInvItemValue(Index, i, 1)
                If (Item(ItemNum).Type = ITEM_TYPE_ARMOR) Or (Item(ItemNum).Type = ITEM_TYPE_WEAPON) Or (Item(ItemNum).Type = ITEM_TYPE_HELMET) Or (Item(ItemNum).Type = ITEM_TYPE_SHIELD) Then
                    Call SetPlayerInvItemDur(Index, i, Item(ItemNum).Data1)
                End If
                Call SendInventoryUpdate(Index, i)
                Has = True
                Exit For
            End If
        End If
    Next i
    
    If Has = False Then
        Call PlayerMsg(Index, "Your inventory is full. Please come back when it is not", BrightRed)
        Exit Sub
    Else
        Call WriteINI(Player(Index).Char(Player(Index).CharNum).Name, "QUEST" & Npc(npcnum).Quest, "1", App.Path + "\qflag.ini")
        Call QuestMsg(Index, "" & Trim(Npc(npcnum).Name) & " says, '" & Trim(Quest(Npc(npcnum).Quest).Start) & "'", SayColor)
    End If

End Sub
Sub GiveRewardItem(ByVal Index As Long, ByVal ItemNum As Long, ByVal ItemVal As Long, ByVal npcnum As Long)
Dim i As Long
Dim Curr As Boolean
Dim Has As Boolean
Dim questnum As Long
    ' Check for subscript out of range
    If IsPlaying(Index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Sub
    End If
    
    If ReadINI(Player(Index).Char(Player(Index).CharNum).Name, "QUEST" & Npc(npcnum).Quest, App.Path + "\qflag.ini") = 2 Then
Call QuestMsg(Index, "A " & Trim(Npc(npcnum).Name) & " says, '" & Trim(Quest(Npc(npcnum).Quest).After) & "'", SayColor)
Exit Sub
End If
    
    If Item(ItemNum).Type = 12 Then Curr = True Else Curr = False
    
    For i = 1 To MAX_INV
        If Curr = True Then
            If GetPlayerInvItemNum(Index, i) = ItemNum Then
                Call SetPlayerInvItemValue(Index, i, GetPlayerInvItemValue(Index, i) + ItemVal)
                Call SendInventoryUpdate(Index, i)
                Has = True
                 Call WriteINI(Player(Index).Char(Player(Index).CharNum).Name, "QUEST" & Npc(npcnum).Quest, "2", App.Path + "\qflag.ini")
                Exit For
            End If
        Else
            If GetPlayerInvItemNum(Index, i) = 0 Then
                Call SetPlayerInvItemNum(Index, i, ItemNum)
                Call SetPlayerInvItemValue(Index, i, 1)
                If (Item(ItemNum).Type = ITEM_TYPE_ARMOR) Or (Item(ItemNum).Type = ITEM_TYPE_WEAPON) Or (Item(ItemNum).Type = ITEM_TYPE_HELMET) Or (Item(ItemNum).Type = ITEM_TYPE_SHIELD) Then
                    Call SetPlayerInvItemDur(Index, i, Item(ItemNum).Data1)
                End If
                Call SendInventoryUpdate(Index, i)
                Has = True
                Exit For
            End If
        End If
    Next i
    
    If Has = False Then
        Call PlayerMsg(Index, "Your inventory is full. Please come back when it is not", BrightRed)
        Exit Sub
    Else
        Call QuestMsg(Index, "----Quest Complete----", BrightGreen)
        Call QuestMsg(Index, "" & Trim(Npc(npcnum).Name) & " says, '" & Trim(Quest(Npc(npcnum).Quest).End) & "'", SayColor)
        Call WriteINI(Player(Index).Char(Player(Index).CharNum).Name, "QUEST" & Npc(npcnum).Quest, "2", App.Path + "\qflag.ini")
        Call DetermineExpType(questnum, Index)
        Call SendPlayerData(Index)
        Call CheckPlayerLevelUp(Index)
        If Item(Quest(Npc(npcnum).Quest).RewardNum).Type = 12 Then
            Call TakeItem(Index, Quest(Npc(npcnum).Quest).ItemReq, Quest(Npc(npcnum).Quest).ItemVal)
        Else
            Call TakeItem(Index, Quest(Npc(npcnum).Quest).ItemReq, 1)
        End If
                        Call SendInventoryUpdate(Index, i)
    End If

End Sub

Sub DetermineExpType(ByVal questnum As Long, ByVal Index As Long)
Dim ExpAmount As Long
Dim npcnum As Long
Dim i As Long
For i = 1 To MAX_QUESTS


'If Quest(i).FirstAidExp > 0 Then
'ExpAmount = Quest(i).QuestExpReward
'Call SetPlayerFirstAidExp(Index, GetPlayerFirstAidExp(Index) + ExpAmount)
'Call PlayerMsg(Index, "You have recieved " & ExpAmount & " Experience in your First Aid Skill !", BrightGreen)
'Call CheckPlayerFirstAidLevelUp(Index)
'End If


Next i
'call sendstats(Index)
End Sub
