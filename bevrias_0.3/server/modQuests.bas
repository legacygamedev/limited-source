Attribute VB_Name = "modQuests"
Sub GiveItemQ(index, item, value)
Dim Slot As String
Slot = 1
Do While Slot < 24
   If GetPlayerInvItemNum(index, Slot) = 0 Then
       Call SetPlayerInvItemNum(index, Slot, item)
       Call SetPlayerInvItemValue(index, Slot, GetPlayerInvItemValue(index, Slot) + value)
       Call SendInventoryUpdate(index, Slot)
      Slot = 24
   End If
    Slot = Slot + 1
Loop
End Sub

Sub TakeItemQ(index, item)
Dim Slot As String
Slot = 1
Do While Slot < 24
      If GetPlayerInvItemNum(index, Slot) = item Then
         Call SetPlayerInvItemNum(index, Slot, 0)
         Call SetPlayerInvItemValue(index, Slot, 0)
         Call SendInventoryUpdate(index, Slot)
         Slot = 24
      End If
      Slot = Slot + 1
Loop
End Sub

Sub TakeCurrencyQ(index, item, value)
Dim Slot As String
Dim amount As String
Dim take As String
Slot = 1
Do While Slot < 24
   If GetPlayerInvItemNum(index, Slot) = item Then
       amount = GetPlayerInvItemValue(index, Slot)
         take = Int(amount - value)
         If take <= 0 Then
            Call SetPlayerInvItemNum(index, Slot, 0)
            Call SetPlayerInvItemValue(index, Slot, 0)
            Call SendInventoryUpdate(index, Slot)
         End If
         If take > 0 Then
            Call SetPlayerInvItemNum(index, Slot, 0)
            Call SetPlayerInvItemValue(index, Slot, 0)
            Call SetPlayerInvItemNum(index, Slot, item)
            Call SetPlayerInvItemValue(index, Slot, take)
            Call SendInventoryUpdate(index, Slot)
         End If
         Slot = 24
      End If
      Slot = Slot + 1
Loop
End Sub

Sub Questone(ByVal index As Long)
Dim Found As Long
Dim Slot As Long
If GetVar(App.Path & "\Scripts\Part.ini", GetPlayerName(index), "Part") = 1 Then
Call PlayerMsg(index, GetVar(App.Path & "\Quests.ini", "QUEST1", "NPCname") & ": " & GetVar(App.Path & "\Quests.ini", "QUEST1", "NPCsay"), 0)
If GetPlayerLevel(index) > GetVar(App.Path & "\Quests.ini", "QUEST1", "LevelReq") Then
If GetVar(App.Path & "\Quests.ini", "QUEST1", "IfItemNr") > 0 Then

Found = 0
Slot = 1
Do While Slot < 24
If GetPlayerInvItemNum(index, Slot) = GetVar(App.Path & "\Quests.ini", "QUEST1", "IfItemNr") Then
Found = 1
Slot = 24
End If
Slot = Slot + 1
Loop

If Found = 1 Then
If GetVar(App.Path & "\Quests.ini", "QUEST1", "Keep") = 1 Then
Call TakeItemQ(index, GetVar(App.Path & "\Quests.ini", "QUEST1", "IfItemNr"))
Call PlayerMsg(index, GetVar(App.Path & "\Quests.ini", "QUEST1", "NPCname") & ": " & GetVar(App.Path & "\Quests.ini", "QUEST1", "FinishedNPCsay"), 14)
Call SetPlayerPOINTS(index, GetPlayerPOINTS(index) + GetVar(App.Path & "\Quests.ini", "QUEST1", "StatPoints"))
Call SetPlayerExp(index, GetPlayerExp(index) + GetVar(App.Path & "\Quests.ini", "QUEST1", "Exp"))
Call GiveItemQ(index, GetVar(App.Path & "\Quests.ini", "QUEST1", "ItemNr"), 1)
Else
Call PlayerMsg(index, GetVar(App.Path & "\Quests.ini", "QUEST1", "NPCname") & ": " & GetVar(App.Path & "\Quests.ini", "QUEST1", "FinishedNPCsay"), 14)
End If
If GetVar(App.Path & "\Quests.ini", "QUEST1", "FullHP") = 1 Then
Call SetPlayerHP(index, GetPlayerMaxHP(index))
Call SendHP(index)
End If
If GetVar(App.Path & "\Quests.ini", "QUEST1", "FullMP") = 1 Then
Call SetPlayerMP(index, GetPlayerMaxMP(index))
Call SendMP(index)
End If
If GetVar(App.Path & "\Quests.ini", "QUEST1", "FullSP") = 1 Then
Call SetPlayerSP(index, GetPlayerMaxSP(index))
Call SendSP(index)
End If
Call PutVar(App.Path & "\Scripts\Part.ini", GetPlayerName(index), "Part", "2")
Call SetPlayerSprite(index, GetVar(App.Path & "\Quests.ini", "QUEST1", "Sprite"))
Call SetPlayerSTR(index, GetPlayerSTR(index) + GetVar(App.Path & "\Quests.ini", "QUEST1", "Strength"))
Call SetPlayerDEF(index, GetPlayerDEF(index) + GetVar(App.Path & "\Quests.ini", "QUEST1", "Defence"))
Call SetPlayerSPEED(index, GetPlayerSPEED(index) + GetVar(App.Path & "\Quests.ini", "QUEST1", "Agility"))
Call SetPlayerMAGI(index, GetPlayerMAGI(index) + GetVar(App.Path & "\Quests.ini", "QUEST1", "Wisdom"))
End If

If Found = 0 Then
Call PlayerMsg(index, GetVar(App.Path & "\Quests.ini", "QUEST1", "NPCname") & ": " & GetVar(App.Path & "\Quests.ini", "QUEST1", "NoItemNPCsay"), 0)
End If
End If

If GetVar(App.Path & "\Quests.ini", "QUEST1", "IfItemNr") = 0 Then
Call PlayerMsg(index, GetVar(App.Path & "\Quests.ini", "QUEST1", "NPCname") & ": " & GetVar(App.Path & "\Quests.ini", "QUEST1", "FinishedNPCsay"), 14)
Call SetPlayerPOINTS(index, GetPlayerPOINTS(index) + GetVar(App.Path & "\Quests.ini", "QUEST1", "StatPoints"))
Call SetPlayerExp(index, GetPlayerExp(index) + GetVar(App.Path & "\Quests.ini", "QUEST1", "Exp"))
Call GiveItemQ(index, GetVar(App.Path & "\Quests.ini", "QUEST1", "ItemNr"), 1)
Call PutVar(App.Path & "\Scripts\Part.ini", GetPlayerName(index), "Part", "2")
If GetVar(App.Path & "\Quests.ini", "QUEST1", "FullHP") = 1 Then
Call SetPlayerHP(index, GetPlayerMaxHP(index))
Call SendHP(index)
End If
If GetVar(App.Path & "\Quests.ini", "QUEST1", "FullMP") = 1 Then
Call SetPlayerMP(index, GetPlayerMaxMP(index))
Call SendMP(index)
End If
If GetVar(App.Path & "\Quests.ini", "QUEST1", "FullSP") = 1 Then
Call SetPlayerSP(index, GetPlayerMaxSP(index))
Call SendSP(index)
End If
Call SetPlayerSTR(index, GetPlayerSTR(index) + GetVar(App.Path & "\Quests.ini", "QUEST1", "Strength"))
Call SetPlayerDEF(index, GetPlayerDEF(index) + GetVar(App.Path & "\Quests.ini", "QUEST1", "Defence"))
Call SetPlayerSPEED(index, GetPlayerSPEED(index) + GetVar(App.Path & "\Quests.ini", "QUEST1", "Agility"))
Call SetPlayerMAGI(index, GetPlayerMAGI(index) + GetVar(App.Path & "\Quests.ini", "QUEST1", "Wisdom"))
End If
End If
End If
If GetVar(App.Path & "\Scripts\Part.ini", GetPlayerName(index), "Part") = 1 Then
Else
Call PlayerMsg(index, GetVar(App.Path & "\Quests.ini", "QUEST1", "NPCname") & ": " & GetVar(App.Path & "\Quests.ini", "QUEST1", "WrongQuestNPCsay"), 0)
End If
End Sub

Sub QuestTwo(ByVal index As Long)
Dim Found As Long
Dim Slot As Long
If GetVar(App.Path & "\Scripts\Part.ini", GetPlayerName(index), "Part") = 1 Then
Call PlayerMsg(index, GetVar(App.Path & "\Quests.ini", "QUEST2", "NPCname") & ": " & GetVar(App.Path & "\Quests.ini", "QUEST2", "NPCsay"), 0)
If GetPlayerLevel(index) > GetVar(App.Path & "\Quests.ini", "QUEST2", "LevelReq") Then
If GetVar(App.Path & "\Quests.ini", "QUEST2", "IfItemNr") > 0 Then

Found = 0
Slot = 1
Do While Slot < 24
If GetPlayerInvItemNum(index, Slot) = GetVar(App.Path & "\Quests.ini", "QUEST2", "IfItemNr") Then
Found = 1
Slot = 24
End If
Slot = Slot + 1
Loop

If Found = 1 Then
If GetVar(App.Path & "\Quests.ini", "QUEST2", "Keep") = 1 Then
Call TakeItemQ(index, GetVar(App.Path & "\Quests.ini", "QUEST2", "IfItemNr"))
Call PlayerMsg(index, GetVar(App.Path & "\Quests.ini", "QUEST2", "NPCname") & ": " & GetVar(App.Path & "\Quests.ini", "QUEST2", "FinishedNPCsay"), 14)
Call SetPlayerPOINTS(index, GetPlayerPOINTS(index) + GetVar(App.Path & "\Quests.ini", "QUEST2", "StatPoints"))
Call SetPlayerExp(index, GetPlayerExp(index) + GetVar(App.Path & "\Quests.ini", "QUEST2", "Exp"))
Call GiveItemQ(index, GetVar(App.Path & "\Quests.ini", "QUEST2", "ItemNr"), 1)
Else
Call PlayerMsg(index, GetVar(App.Path & "\Quests.ini", "QUEST2", "NPCname") & ": " & GetVar(App.Path & "\Quests.ini", "QUEST2", "FinishedNPCsay"), 14)
End If
If GetVar(App.Path & "\Quests.ini", "QUEST2", "FullHP") = 1 Then
Call SetPlayerHP(index, GetPlayerMaxHP(index))
Call SendHP(index)
End If
If GetVar(App.Path & "\Quests.ini", "QUEST2", "FullMP") = 1 Then
Call SetPlayerMP(index, GetPlayerMaxMP(index))
Call SendMP(index)
End If
If GetVar(App.Path & "\Quests.ini", "QUEST2", "FullSP") = 1 Then
Call SetPlayerSP(index, GetPlayerMaxSP(index))
Call SendSP(index)
End If
Call PutVar(App.Path & "\Scripts\Part.ini", GetPlayerName(index), "Part", "3")
Call SetPlayerSprite(index, GetVar(App.Path & "\Quests.ini", "QUEST2", "Sprite"))
Call SetPlayerSTR(index, GetPlayerSTR(index) + GetVar(App.Path & "\Quests.ini", "QUEST2", "Strength"))
Call SetPlayerDEF(index, GetPlayerDEF(index) + GetVar(App.Path & "\Quests.ini", "QUEST2", "Defence"))
Call SetPlayerSPEED(index, GetPlayerSPEED(index) + GetVar(App.Path & "\Quests.ini", "QUEST2", "Agility"))
Call SetPlayerMAGI(index, GetPlayerMAGI(index) + GetVar(App.Path & "\Quests.ini", "QUEST2", "Wisdom"))
End If

If Found = 0 Then
Call PlayerMsg(index, GetVar(App.Path & "\Quests.ini", "QUEST2", "NPCname") & ": " & GetVar(App.Path & "\Quests.ini", "QUEST2", "NoItemNPCsay"), 0)
End If
End If

If GetVar(App.Path & "\Quests.ini", "QUEST2", "IfItemNr") = 0 Then
Call PlayerMsg(index, GetVar(App.Path & "\Quests.ini", "QUEST2", "NPCname") & ": " & GetVar(App.Path & "\Quests.ini", "QUEST2", "FinishedNPCsay"), 14)
Call SetPlayerPOINTS(index, GetPlayerPOINTS(index) + GetVar(App.Path & "\Quests.ini", "QUEST2", "StatPoints"))
Call SetPlayerExp(index, GetPlayerExp(index) + GetVar(App.Path & "\Quests.ini", "QUEST2", "Exp"))
Call GiveItemQ(index, GetVar(App.Path & "\Quests.ini", "QUEST2", "ItemNr"), 1)
Call PutVar(App.Path & "\Scripts\Part.ini", GetPlayerName(index), "Part", "3")
If GetVar(App.Path & "\Quests.ini", "QUEST2", "FullHP") = 1 Then
Call SetPlayerHP(index, GetPlayerMaxHP(index))
Call SendHP(index)
End If
If GetVar(App.Path & "\Quests.ini", "QUEST2", "FullMP") = 1 Then
Call SetPlayerMP(index, GetPlayerMaxMP(index))
Call SendMP(index)
End If
If GetVar(App.Path & "\Quests.ini", "QUEST2", "FullSP") = 1 Then
Call SetPlayerSP(index, GetPlayerMaxSP(index))
Call SendSP(index)
End If
Call SetPlayerSTR(index, GetPlayerSTR(index) + GetVar(App.Path & "\Quests.ini", "QUEST2", "Strength"))
Call SetPlayerDEF(index, GetPlayerDEF(index) + GetVar(App.Path & "\Quests.ini", "QUEST2", "Defence"))
Call SetPlayerSPEED(index, GetPlayerSPEED(index) + GetVar(App.Path & "\Quests.ini", "QUEST2", "Agility"))
Call SetPlayerMAGI(index, GetPlayerMAGI(index) + GetVar(App.Path & "\Quests.ini", "QUEST2", "Wisdom"))
End If
End If
End If
If GetVar(App.Path & "\Scripts\Part.ini", GetPlayerName(index), "Part") = 2 Then
Else
Call PlayerMsg(index, GetVar(App.Path & "\Quests.ini", "QUEST2", "NPCname") & ": " & GetVar(App.Path & "\Quests.ini", "QUEST2", "WrongQuestNPCsay"), 0)
End If
End Sub
Sub QuestThree(ByVal index As Long)
Dim Found As Long
Dim Slot As Long
If GetVar(App.Path & "\Scripts\Part.ini", GetPlayerName(index), "Part") = 3 Then
Call PlayerMsg(index, GetVar(App.Path & "\Quests.ini", "QUEST3", "NPCname") & ": " & GetVar(App.Path & "\Quests.ini", "QUEST3", "NPCsay"), 0)
If GetPlayerLevel(index) > GetVar(App.Path & "\Quests.ini", "QUEST3", "LevelReq") Then
If GetVar(App.Path & "\Quests.ini", "QUEST3", "IfItemNr") > 0 Then

Found = 0
Slot = 1
Do While Slot < 24
If GetPlayerInvItemNum(index, Slot) = GetVar(App.Path & "\Quests.ini", "QUEST3", "IfItemNr") Then
Found = 1
Slot = 24
End If
Slot = Slot + 1
Loop

If Found = 1 Then
If GetVar(App.Path & "\Quests.ini", "QUEST3", "Keep") = 1 Then
Call TakeItemQ(index, GetVar(App.Path & "\Quests.ini", "QUEST3", "IfItemNr"))
Call PlayerMsg(index, GetVar(App.Path & "\Quests.ini", "QUEST3", "NPCname") & ": " & GetVar(App.Path & "\Quests.ini", "QUEST3", "FinishedNPCsay"), 14)
Call SetPlayerPOINTS(index, GetPlayerPOINTS(index) + GetVar(App.Path & "\Quests.ini", "QUEST3", "StatPoints"))
Call SetPlayerExp(index, GetPlayerExp(index) + GetVar(App.Path & "\Quests.ini", "QUEST3", "Exp"))
Call GiveItemQ(index, GetVar(App.Path & "\Quests.ini", "QUEST3", "ItemNr"), 1)
Else
Call PlayerMsg(index, GetVar(App.Path & "\Quests.ini", "QUEST3", "NPCname") & ": " & GetVar(App.Path & "\Quests.ini", "QUEST3", "FinishedNPCsay"), 14)
End If
If GetVar(App.Path & "\Quests.ini", "QUEST3", "FullHP") = 1 Then
Call SetPlayerHP(index, GetPlayerMaxHP(index))
Call SendHP(index)
End If
If GetVar(App.Path & "\Quests.ini", "QUEST3", "FullMP") = 1 Then
Call SetPlayerMP(index, GetPlayerMaxMP(index))
Call SendMP(index)
End If
If GetVar(App.Path & "\Quests.ini", "QUEST3", "FullSP") = 1 Then
Call SetPlayerSP(index, GetPlayerMaxSP(index))
Call SendSP(index)
End If
Call PutVar(App.Path & "\Scripts\Part.ini", GetPlayerName(index), "Part", "4")
Call SetPlayerSprite(index, GetVar(App.Path & "\Quests.ini", "QUEST3", "Sprite"))
Call SetPlayerSTR(index, GetPlayerSTR(index) + GetVar(App.Path & "\Quests.ini", "QUEST3", "Strength"))
Call SetPlayerDEF(index, GetPlayerDEF(index) + GetVar(App.Path & "\Quests.ini", "QUEST3", "Defence"))
Call SetPlayerSPEED(index, GetPlayerSPEED(index) + GetVar(App.Path & "\Quests.ini", "QUEST3", "Agility"))
Call SetPlayerMAGI(index, GetPlayerMAGI(index) + GetVar(App.Path & "\Quests.ini", "QUEST3", "Wisdom"))
End If

If Found = 0 Then
Call PlayerMsg(index, GetVar(App.Path & "\Quests.ini", "QUEST3", "NPCname") & ": " & GetVar(App.Path & "\Quests.ini", "QUEST3", "NoItemNPCsay"), 0)
End If
End If

If GetVar(App.Path & "\Quests.ini", "QUEST3", "IfItemNr") = 0 Then
Call PlayerMsg(index, GetVar(App.Path & "\Quests.ini", "QUEST3", "NPCname") & ": " & GetVar(App.Path & "\Quests.ini", "QUEST3", "FinishedNPCsay"), 14)
Call SetPlayerPOINTS(index, GetPlayerPOINTS(index) + GetVar(App.Path & "\Quests.ini", "QUEST3", "StatPoints"))
Call SetPlayerExp(index, GetPlayerExp(index) + GetVar(App.Path & "\Quests.ini", "QUEST3", "Exp"))
Call GiveItemQ(index, GetVar(App.Path & "\Quests.ini", "QUEST3", "ItemNr"), 1)
Call PutVar(App.Path & "\Scripts\Part.ini", GetPlayerName(index), "Part", "4")
If GetVar(App.Path & "\Quests.ini", "QUEST3", "FullHP") = 1 Then
Call SetPlayerHP(index, GetPlayerMaxHP(index))
Call SendHP(index)
End If
If GetVar(App.Path & "\Quests.ini", "QUEST3", "FullMP") = 1 Then
Call SetPlayerMP(index, GetPlayerMaxMP(index))
Call SendMP(index)
End If
If GetVar(App.Path & "\Quests.ini", "QUEST3", "FullSP") = 1 Then
Call SetPlayerSP(index, GetPlayerMaxSP(index))
Call SendSP(index)
End If
Call SetPlayerSTR(index, GetPlayerSTR(index) + GetVar(App.Path & "\Quests.ini", "QUEST3", "Strength"))
Call SetPlayerDEF(index, GetPlayerDEF(index) + GetVar(App.Path & "\Quests.ini", "QUEST3", "Defence"))
Call SetPlayerSPEED(index, GetPlayerSPEED(index) + GetVar(App.Path & "\Quests.ini", "QUEST3", "Agility"))
Call SetPlayerMAGI(index, GetPlayerMAGI(index) + GetVar(App.Path & "\Quests.ini", "QUEST3", "Wisdom"))
End If
End If
End If
If GetVar(App.Path & "\Scripts\Part.ini", GetPlayerName(index), "Part") = 3 Then
Else
Call PlayerMsg(index, GetVar(App.Path & "\Quests.ini", "QUEST3", "NPCname") & ": " & GetVar(App.Path & "\Quests.ini", "QUEST3", "WrongQuestNPCsay"), 0)
End If
End Sub
Sub QuestFour(ByVal index As Long)
Dim Found As Long
Dim Slot As Long
If GetVar(App.Path & "\Scripts\Part.ini", GetPlayerName(index), "Part") = 4 Then
Call PlayerMsg(index, GetVar(App.Path & "\Quests.ini", "QUEST4", "NPCname") & ": " & GetVar(App.Path & "\Quests.ini", "QUEST4", "NPCsay"), 0)
If GetPlayerLevel(index) > GetVar(App.Path & "\Quests.ini", "QUEST4", "LevelReq") Then
If GetVar(App.Path & "\Quests.ini", "QUEST4", "IfItemNr") > 0 Then

Found = 0
Slot = 1
Do While Slot < 24
If GetPlayerInvItemNum(index, Slot) = GetVar(App.Path & "\Quests.ini", "QUEST4", "IfItemNr") Then
Found = 1
Slot = 24
End If
Slot = Slot + 1
Loop

If Found = 1 Then
If GetVar(App.Path & "\Quests.ini", "QUEST4", "Keep") = 1 Then
Call TakeItemQ(index, GetVar(App.Path & "\Quests.ini", "QUEST4", "IfItemNr"))
Call PlayerMsg(index, GetVar(App.Path & "\Quests.ini", "QUEST4", "NPCname") & ": " & GetVar(App.Path & "\Quests.ini", "QUEST4", "FinishedNPCsay"), 14)
Call SetPlayerPOINTS(index, GetPlayerPOINTS(index) + GetVar(App.Path & "\Quests.ini", "QUEST4", "StatPoints"))
Call SetPlayerExp(index, GetPlayerExp(index) + GetVar(App.Path & "\Quests.ini", "QUEST4", "Exp"))
Call GiveItemQ(index, GetVar(App.Path & "\Quests.ini", "QUEST4", "ItemNr"), 1)
Else
Call PlayerMsg(index, GetVar(App.Path & "\Quests.ini", "QUEST4", "NPCname") & ": " & GetVar(App.Path & "\Quests.ini", "QUEST4", "FinishedNPCsay"), 14)
End If
If GetVar(App.Path & "\Quests.ini", "QUEST1", "FullHP") = 4 Then
Call SetPlayerHP(index, GetPlayerMaxHP(index))
Call SendHP(index)
End If
If GetVar(App.Path & "\Quests.ini", "QUEST1", "FullMP") = 4 Then
Call SetPlayerMP(index, GetPlayerMaxMP(index))
Call SendMP(index)
End If
If GetVar(App.Path & "\Quests.ini", "QUEST1", "FullSP") = 4 Then
Call SetPlayerSP(index, GetPlayerMaxSP(index))
Call SendSP(index)
End If
Call PutVar(App.Path & "\Scripts\Part.ini", GetPlayerName(index), "Part", "5")
Call SetPlayerSprite(index, GetVar(App.Path & "\Quests.ini", "QUEST4", "Sprite"))
Call SetPlayerSTR(index, GetPlayerSTR(index) + GetVar(App.Path & "\Quests.ini", "QUEST4", "Strength"))
Call SetPlayerDEF(index, GetPlayerDEF(index) + GetVar(App.Path & "\Quests.ini", "QUEST4", "Defence"))
Call SetPlayerSPEED(index, GetPlayerSPEED(index) + GetVar(App.Path & "\Quests.ini", "QUEST4", "Agility"))
Call SetPlayerMAGI(index, GetPlayerMAGI(index) + GetVar(App.Path & "\Quests.ini", "QUEST4", "Wisdom"))
End If

If Found = 0 Then
Call PlayerMsg(index, GetVar(App.Path & "\Quests.ini", "QUEST4", "NPCname") & ": " & GetVar(App.Path & "\Quests.ini", "QUEST4", "NoItemNPCsay"), 0)
End If
End If

If GetVar(App.Path & "\Quests.ini", "QUEST4", "IfItemNr") = 0 Then
Call PlayerMsg(index, GetVar(App.Path & "\Quests.ini", "QUEST4", "NPCname") & ": " & GetVar(App.Path & "\Quests.ini", "QUEST4", "FinishedNPCsay"), 14)
Call SetPlayerPOINTS(index, GetPlayerPOINTS(index) + GetVar(App.Path & "\Quests.ini", "QUEST4", "StatPoints"))
Call SetPlayerExp(index, GetPlayerExp(index) + GetVar(App.Path & "\Quests.ini", "QUEST4", "Exp"))
Call GiveItemQ(index, GetVar(App.Path & "\Quests.ini", "QUEST4", "ItemNr"), 1)
Call PutVar(App.Path & "\Scripts\Part.ini", GetPlayerName(index), "Part", "5")
If GetVar(App.Path & "\Quests.ini", "QUEST4", "FullHP") = 1 Then
Call SetPlayerHP(index, GetPlayerMaxHP(index))
Call SendHP(index)
End If
If GetVar(App.Path & "\Quests.ini", "QUEST4", "FullMP") = 1 Then
Call SetPlayerMP(index, GetPlayerMaxMP(index))
Call SendMP(index)
End If
If GetVar(App.Path & "\Quests.ini", "QUEST4", "FullSP") = 1 Then
Call SetPlayerSP(index, GetPlayerMaxSP(index))
Call SendSP(index)
End If
Call SetPlayerSTR(index, GetPlayerSTR(index) + GetVar(App.Path & "\Quests.ini", "QUEST4", "Strength"))
Call SetPlayerDEF(index, GetPlayerDEF(index) + GetVar(App.Path & "\Quests.ini", "QUEST4", "Defence"))
Call SetPlayerSPEED(index, GetPlayerSPEED(index) + GetVar(App.Path & "\Quests.ini", "QUEST4", "Agility"))
Call SetPlayerMAGI(index, GetPlayerMAGI(index) + GetVar(App.Path & "\Quests.ini", "QUEST4", "Wisdom"))
End If
End If
End If
If GetVar(App.Path & "\Scripts\Part.ini", GetPlayerName(index), "Part") = 4 Then
Else
Call PlayerMsg(index, GetVar(App.Path & "\Quests.ini", "QUEST1", "NPCname") & ": " & GetVar(App.Path & "\Quests.ini", "QUEST4", "WrongQuestNPCsay"), 0)
End If
End Sub
Sub QuestFive(ByVal index As Long)
Dim Found As Long
Dim Slot As Long
If GetVar(App.Path & "\Scripts\Part.ini", GetPlayerName(index), "Part") = 5 Then
Call PlayerMsg(index, GetVar(App.Path & "\Quests.ini", "QUEST5", "NPCname") & ": " & GetVar(App.Path & "\Quests.ini", "QUEST5", "NPCsay"), 0)
If GetPlayerLevel(index) > GetVar(App.Path & "\Quests.ini", "QUEST5", "LevelReq") Then
If GetVar(App.Path & "\Quests.ini", "QUEST5", "IfItemNr") > 0 Then

Found = 0
Slot = 1
Do While Slot < 24
If GetPlayerInvItemNum(index, Slot) = GetVar(App.Path & "\Quests.ini", "QUEST5", "IfItemNr") Then
Found = 1
Slot = 24
End If
Slot = Slot + 1
Loop

If Found = 1 Then
If GetVar(App.Path & "\Quests.ini", "QUEST5", "Keep") = 1 Then
Call TakeItemQ(index, GetVar(App.Path & "\Quests.ini", "QUEST5", "IfItemNr"))
Call PlayerMsg(index, GetVar(App.Path & "\Quests.ini", "QUEST5", "NPCname") & ": " & GetVar(App.Path & "\Quests.ini", "QUEST5", "FinishedNPCsay"), 14)
Call SetPlayerPOINTS(index, GetPlayerPOINTS(index) + GetVar(App.Path & "\Quests.ini", "QUEST5", "StatPoints"))
Call SetPlayerExp(index, GetPlayerExp(index) + GetVar(App.Path & "\Quests.ini", "QUEST5", "Exp"))
Call GiveItemQ(index, GetVar(App.Path & "\Quests.ini", "QUEST5", "ItemNr"), 1)
Else
Call PlayerMsg(index, GetVar(App.Path & "\Quests.ini", "QUEST5", "NPCname") & ": " & GetVar(App.Path & "\Quests.ini", "QUEST5", "FinishedNPCsay"), 14)
End If
If GetVar(App.Path & "\Quests.ini", "QUEST5", "FullHP") = 1 Then
Call SetPlayerHP(index, GetPlayerMaxHP(index))
Call SendHP(index)
End If
If GetVar(App.Path & "\Quests.ini", "QUEST5", "FullMP") = 1 Then
Call SetPlayerMP(index, GetPlayerMaxMP(index))
Call SendMP(index)
End If
If GetVar(App.Path & "\Quests.ini", "QUEST5", "FullSP") = 1 Then
Call SetPlayerSP(index, GetPlayerMaxSP(index))
Call SendSP(index)
End If
Call PutVar(App.Path & "\Scripts\Part.ini", GetPlayerName(index), "Part", "6")
Call SetPlayerSprite(index, GetVar(App.Path & "\Quests.ini", "QUEST5", "Sprite"))
Call SetPlayerSTR(index, GetPlayerSTR(index) + GetVar(App.Path & "\Quests.ini", "QUEST5", "Strength"))
Call SetPlayerDEF(index, GetPlayerDEF(index) + GetVar(App.Path & "\Quests.ini", "QUEST5", "Defence"))
Call SetPlayerSPEED(index, GetPlayerSPEED(index) + GetVar(App.Path & "\Quests.ini", "QUEST5", "Agility"))
Call SetPlayerMAGI(index, GetPlayerMAGI(index) + GetVar(App.Path & "\Quests.ini", "QUEST5", "Wisdom"))
End If

If Found = 0 Then
Call PlayerMsg(index, GetVar(App.Path & "\Quests.ini", "QUEST5", "NPCname") & ": " & GetVar(App.Path & "\Quests.ini", "QUEST5", "NoItemNPCsay"), 0)
End If
End If

If GetVar(App.Path & "\Quests.ini", "QUEST5", "IfItemNr") = 0 Then
Call PlayerMsg(index, GetVar(App.Path & "\Quests.ini", "QUEST5", "NPCname") & ": " & GetVar(App.Path & "\Quests.ini", "QUEST5", "FinishedNPCsay"), 14)
Call SetPlayerPOINTS(index, GetPlayerPOINTS(index) + GetVar(App.Path & "\Quests.ini", "QUEST5", "StatPoints"))
Call SetPlayerExp(index, GetPlayerExp(index) + GetVar(App.Path & "\Quests.ini", "QUEST5", "Exp"))
Call GiveItemQ(index, GetVar(App.Path & "\Quests.ini", "QUEST5", "ItemNr"), 1)
Call PutVar(App.Path & "\Scripts\Part.ini", GetPlayerName(index), "Part", "6")
If GetVar(App.Path & "\Quests.ini", "QUEST5", "FullHP") = 1 Then
Call SetPlayerHP(index, GetPlayerMaxHP(index))
Call SendHP(index)
End If
If GetVar(App.Path & "\Quests.ini", "QUEST5", "FullMP") = 1 Then
Call SetPlayerMP(index, GetPlayerMaxMP(index))
Call SendMP(index)
End If
If GetVar(App.Path & "\Quests.ini", "QUEST5", "FullSP") = 1 Then
Call SetPlayerSP(index, GetPlayerMaxSP(index))
Call SendSP(index)
End If
Call SetPlayerSTR(index, GetPlayerSTR(index) + GetVar(App.Path & "\Quests.ini", "QUEST5", "Strength"))
Call SetPlayerDEF(index, GetPlayerDEF(index) + GetVar(App.Path & "\Quests.ini", "QUEST5", "Defence"))
Call SetPlayerSPEED(index, GetPlayerSPEED(index) + GetVar(App.Path & "\Quests.ini", "QUEST5", "Agility"))
Call SetPlayerMAGI(index, GetPlayerMAGI(index) + GetVar(App.Path & "\Quests.ini", "QUEST5", "Wisdom"))
End If
End If
End If
If GetVar(App.Path & "\Scripts\Part.ini", GetPlayerName(index), "Part") = 5 Then
Else
Call PlayerMsg(index, GetVar(App.Path & "\Quests.ini", "QUEST5", "NPCname") & ": " & GetVar(App.Path & "\Quests.ini", "QUEST5", "WrongQuestNPCsay"), 0)
End If
End Sub
Sub QuestSix(ByVal index As Long)
Dim Found As Long
Dim Slot As Long
If GetVar(App.Path & "\Scripts\Part.ini", GetPlayerName(index), "Part") = 6 Then
Call PlayerMsg(index, GetVar(App.Path & "\Quests.ini", "QUEST6", "NPCname") & ": " & GetVar(App.Path & "\Quests.ini", "QUEST6", "NPCsay"), 0)
If GetPlayerLevel(index) > GetVar(App.Path & "\Quests.ini", "QUEST6", "LevelReq") Then
If GetVar(App.Path & "\Quests.ini", "QUEST6", "IfItemNr") > 0 Then

Found = 0
Slot = 1
Do While Slot < 24
If GetPlayerInvItemNum(index, Slot) = GetVar(App.Path & "\Quests.ini", "QUEST6", "IfItemNr") Then
Found = 1
Slot = 24
End If
Slot = Slot + 1
Loop

If Found = 1 Then
If GetVar(App.Path & "\Quests.ini", "QUEST6", "Keep") = 1 Then
Call TakeItemQ(index, GetVar(App.Path & "\Quests.ini", "QUEST6", "IfItemNr"))
Call PlayerMsg(index, GetVar(App.Path & "\Quests.ini", "QUEST6", "NPCname") & ": " & GetVar(App.Path & "\Quests.ini", "QUEST6", "FinishedNPCsay"), 14)
Call SetPlayerPOINTS(index, GetPlayerPOINTS(index) + GetVar(App.Path & "\Quests.ini", "QUEST6", "StatPoints"))
Call SetPlayerExp(index, GetPlayerExp(index) + GetVar(App.Path & "\Quests.ini", "QUEST6", "Exp"))
Call GiveItemQ(index, GetVar(App.Path & "\Quests.ini", "QUEST6", "ItemNr"), 1)
Else
Call PlayerMsg(index, GetVar(App.Path & "\Quests.ini", "QUEST6", "NPCname") & ": " & GetVar(App.Path & "\Quests.ini", "QUEST6", "FinishedNPCsay"), 14)
End If
If GetVar(App.Path & "\Quests.ini", "QUEST6", "FullHP") = 1 Then
Call SetPlayerHP(index, GetPlayerMaxHP(index))
Call SendHP(index)
End If
If GetVar(App.Path & "\Quests.ini", "QUEST6", "FullMP") = 1 Then
Call SetPlayerMP(index, GetPlayerMaxMP(index))
Call SendMP(index)
End If
If GetVar(App.Path & "\Quests.ini", "QUEST6", "FullSP") = 1 Then
Call SetPlayerSP(index, GetPlayerMaxSP(index))
Call SendSP(index)
End If
Call PutVar(App.Path & "\Scripts\Part.ini", GetPlayerName(index), "Part", "7")
Call SetPlayerSprite(index, GetVar(App.Path & "\Quests.ini", "QUEST6", "Sprite"))
Call SetPlayerSTR(index, GetPlayerSTR(index) + GetVar(App.Path & "\Quests.ini", "QUEST6", "Strength"))
Call SetPlayerDEF(index, GetPlayerDEF(index) + GetVar(App.Path & "\Quests.ini", "QUEST6", "Defence"))
Call SetPlayerSPEED(index, GetPlayerSPEED(index) + GetVar(App.Path & "\Quests.ini", "QUEST6", "Agility"))
Call SetPlayerMAGI(index, GetPlayerMAGI(index) + GetVar(App.Path & "\Quests.ini", "QUEST6", "Wisdom"))
End If

If Found = 0 Then
Call PlayerMsg(index, GetVar(App.Path & "\Quests.ini", "QUEST6", "NPCname") & ": " & GetVar(App.Path & "\Quests.ini", "QUEST6", "NoItemNPCsay"), 0)
End If
End If

If GetVar(App.Path & "\Quests.ini", "QUEST6", "IfItemNr") = 0 Then
Call PlayerMsg(index, GetVar(App.Path & "\Quests.ini", "QUEST6", "NPCname") & ": " & GetVar(App.Path & "\Quests.ini", "QUEST6", "FinishedNPCsay"), 14)
Call SetPlayerPOINTS(index, GetPlayerPOINTS(index) + GetVar(App.Path & "\Quests.ini", "QUEST6", "StatPoints"))
Call SetPlayerExp(index, GetPlayerExp(index) + GetVar(App.Path & "\Quests.ini", "QUEST6", "Exp"))
Call GiveItemQ(index, GetVar(App.Path & "\Quests.ini", "QUEST6", "ItemNr"), 1)
Call PutVar(App.Path & "\Scripts\Part.ini", GetPlayerName(index), "Part", "7")
If GetVar(App.Path & "\Quests.ini", "QUEST6", "FullHP") = 1 Then
Call SetPlayerHP(index, GetPlayerMaxHP(index))
Call SendHP(index)
End If
If GetVar(App.Path & "\Quests.ini", "QUEST6", "FullMP") = 1 Then
Call SetPlayerMP(index, GetPlayerMaxMP(index))
Call SendMP(index)
End If
If GetVar(App.Path & "\Quests.ini", "QUEST6", "FullSP") = 1 Then
Call SetPlayerSP(index, GetPlayerMaxSP(index))
Call SendSP(index)
End If
Call SetPlayerSTR(index, GetPlayerSTR(index) + GetVar(App.Path & "\Quests.ini", "QUEST6", "Strength"))
Call SetPlayerDEF(index, GetPlayerDEF(index) + GetVar(App.Path & "\Quests.ini", "QUEST6", "Defence"))
Call SetPlayerSPEED(index, GetPlayerSPEED(index) + GetVar(App.Path & "\Quests.ini", "QUEST6", "Agility"))
Call SetPlayerMAGI(index, GetPlayerMAGI(index) + GetVar(App.Path & "\Quests.ini", "QUEST6", "Wisdom"))
End If
End If
End If
If GetVar(App.Path & "\Scripts\Part.ini", GetPlayerName(index), "Part") = 6 Then
Else
Call PlayerMsg(index, GetVar(App.Path & "\Quests.ini", "QUEST6", "NPCname") & ": " & GetVar(App.Path & "\Quests.ini", "QUEST6", "WrongQuestNPCsay"), 0)
End If
End Sub

Sub QuestSeven(ByVal index As Long)
Dim Found As Long
Dim Slot As Long
If GetVar(App.Path & "\Scripts\Part.ini", GetPlayerName(index), "Part") = 7 Then
Call PlayerMsg(index, GetVar(App.Path & "\Quests.ini", "QUEST7", "NPCname") & ": " & GetVar(App.Path & "\Quests.ini", "QUEST7", "NPCsay"), 0)
If GetPlayerLevel(index) > GetVar(App.Path & "\Quests.ini", "QUEST7", "LevelReq") Then
If GetVar(App.Path & "\Quests.ini", "QUEST7", "IfItemNr") > 0 Then

Found = 0
Slot = 1
Do While Slot < 24
If GetPlayerInvItemNum(index, Slot) = GetVar(App.Path & "\Quests.ini", "QUEST7", "IfItemNr") Then
Found = 1
Slot = 24
End If
Slot = Slot + 1
Loop

If Found = 1 Then
If GetVar(App.Path & "\Quests.ini", "QUEST7", "Keep") = 1 Then
Call TakeItemQ(index, GetVar(App.Path & "\Quests.ini", "QUEST7", "IfItemNr"))
Call PlayerMsg(index, GetVar(App.Path & "\Quests.ini", "QUEST7", "NPCname") & ": " & GetVar(App.Path & "\Quests.ini", "QUEST7", "FinishedNPCsay"), 14)
Call SetPlayerPOINTS(index, GetPlayerPOINTS(index) + GetVar(App.Path & "\Quests.ini", "QUEST7", "StatPoints"))
Call SetPlayerExp(index, GetPlayerExp(index) + GetVar(App.Path & "\Quests.ini", "QUEST7", "Exp"))
Call GiveItemQ(index, GetVar(App.Path & "\Quests.ini", "QUEST7", "ItemNr"), 1)
Else
Call PlayerMsg(index, GetVar(App.Path & "\Quests.ini", "QUEST7", "NPCname") & ": " & GetVar(App.Path & "\Quests.ini", "QUEST7", "FinishedNPCsay"), 14)
End If
If GetVar(App.Path & "\Quests.ini", "QUEST7", "FullHP") = 1 Then
Call SetPlayerHP(index, GetPlayerMaxHP(index))
Call SendHP(index)
End If
If GetVar(App.Path & "\Quests.ini", "QUEST7", "FullMP") = 1 Then
Call SetPlayerMP(index, GetPlayerMaxMP(index))
Call SendMP(index)
End If
If GetVar(App.Path & "\Quests.ini", "QUEST7", "FullSP") = 1 Then
Call SetPlayerSP(index, GetPlayerMaxSP(index))
Call SendSP(index)
End If
Call PutVar(App.Path & "\Scripts\Part.ini", GetPlayerName(index), "Part", "8")
Call SetPlayerSprite(index, GetVar(App.Path & "\Quests.ini", "QUEST7", "Sprite"))
Call SetPlayerSTR(index, GetPlayerSTR(index) + GetVar(App.Path & "\Quests.ini", "QUEST7", "Strength"))
Call SetPlayerDEF(index, GetPlayerDEF(index) + GetVar(App.Path & "\Quests.ini", "QUEST7", "Defence"))
Call SetPlayerSPEED(index, GetPlayerSPEED(index) + GetVar(App.Path & "\Quests.ini", "QUEST7", "Agility"))
Call SetPlayerMAGI(index, GetPlayerMAGI(index) + GetVar(App.Path & "\Quests.ini", "QUEST7", "Wisdom"))
End If

If Found = 0 Then
Call PlayerMsg(index, GetVar(App.Path & "\Quests.ini", "QUEST7", "NPCname") & ": " & GetVar(App.Path & "\Quests.ini", "QUEST7", "NoItemNPCsay"), 0)
End If
End If

If GetVar(App.Path & "\Quests.ini", "QUEST7", "IfItemNr") = 0 Then
Call PlayerMsg(index, GetVar(App.Path & "\Quests.ini", "QUEST7", "NPCname") & ": " & GetVar(App.Path & "\Quests.ini", "QUEST7", "FinishedNPCsay"), 14)
Call SetPlayerPOINTS(index, GetPlayerPOINTS(index) + GetVar(App.Path & "\Quests.ini", "QUEST7", "StatPoints"))
Call SetPlayerExp(index, GetPlayerExp(index) + GetVar(App.Path & "\Quests.ini", "QUEST7", "Exp"))
Call GiveItemQ(index, GetVar(App.Path & "\Quests.ini", "QUEST7", "ItemNr"), 1)
Call PutVar(App.Path & "\Scripts\Part.ini", GetPlayerName(index), "Part", "8")
If GetVar(App.Path & "\Quests.ini", "QUEST7", "FullHP") = 1 Then
Call SetPlayerHP(index, GetPlayerMaxHP(index))
Call SendHP(index)
End If
If GetVar(App.Path & "\Quests.ini", "QUEST7", "FullMP") = 1 Then
Call SetPlayerMP(index, GetPlayerMaxMP(index))
Call SendMP(index)
End If
If GetVar(App.Path & "\Quests.ini", "QUEST7", "FullSP") = 1 Then
Call SetPlayerSP(index, GetPlayerMaxSP(index))
Call SendSP(index)
End If
Call SetPlayerSTR(index, GetPlayerSTR(index) + GetVar(App.Path & "\Quests.ini", "QUEST7", "Strength"))
Call SetPlayerDEF(index, GetPlayerDEF(index) + GetVar(App.Path & "\Quests.ini", "QUEST7", "Defence"))
Call SetPlayerSPEED(index, GetPlayerSPEED(index) + GetVar(App.Path & "\Quests.ini", "QUEST7", "Agility"))
Call SetPlayerMAGI(index, GetPlayerMAGI(index) + GetVar(App.Path & "\Quests.ini", "QUEST7", "Wisdom"))
End If
End If
End If
If GetVar(App.Path & "\Scripts\Part.ini", GetPlayerName(index), "Part") = 7 Then
Else
Call PlayerMsg(index, GetVar(App.Path & "\Quests.ini", "QUEST7", "NPCname") & ": " & GetVar(App.Path & "\Quests.ini", "QUEST7", "WrongQuestNPCsay"), 0)
End If
End Sub

Sub QuestEight(ByVal index As Long)
Dim Found As Long
Dim Slot As Long
If GetVar(App.Path & "\Scripts\Part.ini", GetPlayerName(index), "Part") = 8 Then
Call PlayerMsg(index, GetVar(App.Path & "\Quests.ini", "QUEST8", "NPCname") & ": " & GetVar(App.Path & "\Quests.ini", "QUEST8", "NPCsay"), 0)
If GetPlayerLevel(index) > GetVar(App.Path & "\Quests.ini", "QUEST8", "LevelReq") Then
If GetVar(App.Path & "\Quests.ini", "QUEST8", "IfItemNr") > 0 Then

Found = 0
Slot = 1
Do While Slot < 24
If GetPlayerInvItemNum(index, Slot) = GetVar(App.Path & "\Quests.ini", "QUEST8", "IfItemNr") Then
Found = 1
Slot = 24
End If
Slot = Slot + 1
Loop

If Found = 1 Then
If GetVar(App.Path & "\Quests.ini", "QUEST8", "Keep") = 1 Then
Call TakeItemQ(index, GetVar(App.Path & "\Quests.ini", "QUEST8", "IfItemNr"))
Call PlayerMsg(index, GetVar(App.Path & "\Quests.ini", "QUEST8", "NPCname") & ": " & GetVar(App.Path & "\Quests.ini", "QUEST8", "FinishedNPCsay"), 14)
Call SetPlayerPOINTS(index, GetPlayerPOINTS(index) + GetVar(App.Path & "\Quests.ini", "QUEST8", "StatPoints"))
Call SetPlayerExp(index, GetPlayerExp(index) + GetVar(App.Path & "\Quests.ini", "QUEST8", "Exp"))
Call GiveItemQ(index, GetVar(App.Path & "\Quests.ini", "QUEST8", "ItemNr"), 1)
Else
Call PlayerMsg(index, GetVar(App.Path & "\Quests.ini", "QUEST8", "NPCname") & ": " & GetVar(App.Path & "\Quests.ini", "QUEST8", "FinishedNPCsay"), 14)
End If
If GetVar(App.Path & "\Quests.ini", "QUEST8", "FullHP") = 1 Then
Call SetPlayerHP(index, GetPlayerMaxHP(index))
Call SendHP(index)
End If
If GetVar(App.Path & "\Quests.ini", "QUEST8", "FullMP") = 1 Then
Call SetPlayerMP(index, GetPlayerMaxMP(index))
Call SendMP(index)
End If
If GetVar(App.Path & "\Quests.ini", "QUEST8", "FullSP") = 1 Then
Call SetPlayerSP(index, GetPlayerMaxSP(index))
Call SendSP(index)
End If
Call PutVar(App.Path & "\Scripts\Part.ini", GetPlayerName(index), "Part", "9")
Call SetPlayerSprite(index, GetVar(App.Path & "\Quests.ini", "QUEST8", "Sprite"))
Call SetPlayerSTR(index, GetPlayerSTR(index) + GetVar(App.Path & "\Quests.ini", "QUEST8", "Strength"))
Call SetPlayerDEF(index, GetPlayerDEF(index) + GetVar(App.Path & "\Quests.ini", "QUEST8", "Defence"))
Call SetPlayerSPEED(index, GetPlayerSPEED(index) + GetVar(App.Path & "\Quests.ini", "QUEST8", "Agility"))
Call SetPlayerMAGI(index, GetPlayerMAGI(index) + GetVar(App.Path & "\Quests.ini", "QUEST8", "Wisdom"))
End If

If Found = 0 Then
Call PlayerMsg(index, GetVar(App.Path & "\Quests.ini", "QUEST8", "NPCname") & ": " & GetVar(App.Path & "\Quests.ini", "QUEST8", "NoItemNPCsay"), 0)
End If
End If

If GetVar(App.Path & "\Quests.ini", "QUEST8", "IfItemNr") = 0 Then
Call PlayerMsg(index, GetVar(App.Path & "\Quests.ini", "QUEST8", "NPCname") & ": " & GetVar(App.Path & "\Quests.ini", "QUEST8", "FinishedNPCsay"), 14)
Call SetPlayerPOINTS(index, GetPlayerPOINTS(index) + GetVar(App.Path & "\Quests.ini", "QUEST8", "StatPoints"))
Call SetPlayerExp(index, GetPlayerExp(index) + GetVar(App.Path & "\Quests.ini", "QUEST8", "Exp"))
Call GiveItemQ(index, GetVar(App.Path & "\Quests.ini", "QUEST8", "ItemNr"), 1)
Call PutVar(App.Path & "\Scripts\Part.ini", GetPlayerName(index), "Part", "9")
If GetVar(App.Path & "\Quests.ini", "QUEST8", "FullHP") = 1 Then
Call SetPlayerHP(index, GetPlayerMaxHP(index))
Call SendHP(index)
End If
If GetVar(App.Path & "\Quests.ini", "QUEST8", "FullMP") = 1 Then
Call SetPlayerMP(index, GetPlayerMaxMP(index))
Call SendMP(index)
End If
If GetVar(App.Path & "\Quests.ini", "QUEST8", "FullSP") = 1 Then
Call SetPlayerSP(index, GetPlayerMaxSP(index))
Call SendSP(index)
End If
Call SetPlayerSTR(index, GetPlayerSTR(index) + GetVar(App.Path & "\Quests.ini", "QUEST8", "Strength"))
Call SetPlayerDEF(index, GetPlayerDEF(index) + GetVar(App.Path & "\Quests.ini", "QUEST8", "Defence"))
Call SetPlayerSPEED(index, GetPlayerSPEED(index) + GetVar(App.Path & "\Quests.ini", "QUEST8", "Agility"))
Call SetPlayerMAGI(index, GetPlayerMAGI(index) + GetVar(App.Path & "\Quests.ini", "QUEST8", "Wisdom"))
End If
End If
End If
If GetVar(App.Path & "\Scripts\Part.ini", GetPlayerName(index), "Part") = 8 Then
Else
Call PlayerMsg(index, GetVar(App.Path & "\Quests.ini", "QUEST8", "NPCname") & ": " & GetVar(App.Path & "\Quests.ini", "QUEST8", "WrongQuestNPCsay"), 0)
End If
End Sub

Sub QuestNine(ByVal index As Long)
Dim Found As Long
Dim Slot As Long
If GetVar(App.Path & "\Scripts\Part.ini", GetPlayerName(index), "Part") = 9 Then
Call PlayerMsg(index, GetVar(App.Path & "\Quests.ini", "QUEST9", "NPCname") & ": " & GetVar(App.Path & "\Quests.ini", "QUEST9", "NPCsay"), 0)
If GetPlayerLevel(index) > GetVar(App.Path & "\Quests.ini", "QUEST9", "LevelReq") Then
If GetVar(App.Path & "\Quests.ini", "QUEST9", "IfItemNr") > 0 Then

Found = 0
Slot = 1
Do While Slot < 24
If GetPlayerInvItemNum(index, Slot) = GetVar(App.Path & "\Quests.ini", "QUEST9", "IfItemNr") Then
Found = 1
Slot = 24
End If
Slot = Slot + 1
Loop

If Found = 1 Then
If GetVar(App.Path & "\Quests.ini", "QUEST9", "Keep") = 1 Then
Call TakeItemQ(index, GetVar(App.Path & "\Quests.ini", "QUEST9", "IfItemNr"))
Call PlayerMsg(index, GetVar(App.Path & "\Quests.ini", "QUEST9", "NPCname") & ": " & GetVar(App.Path & "\Quests.ini", "QUEST9", "FinishedNPCsay"), 14)
Call SetPlayerPOINTS(index, GetPlayerPOINTS(index) + GetVar(App.Path & "\Quests.ini", "QUEST9", "StatPoints"))
Call SetPlayerExp(index, GetPlayerExp(index) + GetVar(App.Path & "\Quests.ini", "QUEST9", "Exp"))
Call GiveItemQ(index, GetVar(App.Path & "\Quests.ini", "QUEST9", "ItemNr"), 1)
Else
Call PlayerMsg(index, GetVar(App.Path & "\Quests.ini", "QUEST9", "NPCname") & ": " & GetVar(App.Path & "\Quests.ini", "QUEST9", "FinishedNPCsay"), 14)
End If
If GetVar(App.Path & "\Quests.ini", "QUEST9", "FullHP") = 1 Then
Call SetPlayerHP(index, GetPlayerMaxHP(index))
Call SendHP(index)
End If
If GetVar(App.Path & "\Quests.ini", "QUEST9", "FullMP") = 1 Then
Call SetPlayerMP(index, GetPlayerMaxMP(index))
Call SendMP(index)
End If
If GetVar(App.Path & "\Quests.ini", "QUEST9", "FullSP") = 1 Then
Call SetPlayerSP(index, GetPlayerMaxSP(index))
Call SendSP(index)
End If
Call PutVar(App.Path & "\Scripts\Part.ini", GetPlayerName(index), "Part", "10")
Call SetPlayerSprite(index, GetVar(App.Path & "\Quests.ini", "QUEST9", "Sprite"))
Call SetPlayerSTR(index, GetPlayerSTR(index) + GetVar(App.Path & "\Quests.ini", "QUEST9", "Strength"))
Call SetPlayerDEF(index, GetPlayerDEF(index) + GetVar(App.Path & "\Quests.ini", "QUEST9", "Defence"))
Call SetPlayerSPEED(index, GetPlayerSPEED(index) + GetVar(App.Path & "\Quests.ini", "QUEST9", "Agility"))
Call SetPlayerMAGI(index, GetPlayerMAGI(index) + GetVar(App.Path & "\Quests.ini", "QUEST9", "Wisdom"))
End If

If Found = 0 Then
Call PlayerMsg(index, GetVar(App.Path & "\Quests.ini", "QUEST9", "NPCname") & ": " & GetVar(App.Path & "\Quests.ini", "QUEST9", "NoItemNPCsay"), 0)
End If
End If

If GetVar(App.Path & "\Quests.ini", "QUEST9", "IfItemNr") = 0 Then
Call PlayerMsg(index, GetVar(App.Path & "\Quests.ini", "QUEST9", "NPCname") & ": " & GetVar(App.Path & "\Quests.ini", "QUEST9", "FinishedNPCsay"), 14)
Call SetPlayerPOINTS(index, GetPlayerPOINTS(index) + GetVar(App.Path & "\Quests.ini", "QUEST9", "StatPoints"))
Call SetPlayerExp(index, GetPlayerExp(index) + GetVar(App.Path & "\Quests.ini", "QUEST9", "Exp"))
Call GiveItemQ(index, GetVar(App.Path & "\Quests.ini", "QUEST9", "ItemNr"), 1)
Call PutVar(App.Path & "\Scripts\Part.ini", GetPlayerName(index), "Part", "10")
If GetVar(App.Path & "\Quests.ini", "QUEST9", "FullHP") = 1 Then
Call SetPlayerHP(index, GetPlayerMaxHP(index))
Call SendHP(index)
End If
If GetVar(App.Path & "\Quests.ini", "QUEST9", "FullMP") = 1 Then
Call SetPlayerMP(index, GetPlayerMaxMP(index))
Call SendMP(index)
End If
If GetVar(App.Path & "\Quests.ini", "QUEST9", "FullSP") = 1 Then
Call SetPlayerSP(index, GetPlayerMaxSP(index))
Call SendSP(index)
End If
Call SetPlayerSTR(index, GetPlayerSTR(index) + GetVar(App.Path & "\Quests.ini", "QUEST9", "Strength"))
Call SetPlayerDEF(index, GetPlayerDEF(index) + GetVar(App.Path & "\Quests.ini", "QUEST9", "Defence"))
Call SetPlayerSPEED(index, GetPlayerSPEED(index) + GetVar(App.Path & "\Quests.ini", "QUEST9", "Agility"))
Call SetPlayerMAGI(index, GetPlayerMAGI(index) + GetVar(App.Path & "\Quests.ini", "QUEST9", "Wisdom"))
End If
End If
End If
If GetVar(App.Path & "\Scripts\Part.ini", GetPlayerName(index), "Part") = 9 Then
Else
Call PlayerMsg(index, GetVar(App.Path & "\Quests.ini", "QUEST9", "NPCname") & ": " & GetVar(App.Path & "\Quests.ini", "QUEST9", "WrongQuestNPCsay"), 0)
End If
End Sub
