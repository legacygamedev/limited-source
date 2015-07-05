Attribute VB_Name = "modParty"
Option Explicit
Public Const OFFSET_4 = 4294967296#
Public Const MAXINT_4 = 2147483647
Public Const OFFSET_2 = 65536
Public Const MAXINT_2 = 32767
Sub SetPlayerLFG(ByVal Index As Long, ByVal Looking As Boolean)
Player(Index).Char(Player(Index).CharNum).LookingForParty = Looking
End Sub

Function GetPlayerLFG(ByVal Index As Long) As Byte
GetPlayerLFG = Player(Index).Char(Player(Index).CharNum).LookingForParty
End Function

Sub SetPlayerParty(ByVal Index As Long, ByVal PartyNum As Long)
Dim I As Long

If GetPlayerParty(Index) > 0 Then
Call PlayerMsg(Index, "You cannot join another party since you're already in a party. To leave, simply type /leave.", Yellow)
Exit Sub
End If

If PartyNum = 0 Then
Call PlayerMsg(Index, "You cannot join this party since it no longer exists.", Red)
Exit Sub
End If

Player(Index).Char(Player(Index).CharNum).Party = PartyNum
Player(Index).Char(Player(Index).CharNum).InParty = 1

I = FindBlankPartySlot(PartyNum)
' Let's add the player to the most available slot.
If I > 0 Then
Party(PartyNum).Member(I) = Index
'Call PartyLeaveMsg(Index, PartyNum, GetPlayerName(Index) & " has joined this party.", Yellow)
If Not Party(PartyNum).Leader = Index Then Call PlayerMsg(Index, "You have joined this party.", Yellow)
Call SendParty(PartyNum)
Call UpdatePartyInv(PartyNum)
Else
'Call PartyMsg(PartyNum, GetPlayerName(Index) & " was not able to join this party since it's full.", Red)
Call PlayerMsg(Index, "You cannot join this party since it's full.", Red)
End If
End Sub

Function GetPlayerInParty(ByVal Index As Long) As Long
GetPlayerInParty = Player(Index).Char(Player(Index).CharNum).InParty
End Function

Function GetPlayerParty(ByVal Index As Long) As Long
GetPlayerParty = Player(Index).Char(Player(Index).CharNum).Party
End Function

Sub DeleteParty(ByVal PartyNum As Long)
Dim x As Byte
Dim y As Byte
' Clear out the items.
For x = 1 To MAX_PARTY_INV_SLOTS
For y = 1 To MAX_PARTY_MEMBERS
If Party(PartyNum).Member(y) > 0 Then Call RemovePartyMember(Party(PartyNum).Member(y), PartyNum)
Next y
Next x
' Party is no longer created.
Party(PartyNum).Created = False
Party(PartyNum).TimeCreated = 0
Call UpdatePartyInv(PartyNum)
'Call PartyMsg(PartyNum, "You have been disbanded from your party.", Yellow)

' Clear out the members.
Call SetPartyLeader(0, PartyNum)
End Sub

Function FindBlankPartySlot(ByVal PartyNum As Long) As Byte
Dim I As Byte
FindBlankPartySlot = 0
For I = 1 To MAX_PARTY_MEMBERS
If Party(PartyNum).Member(I) = 0 Then
FindBlankPartySlot = I
Exit Function
End If
Next I
End Function

Sub RollOnItem(ByVal Index As Long, ByVal ItemNum As Byte)
Dim Roll As Long
Dim checker As Byte
Dim x As Byte

'If Party(GetPlayerParty(Index)).PartyItems(ItemNum).num = 0 Then
'Call PlayerMsg(Index, "You cannot roll on an item that doesn't exist.", Yellow)
'Exit Sub
'End If

' Check if they've already rolled on this item, if not, let's roll.
'If Party(GetPlayerParty(Index)).PartyItems(ItemNum).RollValue(Index) = 0 Then
'Roll = Int(Rnd * 997) + 2
'Party(GetPlayerParty(Index)).PartyItems(ItemNum).RollValue(Index) = Roll
'Call SendHighestRoll(GetPlayerParty(Index), ItemNum)
'Call PartyMsg(GetPlayerParty(Index), GetPlayerName(Index) & " has rolled " & Roll & " on " & Trim(Item(Party(GetPlayerParty(Index)).PartyItems(ItemNum).num).Name) & "!", 202, 244, 170)
'Else
' Alert them that they have already rolled on this item.
'Call PlayerMsg(Index, "You have already rolled on this item!", Yellow)
'Exit Sub
'End If

'Call CheckForFlush(GetPlayerParty(Index), ItemNum)
End Sub

Sub NoRollOnItem(ByVal Index As Long, ByVal ItemNum As Byte)

'If Party(GetPlayerParty(Index)).PartyItems(ItemNum).num = 0 Then
'Call PlayerMsg(Index, "You cannot pass rolling on an item that doesn't exist.", Yellow)
'Exit Sub
'End If

' Check if they've already rolled on this item, if not, let's roll.
'If Party(GetPlayerParty(Index)).PartyItems(ItemNum).RollValue(Index) = 0 Then
'Party(GetPlayerParty(Index)).PartyItems(ItemNum).RollValue(Index) = 1
'Call SendHighestRoll(GetPlayerParty(Index), ItemNum)
'Call PartyMsg(GetPlayerParty(Index), GetPlayerName(Index) & " has passed rolling on " & Trim(Item(Party(GetPlayerParty(Index)).PartyItems(ItemNum).num).Name) & "!", 202, 244, 170)
'Else
' Alert them that they have already rolled on this item.
'Call PlayerMsg(Index, "You have already rolled on this item!", Yellow)
'Exit Sub
'End If

'Call CheckForFlush(GetPlayerParty(Index), ItemNum)
End Sub

Sub FlushItem(ByVal HighestRoller As Long, ByVal ItemNum As Byte)
Dim checker As Byte, I As Byte
' Tell the party he received the item.
'Call PartyMsg(GetPlayerParty(HighestRoller), GetPlayerName(HighestRoller) & " has received the " & Trim(Item(Party(GetPlayerParty(HighestRoller)).PartyItems(ItemNum).num).Name) & "!", 202, 244, 170)
'Call EmptyPartyItemSlot(HighestRoller, ItemNum)
End Sub

Sub RemovePartyMember(ByVal Index As Long, ByVal PartyNum As Byte)
Dim I As Long, N As Long
Call RemoveFromParty(Index)

Player(Index).Char(Player(Index).CharNum).InParty = 0
Player(Index).Char(Player(Index).CharNum).Party = 0
Player(Index).Char(Player(Index).CharNum).PartyInvitedTo = 0
Player(Index).Char(Player(Index).CharNum).PartyInvitedToBy = ""

For I = 1 To MAX_PARTY_MEMBERS

If Party(PartyNum).Member(I) = Index Then
For N = 1 To MAX_PARTY_INV_SLOTS

Next N
Call OnPartyRemove(I, PartyNum)
For N = 1 To MAX_PARTY_INV_SLOTS

Next N
Exit Sub
End If
Next I
End Sub

Sub OnPartyRemove(ByVal PartyMember As Byte, ByVal PartyNum As Byte)
Party(PartyNum).Member(PartyMember) = 0
End Sub

Sub NewPartyLeaderNext(ByVal Index As Long, ByVal PartyNum As Byte)
Dim I As Long
' Make sure the new party leader isn't the old one.
For I = 1 To MAX_PARTY_MEMBERS
If Party(PartyNum).Member(I) > 0 Then
If Party(PartyNum).Member(I) <> Index And Player(Party(PartyNum).Member(I)).InGame = True Then
Call SetPartyLeader(Party(PartyNum).Member(I), PartyNum)
Call PartyMsg(PartyNum, GetPlayerName(Party(PartyNum).Leader) & " is now leading this party!", 0, 255, 0)
Exit Sub
End If
End If
Next I
End Sub

Function VerifyOtherPeople(ByVal Index As Long) As Boolean
VerifyOtherPeople = False
If GetPlayerParty(Index) = 0 Then Exit Function
Dim I As Byte
'MsgBox GetPlayerParty(Index)
For I = 1 To MAX_PARTY_MEMBERS
If Party(GetPlayerParty(Index)).Member(I) > 0 And Party(GetPlayerParty(Index)).Member(I) <> Index Then
VerifyOtherPeople = True
Exit Function
End If
Next I
End Function

Sub PartyMsg(ByVal PartyNum As Long, ByVal Msg As String, ByVal Red As Byte, ByVal Green As Byte, ByVal Blue As Byte)
Dim Packet As String
    Packet = "d" & SEP_CHAR & Msg & SEP_CHAR & Red & SEP_CHAR & Green & SEP_CHAR & Blue & SEP_CHAR & END_CHAR
    Call SendDataToParty(PartyNum, Packet)
End Sub

Sub SendDataToParty(ByVal PartyNum As Long, ByVal Data As String)
Dim I As Long
    For I = 1 To MAX_PLAYERS
    If IsPlaying(I) Then
        If GetPlayerParty(I) = PartyNum And IsPlaying(I) Then Call SendDataTo(I, Data)
    End If
    Next I
End Sub

Sub PartyLeaveMsg(ByVal Index As Long, ByVal PartyNum As Long, ByVal Msg As String, ByVal Red As Byte, ByVal Green As Byte, ByVal Blue As Byte)
If PartyNum = 0 Then Exit Sub
Dim Packet As String
    Packet = "d" & SEP_CHAR & Msg & SEP_CHAR & Red & SEP_CHAR & Green & SEP_CHAR & Blue & SEP_CHAR & END_CHAR
    Call SendDataToPartyBut(Index, PartyNum, Packet)
End Sub

Sub SendDataToPartyBut(ByVal Index As Long, ByVal PartyNum As Long, ByVal Data As String)
If PartyNum = 0 Then Exit Sub
Dim I As Long
    
    For I = 1 To MAX_PARTY_MEMBERS
    If Party(PartyNum).Member(I) > 0 Then
    If Party(PartyNum).Member(I) <> Index And Player(Party(PartyNum).Member(I)).InGame = True Then Call SendDataTo(I, Data)
    End If
    Next I
End Sub

Function GetBlankPartyItemSlot(ByVal PartyNum As Byte) As Long
Dim I As Long
GetBlankPartyItemSlot = 0
For I = 1 To MAX_PARTY_INV_SLOTS
'If Party(PartyNum).PartyItems(i).num = 0 Then
'GetBlankPartyItemSlot = i
Exit Function
'End If
Next I
End Function

Function EmptyPartyItemSlot(ByVal Receiver As Long, ByVal ItemSlot As Byte) As Long
' Make sure they have an open slot.
'If FindOpenInvSlot(Receiver, Party(GetPlayerParty(Receiver)).PartyItems(ItemSlot).num) Then
' Give the lucky receiver the item.
'Call GiveItem(Receiver, Party(GetPlayerParty(Receiver)).PartyItems(ItemSlot).num, Party(GetPlayerParty(Receiver)).PartyItems(ItemSlot).Value)
' Remove it from the party inventory.
'Call RemovePartyItem(GetPlayerParty(Receiver), ItemSlot)
'Else
' Inventory full? Let's alert the players.
'Call PartyMsg(GetPlayerParty(Receiver), "This item has been lost since " & GetPlayerName(Receiver) & " has no room left in their inventory!", Yellow)
'End If
End Function

Sub RemovePartyItem(ByVal PartyNum As Byte, ByVal ItemSlot As Byte)
' Remove the main item data.
Dim I As Long
'Party(PartyNum).PartyItems(ItemSlot).num = 0
'Party(PartyNum).PartyItems(ItemSlot).Value = 0
'Party(PartyNum).PartyItems(ItemSlot).RemainingTime = 0
For I = 1 To MAX_PARTY_MEMBERS
' Removed the rolled values.
'Party(PartyNum).PartyItems(ItemSlot).RollValue(i) = 0
Next I
'Call UpdatePartyInv(PartyNum)
End Sub


Sub InvitePlayerToParty(ByVal Inviter As Byte, ByVal Invite As Byte)
If IsPlaying(Invite) = False Then
Call PlayerMsg(Inviter, "This player is not online!", Yellow)
Exit Sub
End If

If GetPlayerParty(Inviter) = 0 Then
Call PlayerMsg(Inviter, "You are not in a party at the moment.", Yellow)
Exit Sub
End If

If GetPlayerParty(Invite) > 0 Then
Call PlayerMsg(Inviter, "This player is already in a party!", Yellow)
Exit Sub
End If

' Make sure it's the leader whose doing the invitation.

If Not Party(GetPlayerParty(Inviter)).Leader = Inviter Then
' Don't allow the inviter to invite this player.
Call PlayerMsg(Inviter, "Only the leader, " & GetPlayerName(Party(GetPlayerParty(Inviter)).Leader) & ", can invited people to join this party.", Yellow)
Exit Sub
End If

Player(Invite).Char(Player(Invite).CharNum).PartyInvitedTo = GetPlayerParty(Inviter)
Player(Invite).Char(Player(Invite).CharNum).PartyInvitedToBy = Trim$(GetPlayerName(Inviter))
Player(Invite).Char(Player(Invite).CharNum).InParty = 0
' Tell them about the offer to join this party.
Call PlayerMsg(Invite, GetPlayerName(Inviter) & " has invited you to join his party. Type /join to accept this offer, or /reject to decline this offer.", Yellow)
Call PlayerMsg(Inviter, GetPlayerName(Invite) & " has been invited to join this party!", Green)

End Sub


Function GetPlayerInvited(ByVal Index As Long) As Long
GetPlayerInvited = Player(Index).Char(Player(Index).CharNum).PartyInvitedTo
End Function

Function FindNewPartySlot() As Byte
Dim I As Long

FindNewPartySlot = 0
For I = 1 To MAX_PARTIES
' We found a party that's not created? Well, this is the one we'll use.
If Party(I).Created = False Then
FindNewPartySlot = I
Exit Function
End If
Next I
End Function

Function CountPartyMembers(ByVal PartyNum As Byte) As Long
Dim I As Long
Dim checker As Byte
CountPartyMembers = 0
For I = 1 To MAX_PARTY_MEMBERS
If Party(PartyNum).Member(I) > 0 Then checker = checker + 1
Next I
CountPartyMembers = checker
End Function

Function CountPartyGuildMembers(ByVal PartyNum As Byte) As Byte
Dim I As Long, checker As Byte
' This function will be for rewarding guild party bonuses later, I'm thinking 2% exp.
CountPartyGuildMembers = 0
For I = 1 To MAX_PARTY_MEMBERS
If I <> Party(PartyNum).Leader Then
If GetPlayerGuildID(Party(PartyNum).Leader) = GetPlayerGuildID(Party(PartyNum).Member(I)) Then checker = checker + 1
End If
Next I
CountPartyGuildMembers = checker
End Function

Sub CreateParty(ByVal Starter As Long)
Dim I As Long
I = 0
' Incase I forget something, this will check if these players are already in a party.
If GetPlayerParty(Starter) > 0 Then
Call PlayerMsg(Starter, "You are already in a party. To leave your current party type /leaveparty.", Yellow)
'Call PlayerMsg(Starter, GetPlayerName(Starter) & " is already in a party.", yellow)
Exit Sub
End If

I = FindNewPartySlot
Party(I).Created = True
Call SetPartyLeader(Starter, I)
Player(Starter).Char(Player(Starter).CharNum).Party = I
Player(Starter).Char(Player(Starter).CharNum).InParty = 1
'Party(I).TimeCreated = LongToUnsigned(timeGetTime)
Party(I).Member(1) = Starter
Call PlayerMsg(Starter, "Your party has been formed! To invite players, use the /invite command followed by the player name.", Yellow)
Call SendParty(GetPlayerParty(Starter))
End Sub


Sub AddItemToPool(ByVal PartyNum As Byte, ByVal ItemNum As Byte, ByVal ItemValue As Byte)
If ItemNum = 0 Then Exit Sub
Dim I As Byte
'i = GetBlankPartyItemSlot(PartyNum)
' Flush out an item first.
'If i = 0 Then Call FlushItemOnTime(PartyNum)

'i = GetBlankPartyItemSlot(PartyNum)
' Now that we have a blank slot, we can add it.
'Call PartyMsg(PartyNum, Trim(Item(ItemNum).Name) & " has been added to your party item pool!", 0, 255, 0)
'Party(PartyNum).PartyItems(i).num = ItemNum
'Party(PartyNum).PartyItems(i).Value = ItemValue
'Party(PartyNum).PartyItems(i).RemainingTime = LongToUnsigned(timeGetTime)
'Call UpdatePartyInv(PartyNum)
End Sub

Sub FlushItemOnTime(ByVal PartyNum As Byte)
Dim I As Long, x As Byte, FindOutHighest As Double
' We'll get the initial items remaining time first to make things easy.
'FindOutHighest = Party(PartyNum).PartyItems(5).RemainingTime
'For i = 1 To MAX_PARTY_INV_SLOTS
'If Party(PartyNum).PartyItems(i).RemainingTime <= FindOutHighest And Party(PartyNum).PartyItems(i).num > 0 Then
'FindOutHighest = Party(PartyNum).PartyItems(i).RemainingTime
'x = i
'End If
'Next i

'Call FlushItem(FindHighestRoller(PartyNum, x), x)
End Sub

Function FindHighestRoller(ByVal PartyNum As Byte, ByVal ItemNum As Byte) As Byte
Dim x As Byte, HighestValue As Long
FindHighestRoller = 0
' Get the highest roll on this item.
'For x = 1 To MAX_PARTY_MEMBERS
'If Party(PartyNum).PartyItems(ItemNum).RollValue(x) > HighestValue Then
'HighestValue = Party(PartyNum).PartyItems(ItemNum).RollValue(x)
'FindHighestRoller = x
'End If
'Next x

'If HighestValue = 0 Then
' There's no rolls on this item, thus we set it to the leader
'Call PartyMsg(PartyNum, "Everyone in this party has not rolled or has passed rolling on " & Trim(Item(Party(PartyNum).PartyItems(ItemNum).num).Name) & ", thus this item will now go to the party leader.", 202, 244, 170)

'FindHighestRoller = Party(PartyNum).Leader
'Exit Function
'ElseIf HighestValue = 1 Then
'Call PartyMsg(PartyNum, "Everyone in this party has passed rolling on " & Trim(Item(Party(PartyNum).PartyItems(ItemNum).num).Name) & ", thus this item will now go to the party leader.", 202, 244, 170)
'FindHighestRoller = Party(PartyNum).Leader
'Exit Function
'End If
End Function

Sub PartyRemoval(ByVal Index As Long, ByVal PartyNum As Long, ByVal PlayerName As String)
Dim I As Long

        ' If they aren't in a party, then we won't even bother trying to remove them.
        If PartyNum = 0 Then
        If Player(Index).InGame = True Then
        Call PlayerMsg(Index, "You aren't in a party at the moment.", Yellow)
        End If
        Exit Sub
        End If

        ' They are, so let's remove them.
        If Party(PartyNum).Leader = Index And VerifyOtherPeople(Index) = False Then
        Call DeleteParty(PartyNum)
        If Player(Index).InGame = True Then Call PlayerMsg(Index, "Your party has been disbanded.", Yellow)
        Exit Sub
        End If
            
            If Party(PartyNum).Leader = Index And VerifyOtherPeople(Index) = True Then
            Call NewPartyLeaderNext(Index, PartyNum)
            For I = 1 To MAX_PARTY_MEMBERS
            If Party(PartyNum).Member(I) = Index Then
            If Player(Index).InGame = True Then Call PlayerMsg(Index, "You have been disbanded from your party.", Yellow)
            Call RemovePartyMember(Index, PartyNum)
           ' Call PartyLeaveMsg(Index, PartyNum, PlayerName & " has been disbanded from this party.", Yellow)
            Call SendParty(PartyNum)
            Exit Sub
            End If
            Next I
            End If
            
            For I = 1 To MAX_PARTY_MEMBERS
            If Party(PartyNum).Member(I) = Index Then
            If Player(Index).InGame = True Then Call PlayerMsg(Index, "You have been disbanded from your party.", Yellow)
           ' Call PartyLeaveMsg(Index, PartyNum, PlayerName & " has been disbanded from this party.", Yellow)
            Call RemovePartyMember(Index, PartyNum)
            Call SendParty(PartyNum)
            Exit Sub
            End If
            Next I
End Sub

Sub SetPartyLeader(ByVal Index As Long, ByVal PartyNum As Long)
Party(PartyNum).Leader = Index
End Sub

Sub RollForFun(ByVal Index As Long)
Dim Roll As Long
If GetPlayerMap(Index) = 0 Or IsPlaying(Index) = False Then Exit Sub
Roll = Int(Rnd * 998) + 1
Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " has rolled a " & Roll & "!", Yellow)
End Sub

Sub SendParty(ByVal PartyNum As Long)
Dim I As Long
If PartyNum = 0 Then Exit Sub
For I = 1 To MAX_PARTY_MEMBERS
If Party(PartyNum).Member(I) > 0 Then Call SendDataToParty(PartyNum, "g8" & SEP_CHAR & I & SEP_CHAR & GetPlayerName(Party(PartyNum).Member(I)) & SEP_CHAR & GetPlayerSprite(Party(PartyNum).Member(I)) & SEP_CHAR & GetPlayerName(Party(PartyNum).Leader) & SEP_CHAR & GetPlayerLevel(Party(PartyNum).Member(I)) & SEP_CHAR & Party(PartyNum).Member(I) & SEP_CHAR & END_CHAR)
Next I
End Sub

Sub UpdatePartyInv(ByVal PartyNum As Long)
Dim Packet As String
Dim I As Long
For I = 1 To MAX_PARTY_INV_SLOTS
'Packet = "g9" & SEP_CHAR & i & SEP_CHAR & Party(PartyNum).PartyItems(i).num & SEP_CHAR & END_CHAR
'Call SendDataToParty(PartyNum, Packet)
Next I
End Sub

Sub CheckForFlush(ByVal PartyNum As Long, ByVal ItemNum As Long)
Dim x As Long
Dim checker As Long

'For x = 1 To MAX_PARTY_MEMBERS
'If Party(PartyNum).PartyItems(ItemNum).RollValue(x) > 0 Then checker = checker + 1
'Next x

' Reward the highest player.
'If checker = CountPartyMembers(PartyNum) Then
'Call FlushItem(FindHighestRoller(PartyNum, ItemNum), ItemNum)
'Exit Sub
'End If
End Sub

Sub RemoveFromParty(ByVal Index As Long)
Dim Packet As String
Packet = "b6" & SEP_CHAR & GetPlayerName(Index) & SEP_CHAR & END_CHAR
Call SendDataToParty(GetPlayerParty(Index), Packet)
End Sub

Function PartyMemberPlaying(ByVal PartyNum As Byte, ByVal Index As Long) As Long
If Index = 0 Or PartyNum = 0 Then Exit Function
If IsPlaying(Index) Then PartyMemberPlaying = 1
End Function

Public Function GetPlayerGuildID(ByVal Index As Long) As Long
    GetPlayerGuildID = Player(Index).Char(Player(Index).CharNum).Guild
End Function

Function LongToUnsigned(Value As Long) As Double
        If Value < 0 Then
          LongToUnsigned = Value + OFFSET_4
        Else
          LongToUnsigned = Value
        End If
End Function

Sub JoinsParty(ByVal Index As Long)
Dim N As Long
Dim o As Long
Dim I As Long

If Player(Index).Invited > 0 Then
                o = 0
                For I = 1 To MAX_PARTY_MEMBERS

                    If Party(Player(Index).Invited).Member(I) = 0 Then
                        If o = 0 Then o = I
                    End If
                Next

                If o <> 0 Then
                    Player(Index).PartyID = Player(Index).Invited
                    Player(Index).InParty = YES
                    Player(Index).Invited = 0
                    Party(Player(Index).PartyID).Member(o) = Index
                    For I = 1 To MAX_PARTY_MEMBERS

                        If Party(Player(Index).PartyID).Member(I) <> 0 And Party(Player(Index).PartyID).Member(I) <> Index Then
                            Call PlayerMsg(Party(Player(Index).PartyID).Member(I), GetPlayerName(Index) & " has joined your party!", Pink)
                        Call SendDataTo(Party(Player(Index).PartyID).Member(I), "partyinfo" & SEP_CHAR & GetPlayerName(I) & SEP_CHAR & GetPlayerLevel(I) & SEP_CHAR & END_CHAR)
                        End If
                    Next
                    Call PlayerMsg(Index, "You have joined the party!", Pink)
                Else
                    Call PlayerMsg(Index, "The party is full!", Pink)
                End If
            Else
                Call PlayerMsg(Index, "You have not been invited into a party!", Pink)
            End If
End Sub
Sub KillParty(ByVal Index As Long)
Dim N As Long
Dim o As Long
Dim I As Long

If Player(Index).PartyID > 0 Then
                Call PlayerMsg(Index, "You have left the party.", Pink)
                N = 0
                For I = 1 To MAX_PARTY_MEMBERS

                    If Party(Player(Index).PartyID).Member(I) = Index Then N = I
                Next
                For I = N To MAX_PARTY_MEMBERS - 1
                    Party(Player(Index).PartyID).Member(I) = Party(Player(Index).PartyID).Member(I + 1)
                Next
                Party(Player(Index).PartyID).Member(MAX_PARTY_MEMBERS) = 0
                N = 0
                For I = 1 To MAX_PARTY_MEMBERS

                    If Party(Player(Index).PartyID).Member(I) <> 0 And Party(Player(Index).PartyID).Member(I) <> Index Then
                        N = N + 1
                        Call PlayerMsg(Party(Player(Index).PartyID).Member(I), GetPlayerName(Index) & " has left the party.", Pink)
                    End If
                Next

                If N < 2 And Index <> 0 Then
                    Call PlayerMsg(Party(Player(Index).PartyID).Member(1), "The party has disbanded.", Pink)
                    Player(Party(Player(Index).PartyID).Member(1)).InParty = NO
                    Player(Party(Player(Index).PartyID).Member(1)).PartyID = 0
                    Party(Player(Index).PartyID).Member(1) = 0
                End If
                Player(Index).InParty = NO
                Player(Index).PartyID = 0
            Else

                If Player(Index).Invited <> 0 Then
                    For I = 1 To MAX_PARTY_MEMBERS

                        If Party(Player(Index).Invited).Member(I) <> 0 And Party(Player(Index).Invited).Member(I) <> Index Then Call PlayerMsg(Index, GetPlayerName(Index) & " has declined the invitation.", Pink)
                    Next
                    Player(Index).Invited = 0
                    Call PlayerMsg(Index, "You have declined the invitation.", Pink)
                Else
                    Call PlayerMsg(Index, "You have not been invited into a party!", Pink)
                End If
            End If
End Sub

Sub PartyChat(ByVal Index As Long)
Dim I As Long

End Sub
