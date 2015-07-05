Attribute VB_Name = "modParty"
Option Explicit

Public Const MAX_PLAYER_PARTY As Byte = 5
Public Const MAX_PARTY As Long = MAX_PLAYERS

Public Type PartyRec
    PartyPlayers(1 To MAX_PLAYER_PARTY) As String
    PartyCount As Long
    Used As Boolean
    HighLevel As Long
End Type
Public Party(1 To MAX_PARTY) As PartyRec

Public Sub Party_Clear_All()
Dim i As Long
    For i = 1 To MAX_PARTY
        Party_Clear i
    Next
End Sub

Public Sub Party_Clear(PartyIndex As Long)
    ZeroMemory ByVal VarPtr(Party(PartyIndex)), LenB(Party(PartyIndex))
End Sub

Public Function Party_Create(ByVal Index As Long) As Boolean
Dim i As Long

    For i = 1 To MAX_PARTY
        ' Check if the party is being used
        If Not Party(i).Used Then
            ' Set the first person to the user - First person is leader
            Party(i).PartyPlayers(1) = Current_Name(Index)
            Party(i).PartyCount = 1
            Party(i).Used = True        ' Need to make sure we set the party to being used
            Party(i).HighLevel = Current_Level(Index)
            
            ' Set the players party info
            Player(Index).InParty = True
            Player(Index).PartyIndex = i
            Party_Create = True
            Exit Function
        End If
    Next
    Party_Create = False
End Function

Public Sub Party_Invite(ByVal Index As Long, ByVal Invitee As Long)
Dim PartyIndex As Long

    ' Can't invite self
    If Index = Invitee Then
        SendPlayerMsg Index, "You can not invite yourself.", ActionColor
        Exit Sub
    End If
    
    ' Check if you're in a party
    If Not Player(Index).InParty Then
        ' Create the party
        If Not Party_Create(Index) Then
            SendPlayerMsg Index, "Could not start a party at this time.", ActionColor
            Exit Sub
        End If
    End If
    
    ' Get the players party index, so we know which party to deal with
    PartyIndex = Player(Index).PartyIndex
    
    ' Check if you are party leader
    If Party(PartyIndex).PartyPlayers(1) <> Current_Name(Index) Then
        SendPlayerMsg Index, "Only party leaders can invite people.", ActionColor
        Exit Sub
    End If
    
    ' Check for an open slot
    If Party(PartyIndex).PartyCount = MAX_PLAYER_PARTY Then
        SendPlayerMsg Index, "Party is full.", ActionColor
        Exit Sub
    End If
    
    ' Check if other player is in party
    If Player(Invitee).InParty Then
        SendPlayerMsg Index, Current_Name(Invitee) + " is currently in a party.", ActionColor
        Exit Sub
    End If
    
    ' Check if other person is already invited to a party
    If Player(Invitee).PartyInvitedBy <> vbNullString Then
        If Player(Invitee).PartyInvitedBy = Current_Name(Index) Then
            SendPlayerMsg Index, "You already invited this player.", ActionColor
        Else
            SendPlayerMsg Index, Current_Name(Invitee) + " has already been invited to a party.", ActionColor
        End If
        Exit Sub
    End If
    
    ' Set the invitees party index
    Player(Invitee).PartyInvitedBy = Current_Name(Index)
    Player(Invitee).PartyIndex = Player(Index).PartyIndex
    
    SendPlayerMsg Index, "You have invited " + Current_Name(Invitee) + " to your party.", ActionColor
    SendPlayerMsg Invitee, Current_Name(Index) + " has invited you to a party.", ActionColor
End Sub

Public Sub Party_SetHighLevel(ByVal PartyIndex As Long)
Dim i As Long, n As Long

    Party(PartyIndex).HighLevel = 0
    ' Find the highest level
    For i = 1 To MAX_PLAYER_PARTY
        If Party(PartyIndex).PartyPlayers(i) <> vbNullString Then
            n = FindPlayer(Party(PartyIndex).PartyPlayers(i))
            If n > 0 Then
                If Current_Level(n) > Party(PartyIndex).HighLevel Then
                    Party(PartyIndex).HighLevel = Current_Level(n)
                End If
            End If
        End If
    Next
End Sub

Public Sub Party_Join(ByVal Index As Long)
Dim PartyIndex As Long
Dim i As Long

    ' Check if you're in a party
    If Player(Index).InParty Then
        SendPlayerMsg Index, "You are currently in a party.", ActionColor
        Exit Sub
    End If
    
    ' Get the players party index, so we know which party to deal with
    PartyIndex = Player(Index).PartyIndex
    
    ' Check if you were invited
    If PartyIndex = 0 Then
        SendPlayerMsg Index, "You have not been invited to a party.", ActionColor
        Exit Sub
    End If
    
    ' Check if the party leader is different then the one who invited you
    If Party(PartyIndex).PartyPlayers(1) <> Player(Index).PartyInvitedBy Then
        SendPlayerMsg Index, "Party error.", ActionColor
        Player(Index).PartyIndex = 0
        Player(Index).PartyInvitedBy = vbNullString
        Exit Sub
    End If
    
    ' Check if somehow the party got filled up
    If Party(PartyIndex).PartyCount = MAX_PLAYER_PARTY Then
        SendPlayerMsg Index, "Party is now full.", ActionColor
        Player(Index).PartyIndex = 0
        Player(Index).PartyInvitedBy = vbNullString
        Exit Sub
    End If
    
    ' Add to party
    ' Find the first open slot
    For i = 1 To MAX_PLAYER_PARTY
        ' Find an empty slot
        If Party(PartyIndex).PartyPlayers(i) = vbNullString Then
            '
            Party(PartyIndex).PartyPlayers(i) = Current_Name(Index)
            Party(PartyIndex).PartyCount = Party(PartyIndex).PartyCount + 1
            
            ' Set the high level
            Party_SetHighLevel PartyIndex
            
            Player(Index).InParty = True
            Player(Index).PartyInvitedBy = vbNullString
            
            SendPartyMsg PartyIndex, Current_Name(Index) + " joined the party.", BrightBlue
            Exit Sub
        End If
    Next
    
    ' If we get here, somehow there was an error - so tell them
    SendPlayerMsg Index, "Party error.", ActionColor
    Player(Index).PartyIndex = 0
    Player(Index).PartyInvitedBy = vbNullString
End Sub

Public Sub Party_Decline(ByVal Index As Long)

    ' Check if in party
    If Player(Index).InParty Then
        SendPlayerMsg Index, "You are currently in a party.", ActionColor
        Exit Sub
    End If
    
    ' Check if had party invite
    If Player(Index).PartyIndex = 0 Then
        SendPlayerMsg Index, "You were not invited to a party.", ActionColor
        Exit Sub
    End If
    
    Dim n As Long
    n = FindPlayer(Player(Index).PartyInvitedBy)
    If n > 0 Then
        SendPlayerMsg n, Current_Name(Index) + " has declined your invitation.", ActionColor
    End If
    
    SendPlayerMsg Index, "You have declined the party invition.", ActionColor
    Player(Index).PartyIndex = 0
    Player(Index).PartyInvitedBy = vbNullString
End Sub

Public Sub Party_Quit(ByVal Index As Long)
Dim PartyIndex As Long
Dim i As Long

    ' Check if in party
    If Not Player(Index).InParty Then
        SendPlayerMsg Index, "You are not in a party.", ActionColor
        Exit Sub
    End If
    
    PartyIndex = Player(Index).PartyIndex
    
    ' Check if you are the party leader
    If Party(PartyIndex).PartyPlayers(1) = Current_Name(Index) Then
        ' Clear all players out of party
        For i = 1 To MAX_PLAYER_PARTY
            If Party(PartyIndex).PartyPlayers(i) <> vbNullString Then
                Player(i).InParty = False
                Player(i).PartyIndex = 0
                Player(i).PartyInvitedBy = vbNullString
                SendPlayerMsg i, "The party has been disbanded.", BrightBlue
            End If
        Next
        Party_Clear PartyIndex
    Else
        For i = 1 To MAX_PLAYER_PARTY
            ' Find the player
            If Party(PartyIndex).PartyPlayers(i) = Current_Name(Index) Then
                ' Clear this player out
                Party(PartyIndex).PartyPlayers(i) = vbNullString
                Party(PartyIndex).PartyCount = Party(PartyIndex).PartyCount - 1
                
                ' Set the high level
                Party_SetHighLevel PartyIndex
            
                Player(Index).InParty = False
                Player(Index).PartyIndex = 0
                Player(Index).PartyInvitedBy = vbNullString
                
                SendPartyMsg PartyIndex, Current_Name(Index) + " has left the party.", BrightBlue
                Exit Sub
            End If
        Next
    End If
End Sub
