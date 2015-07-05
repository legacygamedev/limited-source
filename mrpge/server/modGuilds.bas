Attribute VB_Name = "modGuilds"
Option Explicit

' /createguild name_of_guild
Public Sub CreateGuild(ByVal index As Long, ByVal strGuildName As String)
Dim guildNumber As Long
    If HasItem(index, 2) >= COST_TO_CREATE_GUILD Then
        Call PlayerMsg(index, "Creating Guild... ", RGB_HelpColor)
        Call TakeItem(index, 2, COST_TO_CREATE_GUILD)
        guildNumber = getFreeGuild
        Guild(guildNumber).Founder = player(index).Char(player(index).CharNum).name
        Guild(guildNumber).name = strGuildName
        Guild(guildNumber).Member(1) = player(index).Char(player(index).CharNum).name
        Call SaveGuild(guildNumber)
        player(index).Char(player(index).CharNum).Guild = guildNumber
        player(index).Char(player(index).CharNum).GuildAccess = FOUNDER_OF_GUILD
        Call SavePlayer(index, False)
Call PlayerMsg(index, "A new guild has been formed by the name of " & strGuildName & "! You are the founder!", RGB_AlertColor)
    Else
        Call PlayerMsg(index, "You don't have enough shards. You require " & COST_TO_CREATE_GUILD & "!", RGB_AlertColor)

    End If

End Sub

Public Function findGuild(ByVal strGuildName As String) As Long
Dim i As Long
    For i = 1 To MAX_GUILDS Step 1
        If LCase(Guild(i).name) = LCase(strGuildName) Then
            findGuild = 1
            Exit Function
        End If
    Next i
    findGuild = 0
End Function

Private Function getFreeGuild() As Long
Dim i As Long
    For i = 1 To MAX_GUILDS Step 1
    'MsgBox Len(Guild(i).Name)
        If (Guild(i).name = Chr(0) And Guild(i).Founder = Chr(0)) Or (Guild(i).name = "" And Guild(i).Founder = "") Then
            getFreeGuild = i
            Exit Function
        End If
    Next i
End Function

Public Function getFreeGuildPosition(ByVal guildNo As Long) As Long
Dim i As Long
    For i = 1 To MAX_GUILD_MEMBERS Step 1
    'MsgBox Len(Guild(i).Name)
        If Guild(guildNo).Member(i) = "" Or Guild(guildNo).Member(i) = Chr(0) Then
            getFreeGuildPosition = i
            Exit Function
        End If
    Next i
    getFreeGuildPosition = 0
End Function

Public Sub guildInvite(ByVal index As Long, ByVal chrName As String, ByVal guildID As Long)
    Dim recPlayer As Long
    recPlayer = FindPlayer(chrName)
    If recPlayer > 0 And recPlayer <> index Then
        If player(index).Char(player(index).CharNum).GuildAccess >= LEADER_OF_GUILD And player(recPlayer).Char(player(recPlayer).CharNum).Guild <= 0 Then
            PlayerMsg FindPlayer(chrName), "You have been invited to join '" & Guild(guildID).name & "' guild", RGB_HelpColor
    
            Guild(guildID).InviteList = Guild(guildID).InviteList & "," & chrName
            Call SaveGuild(guildID)
            Call PlayerMsg(index, "Player invited", RGB_HelpColor)
        Else
            If player(recPlayer).Char(player(recPlayer).CharNum).Guild > 0 Then
                PlayerMsg index, "Player already in a guild.", RGB_AlertColor
            Else
                PlayerMsg index, "You have to be a leader or founder of the guild to invite people to the guild.", RGB_AlertColor
            End If
        End If
    Else
        If recPlayer <> index Then
            PlayerMsg index, chrName & " is offline.", RGB_AlertColor
        Else
            PlayerMsg index, "You can't invite yourself!", RGB_AlertColor
        End If
        
    End If
End Sub

Public Sub guildAccept(ByVal index As Long, guildID As Long)
Dim freePos As Long
Dim strSearch As String
Dim strFindVal As String

strSearch = Guild(guildID).InviteList
strFindVal = LCase(Trim(player(index).Char(player(index).CharNum).name))

    If InStr(1, strSearch, strFindVal) > 0 Then
        'has been invited.
        freePos = getFreeGuildPosition(guildID)
        If freePos > 0 Then
        'there is free space
            Guild(guildID).Member(freePos) = player(index).Char(player(index).CharNum).name
            player(index).Char(player(index).CharNum).Guild = guildID
            player(index).Char(player(index).CharNum).GuildAccess = MEMBER_OF_GUILD
        Else
            PlayerMsg index, "Sorry but this guild is full.", RGB_AlertColor
        End If
        'clear the invite list of the name
        Dim heystack As String
        Dim nedle As String
        heystack = Guild(guildID).InviteList
        nedle = "," & LCase(Trim(player(index).Char(player(index).CharNum).name))
        'Debug.Print heystack
        'Debug.Print nedle
        
        heystack = Replace(heystack, nedle, "")
        'Debug.Print heystack
        Guild(guildID).InviteList = heystack
        Call SaveGuild(guildID)
        Call SavePlayer(index, False)
    Else
        PlayerMsg index, "You have not been invited to join this guild.", RGB_AlertColor
    End If
    
End Sub

Public Sub guildRemovePlayer(index As Long, guildID As Long, playername As String)
    Dim recPlayer As Long
    Dim i As Long
    recPlayer = FindPlayer(playername)
    If player(index).Char(player(index).CharNum).GuildAccess >= LEADER_OF_GUILD And player(recPlayer).Char(player(recPlayer).CharNum).Guild = player(index).Char(player(index).CharNum).Guild Then
    'all varified
    For i = 1 To MAX_GUILD_MEMBERS Step 1
        If Trim(LCase(Guild(guildID).Member(i))) = Trim(LCase(playername)) Then
            Guild(guildID).Member(i) = ""
            Exit For
        End If
    Next i
    
    player(recPlayer).Char(player(recPlayer).CharNum).Guild = 0
    player(recPlayer).Char(player(recPlayer).CharNum).GuildAccess = 0
    Call SaveGuild(guildID)
    Call SavePlayer(recPlayer, False)
        PlayerMsg index, "You have removed " & playername & " from the guild", RGB_AlertColor
        PlayerMsg recPlayer, "You been removed from the guild", RGB_AlertColor
    Else
        PlayerMsg index, "You don't have enough power to do that", RGB_AlertColor
    End If
End Sub

Public Sub guildPremotePlayer(index As Long, guildID As Long, playername As String)
    Dim recPlayer As Long
    Dim i As Long
    recPlayer = FindPlayer(playername)
    If recPlayer > 0 Then
        If player(index).Char(player(index).CharNum).GuildAccess >= FOUNDER_OF_GUILD And player(recPlayer).Char(player(recPlayer).CharNum).Guild = player(index).Char(player(index).CharNum).Guild Then
                player(recPlayer).Char(player(recPlayer).CharNum).GuildAccess = LEADER_OF_GUILD
                Call SavePlayer(recPlayer, False)
                PlayerMsg recPlayer, "You have been premoted to a leader of the guild. You can now add new members.", RGB_AlertColor
                PlayerMsg index, "You have promoted " & playername & ".", RGB_AlertColor
        Else
            PlayerMsg index, "You are not the founder.", RGB_AlertColor
        End If
    End If
End Sub

Public Function guildGetMemberNo(guildID)
Dim membercount As Long
    Dim i As Long
    membercount = 0
    For i = 1 To MAX_GUILD_MEMBERS
        If Guild(guildID).Member(i) <> "" And Guild(guildID).Member(i) <> Chr(0) Then
            membercount = membercount + 1
        End If
    Next i
    guildGetMemberNo = membercount
End Function

Public Sub guildDemotePlayer(index As Long, guildID As Long, playername As String)
    Dim recPlayer As Long
    Dim i As Long
    recPlayer = FindPlayer(playername)
    If recPlayer > 0 Then
        If player(index).Char(player(index).CharNum).GuildAccess >= FOUNDER_OF_GUILD And player(recPlayer).Char(player(recPlayer).CharNum).Guild = player(index).Char(player(index).CharNum).Guild Then
            player(recPlayer).Char(player(recPlayer).CharNum).GuildAccess = MEMBER_OF_GUILD
            Call SavePlayer(recPlayer, False)
            PlayerMsg recPlayer, "You have been demoted to a member of the guild. You can no longer add new members.", RGB_AlertColor
            PlayerMsg index, "You have demoted " & playername & ".", RGB_AlertColor
        Else
            PlayerMsg index, "You are not the founder.", RGB_AlertColor
        End If
    End If
End Sub

Public Sub guildDispand(ByVal index As Long)
    If player(index).Char(player(index).CharNum).Guild > 0 And player(index).Char(player(index).CharNum).GuildAccess >= FOUNDER_OF_GUILD Then
    If guildGetMemberNo(player(index).Char(player(index).CharNum).Guild) <= 1 Then
        Call removeGuild(player(index).Char(player(index).CharNum).Guild)
        PlayerMsg index, "You have just dispanded the guild.", RGB_AlertColor
    Else
        PlayerMsg index, "There are still players in the guild. Please remove them first", RGB_AlertColor
    End If
    Else
        PlayerMsg index, "You are not the founder.", RGB_AlertColor
    End If
End Sub
Private Sub removeGuild(guildID)
Dim i As Long
Dim strRemPlayer As Long
Dim filename As String
    For i = 1 To MAX_GUILD_MEMBERS Step 1
        If Guild(guildID).Member(i) <> "" And Guild(guildID).Member(i) <> Chr(0) Then
            filename = App.Path & "\accounts\" & Trim(Guild(guildID).Member(i)) & ".ini"
            Guild(guildID).Member(i) = Chr(0)
            Guild(guildID).Leaders(i) = Chr(0)
            'clear player files that are not online
            Call PutVar(filename, "CHAR" & i, "Guild", Chr(0))
            Call PutVar(filename, "CHAR" & i, "GuildAccess", Chr(0))
            strRemPlayer = FindPlayer(Guild(guildID).Member(i))
            If strRemPlayer > 0 Then
                player(strRemPlayer).Char(player(strRemPlayer).CharNum).Guild = Chr(0)
                player(strRemPlayer).Char(player(strRemPlayer).CharNum).GuildAccess = Chr(0)
                Call SavePlayer(strRemPlayer, False)
            End If
        End If
    Next i
    Guild(guildID).Description = Chr(0)
    Guild(guildID).Founder = Chr(0)
    Guild(guildID).InviteList = Chr(0)
    Guild(guildID).name = Chr(0)
    Call SaveGuild(guildID)
End Sub

Public Sub GuildInfo(index As Long)
Dim lngGuild As Long
Dim i As Long
    If player(index).Char(player(index).CharNum).Guild > 0 And player(index).Char(player(index).CharNum).GuildAccess >= MEMBER_OF_GUILD Then
    lngGuild = player(index).Char(player(index).CharNum).Guild
        Call PlayerMsg(index, ":: Guild info ::", RGB_HelpColor)
        Call PlayerMsg(index, "Name: " & Guild(lngGuild).name, RGB_HelpColor)
        Call PlayerMsg(index, "Description: " & Guild(lngGuild).Description, RGB_HelpColor)
        Call PlayerMsg(index, "Founder: " & Guild(lngGuild).Founder, RGB_AlertColor)
        Call PlayerMsg(index, "Members: ", RGB_HelpColor)
        For i = 1 To MAX_GUILD_MEMBERS
            If Guild(lngGuild).Member(i) <> "" And Guild(lngGuild).Member(i) <> Chr(0) Then
                Call PlayerMsg(index, Guild(lngGuild).Member(i), RGB_HelpColor)
            End If
        Next i
        Call PlayerMsg(index, "Leaders: ", RGB_HelpColor)
        For i = 1 To MAX_GUILD_MEMBERS
            If Guild(lngGuild).Member(i) <> "" And Guild(lngGuild).Member(i) <> Chr(0) Then
                Call PlayerMsg(index, Guild(lngGuild).Leaders(i), RGB_HelpColor)
            End If
        Next i
        
    End If
End Sub

Public Sub guildeditDiscription(ByVal index As Long, ByVal newdesc As String)
    If player(index).Char(player(index).CharNum).Guild > 0 And player(index).Char(player(index).CharNum).GuildAccess >= FOUNDER_OF_GUILD Then
        Guild(player(index).Char(player(index).CharNum).Guild).Description = newdesc
        Call PlayerMsg(index, "Guild description updated.", RGB_HelpColor)
    Else
        Call PlayerMsg(index, "You are not powerfull enough!", RGB_AlertColor)
    End If
End Sub

Public Function getPlayerGuildID(ByVal index As Long) As Long
    getPlayerGuildID = player(index).Char(player(index).CharNum).Guild
End Function

Public Sub sendGuildMsg(ByVal msg As String, ByVal guildID As Long)
Dim i As Long
    For i = 1 To MAX_GUILD_MEMBERS
        If Guild(guildID).Member(i) <> "" And Guild(guildID).Member(i) <> Chr(0) Then
            Call GuildMsg(FindPlayer(Guild(guildID).Member(i)), Guild(guildID).Member(i) & ": " & msg, RGB_GuildSay)
        End If
    Next i
End Sub
