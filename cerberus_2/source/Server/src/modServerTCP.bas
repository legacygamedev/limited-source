Attribute VB_Name = "modServerTCP"
'   This file is part of the Cerberus Engine 2nd Edition.
'
'    The Cerberus Engine 2nd Edition is free software; you can redistribute it
'    and/or modify it under the terms of the GNU General Public License as
'    published by the Free Software Foundation; either version 2 of the License,
'    or (at your option) any later version.
'
'    Cerberus 2nd Edition is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License
'    along with Cerberus 2nd Edition; if not, write to the Free Software
'    Foundation, Inc., 51 Franklin St, Fifth Floor, Boston, MA  02110-1301  USA

Option Explicit

Sub UpdateCaption()
    frmCServer.Caption = GAME_NAME & " - CServer - " & "IP: " & GameServer.LocalAddress & " - Port: " & STR(GameServer.LocalPort) & " - Players Online: " & TotalOnlinePlayers
    Exit Sub
End Sub


' ********************************
' ** Connection/Security checks **
' ********************************

Function IsPlaying(ByVal Index As Long) As Boolean
    If IsConnected(Index) And Player(Index).InGame = True Then
        IsPlaying = True
    Else
        IsPlaying = False
    End If
End Function

Function IsLoggedIn(ByVal Index As Long) As Boolean
    If IsConnected(Index) And Trim(Player(Index).Login) <> "" Then
        IsLoggedIn = True
    Else
        IsLoggedIn = False
    End If
End Function

Function IsBanned(ByVal IP As String) As Boolean
Dim FileName As String, fIP As String, fName As String
Dim f As Long

    IsBanned = False
    
    FileName = App.Path & "\logs\banlist.txt"
    
    ' Check if file exists
    If Not FileExist("logs\banlist.txt") Then
        f = FreeFile
        Open FileName For Output As #f
        Close #f
    End If
    
    f = FreeFile
    Open FileName For Input As #f
    
    Do While Not EOF(f)
        Input #f, fIP
        Input #f, fName
    
        ' Is banned?
        If Trim(LCase(fIP)) = Trim(LCase(Mid(IP, 1, Len(fIP)))) Then
            IsBanned = True
            Close #f
            Exit Function
        End If
    Loop
    
    Close #f
End Function

Sub HackingAttempt(ByVal Index As Long, ByVal Reason As String)
    If Index > 0 Then
        If IsPlaying(Index) Then
            Call AddLog(GetPlayerLogin(Index) & "/" & GetPlayerName(Index) & " has been booted for (" & Reason & ")", HACK_LOG)
            Call TextAdd(frmCServer.txtText, GetPlayerLogin(Index) & "/" & GetPlayerName(Index) & " has been booted for (" & Reason & ")", True)
        End If
    
        Call AlertMsg(Index, "You have lost your connection with " & GAME_NAME & ".")
    End If
End Sub

Function IsConnected(ByVal Index As Long) As Boolean
        If GameServer.Sockets.Item(Index).Socket Is Nothing Then
        IsConnected = False
    Else
        IsConnected = True
    End If
End Function

Sub AcceptConnection(Socket As JBSOCKETSERVERLib.ISocket)
Dim i As Long

    i = FindOpenPlayerSlot

    If i <> 0 Then

        Socket.UserData = i
        Set GameServer.Sockets.Item(CStr(i)).Socket = Socket
        Call SocketConnected(i)
        Socket.RequestRead
    Else
        Socket.Close
    End If
End Sub

Sub SocketConnected(ByVal Index As Long)
    If Index <> 0 Then
        '' Are they trying to connect more then one connection?
        ''If Not IsMultiIPOnline(GetPlayerIP(Index)) Then
            If Not IsBanned(GetPlayerIP(Index)) Then
                Call TextAdd(frmCServer.txtText, "Received connection from " & GetPlayerIP(Index) & ".", True)
            Else
                Call AlertMsg(Index, "You have been banned from " & GAME_NAME & ", and can no longer play.")
            End If
            ' Set The High Index
            Call SetHighIndex
            Call SendHighIndex
        ''Else
           '' Tried multiple connections
            ''Call AlertMsg(Index, GAME_NAME & " does not allow multiple IP's anymore.")
        ''End If
    End If
End Sub

Sub IncomingData(Socket As JBSOCKETSERVERLib.ISocket, Data As JBSOCKETSERVERLib.IData)
Dim Buffer As String
Dim dbytes() As Byte
Dim Packet As String
Dim Top As String * 3
Dim Start As Integer
Dim Index As Long
Dim DataLength As Long

    dbytes = Data.Read
    Socket.RequestRead
    Buffer = StrConv(dbytes(), vbUnicode)
    DataLength = Len(Buffer)
    Index = CLng(Socket.UserData)
    If Buffer = "top" Then
        Top = STR(TotalOnlinePlayers)
        Call SendDataTo(Index, Top)
        Call CloseSocket(Index)
    End If
            
    Player(Index).Buffer = Player(Index).Buffer & Buffer
        
    Start = InStr(Player(Index).Buffer, END_CHAR)
    Do While Start > 0
        Packet = Mid(Player(Index).Buffer, 1, Start - 1)
        Player(Index).Buffer = Mid(Player(Index).Buffer, Start + 1, Len(Player(Index).Buffer))
        Player(Index).DataPackets = Player(Index).DataPackets + 1
        Start = InStr(Player(Index).Buffer, END_CHAR)
        If Len(Packet) > 0 Then
            Call HandleData(Index, Packet)
        End If
    Loop
                
    ' Check if elapsed time has passed
    Player(Index).DataBytes = Player(Index).DataBytes + DataLength
    If GetTickCount >= Player(Index).DataTimer + 1000 Then
        Player(Index).DataTimer = GetTickCount
        Player(Index).DataBytes = 0
        Player(Index).DataPackets = 0
        Exit Sub
    End If
        
    ' Check for data flooding
    If Player(Index).DataBytes > 2000 And GetPlayerAccess(Index) <= 0 Then
        Call HackingAttempt(Index, "Data Flooding")
        Exit Sub
    End If
        
    ' Check for packet flooding
    If Player(Index).DataPackets > 25 And GetPlayerAccess(Index) <= 0 Then
        Call HackingAttempt(Index, "Packet Flooding")
        Exit Sub
    End If
End Sub

Sub CloseSocket(ByVal Index As Long)
    ' Make sure player was/is playing the game, And If so, save'm.
    If Index > 0 Then
        Call LeftGame(Index)
           
        Call TextAdd(frmCServer.txtText, "Connection from " & GetPlayerIP(Index) & " has been terminated.", True)
        Call AddLog("Connection from " & GetPlayerIP(Index) & " has been terminated.", PLAYER_LOG)
           
        Call GameServer.Sockets.Item(Index).ShutDown(ShutdownBoth)
        Set GameServer.Sockets.Item(Index).Socket = Nothing
           
        Call UpdateCaption
        Call SetHighIndex
        Call SendHighIndex
        Call ClearPlayer(Index)
    End If
End Sub


' ********************
' ** Packet Sending **
' ********************

Sub SendDataTo(ByVal Index As Long, ByVal Data As String)
Dim dbytes() As Byte

    dbytes = StrConv(Data, vbFromUnicode)
    If IsConnected(Index) Then
        GameServer.Sockets.Item(Index).WriteBytes dbytes
        DoEvents
    End If
End Sub

Sub SendDataToAll(ByVal Data As String)
Dim i As Long

    For i = 1 To HighIndex
        If IsPlaying(i) Then
            Call SendDataTo(i, Data)
        End If
    Next i
End Sub

Sub SendDataToAllBut(ByVal Index As Long, ByVal Data As String)
Dim i As Long

    For i = 1 To HighIndex
        If IsPlaying(i) And i <> Index Then
            Call SendDataTo(i, Data)
        End If
    Next i
End Sub

Sub SendDataToMap(ByVal MapNum As Long, ByVal Data As String)
Dim i As Long

    For i = 1 To HighIndex
        If IsPlaying(i) Then
            If GetPlayerMap(i) = MapNum Then
                Call SendDataTo(i, Data)
            End If
        End If
    Next i
End Sub

Sub SendDataToMapBut(ByVal Index As Long, ByVal MapNum As Long, ByVal Data As String)
Dim i As Long

    For i = 1 To HighIndex
        If IsPlaying(i) Then
            If GetPlayerMap(i) = MapNum And i <> Index Then
                Call SendDataTo(i, Data)
            End If
        End If
    Next i
End Sub


' ***************
' ** Messaging **
' ***************

Sub SendWhosOnline(ByVal Index As Long)
Dim s As String
Dim n As Long, i As Long

    s = ""
    n = 0
    For i = 1 To HighIndex
        If IsPlaying(i) And i <> Index Then
            s = s & GetPlayerName(i) & ", "
            n = n + 1
        End If
    Next i
            
    If n = 0 Then
        s = "There are no other players online."
    Else
        s = Mid(s, 1, Len(s) - 2)
        s = "There are " & n & " other players online: " & s & "."
    End If
        
    Call PlayerMsg(Index, s, WhoColor)
End Sub

Sub GlobalMsg(ByVal Msg As String, ByVal Color As Byte)
Dim Packet As String

    Packet = "GLOBALMSG" & SEP_CHAR & Msg & SEP_CHAR & Color & SEP_CHAR & END_CHAR
    
    Call SendDataToAll(Packet)
End Sub

Sub AdminMsg(ByVal Msg As String, ByVal Color As Byte)
Dim Packet As String
Dim i As Long

    Packet = "ADMINMSG" & SEP_CHAR & Msg & SEP_CHAR & Color & SEP_CHAR & END_CHAR
    For i = 1 To HighIndex
        If IsPlaying(i) And GetPlayerAccess(i) > 0 Then
            Call SendDataTo(i, Packet)
        End If
    Next i
End Sub

Sub PlayerMsg(ByVal Index As Long, ByVal Msg As String, ByVal Color As Byte)
Dim Packet As String

    Packet = "PLAYERMSG" & SEP_CHAR & Msg & SEP_CHAR & Color & SEP_CHAR & END_CHAR
    
    Call SendDataTo(Index, Packet)
End Sub

Sub MapMsg(ByVal MapNum As Long, ByVal Msg As String, ByVal Color As Byte)
Dim Packet As String
Dim Text As String

    Packet = "MAPMSG" & SEP_CHAR & Msg & SEP_CHAR & Color & SEP_CHAR & END_CHAR
    
    Call SendDataToMap(MapNum, Packet)
End Sub

Sub AlertMsg(ByVal Index As Long, ByVal Msg As String)
Dim Packet As String

    Packet = "ALERTMSG" & SEP_CHAR & Msg & SEP_CHAR & END_CHAR
    
    Call SendDataTo(Index, Packet)
    Call GameServer.Sockets.Item(Index).ShutDown(ShutdownWrite)
End Sub


' ****************************
' ** Account/Player packets **
' ****************************

Sub SendChars(ByVal Index As Long)
Dim Packet As String
Dim i As Long
    
    Packet = "ALLCHARS" & SEP_CHAR
    For i = 1 To MAX_CHARS
        Packet = Packet & Trim(Player(Index).Char(i).Name) & SEP_CHAR & Trim(Class(Player(Index).Char(i).Class).Name) & SEP_CHAR & Player(Index).Char(i).Level & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR
    
    Call SendDataTo(Index, Packet)
End Sub

Sub SendPlayerData(ByVal Index As Long)
Dim Packet As String

    ' Send index's player data to everyone including himself on th emap
    Packet = "PLAYERDATA" & SEP_CHAR & Index & SEP_CHAR & GetPlayerName(Index) & SEP_CHAR & GetPlayerSprite(Index) & SEP_CHAR & GetPlayerMap(Index) & SEP_CHAR & GetPlayerX(Index) & SEP_CHAR & GetPlayerY(Index) & SEP_CHAR & GetPlayerDir(Index) & SEP_CHAR & GetPlayerAccess(Index) & SEP_CHAR & GetPlayerPK(Index) & SEP_CHAR & END_CHAR
    Call SendDataToMap(GetPlayerMap(Index), Packet)
End Sub

Sub SendInventory(ByVal Index As Long)
Dim Packet As String
Dim i As Long

    Packet = "PLAYERINV" & SEP_CHAR
    For i = 1 To MAX_INV
        Packet = Packet & GetPlayerInvItemNum(Index, i) & SEP_CHAR & GetPlayerInvItemValue(Index, i) & SEP_CHAR & GetPlayerInvItemDur(Index, i) & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR
    
    Call SendDataTo(Index, Packet)
End Sub

Sub SendInventoryUpdate(ByVal Index As Long, ByVal InvSlot As Long)
Dim Packet As String
    
    Packet = "PLAYERINVUPDATE" & SEP_CHAR & InvSlot & SEP_CHAR & GetPlayerInvItemNum(Index, InvSlot) & SEP_CHAR & GetPlayerInvItemValue(Index, InvSlot) & SEP_CHAR & GetPlayerInvItemDur(Index, InvSlot) & SEP_CHAR & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendWornEquipment(ByVal Index As Long)
Dim Packet As String
    
    Packet = "PLAYERWORNEQ" & SEP_CHAR & GetPlayerWeaponSlot(Index) & SEP_CHAR & GetPlayerArmorSlot(Index) & SEP_CHAR & GetPlayerHelmetSlot(Index) & SEP_CHAR & GetPlayerShieldSlot(Index) & SEP_CHAR & GetPlayerAmuletSlot(Index) & SEP_CHAR & GetPlayerRingSlot(Index) & SEP_CHAR & GetPlayerArrowSlot(Index) & SEP_CHAR & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendHP(ByVal Index As Long)
Dim Packet As String

    Packet = "PLAYERHP" & SEP_CHAR & GetPlayerMaxHP(Index) & SEP_CHAR & GetPlayerHP(Index) & SEP_CHAR & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendMP(ByVal Index As Long)
Dim Packet As String

    Packet = "PLAYERMP" & SEP_CHAR & GetPlayerMaxMP(Index) & SEP_CHAR & GetPlayerMP(Index) & SEP_CHAR & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendSP(ByVal Index As Long)
Dim Packet As String

    Packet = "PLAYERSP" & SEP_CHAR & GetPlayerMaxSP(Index) & SEP_CHAR & GetPlayerSP(Index) & SEP_CHAR & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendStats(ByVal Index As Long)
Dim Packet As String
    
    Packet = "PLAYERSTATS" & SEP_CHAR & GetPlayerSTR(Index) & SEP_CHAR & GetPlayerDEF(Index) & SEP_CHAR & GetPlayerSPEED(Index) & SEP_CHAR & GetPlayerMAGI(Index) & SEP_CHAR & GetPlayerDEX(Index) & SEP_CHAR & GetPlayerLevel(Index) & SEP_CHAR & GetPlayerExp(Index) & SEP_CHAR & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendPlayerXY(ByVal Index As Long)
Dim Packet As String

    Packet = "PLAYERXY" & SEP_CHAR & GetPlayerX(Index) & SEP_CHAR & GetPlayerY(Index) & SEP_CHAR & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendNpcQuests(ByVal Index As Long, ByVal NpcNum As Long)
Dim Packet As String
Dim n As Long

    Packet = "NPCQUESTS" & SEP_CHAR & NpcNum & SEP_CHAR
    For n = 1 To MAX_NPC_QUESTS
        Packet = Packet & Npc(NpcNum).QuestNPC(n) & SEP_CHAR
    Next n
    Packet = Packet & END_CHAR
    
    Call SendDataTo(Index, Packet)
End Sub

Sub SendTrade(ByVal Index As Long, ByVal ShopNum As Long)
Dim Packet As String
Dim i As Long, x As Long, y As Long, n As Long

    Packet = "TRADE" & SEP_CHAR & ShopNum & SEP_CHAR & Shop(ShopNum).FixesItems
    For i = 1 To MAX_TRADES
        For n = 1 To MAX_GIVE_ITEMS
            Packet = Packet & SEP_CHAR & Shop(ShopNum).TradeItem(i).GiveItem(n)
        Next n
        For n = 1 To MAX_GIVE_VALUE
            Packet = Packet & SEP_CHAR & Shop(ShopNum).TradeItem(i).GiveValue(n)
        Next n
        For n = 1 To MAX_GET_ITEMS
            Packet = Packet & SEP_CHAR & Shop(ShopNum).TradeItem(i).GetItem(n)
        Next n
        For n = 1 To MAX_GET_VALUE
            Packet = Packet & SEP_CHAR & Shop(ShopNum).TradeItem(i).GetValue(n)
        Next n
        Packet = Packet & SEP_CHAR & Shop(ShopNum).ItemStock(i)
        
        '' Item #
        'x = Shop(ShopNum).TradeItem(i).GetItem
        
        'If Item(x).Type = ITEM_TYPE_SPELL Then
            '' Spell class requirement
            'y = Spell(Item(x).Data1).ClassReq
            
            'If y = 0 Then
                'Call PlayerMsg(Index, Trim(Item(x).Name) & " can be used by all classes.", Yellow)
            'Else
                'Call PlayerMsg(Index, Trim(Item(x).Name) & " can only be used by a " & GetClassName(y - 1) & ".", Yellow)
            'End If
        'End If
    Next i
    Packet = Packet & SEP_CHAR & END_CHAR
    
    Call SendDataTo(Index, Packet)
End Sub

Sub SendPlayerSkills(ByVal Index As Long)
Dim Packet As String
Dim i As Long

    Packet = "SKILLS" & SEP_CHAR
    For i = 1 To MAX_PLAYER_SKILLS
        Packet = Packet & GetPlayerSkill(Index, i) & SEP_CHAR & GetPlayerSkillLevel(Index, i) & SEP_CHAR & GetPlayerSkillExp(Index, i) & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR
    
    Call SendDataTo(Index, Packet)
End Sub

Sub SendPlayerSkillsLevel(Index As Long, SkillSlot As Long)
Dim Packet As String

    Packet = "PLAYERSKILLSLVL" & SEP_CHAR & SkillSlot & SEP_CHAR & GetPlayerSkillLevel(Index, SkillSlot) & SEP_CHAR & END_CHAR
    
    Call SendDataTo(Index, Packet)
End Sub

Sub SendPlayerSkillsExp(Index As Long, SkillSlot As Long)
Dim Packet As String

    Packet = "PLAYERSKILLSEXP" & SEP_CHAR & SkillSlot & SEP_CHAR & GetPlayerSkillExp(Index, SkillSlot) & SEP_CHAR & END_CHAR
    
    Call SendDataTo(Index, Packet)
End Sub

Sub SendPlayerSpells(ByVal Index As Long)
Dim Packet As String
Dim i As Long

    Packet = "SPELLS" & SEP_CHAR
    For i = 1 To MAX_PLAYER_SPELLS
        Packet = Packet & GetPlayerSpell(Index, i) & SEP_CHAR & GetPlayerSpellLevel(Index, i) & SEP_CHAR & GetPlayerSpellExp(Index, i) & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR
    
    Call SendDataTo(Index, Packet)
End Sub

Sub SendPlayerSpellsLevel(Index As Long, SpellSlot As Long)
Dim Packet As String

    Packet = "PLAYERSPELLSLVL" & SEP_CHAR & SpellSlot & SEP_CHAR & GetPlayerSpellLevel(Index, SpellSlot) & SEP_CHAR & END_CHAR
    
    Call SendDataTo(Index, Packet)
End Sub

Sub SendPlayerSpellsExp(Index As Long, SpellSlot As Long)
Dim Packet As String

    Packet = "PLAYERSPELLSEXP" & SEP_CHAR & SpellSlot & SEP_CHAR & GetPlayerSpellExp(Index, SpellSlot) & SEP_CHAR & END_CHAR
    
    Call SendDataTo(Index, Packet)
End Sub

Sub SendPlayerQuests(ByVal Index As Long)
Dim Packet As String
Dim i As Long

    Packet = "QUESTS" & SEP_CHAR
    For i = 1 To MAX_PLAYER_QUESTS
        Packet = Packet & GetPlayerQuest(Index, i) & SEP_CHAR & GetPlayerQuestMap(Index, i) & SEP_CHAR & GetPlayerQuestBy(Index, i) & SEP_CHAR & GetPlayerQuestValue(Index, i) & SEP_CHAR & GetPlayerQuestCount(Index, i) & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR
    
    Call SendDataTo(Index, Packet)
End Sub

Sub UpdatePlayerQuest(Index As Long, QuestSlot As Long)
Dim Packet As String

    Packet = "PLAYERQUEST" & SEP_CHAR & QuestSlot & SEP_CHAR & GetPlayerQuest(Index, QuestSlot) & SEP_CHAR & GetPlayerQuestMap(Index, QuestSlot) & SEP_CHAR & GetPlayerQuestBy(Index, QuestSlot) & SEP_CHAR & GetPlayerQuestValue(Index, QuestSlot) & SEP_CHAR & GetPlayerQuestCount(Index, QuestSlot) & SEP_CHAR & END_CHAR
    
    Call SendDataTo(Index, Packet)
End Sub

Sub SendLeftGame(ByVal Index As Long)
Dim Packet As String

    Packet = "PLAYERDATA" & SEP_CHAR & Index & SEP_CHAR & "" & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & 0 & END_CHAR
    Call SendDataToAllBut(Index, Packet)
End Sub


' *****************
' ** Map packets **
' *****************

Sub SendMap(ByVal Index As Long, ByVal MapNum As Long)
Dim Packet As String, P1 As String, P2 As String
Dim x As Long
Dim y As Long

    Packet = "MAPDATA" & SEP_CHAR & MapNum & SEP_CHAR & Trim(Map(MapNum).Name) & SEP_CHAR & Map(MapNum).Revision & SEP_CHAR & Trim(Map(MapNum).Owner) & SEP_CHAR & Map(MapNum).Moral & SEP_CHAR & Map(MapNum).Up & SEP_CHAR & Map(MapNum).Down & SEP_CHAR & Map(MapNum).Left & SEP_CHAR & Map(MapNum).Right & SEP_CHAR & Map(MapNum).Music & SEP_CHAR & Map(MapNum).BootMap & SEP_CHAR & Map(MapNum).BootX & SEP_CHAR & Map(MapNum).BootY & SEP_CHAR & Map(MapNum).Indoors & SEP_CHAR
    
    For x = 1 To MAX_MAP_NPCS
        Packet = Packet & Map(MapNum).NSpawn(x).NSx & SEP_CHAR & Map(MapNum).NSpawn(x).NSy & SEP_CHAR
    Next x
    
    For x = 1 To MAX_MAP_RESOURCES
        Packet = Packet & Map(MapNum).RSpawn(x).RSx & SEP_CHAR & Map(MapNum).RSpawn(x).RSy & SEP_CHAR
    Next x
    
    For y = 0 To MAX_MAPY
        For x = 0 To MAX_MAPX
            With Map(MapNum).Tile(x, y)
                Packet = Packet & .Tileset & SEP_CHAR & .Ground & SEP_CHAR & .Mask & SEP_CHAR & .Mask2 & SEP_CHAR & .Anim & SEP_CHAR & .Fringe & SEP_CHAR & .Fringe2 & SEP_CHAR & .FAnim & SEP_CHAR & .Light & SEP_CHAR & .Type & SEP_CHAR & .Data1 & SEP_CHAR & .Data2 & SEP_CHAR & .Data3 & SEP_CHAR & .WalkUp & SEP_CHAR & .WalkDown & SEP_CHAR & .WalkLeft & SEP_CHAR & .WalkRight & SEP_CHAR & .Build & SEP_CHAR
            End With
        Next x
    Next y
    
    For x = 1 To MAX_MAP_NPCS
        Packet = Packet & Map(MapNum).Npc(x) & SEP_CHAR
    Next x
    
    For x = 1 To MAX_MAP_RESOURCES
        Packet = Packet & Map(MapNum).Resource(x) & SEP_CHAR
    Next x
    
    Packet = Packet & END_CHAR
    
    x = Int(Len(Packet) / 2)
    P1 = Mid(Packet, 1, x)
    P2 = Mid(Packet, x + 1, Len(Packet) - x)
    Call SendDataTo(Index, Packet)
End Sub

Sub SendJoinMap(ByVal Index As Long)
Dim Packet As String
Dim i As Long

    Packet = ""
    
    ' Send all players on current map to index
    For i = 1 To HighIndex
        Packet = ""
        If IsPlaying(i) And i <> Index And GetPlayerMap(i) = GetPlayerMap(Index) Then
            Packet = Packet & "PLAYERDATA" & SEP_CHAR & i & SEP_CHAR & GetPlayerName(i) & SEP_CHAR & GetPlayerSprite(i) & SEP_CHAR & GetPlayerMap(i) & SEP_CHAR & GetPlayerX(i) & SEP_CHAR & GetPlayerY(i) & SEP_CHAR & GetPlayerDir(i) & SEP_CHAR & GetPlayerAccess(i) & SEP_CHAR & GetPlayerPK(i) & SEP_CHAR & END_CHAR
            Call SendDataTo(Index, Packet)
        End If
    Next i
    
    ' Send index's player data to everyone on the map including himself
    Packet = "PLAYERDATA" & SEP_CHAR & Index & SEP_CHAR & GetPlayerName(Index) & SEP_CHAR & GetPlayerSprite(Index) & SEP_CHAR & GetPlayerMap(Index) & SEP_CHAR & GetPlayerX(Index) & SEP_CHAR & GetPlayerY(Index) & SEP_CHAR & GetPlayerDir(Index) & SEP_CHAR & GetPlayerAccess(Index) & SEP_CHAR & GetPlayerPK(Index) & SEP_CHAR & END_CHAR
    Call SendDataToMap(GetPlayerMap(Index), Packet)
End Sub

Sub SendLeaveMap(ByVal Index As Long, ByVal MapNum As Long)
Dim Packet As String

    Packet = "PLAYERDATA" & SEP_CHAR & Index & SEP_CHAR & GetPlayerName(Index) & SEP_CHAR & GetPlayerSprite(Index) & SEP_CHAR & 0 & SEP_CHAR & GetPlayerX(Index) & SEP_CHAR & GetPlayerY(Index) & SEP_CHAR & GetPlayerDir(Index) & SEP_CHAR & GetPlayerAccess(Index) & SEP_CHAR & GetPlayerPK(Index) & SEP_CHAR & END_CHAR
    Call SendDataToMapBut(Index, MapNum, Packet)
End Sub

Sub SendMapItemsTo(ByVal Index As Long, ByVal MapNum As Long)
Dim Packet As String
Dim i As Long

    Packet = "MAPITEMDATA" & SEP_CHAR
    For i = 1 To MAX_MAP_ITEMS
        Packet = Packet & MapItem(MapNum, i).Num & SEP_CHAR & MapItem(MapNum, i).Value & SEP_CHAR & MapItem(MapNum, i).Dur & SEP_CHAR & MapItem(MapNum, i).x & SEP_CHAR & MapItem(MapNum, i).y & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR
    
    Call SendDataTo(Index, Packet)
End Sub

Sub SendMapItemsToAll(ByVal MapNum As Long)
Dim Packet As String
Dim i As Long

    Packet = "MAPITEMDATA" & SEP_CHAR
    For i = 1 To MAX_MAP_ITEMS
        Packet = Packet & MapItem(MapNum, i).Num & SEP_CHAR & MapItem(MapNum, i).Value & SEP_CHAR & MapItem(MapNum, i).Dur & SEP_CHAR & MapItem(MapNum, i).x & SEP_CHAR & MapItem(MapNum, i).y & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR
    
    Call SendDataToMap(MapNum, Packet)
End Sub

Sub SendMapNpcsTo(ByVal Index As Long, ByVal MapNum As Long)
Dim Packet As String
Dim i As Long

    Packet = "MAPNPCDATA" & SEP_CHAR
    For i = 1 To MAX_MAP_NPCS
        Packet = Packet & MapNpc(MapNum, i).Num & SEP_CHAR & MapNpc(MapNum, i).x & SEP_CHAR & MapNpc(MapNum, i).y & SEP_CHAR & MapNpc(MapNum, i).Dir & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR
    
    Call SendDataTo(Index, Packet)
End Sub

Sub SendMapNpcsToMap(ByVal MapNum As Long)
Dim Packet As String
Dim i As Long

    Packet = "MAPNPCDATA" & SEP_CHAR
    For i = 1 To MAX_MAP_NPCS
        Packet = Packet & MapNpc(MapNum, i).Num & SEP_CHAR & MapNpc(MapNum, i).x & SEP_CHAR & MapNpc(MapNum, i).y & SEP_CHAR & MapNpc(MapNum, i).Dir & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR
    
    Call SendDataToMap(MapNum, Packet)
End Sub

Sub SendMapResourcesTo(ByVal Index As Long, ByVal MapNum As Long)
Dim Packet As String
Dim i As Long

    Packet = "MAPRESOURCEDATA" & SEP_CHAR
    For i = 1 To MAX_MAP_RESOURCES
        Packet = Packet & MapResource(MapNum, i).Num & SEP_CHAR & MapResource(MapNum, i).x & SEP_CHAR & MapResource(MapNum, i).y & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR
    
    Call SendDataTo(Index, Packet)
End Sub

Sub SendMapResourcesToMap(ByVal MapNum As Long)
Dim Packet As String
Dim i As Long

    Packet = "MAPRESOURCEDATA" & SEP_CHAR
    For i = 1 To MAX_MAP_RESOURCES
        Packet = Packet & MapResource(MapNum, i).Num & SEP_CHAR & MapResource(MapNum, i).x & SEP_CHAR & MapResource(MapNum, i).y & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR
    
    Call SendDataToMap(MapNum, Packet)
End Sub


' *********************
' ** General packets **
' *********************

Sub SendHighIndex()
Dim i As Long
    For i = 1 To HighIndex
        Call SendDataTo(i, "HighIndex" & SEP_CHAR & HighIndex & SEP_CHAR & END_CHAR)
    Next i
End Sub

Sub SendWelcome(ByVal Index As Long)
Dim MOTD As String
Dim f As Long

    ' Send them welcome
    Call PlayerMsg(Index, "Welcome to " & GAME_NAME & "!  Version " & CLIENT_MAJOR & "." & CLIENT_MINOR & "." & CLIENT_REVISION, BrightBlue)
    Call PlayerMsg(Index, "Type /help for help on commands.  Use arrow keys to move, hold down shift to run, and use ctrl to attack.", Cyan)
    ' Send them MOTD
    MOTD = GetVar(App.Path & "\data\motd.ini", "MOTD", "Msg")
    If Trim(MOTD) <> "" Then
        Call PlayerMsg(Index, "MOTD: " & MOTD, BrightCyan)
    End If
    
    ' Send whos online
    Call SendWhosOnline(Index)
End Sub

Sub SendClasses(ByVal Index As Long)
Dim Packet As String
Dim i As Long

    Packet = "CLASSESDATA" & SEP_CHAR & Max_Classes & SEP_CHAR
    For i = 0 To Max_Classes
        Packet = Packet & GetClassName(i) & SEP_CHAR & GetClassMaxHP(i) & SEP_CHAR & GetClassMaxMP(i) & SEP_CHAR & GetClassMaxSP(i) & SEP_CHAR & Class(i).STR & SEP_CHAR & Class(i).DEF & SEP_CHAR & Class(i).SPEED & SEP_CHAR & Class(i).MAGI & SEP_CHAR & Class(i).DEX & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR
    
    Call SendDataTo(Index, Packet)
End Sub

Sub SendNewCharClasses(ByVal Index As Long)
Dim Packet As String
Dim i As Long

    Packet = "NEWCHARCLASSES" & SEP_CHAR & Max_Classes & SEP_CHAR
    For i = 0 To Max_Classes
        Packet = Packet & GetClassName(i) & SEP_CHAR & GetClassMaxHP(i) & SEP_CHAR & GetClassMaxMP(i) & SEP_CHAR & GetClassMaxSP(i) & SEP_CHAR & Class(i).STR & SEP_CHAR & Class(i).DEF & SEP_CHAR & Class(i).SPEED & SEP_CHAR & Class(i).MAGI & SEP_CHAR & Class(i).DEX & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR
    
    Call SendDataTo(Index, Packet)
End Sub

Sub SendItems(ByVal Index As Long)
Dim Packet As String
Dim i As Long

    For i = 1 To MAX_ITEMS
        If Trim(Item(i).Name) <> "" Then
            Call SendUpdateItemTo(Index, i)
        End If
    Next i
End Sub

Sub SendNpcs(ByVal Index As Long)
Dim Packet As String
Dim i As Long

    For i = 1 To MAX_NPCS
        If Trim(Npc(i).Name) <> "" Then
            Call SendUpdateNpcTo(Index, i)
        End If
    Next i
End Sub

Sub SendShops(ByVal Index As Long)
Dim i As Long

    For i = 1 To MAX_SHOPS
        If Trim(Shop(i).Name) <> "" Then
            Call SendUpdateShopTo(Index, i)
        End If
    Next i
End Sub

Sub SendSkills(ByVal Index As Long)
Dim i As Long

    For i = 1 To MAX_SKILLS
        If Trim(Skill(i).Name) <> "" Then
            Call SendUpdateSkillTo(Index, i)
        End If
    Next i
End Sub

Sub SendSpells(ByVal Index As Long)
Dim i As Long

    For i = 1 To MAX_SPELLS
        If Trim(Spell(i).Name) <> "" Then
            Call SendUpdateSpellTo(Index, i)
        End If
    Next i
End Sub

Sub SendQuests(ByVal Index As Long)
Dim i As Long

    For i = 1 To MAX_QUESTS
        If Trim(Quest(i).Name) <> "" Then
            Call SendUpdateQuestTo(Index, i)
        End If
    Next i
End Sub

Sub SendGUIS(ByVal Index As Long)
Dim i As Long

    For i = 1 To MAX_GUIS
        If Trim(GUI(i).Name) <> "" Then
            Call SendUpdateGUITo(Index, i)
        End If
    Next i
End Sub


' ********************
' ** Editor packets **
' ********************

Sub SendUpdateItemToAll(ByVal ItemNum As Long)
Dim Packet As String

    Packet = "UPDATEITEM" & SEP_CHAR & ItemNum & SEP_CHAR & Trim(Item(ItemNum).Name) & SEP_CHAR & Item(ItemNum).Pic & SEP_CHAR & Item(ItemNum).Type & SEP_CHAR & Item(ItemNum).Data1 & SEP_CHAR & Item(ItemNum).Data2 & SEP_CHAR & Item(ItemNum).Data3 & SEP_CHAR & END_CHAR
    Call SendDataToAll(Packet)
End Sub

Sub SendUpdateItemTo(ByVal Index As Long, ByVal ItemNum As Long)
Dim Packet As String

    Packet = "UPDATEITEM" & SEP_CHAR & ItemNum & SEP_CHAR & Trim(Item(ItemNum).Name) & SEP_CHAR & Item(ItemNum).Pic & SEP_CHAR & Item(ItemNum).Type & SEP_CHAR & Item(ItemNum).Data1 & SEP_CHAR & Item(ItemNum).Data2 & SEP_CHAR & Item(ItemNum).Data3 & SEP_CHAR & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendEditItemTo(ByVal Index As Long, ByVal ItemNum As Long)
Dim Packet As String

    Packet = "EDITITEM" & SEP_CHAR & ItemNum & SEP_CHAR & Trim(Item(ItemNum).Name) & SEP_CHAR & Item(ItemNum).Pic & SEP_CHAR & Item(ItemNum).Type & SEP_CHAR & Item(ItemNum).Data1 & SEP_CHAR & Item(ItemNum).Data2 & SEP_CHAR & Item(ItemNum).Data3 & SEP_CHAR & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendUpdateNpcToAll(ByVal NpcNum As Long)
Dim Packet As String

    Packet = "UPDATENPC" & SEP_CHAR & NpcNum & SEP_CHAR & Trim(Npc(NpcNum).Name) & SEP_CHAR & Npc(NpcNum).Sprite & SEP_CHAR & Npc(NpcNum).Behavior & SEP_CHAR & Npc(NpcNum).Big & SEP_CHAR & END_CHAR
    Call SendDataToAll(Packet)
End Sub

Sub SendUpdateNpcTo(ByVal Index As Long, ByVal NpcNum As Long)
Dim Packet As String

    Packet = "UPDATENPC" & SEP_CHAR & NpcNum & SEP_CHAR & Trim(Npc(NpcNum).Name) & SEP_CHAR & Npc(NpcNum).Sprite & SEP_CHAR & Npc(NpcNum).Behavior & SEP_CHAR & Npc(NpcNum).Big & SEP_CHAR & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendEditNpcTo(ByVal Index As Long, ByVal NpcNum As Long)
Dim Packet As String
Dim i As Long

    Packet = "EDITNPC" & SEP_CHAR & NpcNum & SEP_CHAR & Trim(Npc(NpcNum).Name) & SEP_CHAR & Npc(NpcNum).Sprite & SEP_CHAR & Npc(NpcNum).SpawnSecs & SEP_CHAR & Npc(NpcNum).Behavior & SEP_CHAR & Npc(NpcNum).Range & SEP_CHAR & Npc(NpcNum).STR & SEP_CHAR & Npc(NpcNum).DEF & SEP_CHAR & Npc(NpcNum).SPEED & SEP_CHAR & Npc(NpcNum).MAGI
    Packet = Packet & SEP_CHAR & Npc(NpcNum).Big & SEP_CHAR & Npc(NpcNum).MaxHp & SEP_CHAR & Npc(NpcNum).Respawn & SEP_CHAR & Npc(NpcNum).HitOnlyWith & SEP_CHAR & Npc(NpcNum).ShopLink & SEP_CHAR & Npc(NpcNum).ExpType & SEP_CHAR & Npc(NpcNum).EXP & SEP_CHAR
    For i = 1 To MAX_NPC_QUESTS
        Packet = Packet & Npc(NpcNum).QuestNPC(i) & SEP_CHAR
    Next i
    For i = 1 To MAX_NPC_DROPS
        Packet = Packet & Npc(NpcNum).ItemNPC(i).Chance
        Packet = Packet & SEP_CHAR & Npc(NpcNum).ItemNPC(i).ItemNum
        Packet = Packet & SEP_CHAR & Npc(NpcNum).ItemNPC(i).ItemValue & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendUpdateShopToAll(ByVal ShopNum As Long)
Dim Packet As String

    Packet = "UPDATESHOP" & SEP_CHAR & ShopNum & SEP_CHAR & Trim(Shop(ShopNum).Name) & SEP_CHAR & END_CHAR
    Call SendDataToAll(Packet)
End Sub

Sub SendUpdateShopTo(ByVal Index As Long, ByVal ShopNum)
Dim Packet As String

    Packet = "UPDATESHOP" & SEP_CHAR & ShopNum & SEP_CHAR & Trim(Shop(ShopNum).Name) & SEP_CHAR & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendEditShopTo(ByVal Index As Long, ByVal ShopNum As Long)
Dim Packet As String
Dim i As Long
Dim n As Long

    Packet = "EDITSHOP" & SEP_CHAR & ShopNum & SEP_CHAR & Trim(Shop(ShopNum).Name) & SEP_CHAR & Shop(ShopNum).FixesItems
    For i = 1 To MAX_TRADES
        For n = 1 To MAX_GIVE_ITEMS
            Packet = Packet & SEP_CHAR & Shop(ShopNum).TradeItem(i).GiveItem(n)
        Next n
        For n = 1 To MAX_GIVE_VALUE
            Packet = Packet & SEP_CHAR & Shop(ShopNum).TradeItem(i).GiveValue(n)
        Next n
        For n = 1 To MAX_GET_ITEMS
            Packet = Packet & SEP_CHAR & Shop(ShopNum).TradeItem(i).GetItem(n)
        Next n
        For n = 1 To MAX_GET_VALUE
            Packet = Packet & SEP_CHAR & Shop(ShopNum).TradeItem(i).GetValue(n)
        Next n
        Packet = Packet & SEP_CHAR & Shop(ShopNum).ItemStock(i)
    Next i
    Packet = Packet & SEP_CHAR & END_CHAR

    Call SendDataTo(Index, Packet)
End Sub

Sub SendUpdateSkillToAll(ByVal SkillNum As Long)
Dim Packet As String

    Packet = "UPDATESKILL" & SEP_CHAR & SkillNum & SEP_CHAR & Trim(Skill(SkillNum).Name) & SEP_CHAR & Skill(SkillNum).SkillSprite & SEP_CHAR & Skill(SkillNum).ClassReq & SEP_CHAR & Skill(SkillNum).LevelReq & SEP_CHAR & Skill(SkillNum).Type & SEP_CHAR & Skill(SkillNum).Data1 & SEP_CHAR & Skill(SkillNum).Data2 & SEP_CHAR & Skill(SkillNum).Data3 & SEP_CHAR & END_CHAR
    Call SendDataToAll(Packet)
End Sub

Sub SendUpdateSkillTo(ByVal Index As Long, ByVal SkillNum As Long)
Dim Packet As String

    Packet = "UPDATESKILL" & SEP_CHAR & SkillNum & SEP_CHAR & Trim(Skill(SkillNum).Name) & SEP_CHAR & Skill(SkillNum).SkillSprite & SEP_CHAR & Skill(SkillNum).ClassReq & SEP_CHAR & Skill(SkillNum).LevelReq & SEP_CHAR & Skill(SkillNum).Type & SEP_CHAR & Skill(SkillNum).Data1 & SEP_CHAR & Skill(SkillNum).Data2 & SEP_CHAR & Skill(SkillNum).Data3 & SEP_CHAR & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendEditSkillTo(ByVal Index As Long, ByVal SkillNum As Long)
Dim Packet As String

    Packet = "EDITSKILL" & SEP_CHAR & SkillNum & SEP_CHAR & Trim(Skill(SkillNum).Name) & SEP_CHAR & Skill(SkillNum).SkillSprite & SEP_CHAR & Skill(SkillNum).ClassReq & SEP_CHAR & Skill(SkillNum).LevelReq & SEP_CHAR & Skill(SkillNum).Type & SEP_CHAR & Skill(SkillNum).Data1 & SEP_CHAR & Skill(SkillNum).Data2 & SEP_CHAR & Skill(SkillNum).Data3 & SEP_CHAR & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendUpdateSpellToAll(ByVal SpellNum As Long)
Dim Packet As String

    Packet = "UPDATESPELL" & SEP_CHAR & SpellNum & SEP_CHAR & Trim(Spell(SpellNum).Name) & SEP_CHAR & Spell(SpellNum).SpellSprite & SEP_CHAR & Spell(SpellNum).ClassReq & SEP_CHAR & Spell(SpellNum).LevelReq & SEP_CHAR & Spell(SpellNum).Type & SEP_CHAR & Spell(SpellNum).Data1 & SEP_CHAR & Spell(SpellNum).Data2 & SEP_CHAR & Spell(SpellNum).Data3 & SEP_CHAR & END_CHAR
    Call SendDataToAll(Packet)
End Sub

Sub SendUpdateSpellTo(ByVal Index As Long, ByVal SpellNum As Long)
Dim Packet As String

    Packet = "UPDATESPELL" & SEP_CHAR & SpellNum & SEP_CHAR & Trim(Spell(SpellNum).Name) & SEP_CHAR & Spell(SpellNum).SpellSprite & SEP_CHAR & Spell(SpellNum).ClassReq & SEP_CHAR & Spell(SpellNum).LevelReq & SEP_CHAR & Spell(SpellNum).Type & SEP_CHAR & Spell(SpellNum).Data1 & SEP_CHAR & Spell(SpellNum).Data2 & SEP_CHAR & Spell(SpellNum).Data3 & SEP_CHAR & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendEditSpellTo(ByVal Index As Long, ByVal SpellNum As Long)
Dim Packet As String

    Packet = "EDITSPELL" & SEP_CHAR & SpellNum & SEP_CHAR & Trim(Spell(SpellNum).Name) & SEP_CHAR & Spell(SpellNum).SpellSprite & SEP_CHAR & Spell(SpellNum).ClassReq & SEP_CHAR & Spell(SpellNum).LevelReq & SEP_CHAR & Spell(SpellNum).Type & SEP_CHAR & Spell(SpellNum).Data1 & SEP_CHAR & Spell(SpellNum).Data2 & SEP_CHAR & Spell(SpellNum).Data3 & SEP_CHAR & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendUpdateQuestToAll(ByVal QuestNum As Long)
Dim Packet As String

    'Packet = "UPDATEQUEST" & SEP_CHAR & QuestNum & SEP_CHAR & Trim(Quest(QuestNum).Name) & SEP_CHAR & Quest(QuestNum).SetBy & SEP_CHAR & Quest(QuestNum).ClassReq & SEP_CHAR & Quest(QuestNum).LevelMin & SEP_CHAR & Quest(QuestNum).LevelMax & SEP_CHAR & Quest(QuestNum).Type & SEP_CHAR & Quest(QuestNum).Reward & SEP_CHAR & Quest(QuestNum).RewardValue & SEP_CHAR & Quest(QuestNum).Data1 & SEP_CHAR & Quest(QuestNum).Data2 & SEP_CHAR & Quest(QuestNum).Data3 & SEP_CHAR & Trim(Quest(QuestNum).Description) & SEP_CHAR & END_CHAR
    Packet = "UPDATEQUEST" & SEP_CHAR & QuestNum & SEP_CHAR & Trim(Quest(QuestNum).Name) & SEP_CHAR & Quest(QuestNum).ClassReq & SEP_CHAR & Quest(QuestNum).LevelMin & SEP_CHAR & Quest(QuestNum).LevelMax & SEP_CHAR & Quest(QuestNum).Reward & SEP_CHAR & Quest(QuestNum).RewardValue & SEP_CHAR & Trim(Quest(QuestNum).Description) & SEP_CHAR & END_CHAR
    Call SendDataToAll(Packet)
End Sub

Sub SendUpdateQuestTo(ByVal Index As Long, ByVal QuestNum As Long)
Dim Packet As String

    'Packet = "UPDATEQUEST" & SEP_CHAR & QuestNum & SEP_CHAR & Trim(Quest(QuestNum).Name) & SEP_CHAR & Quest(QuestNum).SetBy & SEP_CHAR & Quest(QuestNum).ClassReq & SEP_CHAR & Quest(QuestNum).LevelMin & SEP_CHAR & Quest(QuestNum).LevelMax & SEP_CHAR & Quest(QuestNum).Type & SEP_CHAR & Quest(QuestNum).Reward & SEP_CHAR & Quest(QuestNum).RewardValue & SEP_CHAR & Quest(QuestNum).Data1 & SEP_CHAR & Quest(QuestNum).Data2 & SEP_CHAR & Quest(QuestNum).Data3 & SEP_CHAR & Trim(Quest(QuestNum).Description) & SEP_CHAR & END_CHAR
    Packet = "UPDATEQUEST" & SEP_CHAR & QuestNum & SEP_CHAR & Trim(Quest(QuestNum).Name) & SEP_CHAR & Quest(QuestNum).ClassReq & SEP_CHAR & Quest(QuestNum).LevelMin & SEP_CHAR & Quest(QuestNum).LevelMax & SEP_CHAR & Quest(QuestNum).Reward & SEP_CHAR & Quest(QuestNum).RewardValue & SEP_CHAR & Trim(Quest(QuestNum).Description) & SEP_CHAR & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendEditQuestTo(ByVal Index As Long, ByVal QuestNum As Long)
Dim Packet As String

    Packet = "EDITQUEST" & SEP_CHAR & QuestNum & SEP_CHAR & Trim(Quest(QuestNum).Name) & SEP_CHAR & Quest(QuestNum).SetBy & SEP_CHAR & Quest(QuestNum).ClassReq & SEP_CHAR & Quest(QuestNum).LevelMin & SEP_CHAR & Quest(QuestNum).LevelMax & SEP_CHAR & Quest(QuestNum).Type & SEP_CHAR & Quest(QuestNum).Reward & SEP_CHAR & Quest(QuestNum).RewardValue & SEP_CHAR & Quest(QuestNum).Data1 & SEP_CHAR & Quest(QuestNum).Data2 & SEP_CHAR & Quest(QuestNum).Data3 & SEP_CHAR & Trim(Quest(QuestNum).Description) & SEP_CHAR & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendUpdateGUIToAll(ByVal GUINum As Long)
Dim Packet As String

    Packet = "UPDATEGUI" & SEP_CHAR & GUINum & SEP_CHAR & Trim(GUI(GUINum).Name) & SEP_CHAR & Trim(GUI(GUINum).Designer) & SEP_CHAR & END_CHAR
    Call SendDataToAll(Packet)
End Sub

Sub SendUpdateGUITo(ByVal Index As Long, ByVal GUINum As Long)
Dim Packet As String

    Packet = "UPDATEGUI" & SEP_CHAR & GUINum & SEP_CHAR & Trim(GUI(GUINum).Name) & SEP_CHAR & Trim(GUI(GUINum).Designer) & SEP_CHAR & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendEditGUITo(ByVal Index As Long, ByVal GUINum As Long)
Dim Packet As String
Dim i As Long

    With GUI(GUINum)
        Packet = "EDITGUI" & SEP_CHAR & GUINum & SEP_CHAR & Trim(.Name) & SEP_CHAR & Trim(.Designer) & SEP_CHAR & .Revision
        For i = 1 To 7
            Packet = Packet & SEP_CHAR & .Background(i).Data1 & SEP_CHAR & .Background(i).Data2 & SEP_CHAR & .Background(i).Data3 & SEP_CHAR & .Background(i).Data4 & SEP_CHAR & .Background(i).Data5
        Next i
        For i = 1 To 5
            Packet = Packet & SEP_CHAR & .Menu(i).Data1 & SEP_CHAR & .Menu(i).Data2 & SEP_CHAR & .Menu(i).Data3 & SEP_CHAR & .Menu(i).Data4
        Next i
        For i = 1 To 4
            Packet = Packet & SEP_CHAR & .Login(i).Data1 & SEP_CHAR & .Login(i).Data2 & SEP_CHAR & .Login(i).Data3 & SEP_CHAR & .Login(i).Data4
        Next i
        For i = 1 To 4
            Packet = Packet & SEP_CHAR & .NewAcc(i).Data1 & SEP_CHAR & .NewAcc(i).Data2 & SEP_CHAR & .NewAcc(i).Data3 & SEP_CHAR & .NewAcc(i).Data4
        Next i
        For i = 1 To 4
            Packet = Packet & SEP_CHAR & .DelAcc(i).Data1 & SEP_CHAR & .DelAcc(i).Data2 & SEP_CHAR & .DelAcc(i).Data3 & SEP_CHAR & .DelAcc(i).Data4
        Next i
        For i = 1 To 2
            Packet = Packet & SEP_CHAR & .Credits(i).Data1 & SEP_CHAR & .Credits(i).Data2 & SEP_CHAR & .Credits(i).Data3 & SEP_CHAR & .Credits(i).Data4
        Next i
        For i = 1 To 5
            Packet = Packet & SEP_CHAR & .Chars(i).Data1 & SEP_CHAR & .Chars(i).Data2 & SEP_CHAR & .Chars(i).Data3 & SEP_CHAR & .Chars(i).Data4
        Next i
        For i = 1 To 14
            Packet = Packet & SEP_CHAR & .NewChar(i).Data1 & SEP_CHAR & .NewChar(i).Data2 & SEP_CHAR & .NewChar(i).Data3 & SEP_CHAR & .NewChar(i).Data4
        Next i
        Packet = Packet & SEP_CHAR & END_CHAR
    End With
    
    Call SendDataTo(Index, Packet)
End Sub
