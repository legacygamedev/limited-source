Attribute VB_Name = "modServerTCP"
Option Explicit

Sub UpdateCaption()
    frmServer.Caption = "Mirage Server <IP " & frmServer.Socket(0).LocalIP & " Port " & str(frmServer.Socket(0).LocalPort) & "> (" & TotalOnlinePlayers & ")"
End Sub

Function IsConnected(ByVal Index As Long) As Boolean
    If frmServer.Socket(Index).State = sckConnected Then
        IsConnected = True
    Else
        IsConnected = False
    End If
End Function

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

Public Function IsMultiComputers(ByVal HDModel As String, ByVal HDSerial As String) As Boolean
Dim I As Long

IsMultiComputers = False
For I = 1 To MAX_PLAYERS
    If IsConnected(I) And (Player(I).HDModel = HDModel Or Player(I).HDSerial = HDSerial) Then
        IsMultiComputers = True
        Exit Function
    End If
Next I
End Function

Function IsMultiAccounts(ByVal Login As String) As Boolean
Dim I As Long

IsMultiAccounts = False
For I = 1 To MAX_PLAYERS
    If IsConnected(I) And LCase(Trim(Player(I).Login)) = LCase(Trim(Login)) Then
        IsMultiAccounts = True
        Exit Function
    End If
Next I
End Function

Function IsMultiIPOnline(ByVal IP As String) As Boolean
Dim I As Long
Dim N As Long

    N = 0
    IsMultiIPOnline = False
    For I = 1 To MAX_PLAYERS
        If IsConnected(I) And Trim(GetPlayerIP(I)) = Trim(IP) Then
            N = N + 1
            
            If (N > 1) Then
                IsMultiIPOnline = True
                Exit Function
            End If
        End If
    Next I
End Function


Sub SendDataTo(ByVal Index As Long, ByVal Data As String)
Dim I As Long, N As Long, startc As Long

    If IsConnected(Index) Then
        Debug.Print "S: " & Data
        frmServer.Socket(Index).SendData Data
        DoEvents
    End If
End Sub

Sub SendDataToAll(ByVal Data As String)
Dim I As Long

    For I = 1 To MAX_PLAYERS
        If IsPlaying(I) Then
            Call SendDataTo(I, Data)
        End If
    Next I
End Sub

Sub SendDataToAllBut(ByVal Index As Long, ByVal Data As String)
Dim I As Long

    For I = 1 To MAX_PLAYERS
        If IsPlaying(I) And I <> Index Then
            Call SendDataTo(I, Data)
        End If
    Next I
End Sub

Sub SendDataToMap(ByVal MapNum As Long, ByVal Data As String)
Dim I As Long

    For I = 1 To MAX_PLAYERS
        If IsPlaying(I) Then
            If GetPlayerMap(I) = MapNum Then
                Call SendDataTo(I, Data)
            End If
        End If
    Next I
End Sub

Sub SendDataToMapBut(ByVal Index As Long, ByVal MapNum As Long, ByVal Data As String)
Dim I As Long

    For I = 1 To MAX_PLAYERS
        If IsPlaying(I) Then
            If GetPlayerMap(I) = MapNum And I <> Index Then
                Call SendDataTo(I, Data)
            End If
        End If
    Next I
End Sub

Sub GlobalMsg(ByVal Msg As String, ByVal Color As Byte)
Dim Packet As String

    Packet = "GLOBALMSG" & SEP_CHAR & Msg & SEP_CHAR & Color & SEP_CHAR & END_CHAR
    
    Call SendDataToAll(Packet)
End Sub

Sub AdminMsg(ByVal Msg As String, ByVal Color As Byte)
Dim Packet As String
Dim I As Long

    Packet = "ADMINMSG" & SEP_CHAR & Msg & SEP_CHAR & Color & SEP_CHAR & END_CHAR
    For I = 1 To MAX_PLAYERS
        If IsPlaying(I) And GetPlayerAccess(I) > 0 Then
            Call SendDataTo(I, Packet)
        End If
    Next I
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
    Call CloseSocket(Index)
End Sub

Sub HackingAttempt(ByVal Index As Long, ByVal Reason As String)
    If Index > 0 Then
        If IsPlaying(Index) Then
            Call AdminMsg(GetPlayerLogin(Index) & "/" & GetPlayerName(Index) & " has been booted for (" & Reason & ")", White)
        End If
    
        Call AlertMsg(Index, "You have lost your connection with " & GAME_NAME & ".")
    End If
End Sub

Sub AcceptConnection(ByVal Index As Long, ByVal SocketId As Long)
Dim I As Long

    If (Index = 0) Then
        I = FindOpenPlayerSlot
        
        If I <> 0 Then
            ' Whoho, we can connect them
            frmServer.Socket(I).Close
            frmServer.Socket(I).Accept SocketId
            Call SocketConnected(I)
        End If
    End If
End Sub

Sub SocketConnected(ByVal Index As Long)
    If Index <> 0 Then
        ' Are they trying to connect more then one connection?
        'If Not IsMultiIPOnline(GetPlayerIP(Index)) Then
            If Not IsBanned(GetPlayerIP(Index)) Then
                Call TextAdd(frmServer.txtText, "Received connection from " & GetPlayerIP(Index) & ".", True)
            Else
                Call AlertMsg(Index, "You have been banned from " & GAME_NAME & ", and can no longer play.")
            End If
        'Else
           ' Tried multiple connections
        '    Call AlertMsg(Index, GAME_NAME & " does not allow multiple IP's anymore.")
        'End If
    End If
End Sub

Sub IncomingData(ByVal Index As Long, ByVal DataLength As Long)
On Error Resume Next

Dim Buffer As String
Dim Packet As String
Dim Start As Integer

    If Index > 0 Then
        frmServer.Socket(Index).GetData Buffer, vbString, DataLength
                   
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
        
        'Below removed due to floodguard flaw.
        
        ' Check for data flooding
        'If Player(Index).DataBytes > 1000 And GetPlayerAccess(Index) <= 0 Then
        '    Call HackingAttempt(Index, "Data Flooding")
        '    Exit Sub
        'End If
        
        ' Check for packet flooding
        'If Player(Index).DataPackets > 25 And GetPlayerAccess(Index) <= 0 Then
        '    Call HackingAttempt(Index, "Packet Flooding")
        '    Exit Sub
        'End If
    End If
End Sub


Sub CloseSocket(ByVal Index As Long)
    ' Make sure player was/is playing the game, and if so, save'm.
    If Index > 0 Then
        Call LeftGame(Index)
    
        Call TextAdd(frmServer.txtText, "Connection from " & GetPlayerIP(Index) & " has been terminated.", True)
        
        frmServer.Socket(Index).Close
            
        Call UpdateCaption
        Call ClearPlayer(Index)
    End If
End Sub

Sub SendWhosOnline(ByVal Index As Long)
Dim s As String
Dim N As Long, I As Long

    s = ""
    N = 0
    For I = 1 To MAX_PLAYERS
        If IsPlaying(I) And I <> Index Then
            s = s & GetPlayerName(I) & ", "
            N = N + 1
        End If
    Next I
            
    If N = 0 Then
        s = "There are no other players online."
    Else
        s = Mid(s, 1, Len(s) - 2)
        s = "There are " & N & " other players online: " & s & "."
    End If
        
    Call PlayerMsg(Index, s, WhoColor)
End Sub

Sub SendChars(ByVal Index As Long)
Dim Packet As String
Dim I As Long
    
    Packet = "ALLCHARS" & SEP_CHAR
    For I = 1 To MAX_CHARS
        Packet = Packet & Trim(Player(Index).Char(I).Name) & SEP_CHAR & Trim(Class(Player(Index).Char(I).Class).Name) & SEP_CHAR & Player(Index).Char(I).Level & SEP_CHAR
    Next I
    Packet = Packet & END_CHAR
    
    Call SendDataTo(Index, Packet)
End Sub

Sub SendJoinMap(ByVal Index As Long)
Dim Packet As String
Dim I As Long

    Packet = ""
    
    ' Send all players on current map to index
    For I = 1 To MAX_PLAYERS
        If IsPlaying(I) And I <> Index And GetPlayerMap(I) = GetPlayerMap(Index) Then
            Packet = Packet & "PLAYERDATA" & SEP_CHAR & I & SEP_CHAR & GetPlayerName(I) & SEP_CHAR & GetPlayerSprite(I) & SEP_CHAR & GetPlayerMap(I) & SEP_CHAR & GetPlayerX(I) & SEP_CHAR & GetPlayerY(I) & SEP_CHAR & GetPlayerDir(I) & SEP_CHAR & GetPlayerAccess(I) & SEP_CHAR & GetPlayerPK(I) & SEP_CHAR & END_CHAR
            Call SendDataTo(Index, Packet)
        End If
    Next I
    
    ' Send index's player data to everyone on the map including himself
    Packet = "PLAYERDATA" & SEP_CHAR & Index & SEP_CHAR & GetPlayerName(Index) & SEP_CHAR & GetPlayerSprite(Index) & SEP_CHAR & GetPlayerMap(Index) & SEP_CHAR & GetPlayerX(Index) & SEP_CHAR & GetPlayerY(Index) & SEP_CHAR & GetPlayerDir(Index) & SEP_CHAR & GetPlayerAccess(Index) & SEP_CHAR & GetPlayerPK(Index) & SEP_CHAR & END_CHAR
    Call SendDataToMap(GetPlayerMap(Index), Packet)
End Sub

Sub SendLeaveMap(ByVal Index As Long, ByVal MapNum As Long)
Dim Packet As String

    Packet = "PLAYERDATA" & SEP_CHAR & Index & SEP_CHAR & GetPlayerName(Index) & SEP_CHAR & GetPlayerSprite(Index) & SEP_CHAR & 0 & SEP_CHAR & GetPlayerX(Index) & SEP_CHAR & GetPlayerY(Index) & SEP_CHAR & GetPlayerDir(Index) & SEP_CHAR & GetPlayerAccess(Index) & SEP_CHAR & GetPlayerPK(Index) & SEP_CHAR & END_CHAR
    Call SendDataToMapBut(Index, MapNum, Packet)
End Sub

Sub SendPlayerData(ByVal Index As Long)
Dim Packet As String

    ' Send index's player data to everyone including himself on th emap
    Packet = "PLAYERDATA" & SEP_CHAR & Index & SEP_CHAR & GetPlayerName(Index) & SEP_CHAR & GetPlayerSprite(Index) & SEP_CHAR & GetPlayerMap(Index) & SEP_CHAR & GetPlayerX(Index) & SEP_CHAR & GetPlayerY(Index) & SEP_CHAR & GetPlayerDir(Index) & SEP_CHAR & GetPlayerAccess(Index) & SEP_CHAR & GetPlayerPK(Index) & SEP_CHAR & END_CHAR
    Call SendDataToMap(GetPlayerMap(Index), Packet)
End Sub

Sub SendMap(ByVal Index As Long, ByVal MapNum As Long)
Dim Packet As String, P1 As String, P2 As String
Dim X As Long
Dim Y As Long

    Packet = "MAPDATA" & SEP_CHAR & MapNum & SEP_CHAR & Trim(Map(MapNum).Name) & SEP_CHAR & Map(MapNum).Revision & SEP_CHAR & Map(MapNum).Moral & SEP_CHAR & Map(MapNum).Up & SEP_CHAR & Map(MapNum).Down & SEP_CHAR & Map(MapNum).Left & SEP_CHAR & Map(MapNum).Right & SEP_CHAR & Map(MapNum).Music & SEP_CHAR & Map(MapNum).BootMap & SEP_CHAR & Map(MapNum).BootX & SEP_CHAR & Map(MapNum).BootY & SEP_CHAR & Map(MapNum).Shop & SEP_CHAR
    
    For Y = 0 To MAX_MAPY
        For X = 0 To MAX_MAPX
            With Map(MapNum).Tile(X, Y)
                Packet = Packet & .Ground & SEP_CHAR & .Mask & SEP_CHAR & .Anim & SEP_CHAR & .Fringe & SEP_CHAR & .Type & SEP_CHAR & .Data1 & SEP_CHAR & .Data2 & SEP_CHAR & .Data3 & SEP_CHAR
            End With
        Next X
    Next Y
    
    For X = 1 To MAX_MAP_NPCS
        Packet = Packet & Map(MapNum).Npc(X) & SEP_CHAR
    Next X
    
    Packet = Packet & END_CHAR
    
    X = Int(Len(Packet) / 2)
    P1 = Mid(Packet, 1, X)
    P2 = Mid(Packet, X + 1, Len(Packet) - X)
    Call SendDataTo(Index, Packet)
End Sub

Sub SendMapItemsTo(ByVal Index As Long, ByVal MapNum As Long)
Dim Packet As String
Dim I As Long

    Packet = "MAPITEMDATA" & SEP_CHAR
    For I = 1 To MAX_MAP_ITEMS
        Packet = Packet & MapItem(MapNum, I).Num & SEP_CHAR & MapItem(MapNum, I).Value & SEP_CHAR & MapItem(MapNum, I).Dur & SEP_CHAR & MapItem(MapNum, I).X & SEP_CHAR & MapItem(MapNum, I).Y & SEP_CHAR
    Next I
    Packet = Packet & END_CHAR
    
    Call SendDataTo(Index, Packet)
End Sub

Sub SendMapItemsToAll(ByVal MapNum As Long)
Dim Packet As String
Dim I As Long

    Packet = "MAPITEMDATA" & SEP_CHAR
    For I = 1 To MAX_MAP_ITEMS
        Packet = Packet & MapItem(MapNum, I).Num & SEP_CHAR & MapItem(MapNum, I).Value & SEP_CHAR & MapItem(MapNum, I).Dur & SEP_CHAR & MapItem(MapNum, I).X & SEP_CHAR & MapItem(MapNum, I).Y & SEP_CHAR
    Next I
    Packet = Packet & END_CHAR
    
    Call SendDataToMap(MapNum, Packet)
End Sub

Sub SendMapNpcsTo(ByVal Index As Long, ByVal MapNum As Long)
Dim Packet As String
Dim I As Long

    Packet = "MAPNPCDATA" & SEP_CHAR
    For I = 1 To MAX_MAP_NPCS
        Packet = Packet & MapNpc(MapNum, I).Num & SEP_CHAR & MapNpc(MapNum, I).X & SEP_CHAR & MapNpc(MapNum, I).Y & SEP_CHAR & MapNpc(MapNum, I).Dir & SEP_CHAR
    Next I
    Packet = Packet & END_CHAR
    
    Call SendDataTo(Index, Packet)
End Sub

Sub SendMapNpcsToMap(ByVal MapNum As Long)
Dim Packet As String
Dim I As Long

    Packet = "MAPNPCDATA" & SEP_CHAR
    For I = 1 To MAX_MAP_NPCS
        Packet = Packet & MapNpc(MapNum, I).Num & SEP_CHAR & MapNpc(MapNum, I).X & SEP_CHAR & MapNpc(MapNum, I).Y & SEP_CHAR & MapNpc(MapNum, I).Dir & SEP_CHAR
    Next I
    Packet = Packet & END_CHAR
    
    Call SendDataToMap(MapNum, Packet)
End Sub

Sub SendItems(ByVal Index As Long)
Dim Packet As String
Dim I As Long

'Just Send the item cache :P
Call SendDataTo(Index, Cache(1).nCache)

'For i = 1 To MAX_ITEMS
'    If Trim(Item(i).Name) <> "" Then
'        Call SendUpdateItemTo(Index, i)
'    End If
'Next i
End Sub

Sub SendNpcs(ByVal Index As Long)
Dim Packet As String
Dim I As Long

'Just Send the item cache :P
Call SendDataTo(Index, Cache(2).nCache)

'For i = 1 To MAX_NPCS
'    If Trim(Npc(i).Name) <> "" Then
'        Call SendUpdateNpcTo(Index, i)
'    End If
'Next i
End Sub

Sub SendInventory(ByVal Index As Long)
Dim Packet As String
Dim I As Long

    Packet = "PLAYERINV" & SEP_CHAR
    For I = 1 To MAX_INV
        Packet = Packet & GetPlayerInvItemNum(Index, I) & SEP_CHAR & GetPlayerInvItemValue(Index, I) & SEP_CHAR & GetPlayerInvItemDur(Index, I) & SEP_CHAR
    Next I
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
    
    Packet = "PLAYERWORNEQ" & SEP_CHAR & GetPlayerArmorSlot(Index) & SEP_CHAR & GetPlayerWeaponSlot(Index) & SEP_CHAR & GetPlayerHelmetSlot(Index) & SEP_CHAR & GetPlayerShieldSlot(Index) & SEP_CHAR & END_CHAR
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
    
    Packet = "PLAYERSTATS" & SEP_CHAR & GetPlayerSTR(Index) & SEP_CHAR & GetPlayerDEF(Index) & SEP_CHAR & GetPlayerSPEED(Index) & SEP_CHAR & GetPlayerMAGI(Index) & SEP_CHAR & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendWelcome(ByVal Index As Long)
Dim F As Long

    ' Send them welcome
    Call PlayerMsg(Index, "Welcome to " & GAME_NAME & " version " & CLIENT_MAJOR & "." & CLIENT_MINOR & "." & CLIENT_REVISION, BrightBlue)
    Call PlayerMsg(Index, "Type /help for help on commands.  Use arrow keys to move, hold down shift to run, and use ctrl to attack.", Cyan)
    ' Send them MOTD
    If MOTD <> "" Then
        Call PlayerMsg(Index, "MOTD: " & MOTD, BrightCyan)
    End If
    
    ' Send whos online
    Call SendWhosOnline(Index)
End Sub

Sub SendClasses(ByVal Index As Long)
Dim Packet As String
Dim I As Long

    Packet = "CLASSESDATA" & SEP_CHAR & Max_Classes & SEP_CHAR
    For I = 0 To Max_Classes
        Packet = Packet & GetClassName(I) & SEP_CHAR & GetClassMaxHP(I) & SEP_CHAR & GetClassMaxMP(I) & SEP_CHAR & GetClassMaxSP(I) & SEP_CHAR & Class(I).str & SEP_CHAR & Class(I).DEF & SEP_CHAR & Class(I).SPEED & SEP_CHAR & Class(I).MAGI & SEP_CHAR
    Next I
    Packet = Packet & END_CHAR
    
    Call SendDataTo(Index, Packet)
End Sub

Sub SendNewCharClasses(ByVal Index As Long)
Dim Packet As String
Dim I As Long

    Packet = "NEWCHARCLASSES" & SEP_CHAR & Max_Classes & SEP_CHAR
    For I = 0 To Max_Classes
        Packet = Packet & GetClassName(I) & SEP_CHAR & GetClassMaxHP(I) & SEP_CHAR & GetClassMaxMP(I) & SEP_CHAR & GetClassMaxSP(I) & SEP_CHAR & Class(I).str & SEP_CHAR & Class(I).DEF & SEP_CHAR & Class(I).SPEED & SEP_CHAR & Class(I).MAGI & SEP_CHAR
    Next I
    Packet = Packet & END_CHAR
    
    Call SendDataTo(Index, Packet)
End Sub

Sub SendLeftGame(ByVal Index As Long)
Dim Packet As String

    Packet = "PLAYERDATA" & SEP_CHAR & Index & SEP_CHAR & "" & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & 0 & END_CHAR
    Call SendDataToAllBut(Index, Packet)
End Sub

Sub SendPlayerXY(ByVal Index As Long)
Dim Packet As String

    Packet = "PLAYERXY" & SEP_CHAR & GetPlayerX(Index) & SEP_CHAR & GetPlayerY(Index) & SEP_CHAR & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendUpdateItemToAll(ByVal ItemNum As Long)
Dim Packet As String

Packet = "UPDATEITEM" & SEP_CHAR & ItemNum & SEP_CHAR & Trim(Item(ItemNum).Name) & SEP_CHAR & Item(ItemNum).Pic & SEP_CHAR & Item(ItemNum).Type & SEP_CHAR & Item(ItemNum).Data1 & SEP_CHAR & Item(ItemNum).Data2 & SEP_CHAR & Item(ItemNum).Data3 & SEP_CHAR & END_CHAR
Call SendDataToAll(Packet)
End Sub

Sub SendUpdateItemTo(ByVal Index As Long, ByVal ItemNum As Long)
Dim Packet As String

Packet = "UPDATEITEM" & SEP_CHAR & ItemNum & SEP_CHAR & Trim(Item(ItemNum).Name) & SEP_CHAR & Item(ItemNum).Pic & SEP_CHAR & Item(ItemNum).Type & SEP_CHAR & END_CHAR
Call SendDataTo(Index, Packet)
End Sub

Sub SendEditItemTo(ByVal Index As Long, ByVal ItemNum As Long)
Dim Packet As String

Packet = "EDITITEM" & SEP_CHAR & ItemNum & SEP_CHAR & Trim(Item(ItemNum).Name) & SEP_CHAR & Item(ItemNum).Pic & SEP_CHAR & Item(ItemNum).Type & SEP_CHAR & Item(ItemNum).Data1 & SEP_CHAR & Item(ItemNum).Data2 & SEP_CHAR & Item(ItemNum).Data3 & SEP_CHAR & Item(ItemNum).Unbreakable & SEP_CHAR & Item(ItemNum).Locked & SEP_CHAR & Item(ItemNum).Disabled & SEP_CHAR & GetItemAssigned(ItemNum) & SEP_CHAR & END_CHAR
Call SendDataTo(Index, Packet)
End Sub

Sub SendUpdateNpcToAll(ByVal NpcNum As Long)
Dim Packet As String

    Packet = "UPDATENPC" & SEP_CHAR & NpcNum & SEP_CHAR & Trim(Npc(NpcNum).Name) & SEP_CHAR & Npc(NpcNum).Sprite & SEP_CHAR & END_CHAR
    Call SendDataToAll(Packet)
End Sub

Sub SendUpdateNpcTo(ByVal Index As Long, ByVal NpcNum As Long)
Dim Packet As String

    Packet = "UPDATENPC" & SEP_CHAR & NpcNum & SEP_CHAR & Trim(Npc(NpcNum).Name) & SEP_CHAR & Npc(NpcNum).Sprite & SEP_CHAR & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendEditNpcTo(ByVal Index As Long, ByVal NpcNum As Long)
Dim Packet As String

    Packet = "EDITNPC" & SEP_CHAR & NpcNum & SEP_CHAR & Trim(Npc(NpcNum).Name) & SEP_CHAR & Trim(Npc(NpcNum).AttackSay) & SEP_CHAR & Npc(NpcNum).Sprite & SEP_CHAR & Npc(NpcNum).SpawnSecs & SEP_CHAR & Npc(NpcNum).Behavior & SEP_CHAR & Npc(NpcNum).Range & SEP_CHAR & Npc(NpcNum).DropChance & SEP_CHAR & Npc(NpcNum).DropItem & SEP_CHAR & Npc(NpcNum).DropItemValue & SEP_CHAR & Npc(NpcNum).str & SEP_CHAR & Npc(NpcNum).DEF & SEP_CHAR & Npc(NpcNum).SPEED & SEP_CHAR & Npc(NpcNum).MAGI & SEP_CHAR & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendShops(ByVal Index As Long)
Dim I As Long

    For I = 1 To MAX_SHOPS
        If Trim(Shop(I).Name) <> "" Then
            Call SendUpdateShopTo(Index, I)
        End If
    Next I
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
Dim I As Long

    Packet = "EDITSHOP" & SEP_CHAR & ShopNum & SEP_CHAR & Trim(Shop(ShopNum).Name) & SEP_CHAR & Trim(Shop(ShopNum).JoinSay) & SEP_CHAR & Trim(Shop(ShopNum).LeaveSay) & SEP_CHAR & Shop(ShopNum).FixesItems & SEP_CHAR
    For I = 1 To MAX_TRADES
        Packet = Packet & Shop(ShopNum).TradeItem(I).GiveItem & SEP_CHAR & Shop(ShopNum).TradeItem(I).GiveValue & SEP_CHAR & Shop(ShopNum).TradeItem(I).GetItem & SEP_CHAR & Shop(ShopNum).TradeItem(I).GetValue & SEP_CHAR
    Next I
    Packet = Packet & END_CHAR

    Call SendDataTo(Index, Packet)
End Sub

Sub SendSpells(ByVal Index As Long)
Dim I As Long

    For I = 1 To MAX_SPELLS
        If Trim(Spell(I).Name) <> "" Then
            Call SendUpdateSpellTo(Index, I)
        End If
    Next I
End Sub

Sub SendUpdateSpellToAll(ByVal SpellNum As Long)
Dim Packet As String

    Packet = "UPDATESPELL" & SEP_CHAR & SpellNum & SEP_CHAR & Trim(Spell(SpellNum).Name) & SEP_CHAR & END_CHAR
    Call SendDataToAll(Packet)
End Sub

Sub SendUpdateSpellTo(ByVal Index As Long, ByVal SpellNum As Long)
Dim Packet As String

    Packet = "UPDATESPELL" & SEP_CHAR & SpellNum & SEP_CHAR & Trim(Spell(SpellNum).Name) & SEP_CHAR & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendEditSpellTo(ByVal Index As Long, ByVal SpellNum As Long)
Dim Packet As String

    Packet = "EDITSPELL" & SEP_CHAR & SpellNum & SEP_CHAR & Trim(Spell(SpellNum).Name) & SEP_CHAR & Spell(SpellNum).ClassReq & SEP_CHAR & Spell(SpellNum).LevelReq & SEP_CHAR & Spell(SpellNum).Type & SEP_CHAR & Spell(SpellNum).Data1 & SEP_CHAR & Spell(SpellNum).Data2 & SEP_CHAR & Spell(SpellNum).Data3 & SEP_CHAR & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendTrade(ByVal Index As Long, ByVal ShopNum As Long)
Dim Packet As String
Dim I As Long, X As Long, Y As Long

    Packet = "TRADE" & SEP_CHAR & ShopNum & SEP_CHAR & Shop(ShopNum).FixesItems & SEP_CHAR
    For I = 1 To MAX_TRADES
        Packet = Packet & Shop(ShopNum).TradeItem(I).GiveItem & SEP_CHAR & Shop(ShopNum).TradeItem(I).GiveValue & SEP_CHAR & Shop(ShopNum).TradeItem(I).GetItem & SEP_CHAR & Shop(ShopNum).TradeItem(I).GetValue & SEP_CHAR
        
        ' Item #
        X = Shop(ShopNum).TradeItem(I).GetItem
        
        If Item(X).Type = ITEM_TYPE_SPELL Then
            ' Spell class requirement
            Y = Spell(Item(X).Data1).ClassReq
            
            If Y = 0 Then
                Call PlayerMsg(Index, Trim(Item(X).Name) & " can be used by all classes.", Yellow)
            Else
                Call PlayerMsg(Index, Trim(Item(X).Name) & " can only be used by a " & GetClassName(Y - 1) & ".", Yellow)
            End If
        End If
    Next I
    Packet = Packet & END_CHAR
    
    Call SendDataTo(Index, Packet)
End Sub

Sub SendPlayerSpells(ByVal Index As Long)
Dim Packet As String
Dim I As Long

    Packet = "SPELLS" & SEP_CHAR
    For I = 1 To MAX_PLAYER_SPELLS
        Packet = Packet & GetPlayerSpell(Index, I) & SEP_CHAR
    Next I
    Packet = Packet & END_CHAR
    
    Call SendDataTo(Index, Packet)
End Sub

Sub SendWeatherTo(ByVal Index As Long)
Dim Packet As String

    Packet = "WEATHER" & SEP_CHAR & GameWeather & SEP_CHAR & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendWeatherToAll()
Dim I As Long

    For I = 1 To MAX_PLAYERS
        If IsPlaying(I) Then
            Call SendWeatherTo(I)
        End If
    Next I
End Sub

Sub SendTimeTo(ByVal Index As Long)
Dim Packet As String

    Packet = "TIME" & SEP_CHAR & GameTime & SEP_CHAR & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendTimeToAll()
Dim I As Long

    For I = 1 To MAX_PLAYERS
        If IsPlaying(I) Then
            Call SendTimeTo(I)
        End If
    Next I
End Sub

