Attribute VB_Name = "modServerTCP"
Option Explicit

Sub UpdateCaption()
    frmServer.Caption = "Mirage Server <IP " & frmServer.Socket(0).LocalIP & " Port " & STR(frmServer.Socket(0).LocalPort) & "> (" & TotalOnlinePlayers & ")"
End Sub

Function IsConnected(ByVal index As Long) As Boolean
    If frmServer.Socket(index).State = sckConnected Then
        IsConnected = True
    Else
        IsConnected = False
    End If
End Function

Function IsPlaying(ByVal index As Long) As Boolean
    If IsConnected(index) And Player(index).InGame = True Then
        IsPlaying = True
    Else
        IsPlaying = False
    End If
End Function

Function IsLoggedIn(ByVal index As Long) As Boolean
    If IsConnected(index) And Trim(Player(index).Login) <> "" Then
        IsLoggedIn = True
    Else
        IsLoggedIn = False
    End If
End Function

Function IsMultiAccounts(ByVal Login As String) As Boolean
Dim i As Long

    IsMultiAccounts = False
    For i = 1 To MAX_PLAYERS
        If IsConnected(i) And LCase(Trim(Player(i).Login)) = LCase(Trim(Login)) Then
            IsMultiAccounts = True
            Exit Function
        End If
    Next i
End Function

Function IsMultiIPOnline(ByVal IP As String) As Boolean
Dim i As Long
Dim n As Long

    n = 0
    IsMultiIPOnline = False
    For i = 1 To MAX_PLAYERS
        If IsConnected(i) And Trim(GetPlayerIP(i)) = Trim(IP) Then
            n = n + 1
            
            If (n > 1) Then
                IsMultiIPOnline = True
                Exit Function
            End If
        End If
    Next i
End Function

Function IsBanned(ByVal IP As String) As Boolean
Dim FileName As String, fIP As String, fName As String
Dim f As Long

    IsBanned = False
    
    FileName = App.Path & "\banlist.txt"
    
    ' Check if file exists
    If Not FileExist("banlist.txt") Then
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

Sub SendDataTo(ByVal index As Long, ByVal Data As String)
Dim i As Long, n As Long, startc As Long

    If IsConnected(index) Then
        frmServer.Socket(index).SendData Data
        DoEvents
    End If
End Sub

Sub SendDataToAll(ByVal Data As String)
Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            Call SendDataTo(i, Data)
        End If
    Next i
End Sub

Sub SendDataToAllBut(ByVal index As Long, ByVal Data As String)
Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) And i <> index Then
            Call SendDataTo(i, Data)
        End If
    Next i
End Sub

Sub SendDataToMap(ByVal MapNum As Long, ByVal Data As String)
Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            If GetPlayerMap(i) = MapNum Then
                Call SendDataTo(i, Data)
            End If
        End If
    Next i
End Sub

Sub SendDataToMapBut(ByVal index As Long, ByVal MapNum As Long, ByVal Data As String)
Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            If GetPlayerMap(i) = MapNum And i <> index Then
                Call SendDataTo(i, Data)
            End If
        End If
    Next i
End Sub

Sub GlobalMsg(ByVal Msg As String, ByVal Color As Byte)
Dim Packet As String

    Packet = "GLOBALMSG" & SEP_CHAR & Msg & SEP_CHAR & Color & END_CHAR
    
    Call SendDataToAll(Packet)
End Sub

Sub AdminMsg(ByVal Msg As String, ByVal Color As Byte)
Dim Packet As String
Dim i As Long

    Packet = "ADMINMSG" & SEP_CHAR & Msg & SEP_CHAR & Color & END_CHAR
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) And GetPlayerAccess(i) > 0 Then
            Call SendDataTo(i, Packet)
        End If
    Next i
End Sub

Sub PlayerMsg(ByVal index As Long, ByVal Msg As String, ByVal Color As Byte)
Dim Packet As String

    Packet = "PLAYERMSG" & SEP_CHAR & Msg & SEP_CHAR & Color & END_CHAR
    
    Call SendDataTo(index, Packet)
End Sub

Sub MapMsg(ByVal MapNum As Long, ByVal Msg As String, ByVal Color As Byte)
Dim Packet As String
Dim Text As String

    Packet = "MAPMSG" & SEP_CHAR & Msg & SEP_CHAR & Color & END_CHAR
    
    Call SendDataToMap(MapNum, Packet)
End Sub

Sub AlertMsg(ByVal index As Long, ByVal Msg As String)
Dim Packet As String

    Packet = "ALERTMSG" & SEP_CHAR & Msg & END_CHAR
    
    Call SendDataTo(index, Packet)
    Call CloseSocket(index)
End Sub

Sub HackingAttempt(ByVal index As Long, ByVal Reason As String)
    If index > 0 Then
        If IsPlaying(index) Then
            Call GlobalMsg(GetPlayerLogin(index) & "/" & GetPlayerName(index) & " has been booted for (" & Reason & ")", White)
        End If
    
        Call AlertMsg(index, "You have lost your connection with " & GAME_NAME & ".")
    End If
End Sub

Sub AcceptConnection(ByVal index As Long, ByVal SocketId As Long)
Dim i As Long

    If (index = 0) Then
        i = FindOpenPlayerSlot
        
        If i <> 0 Then
            ' Whoho, we can connect them
            frmServer.Socket(i).Close
            frmServer.Socket(i).Accept SocketId
            Call SocketConnected(i)
        End If
    End If
End Sub

Sub SocketConnected(ByVal index As Long)
    If index <> 0 Then
        ' Are they trying to connect more then one connection?
        'If Not IsMultiIPOnline(GetPlayerIP(Index)) Then
            If Not IsBanned(GetPlayerIP(index)) Then
                Call TextAdd(frmServer.txtText, "Received connection from " & GetPlayerIP(index) & ".", True)
            Else
                Call AlertMsg(index, "You have been banned from " & GAME_NAME & ", and can no longer play.")
            End If
        'Else
           ' Tried multiple connections
        '    Call AlertMsg(Index, GAME_NAME & " does not allow multiple IP's anymore.")
        'End If
    End If
End Sub

Sub IncomingData(ByVal index As Long, ByVal DataLength As Long)
On Error Resume Next

Dim Buffer As String
Dim Packet As String
Dim top As String * 3
Dim Start As Integer

    If index > 0 Then
        frmServer.Socket(index).GetData Buffer, vbString, DataLength
        
        If Buffer = "top" Then
            top = STR(TotalOnlinePlayers)
            Call SendDataTo(index, top)
            Call CloseSocket(index)
        End If
            
        Player(index).Buffer = Player(index).Buffer & Buffer
        
        Start = InStr(Player(index).Buffer, END_CHAR)
        Do While Start > 0
            Packet = Mid(Player(index).Buffer, 1, Start - 1)
            Player(index).Buffer = Mid(Player(index).Buffer, Start + 1, Len(Player(index).Buffer))
            Player(index).DataPackets = Player(index).DataPackets + 1
            Start = InStr(Player(index).Buffer, END_CHAR)
            If Len(Packet) > 0 Then
                Call HandleData(index, Packet)
            End If
        Loop
                
        ' Check if elapsed time has passed
        Player(index).DataBytes = Player(index).DataBytes + DataLength
        If GetTickCount >= Player(index).DataTimer + 1000 Then
            Player(index).DataTimer = GetTickCount
            Player(index).DataBytes = 0
            Player(index).DataPackets = 0
            Exit Sub
        End If
        
        ' Check for data flooding
        If Player(index).DataBytes > 1000 And GetPlayerAccess(index) <= 0 Then
            Call HackingAttempt(index, "Data Flooding")
            Exit Sub
        End If
        
        ' Check for packet flooding
        If Player(index).DataPackets > 25 And GetPlayerAccess(index) <= 0 Then
            Call HackingAttempt(index, "Packet Flooding")
            Exit Sub
        End If
    End If
End Sub



Sub CloseSocket(ByVal index As Long)
    ' Make sure player was/is playing the game, and if so, save'm.
    If index > 0 Then
        Call LeftGame(index)
    
        Call TextAdd(frmServer.txtText, "Connection from " & GetPlayerIP(index) & " has been terminated.", True)
        
        frmServer.Socket(index).Close
            
        Call UpdateCaption
        Call ClearPlayer(index)
    End If
End Sub

Sub SendWhosOnline(ByVal index As Long)
Dim s As String
Dim n As Long, i As Long

    s = ""
    n = 0
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) And i <> index Then
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
        
    Call PlayerMsg(index, s, WhoColor)
End Sub

Sub SendChars(ByVal index As Long)
Dim Packet As String
Dim i As Long
    
    Packet = "ALLCHARS" & SEP_CHAR
    For i = 1 To MAX_CHARS
        Packet = Packet & Trim(Player(index).Char(i).Name) & SEP_CHAR & Trim(Class(Player(index).Char(i).Class).Name) & SEP_CHAR & Player(index).Char(i).Level & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR
    
    Call SendDataTo(index, Packet)
End Sub

Sub SendJoinMap(ByVal index As Long)
Dim Packet As String
Dim i As Long

    Packet = ""
    
    ' Send all players on current map to index
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) And i <> index And GetPlayerMap(i) = GetPlayerMap(index) Then
            Packet = Packet & "PLAYERDATA" & SEP_CHAR & i & SEP_CHAR & GetPlayerName(i) & SEP_CHAR & GetPlayerSprite(i) & SEP_CHAR & GetPlayerMap(i) & SEP_CHAR & GetPlayerX(i) & SEP_CHAR & GetPlayerY(i) & SEP_CHAR & GetPlayerDir(i) & SEP_CHAR & GetPlayerAccess(i) & SEP_CHAR & GetPlayerPK(i) & END_CHAR
            Call SendDataTo(index, Packet)
        End If
    Next i
    
    ' Send index's player data to everyone on the map including himself
    Packet = "PLAYERDATA" & SEP_CHAR & index & SEP_CHAR & GetPlayerName(index) & SEP_CHAR & GetPlayerSprite(index) & SEP_CHAR & GetPlayerMap(index) & SEP_CHAR & GetPlayerX(index) & SEP_CHAR & GetPlayerY(index) & SEP_CHAR & GetPlayerDir(index) & SEP_CHAR & GetPlayerAccess(index) & SEP_CHAR & GetPlayerPK(index) & END_CHAR
    Call SendDataToMap(GetPlayerMap(index), Packet)
End Sub

Sub SendLeaveMap(ByVal index As Long, ByVal MapNum As Long)
Dim Packet As String

    Packet = "PLAYERDATA" & SEP_CHAR & index & SEP_CHAR & GetPlayerName(index) & SEP_CHAR & GetPlayerSprite(index) & SEP_CHAR & 0 & SEP_CHAR & GetPlayerX(index) & SEP_CHAR & GetPlayerY(index) & SEP_CHAR & GetPlayerDir(index) & SEP_CHAR & GetPlayerAccess(index) & SEP_CHAR & GetPlayerPK(index) & END_CHAR
    Call SendDataToMapBut(index, MapNum, Packet)
End Sub

Sub SendPlayerData(ByVal index As Long)
Dim Packet As String

    ' Send index's player data to everyone including himself on th emap
    Packet = "PLAYERDATA" & SEP_CHAR & index & SEP_CHAR & GetPlayerName(index) & SEP_CHAR & GetPlayerSprite(index) & SEP_CHAR & GetPlayerMap(index) & SEP_CHAR & GetPlayerX(index) & SEP_CHAR & GetPlayerY(index) & SEP_CHAR & GetPlayerDir(index) & SEP_CHAR & GetPlayerAccess(index) & SEP_CHAR & GetPlayerPK(index) & END_CHAR
    Call SendDataToMap(GetPlayerMap(index), Packet)
End Sub

Sub SendMap(ByVal index As Long, ByVal MapNum As Long)
Dim Packet As String, P1 As String, P2 As String
Dim x As Long
Dim y As Long

    Packet = "MAPDATA" & SEP_CHAR & MapNum & SEP_CHAR & Trim(Map(MapNum).Name) & SEP_CHAR & Map(MapNum).Revision & SEP_CHAR & Map(MapNum).Moral & SEP_CHAR & Map(MapNum).Up & SEP_CHAR & Map(MapNum).Down & SEP_CHAR & Map(MapNum).Left & SEP_CHAR & Map(MapNum).Right & SEP_CHAR & Map(MapNum).Music & SEP_CHAR & Map(MapNum).BootMap & SEP_CHAR & Map(MapNum).BootX & SEP_CHAR & Map(MapNum).BootY & SEP_CHAR & Map(MapNum).Shop & SEP_CHAR
    
  For y = 0 To MAX_MAPY
For x = 0 To MAX_MAPX
With Map(MapNum).Tile(x, y)
Packet = Packet & .Ground & SEP_CHAR & .Mask & SEP_CHAR & .Anim & SEP_CHAR & .Mask2 & SEP_CHAR & .M2Anim & SEP_CHAR & .Fringe & SEP_CHAR & .FAnim & SEP_CHAR & .Fringe2 & SEP_CHAR & .F2Anim & SEP_CHAR & .Type & SEP_CHAR & .Data1 & SEP_CHAR & .Data2 & SEP_CHAR & .Data3 & SEP_CHAR
End With
Next x
Next y
    
    For x = 1 To MAX_MAP_NPCS
        Packet = Packet & Map(MapNum).Npc(x) & SEP_CHAR
    Next x
    
    Packet = Packet & END_CHAR
    
    x = Int(Len(Packet) / 2)
    P1 = Mid(Packet, 1, x)
    P2 = Mid(Packet, x + 1, Len(Packet) - x)
    Call SendDataTo(index, Packet)
End Sub

Sub SendMapItemsTo(ByVal index As Long, ByVal MapNum As Long)
Dim Packet As String
Dim i As Long

    Packet = "MAPITEMDATA" & SEP_CHAR
    For i = 1 To MAX_MAP_ITEMS
        Packet = Packet & MapItem(MapNum, i).Num & SEP_CHAR & MapItem(MapNum, i).Value & SEP_CHAR & MapItem(MapNum, i).Dur & SEP_CHAR & MapItem(MapNum, i).x & SEP_CHAR & MapItem(MapNum, i).y & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR
    
    Call SendDataTo(index, Packet)
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

Sub SendMapNpcsTo(ByVal index As Long, ByVal MapNum As Long)
Dim Packet As String
Dim i As Long

    Packet = "MAPNPCDATA" & SEP_CHAR
    For i = 1 To MAX_MAP_NPCS
        Packet = Packet & MapNpc(MapNum, i).Num & SEP_CHAR & MapNpc(MapNum, i).x & SEP_CHAR & MapNpc(MapNum, i).y & SEP_CHAR & MapNpc(MapNum, i).Dir & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR
    
    Call SendDataTo(index, Packet)
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

Sub SendItems(ByVal index As Long)
Dim Packet As String
Dim i As Long

    For i = 1 To MAX_ITEMS
        If Trim(Item(i).Name) <> "" Then
            Call SendUpdateItemTo(index, i)
        End If
    Next i
End Sub

Sub SendNpcs(ByVal index As Long)
Dim Packet As String
Dim i As Long

    For i = 1 To MAX_NPCS
        If Trim(Npc(i).Name) <> "" Then
            Call SendUpdateNpcTo(index, i)
        End If
    Next i
End Sub

Sub SendInventory(ByVal index As Long)
Dim Packet As String
Dim i As Long

    Packet = "PLAYERINV" & SEP_CHAR
    For i = 1 To MAX_INV
        Packet = Packet & GetPlayerInvItemNum(index, i) & SEP_CHAR & GetPlayerInvItemValue(index, i) & SEP_CHAR & GetPlayerInvItemDur(index, i) & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR
    
    Call SendDataTo(index, Packet)
End Sub

Sub SendInventoryUpdate(ByVal index As Long, ByVal InvSlot As Long)
Dim Packet As String
    
    Packet = "PLAYERINVUPDATE" & SEP_CHAR & InvSlot & SEP_CHAR & GetPlayerInvItemNum(index, InvSlot) & SEP_CHAR & GetPlayerInvItemValue(index, InvSlot) & SEP_CHAR & GetPlayerInvItemDur(index, InvSlot) & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendWornEquipment(ByVal index As Long)
Dim Packet As String
    
    Packet = "PLAYERWORNEQ" & SEP_CHAR & GetPlayerArmorSlot(index) & SEP_CHAR & GetPlayerWeaponSlot(index) & SEP_CHAR & GetPlayerHelmetSlot(index) & SEP_CHAR & GetPlayerShieldSlot(index) & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendHP(ByVal index As Long)
Dim Packet As String

    Packet = "PLAYERHP" & SEP_CHAR & GetPlayerMaxHP(index) & SEP_CHAR & GetPlayerHP(index) & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendMP(ByVal index As Long)
Dim Packet As String

    Packet = "PLAYERMP" & SEP_CHAR & GetPlayerMaxMP(index) & SEP_CHAR & GetPlayerMP(index) & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendSP(ByVal index As Long)
Dim Packet As String

    Packet = "PLAYERSP" & SEP_CHAR & GetPlayerMaxSP(index) & SEP_CHAR & GetPlayerSP(index) & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendStats(ByVal index As Long)
Dim Packet As String
    
    Packet = "PLAYERSTATS" & SEP_CHAR & GetPlayerSTR(index) & SEP_CHAR & GetPlayerDEF(index) & SEP_CHAR & GetPlayerSPEED(index) & SEP_CHAR & GetPlayerMAGI(index) & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendWelcome(ByVal index As Long)
Dim MOTD As String
Dim f As Long

    ' Send them welcome
    Call PlayerMsg(index, "Welcome to " & GAME_NAME & "!  Programmed from scratch by yours truely Consty!  Version " & CLIENT_MAJOR & "." & CLIENT_MINOR & "." & CLIENT_REVISION, BrightBlue)
    Call PlayerMsg(index, "Type /help for help on commands.  Use arrow keys to move, hold down shift to run, and use ctrl to attack.", Cyan)
    ' Send them MOTD
    MOTD = GetVar(App.Path & "\motd.ini", "MOTD", "Msg")
    If Trim(MOTD) <> "" Then
        Call PlayerMsg(index, "MOTD: " & MOTD, BrightCyan)
    End If
    
    ' Send whos online
    Call SendWhosOnline(index)
End Sub

Sub SendClasses(ByVal index As Long)
Dim Packet As String
Dim i As Long

    Packet = "CLASSESDATA" & SEP_CHAR & Max_Classes & SEP_CHAR
    For i = 0 To Max_Classes
        Packet = Packet & GetClassName(i) & SEP_CHAR & GetClassMaxHP(i) & SEP_CHAR & GetClassMaxMP(i) & SEP_CHAR & GetClassMaxSP(i) & SEP_CHAR & Class(i).STR & SEP_CHAR & Class(i).DEF & SEP_CHAR & Class(i).SPEED & SEP_CHAR & Class(i).MAGI & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR
    
    Call SendDataTo(index, Packet)
End Sub

Sub SendNewCharClasses(ByVal index As Long)
Dim Packet As String
Dim i As Long

    Packet = "NEWCHARCLASSES" & SEP_CHAR & Max_Classes & SEP_CHAR
    For i = 0 To Max_Classes
        Packet = Packet & GetClassName(i) & SEP_CHAR & GetClassMaxHP(i) & SEP_CHAR & GetClassMaxMP(i) & SEP_CHAR & GetClassMaxSP(i) & SEP_CHAR & Class(i).STR & SEP_CHAR & Class(i).DEF & SEP_CHAR & Class(i).SPEED & SEP_CHAR & Class(i).MAGI & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR
    
    Call SendDataTo(index, Packet)
End Sub

Sub SendLeftGame(ByVal index As Long)
Dim Packet As String

    Packet = "PLAYERDATA" & SEP_CHAR & index & SEP_CHAR & "" & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & 0 & END_CHAR
    Call SendDataToAllBut(index, Packet)
End Sub

Sub SendPlayerXY(ByVal index As Long)
Dim Packet As String

    Packet = "PLAYERXY" & SEP_CHAR & GetPlayerX(index) & SEP_CHAR & GetPlayerY(index) & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendUpdateItemToAll(ByVal ItemNum As Long)
Dim Packet As String

    Packet = "UPDATEITEM" & SEP_CHAR & ItemNum & SEP_CHAR & Trim(Item(ItemNum).Name) & SEP_CHAR & Item(ItemNum).Pic & SEP_CHAR & Item(ItemNum).Type & SEP_CHAR & Item(ItemNum).Data1 & SEP_CHAR & Item(ItemNum).Data2 & SEP_CHAR & Item(ItemNum).Data3 & END_CHAR
    Call SendDataToAll(Packet)
End Sub

Sub SendUpdateItemTo(ByVal index As Long, ByVal ItemNum As Long)
Dim Packet As String

    Packet = "UPDATEITEM" & SEP_CHAR & ItemNum & SEP_CHAR & Trim(Item(ItemNum).Name) & SEP_CHAR & Item(ItemNum).Pic & SEP_CHAR & Item(ItemNum).Type & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendEditItemTo(ByVal index As Long, ByVal ItemNum As Long)
Dim Packet As String

    Packet = "EDITITEM" & SEP_CHAR & ItemNum & SEP_CHAR & Trim(Item(ItemNum).Name) & SEP_CHAR & Item(ItemNum).Pic & SEP_CHAR & Item(ItemNum).Type & SEP_CHAR & Item(ItemNum).Data1 & SEP_CHAR & Item(ItemNum).Data2 & SEP_CHAR & Item(ItemNum).Data3 & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendUpdateNpcToAll(ByVal npcnum As Long)
Dim Packet As String

    Packet = "UPDATENPC" & SEP_CHAR & npcnum & SEP_CHAR & Trim(Npc(npcnum).Name) & SEP_CHAR & Npc(npcnum).Sprite & END_CHAR
    Call SendDataToAll(Packet)
End Sub

Sub SendUpdateNpcTo(ByVal index As Long, ByVal npcnum As Long)
Dim Packet As String

    Packet = "UPDATENPC" & SEP_CHAR & npcnum & SEP_CHAR & Trim(Npc(npcnum).Name) & SEP_CHAR & Npc(npcnum).Sprite & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendEditNpcTo(ByVal index As Long, ByVal npcnum As Long)
Dim Packet As String

    Packet = "EDITNPC" & SEP_CHAR & npcnum & SEP_CHAR & Trim(Npc(npcnum).Name) & SEP_CHAR & Trim(Npc(npcnum).AttackSay) & SEP_CHAR & Npc(npcnum).Sprite & SEP_CHAR & Npc(npcnum).SpawnSecs & SEP_CHAR & Npc(npcnum).Behavior & SEP_CHAR & Npc(npcnum).Range & SEP_CHAR & Npc(npcnum).DropChance & SEP_CHAR & Npc(npcnum).DropItem & SEP_CHAR & Npc(npcnum).DropItemValue & SEP_CHAR & Npc(npcnum).STR & SEP_CHAR & Npc(npcnum).DEF & SEP_CHAR & Npc(npcnum).SPEED & SEP_CHAR & Npc(npcnum).MAGI & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendShops(ByVal index As Long)
Dim i As Long

    For i = 1 To MAX_SHOPS
        If Trim(Shop(i).Name) <> "" Then
            Call SendUpdateShopTo(index, i)
        End If
    Next i
End Sub

Sub SendUpdateShopToAll(ByVal ShopNum As Long)
Dim Packet As String

    Packet = "UPDATESHOP" & SEP_CHAR & ShopNum & SEP_CHAR & Trim(Shop(ShopNum).Name) & END_CHAR
    Call SendDataToAll(Packet)
End Sub

Sub SendUpdateShopTo(ByVal index As Long, ByVal ShopNum)
Dim Packet As String

    Packet = "UPDATESHOP" & SEP_CHAR & ShopNum & SEP_CHAR & Trim(Shop(ShopNum).Name) & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendEditShopTo(ByVal index As Long, ByVal ShopNum As Long)
Dim Packet As String
Dim i As Long

    Packet = "EDITSHOP" & SEP_CHAR & ShopNum & SEP_CHAR & Trim(Shop(ShopNum).Name) & SEP_CHAR & Trim(Shop(ShopNum).JoinSay) & SEP_CHAR & Trim(Shop(ShopNum).LeaveSay) & SEP_CHAR & Shop(ShopNum).FixesItems & SEP_CHAR
    For i = 1 To MAX_TRADES
        Packet = Packet & Shop(ShopNum).TradeItem(i).GiveItem & SEP_CHAR & Shop(ShopNum).TradeItem(i).GiveValue & SEP_CHAR & Shop(ShopNum).TradeItem(i).GetItem & SEP_CHAR & Shop(ShopNum).TradeItem(i).GetValue & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR

    Call SendDataTo(index, Packet)
End Sub

Sub SendSpells(ByVal index As Long)
Dim i As Long

    For i = 1 To MAX_SPELLS
        If Trim(Spell(i).Name) <> "" Then
            Call SendUpdateSpellTo(index, i)
        End If
    Next i
End Sub

Sub SendUpdateSpellToAll(ByVal SpellNum As Long)
Dim Packet As String

    Packet = "UPDATESPELL" & SEP_CHAR & SpellNum & SEP_CHAR & Trim(Spell(SpellNum).Name) & END_CHAR
    Call SendDataToAll(Packet)
End Sub

Sub SendUpdateSpellTo(ByVal index As Long, ByVal SpellNum As Long)
Dim Packet As String

    Packet = "UPDATESPELL" & SEP_CHAR & SpellNum & SEP_CHAR & Trim(Spell(SpellNum).Name) & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendEditSpellTo(ByVal index As Long, ByVal SpellNum As Long)
Dim Packet As String

    Packet = "EDITSPELL" & SEP_CHAR & SpellNum & SEP_CHAR & Trim(Spell(SpellNum).Name) & SEP_CHAR & Spell(SpellNum).ClassReq & SEP_CHAR & Spell(SpellNum).LevelReq & SEP_CHAR & Spell(SpellNum).Type & SEP_CHAR & Spell(SpellNum).Data1 & SEP_CHAR & Spell(SpellNum).Data2 & SEP_CHAR & Spell(SpellNum).Data3 & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendTrade(ByVal index As Long, ByVal ShopNum As Long)
Dim Packet As String
Dim i As Long, x As Long, y As Long

    Packet = "TRADE" & SEP_CHAR & ShopNum & SEP_CHAR & Shop(ShopNum).FixesItems & SEP_CHAR
    For i = 1 To MAX_TRADES
        Packet = Packet & Shop(ShopNum).TradeItem(i).GiveItem & SEP_CHAR & Shop(ShopNum).TradeItem(i).GiveValue & SEP_CHAR & Shop(ShopNum).TradeItem(i).GetItem & SEP_CHAR & Shop(ShopNum).TradeItem(i).GetValue & SEP_CHAR
        
        ' Item #
        x = Shop(ShopNum).TradeItem(i).GetItem
        
        If Item(x).Type = ITEM_TYPE_SPELL Then
            ' Spell class requirement
            y = Spell(Item(x).Data1).ClassReq
            
            If y = 0 Then
                Call PlayerMsg(index, Trim(Item(x).Name) & " can be used by all classes.", Yellow)
            Else
                Call PlayerMsg(index, Trim(Item(x).Name) & " can only be used by a " & GetClassName(y - 1) & ".", Yellow)
            End If
        End If
    Next i
    Packet = Packet & END_CHAR
    
    Call SendDataTo(index, Packet)
End Sub

Sub SendPlayerSpells(ByVal index As Long)
Dim Packet As String
Dim i As Long

    Packet = "SPELLS" & SEP_CHAR
    For i = 1 To MAX_PLAYER_SPELLS
        Packet = Packet & GetPlayerSpell(index, i) & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR
    
    Call SendDataTo(index, Packet)
End Sub

Sub SendWeatherTo(ByVal index As Long)
Dim Packet As String

    Packet = "WEATHER" & SEP_CHAR & GameWeather & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendWeatherToAll()
Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            Call SendWeatherTo(i)
        End If
    Next i
End Sub

Sub SendTimeTo(ByVal index As Long)
Dim Packet As String

    Packet = "TIME" & SEP_CHAR & GameTime & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendTimeToAll()
Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            Call SendTimeTo(i)
        End If
    Next i
End Sub
