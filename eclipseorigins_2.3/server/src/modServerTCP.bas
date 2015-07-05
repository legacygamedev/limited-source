Attribute VB_Name = "modServerTCP"
Option Explicit

Sub UpdateCaption()
    frmServer.Caption = Options.Game_Name & " <IP " & frmServer.Socket(0).LocalIP & " Port " & CStr(frmServer.Socket(0).LocalPort) & "> (" & TotalOnlinePlayers & ")"
End Sub

Sub CreateFullMapCache()
    Dim I As Long

    For I = 1 To MAX_MAPS
        Call MapCache_Create(I)
    Next

End Sub

Function IsConnected(ByVal index As Long) As Boolean

    If frmServer.Socket(index).State = sckConnected Then
        IsConnected = True
    End If

End Function

Function IsPlaying(ByVal index As Long) As Boolean

    If IsConnected(index) Then
        If TempPlayer(index).InGame Then
            IsPlaying = True
        End If
    End If

End Function

Function IsLoggedIn(ByVal index As Long) As Boolean

    If IsConnected(index) Then
        If LenB(Trim$(Player(index).Login)) > 0 Then
            IsLoggedIn = True
        End If
    End If

End Function

Function IsMultiAccounts(ByVal Login As String) As Boolean
    Dim I As Long

    For I = 1 To Player_HighIndex

        If IsConnected(I) Then
            If LCase$(Trim$(Player(I).Login)) = LCase$(Login) Then
                IsMultiAccounts = True
                Exit Function
            End If
        End If

    Next

End Function

Function IsMultiIPOnline(ByVal IP As String) As Boolean
    Dim I As Long
    Dim n As Long

    For I = 1 To Player_HighIndex

        If IsConnected(I) Then
            If Trim$(GetPlayerIP(I)) = IP Then
                n = n + 1

                If (n > 1) Then
                    IsMultiIPOnline = True
                    Exit Function
                End If
            End If
        End If

    Next

End Function

Function IsBanned(ByVal IP As String) As Boolean
    Dim filename As String
    Dim fIP As String
    Dim fName As String
    Dim F As Long
    filename = App.Path & "\data\banlist.txt"

    ' Check if file exists
    If Not FileExist("data\banlist.txt") Then
        F = FreeFile
        Open filename For Output As #F
        Close #F
    End If

    F = FreeFile
    Open filename For Input As #F

    Do While Not EOF(F)
        Input #F, fIP
        Input #F, fName

        ' Is banned?
        If Trim$(LCase$(fIP)) = Trim$(LCase$(Mid$(IP, 1, Len(fIP)))) Then
            IsBanned = True
            Close #F
            Exit Function
        End If

    Loop

    Close #F
End Function

Sub SendDataTo(ByVal index As Long, ByRef Data() As Byte)
Dim Buffer As clsBuffer
Dim TempData() As Byte

    If IsConnected(index) Then
        Set Buffer = New clsBuffer
        TempData = Data
        
        Buffer.PreAllocate 4 + (UBound(TempData) - LBound(TempData)) + 1
        Buffer.WriteLong (UBound(TempData) - LBound(TempData)) + 1
        Buffer.WriteBytes TempData()
              
        frmServer.Socket(index).SendData Buffer.ToArray()
    End If
End Sub

Sub SendDataToAll(ByRef Data() As Byte)
    Dim I As Long

    For I = 1 To Player_HighIndex

        If IsPlaying(I) Then
            Call SendDataTo(I, Data)
        End If

    Next

End Sub

Sub SendDataToAllBut(ByVal index As Long, ByRef Data() As Byte)
    Dim I As Long

    For I = 1 To Player_HighIndex

        If IsPlaying(I) Then
            If I <> index Then
                Call SendDataTo(I, Data)
            End If
        End If

    Next

End Sub

Sub SendDataToMap(ByVal mapNum As Long, ByRef Data() As Byte)
    Dim I As Long

    For I = 1 To Player_HighIndex

        If IsPlaying(I) Then
            If GetPlayerMap(I) = mapNum Then
                Call SendDataTo(I, Data)
            End If
        End If

    Next

End Sub

Sub SendDataToMapBut(ByVal index As Long, ByVal mapNum As Long, ByRef Data() As Byte)
    Dim I As Long

    For I = 1 To Player_HighIndex

        If IsPlaying(I) Then
            If GetPlayerMap(I) = mapNum Then
                If I <> index Then
                    Call SendDataTo(I, Data)
                End If
            End If
        End If

    Next

End Sub

Sub SendDataToParty(ByVal partyNum As Long, ByRef Data() As Byte)
Dim I As Long

    For I = 1 To Party(partyNum).MemberCount
        If Party(partyNum).Member(I) > 0 Then
            Call SendDataTo(Party(partyNum).Member(I), Data)
        End If
    Next
End Sub

Public Sub GlobalMsg(ByVal Msg As String, ByVal color As Byte)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SGlobalMsg
    Buffer.WriteString Msg
    Buffer.WriteLong color
    SendDataToAll Buffer.ToArray
    
    Set Buffer = Nothing
End Sub

Public Sub AdminMsg(ByVal Msg As String, ByVal color As Byte)
    Dim Buffer As clsBuffer
    Dim I As Long
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SAdminMsg
    Buffer.WriteString Msg
    Buffer.WriteLong color

    For I = 1 To Player_HighIndex
        If IsPlaying(I) And GetPlayerAccess(I) > 0 Then
            SendDataTo I, Buffer.ToArray
        End If
    Next
    
    Set Buffer = Nothing
End Sub

Public Sub PlayerMsg(ByVal index As Long, ByVal Msg As String, ByVal color As Byte)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SPlayerMsg
    Buffer.WriteString Msg
    Buffer.WriteLong color
    SendDataTo index, Buffer.ToArray
    
    Set Buffer = Nothing
End Sub

Public Sub MapMsg(ByVal mapNum As Long, ByVal Msg As String, ByVal color As Byte)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer

    Buffer.WriteLong SMapMsg
    Buffer.WriteString Msg
    Buffer.WriteLong color
    SendDataToMap mapNum, Buffer.ToArray
    
    Set Buffer = Nothing
End Sub

Public Sub AlertMsg(ByVal index As Long, ByVal Msg As String)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer

    Buffer.WriteLong SAlertMsg
    Buffer.WriteString Msg
    SendDataTo index, Buffer.ToArray
    DoEvents
    Call CloseSocket(index)
    
    Set Buffer = Nothing
End Sub

Public Sub PartyMsg(ByVal partyNum As Long, ByVal Msg As String, ByVal color As Byte)
Dim I As Long
    ' send message to all people
    For I = 1 To MAX_PARTY_MEMBERS
        ' exist?
        If Party(partyNum).Member(I) > 0 Then
            ' make sure they're logged on
            If IsConnected(Party(partyNum).Member(I)) And IsPlaying(Party(partyNum).Member(I)) Then
                PlayerMsg Party(partyNum).Member(I), Msg, color
            End If
        End If
    Next
End Sub

Sub HackingAttempt(ByVal index As Long, ByVal Reason As String)

    If index > 0 Then
        If IsPlaying(index) Then
            Call GlobalMsg(GetPlayerLogin(index) & "/" & GetPlayerName(index) & " has been booted for (" & Reason & ")", White)
        End If

        Call AlertMsg(index, "You have lost your connection with " & Options.Game_Name & ".")
    End If

End Sub

Sub AcceptConnection(ByVal index As Long, ByVal SocketId As Long)
    Dim I As Long

    If (index = 0) Then
        I = FindOpenPlayerSlot

        If I <> 0 Then
            ' we can connect them
            frmServer.Socket(I).Close
            frmServer.Socket(I).Accept SocketId
            Call SocketConnected(I)
        End If
    End If

End Sub

Sub SocketConnected(ByVal index As Long)
Dim I As Long

    If index <> 0 Then
        ' make sure they're not banned
        If Not IsBanned(GetPlayerIP(index)) Then
            Call TextAdd("Received connection from " & GetPlayerIP(index) & ".")
        Else
            Call AlertMsg(index, "You have been banned from " & Options.Game_Name & ", and can no longer play.")
        End If
        ' re-set the high index
        Player_HighIndex = 0
        For I = MAX_PLAYERS To 1 Step -1
            If IsConnected(I) Then
                Player_HighIndex = I
                Exit For
            End If
        Next
        ' send the new highindex to all logged in players
        SendHighIndex
    End If
End Sub

Sub IncomingData(ByVal index As Long, ByVal DataLength As Long)
Dim Buffer() As Byte
Dim pLength As Long

    If GetPlayerAccess(index) <= 0 Then
        ' Check for data flooding
        If TempPlayer(index).DataBytes > 1000 Then
            Exit Sub
        End If
    
        ' Check for packet flooding
        If TempPlayer(index).DataPackets > 25 Then
            Exit Sub
        End If
    End If
            
    ' Check if elapsed time has passed
    TempPlayer(index).DataBytes = TempPlayer(index).DataBytes + DataLength
    If GetTickCount >= TempPlayer(index).DataTimer Then
        TempPlayer(index).DataTimer = GetTickCount + 1000
        TempPlayer(index).DataBytes = 0
        TempPlayer(index).DataPackets = 0
    End If
    
    ' Get the data from the socket now
    frmServer.Socket(index).GetData Buffer(), vbUnicode, DataLength
    TempPlayer(index).Buffer.WriteBytes Buffer()
    
    If TempPlayer(index).Buffer.Length >= 4 Then
        pLength = TempPlayer(index).Buffer.ReadLong(False)
    
        If pLength < 0 Then
            Exit Sub
        End If
    End If
    
    Do While pLength > 0 And pLength <= TempPlayer(index).Buffer.Length - 4
        If pLength <= TempPlayer(index).Buffer.Length - 4 Then
            TempPlayer(index).DataPackets = TempPlayer(index).DataPackets + 1
            TempPlayer(index).Buffer.ReadLong
            HandleData index, TempPlayer(index).Buffer.ReadBytes(pLength)
        End If
        
        pLength = 0
        If TempPlayer(index).Buffer.Length >= 4 Then
            pLength = TempPlayer(index).Buffer.ReadLong(False)
        
            If pLength < 0 Then
                Exit Sub
            End If
        End If
    Loop
            
    TempPlayer(index).Buffer.Trim
End Sub

Sub CloseSocket(ByVal index As Long)

    If index > 0 Then
        Call LeftGame(index)
        Call TextAdd("Connection from " & GetPlayerIP(index) & " has been terminated.")
        frmServer.Socket(index).Close
        Call UpdateCaption
        Call ClearPlayer(index)
    End If

End Sub

Public Sub MapCache_Create(ByVal mapNum As Long)
    Dim MapData As String
    Dim x As Long
    Dim y As Long
    Dim I As Long, z As Long, w As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong mapNum
    Buffer.WriteString Trim$(Map(mapNum).Name)
    Buffer.WriteString Trim$(Map(mapNum).Music)
    Buffer.WriteLong Map(mapNum).Revision
    Buffer.WriteByte Map(mapNum).Moral
    Buffer.WriteLong Map(mapNum).Up
    Buffer.WriteLong Map(mapNum).Down
    Buffer.WriteLong Map(mapNum).Left
    Buffer.WriteLong Map(mapNum).Right
    Buffer.WriteLong Map(mapNum).BootMap
    Buffer.WriteByte Map(mapNum).BootX
    Buffer.WriteByte Map(mapNum).BootY
    Buffer.WriteByte Map(mapNum).MaxX
    Buffer.WriteByte Map(mapNum).MaxY

    For x = 0 To Map(mapNum).MaxX
        For y = 0 To Map(mapNum).MaxY

            With Map(mapNum).Tile(x, y)
                For I = 1 To MapLayer.Layer_Count - 1
                    Buffer.WriteLong .Layer(I).x
                    Buffer.WriteLong .Layer(I).y
                    Buffer.WriteLong .Layer(I).Tileset
                Next
                Buffer.WriteByte .Type
                Buffer.WriteLong .Data1
                Buffer.WriteLong .Data2
                Buffer.WriteLong .Data3
                Buffer.WriteByte .DirBlock
            End With

        Next
    Next

    For x = 1 To MAX_MAP_NPCS
        Buffer.WriteLong Map(mapNum).Npc(x)
    Next

    MapCache(mapNum).Data = Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

' *****************************
' ** Outgoing Server Packets **
' *****************************
Sub SendWhosOnline(ByVal index As Long)
    Dim s As String
    Dim n As Long
    Dim I As Long

    For I = 1 To Player_HighIndex

        If IsPlaying(I) Then
            If I <> index Then
                s = s & GetPlayerName(I) & ", "
                n = n + 1
            End If
        End If

    Next

    If n = 0 Then
        s = "There are no other players online."
    Else
        s = Mid$(s, 1, Len(s) - 2)
        s = "There are " & n & " other players online: " & s & "."
    End If

    Call PlayerMsg(index, s, WhoColor)
End Sub

Function PlayerData(ByVal index As Long) As Byte()
    Dim Buffer As clsBuffer, I As Long

    If index > MAX_PLAYERS Then Exit Function
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SPlayerData
    Buffer.WriteLong index
    Buffer.WriteString GetPlayerName(index)
    Buffer.WriteLong GetPlayerLevel(index)
    Buffer.WriteLong GetPlayerPOINTS(index)
    Buffer.WriteLong GetPlayerSprite(index)
    Buffer.WriteLong GetPlayerMap(index)
    Buffer.WriteLong GetPlayerX(index)
    Buffer.WriteLong GetPlayerY(index)
    Buffer.WriteLong GetPlayerDir(index)
    Buffer.WriteLong GetPlayerAccess(index)
    Buffer.WriteLong GetPlayerPK(index)
    Buffer.WriteLong GetPlayerClass(index)
    
    For I = 1 To Stats.Stat_Count - 1
        Buffer.WriteLong GetPlayerStat(index, I)
    Next
    
    PlayerData = Buffer.ToArray()
    Set Buffer = Nothing
End Function

Sub SendJoinMap(ByVal index As Long)
    Dim packet As String
    Dim I As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer

    ' Send all players on current map to index
    For I = 1 To Player_HighIndex
        If IsPlaying(I) Then
            If I <> index Then
                If GetPlayerMap(I) = GetPlayerMap(index) Then
                    SendDataTo index, PlayerData(I)
                End If
            End If
        End If
    Next

    ' Send index's player data to everyone on the map including himself
    SendDataToMap GetPlayerMap(index), PlayerData(index)
    
    Set Buffer = Nothing
End Sub

Sub SendLeaveMap(ByVal index As Long, ByVal mapNum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SLeft
    Buffer.WriteLong index
    SendDataToMapBut index, mapNum, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendPlayerData(ByVal index As Long)
    Dim packet As String
    SendDataToMap GetPlayerMap(index), PlayerData(index)
End Sub

Sub SendMap(ByVal index As Long, ByVal mapNum As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate (UBound(MapCache(mapNum).Data) - LBound(MapCache(mapNum).Data)) + 5
    Buffer.WriteLong SMapData
    Buffer.WriteBytes MapCache(mapNum).Data()
    SendDataTo index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendMapItemsTo(ByVal index As Long, ByVal mapNum As Long)
    Dim packet As String
    Dim I As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SMapItemData

    For I = 1 To MAX_MAP_ITEMS
        Buffer.WriteString MapItem(mapNum, I).playerName
        Buffer.WriteLong MapItem(mapNum, I).Num
        Buffer.WriteLong MapItem(mapNum, I).Value
        Buffer.WriteLong MapItem(mapNum, I).x
        Buffer.WriteLong MapItem(mapNum, I).y
    Next

    SendDataTo index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendMapItemsToAll(ByVal mapNum As Long)
    Dim packet As String
    Dim I As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SMapItemData

    For I = 1 To MAX_MAP_ITEMS
        Buffer.WriteString MapItem(mapNum, I).playerName
        Buffer.WriteLong MapItem(mapNum, I).Num
        Buffer.WriteLong MapItem(mapNum, I).Value
        Buffer.WriteLong MapItem(mapNum, I).x
        Buffer.WriteLong MapItem(mapNum, I).y
    Next

    SendDataToMap mapNum, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendMapNpcVitals(ByVal mapNum As Long, ByVal mapNpcNum As Long)
    Dim I As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SMapNpcVitals
    Buffer.WriteLong mapNpcNum
    For I = 1 To Vitals.Vital_Count - 1
        Buffer.WriteLong MapNpc(mapNum).Npc(mapNpcNum).Vital(I)
    Next

    SendDataToMap mapNum, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendMapNpcsTo(ByVal index As Long, ByVal mapNum As Long)
    Dim packet As String
    Dim I As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SMapNpcData

    For I = 1 To MAX_MAP_NPCS
        Buffer.WriteLong MapNpc(mapNum).Npc(I).Num
        Buffer.WriteLong MapNpc(mapNum).Npc(I).x
        Buffer.WriteLong MapNpc(mapNum).Npc(I).y
        Buffer.WriteLong MapNpc(mapNum).Npc(I).Dir
        Buffer.WriteLong MapNpc(mapNum).Npc(I).Vital(HP)
    Next

    SendDataTo index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendMapNpcsToMap(ByVal mapNum As Long)
    Dim packet As String
    Dim I As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SMapNpcData

    For I = 1 To MAX_MAP_NPCS
        Buffer.WriteLong MapNpc(mapNum).Npc(I).Num
        Buffer.WriteLong MapNpc(mapNum).Npc(I).x
        Buffer.WriteLong MapNpc(mapNum).Npc(I).y
        Buffer.WriteLong MapNpc(mapNum).Npc(I).Dir
        Buffer.WriteLong MapNpc(mapNum).Npc(I).Vital(HP)
    Next

    SendDataToMap mapNum, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendItems(ByVal index As Long)
    Dim I As Long

    For I = 1 To MAX_ITEMS

        If LenB(Trim$(Item(I).Name)) > 0 Then
            Call SendUpdateItemTo(index, I)
        End If

    Next

End Sub

Sub SendAnimations(ByVal index As Long)
    Dim I As Long

    For I = 1 To MAX_ANIMATIONS

        If LenB(Trim$(Animation(I).Name)) > 0 Then
            Call SendUpdateAnimationTo(index, I)
        End If

    Next

End Sub

Sub SendNpcs(ByVal index As Long)
    Dim I As Long

    For I = 1 To MAX_NPCS

        If LenB(Trim$(Npc(I).Name)) > 0 Then
            Call SendUpdateNpcTo(index, I)
        End If

    Next

End Sub

Sub SendResources(ByVal index As Long)
    Dim I As Long

    For I = 1 To MAX_RESOURCES

        If LenB(Trim$(Resource(I).Name)) > 0 Then
            Call SendUpdateResourceTo(index, I)
        End If

    Next

End Sub

Sub SendInventory(ByVal index As Long)
    Dim packet As String
    Dim I As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SPlayerInv

    For I = 1 To MAX_INV
        Buffer.WriteLong GetPlayerInvItemNum(index, I)
        Buffer.WriteLong GetPlayerInvItemValue(index, I)
    Next

    SendDataTo index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendInventoryUpdate(ByVal index As Long, ByVal invSlot As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SPlayerInvUpdate
    Buffer.WriteLong invSlot
    Buffer.WriteLong GetPlayerInvItemNum(index, invSlot)
    Buffer.WriteLong GetPlayerInvItemValue(index, invSlot)
    SendDataTo index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendWornEquipment(ByVal index As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SPlayerWornEq
    Buffer.WriteLong GetPlayerEquipment(index, Armor)
    Buffer.WriteLong GetPlayerEquipment(index, Weapon)
    Buffer.WriteLong GetPlayerEquipment(index, Helmet)
    Buffer.WriteLong GetPlayerEquipment(index, Shield)
    SendDataTo index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendMapEquipment(ByVal index As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SMapWornEq
    Buffer.WriteLong index
    Buffer.WriteLong GetPlayerEquipment(index, Armor)
    Buffer.WriteLong GetPlayerEquipment(index, Weapon)
    Buffer.WriteLong GetPlayerEquipment(index, Helmet)
    Buffer.WriteLong GetPlayerEquipment(index, Shield)
    
    SendDataToMap GetPlayerMap(index), Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendMapEquipmentTo(ByVal PlayerNum As Long, ByVal index As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SMapWornEq
    Buffer.WriteLong PlayerNum
    Buffer.WriteLong GetPlayerEquipment(PlayerNum, Armor)
    Buffer.WriteLong GetPlayerEquipment(PlayerNum, Weapon)
    Buffer.WriteLong GetPlayerEquipment(PlayerNum, Helmet)
    Buffer.WriteLong GetPlayerEquipment(PlayerNum, Shield)
    
    SendDataTo index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendVital(ByVal index As Long, ByVal Vital As Vitals)
    Dim packet As String
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer

    Select Case Vital
        Case HP
            Buffer.WriteLong SPlayerHp
            Buffer.WriteLong GetPlayerMaxVital(index, Vitals.HP)
            Buffer.WriteLong GetPlayerVital(index, Vitals.HP)
        Case MP
            Buffer.WriteLong SPlayerMp
            Buffer.WriteLong GetPlayerMaxVital(index, Vitals.MP)
            Buffer.WriteLong GetPlayerVital(index, Vitals.MP)
    End Select

    SendDataTo index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendEXP(ByVal index As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SPlayerEXP
    Buffer.WriteLong GetPlayerExp(index)
    Buffer.WriteLong GetPlayerNextLevel(index)
    
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendStats(ByVal index As Long)
Dim I As Long
Dim packet As String
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerStats
    For I = 1 To Stats.Stat_Count - 1
        Buffer.WriteLong GetPlayerStat(index, I)
    Next
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendWelcome(ByVal index As Long)

    ' Send them MOTD
    If LenB(Options.MOTD) > 0 Then
        Call PlayerMsg(index, Options.MOTD, BrightCyan)
    End If

    ' Send whos online
    Call SendWhosOnline(index)
End Sub

Sub SendClasses(ByVal index As Long)
    Dim packet As String
    Dim I As Long, n As Long, q As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SClassesData
    Buffer.WriteLong Max_Classes

    For I = 1 To Max_Classes
        Buffer.WriteString GetClassName(I)
        Buffer.WriteLong GetClassMaxVital(I, Vitals.HP)
        Buffer.WriteLong GetClassMaxVital(I, Vitals.MP)
        
        ' set sprite array size
        n = UBound(Class(I).MaleSprite)
        
        ' send array size
        Buffer.WriteLong n
        
        ' loop around sending each sprite
        For q = 0 To n
            Buffer.WriteLong Class(I).MaleSprite(q)
        Next
        
        ' set sprite array size
        n = UBound(Class(I).FemaleSprite)
        
        ' send array size
        Buffer.WriteLong n
        
        ' loop around sending each sprite
        For q = 0 To n
            Buffer.WriteLong Class(I).FemaleSprite(q)
        Next
        
        For q = 1 To Stats.Stat_Count - 1
            Buffer.WriteLong Class(I).stat(q)
        Next
    Next

    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendNewCharClasses(ByVal index As Long)
    Dim packet As String
    Dim I As Long, n As Long, q As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SNewCharClasses
    Buffer.WriteLong Max_Classes

    For I = 1 To Max_Classes
        Buffer.WriteString GetClassName(I)
        Buffer.WriteLong GetClassMaxVital(I, Vitals.HP)
        Buffer.WriteLong GetClassMaxVital(I, Vitals.MP)
        
        ' set sprite array size
        n = UBound(Class(I).MaleSprite)
        ' send array size
        Buffer.WriteLong n
        ' loop around sending each sprite
        For q = 0 To n
            Buffer.WriteLong Class(I).MaleSprite(q)
        Next
        
        ' set sprite array size
        n = UBound(Class(I).FemaleSprite)
        ' send array size
        Buffer.WriteLong n
        ' loop around sending each sprite
        For q = 0 To n
            Buffer.WriteLong Class(I).FemaleSprite(q)
        Next
        
        For q = 1 To Stats.Stat_Count - 1
            Buffer.WriteLong Class(I).stat(q)
        Next
    Next

    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendLeftGame(ByVal index As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerData
    Buffer.WriteLong index
    Buffer.WriteString vbNullString
    Buffer.WriteLong 0
    Buffer.WriteLong 0
    Buffer.WriteLong 0
    Buffer.WriteLong 0
    Buffer.WriteLong 0
    Buffer.WriteLong 0
    Buffer.WriteLong 0
    SendDataToAllBut index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPlayerXY(ByVal index As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerXY
    Buffer.WriteLong GetPlayerX(index)
    Buffer.WriteLong GetPlayerY(index)
    Buffer.WriteLong GetPlayerDir(index)
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPlayerXYToMap(ByVal index As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerXYMap
    Buffer.WriteLong index
    Buffer.WriteLong GetPlayerX(index)
    Buffer.WriteLong GetPlayerY(index)
    Buffer.WriteLong GetPlayerDir(index)
    SendDataToMap GetPlayerMap(index), Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendUpdateItemToAll(ByVal itemnum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim ItemSize As Long
    Dim ItemData() As Byte
    Set Buffer = New clsBuffer
    ItemSize = LenB(Item(itemnum))
    
    ReDim ItemData(ItemSize - 1)
    
    CopyMemory ItemData(0), ByVal VarPtr(Item(itemnum)), ItemSize
    
    Buffer.WriteLong SUpdateItem
    Buffer.WriteLong itemnum
    Buffer.WriteBytes ItemData
    
    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendUpdateItemTo(ByVal index As Long, ByVal itemnum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim ItemSize As Long
    Dim ItemData() As Byte
    Set Buffer = New clsBuffer
    ItemSize = LenB(Item(itemnum))
    ReDim ItemData(ItemSize - 1)
    CopyMemory ItemData(0), ByVal VarPtr(Item(itemnum)), ItemSize
    Buffer.WriteLong SUpdateItem
    Buffer.WriteLong itemnum
    Buffer.WriteBytes ItemData
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendUpdateAnimationToAll(ByVal AnimationNum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim AnimationSize As Long
    Dim AnimationData() As Byte
    Set Buffer = New clsBuffer
    AnimationSize = LenB(Animation(AnimationNum))
    ReDim AnimationData(AnimationSize - 1)
    CopyMemory AnimationData(0), ByVal VarPtr(Animation(AnimationNum)), AnimationSize
    Buffer.WriteLong SUpdateAnimation
    Buffer.WriteLong AnimationNum
    Buffer.WriteBytes AnimationData
    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendUpdateAnimationTo(ByVal index As Long, ByVal AnimationNum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim AnimationSize As Long
    Dim AnimationData() As Byte
    Set Buffer = New clsBuffer
    AnimationSize = LenB(Animation(AnimationNum))
    ReDim AnimationData(AnimationSize - 1)
    CopyMemory AnimationData(0), ByVal VarPtr(Animation(AnimationNum)), AnimationSize
    Buffer.WriteLong SUpdateAnimation
    Buffer.WriteLong AnimationNum
    Buffer.WriteBytes AnimationData
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendUpdateNpcToAll(ByVal npcNum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim NPCSize As Long
    Dim NPCData() As Byte
    Set Buffer = New clsBuffer
    NPCSize = LenB(Npc(npcNum))
    ReDim NPCData(NPCSize - 1)
    CopyMemory NPCData(0), ByVal VarPtr(Npc(npcNum)), NPCSize
    Buffer.WriteLong SUpdateNpc
    Buffer.WriteLong npcNum
    Buffer.WriteBytes NPCData
    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendUpdateNpcTo(ByVal index As Long, ByVal npcNum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim NPCSize As Long
    Dim NPCData() As Byte
    Set Buffer = New clsBuffer
    NPCSize = LenB(Npc(npcNum))
    ReDim NPCData(NPCSize - 1)
    CopyMemory NPCData(0), ByVal VarPtr(Npc(npcNum)), NPCSize
    Buffer.WriteLong SUpdateNpc
    Buffer.WriteLong npcNum
    Buffer.WriteBytes NPCData
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendUpdateResourceToAll(ByVal ResourceNum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim ResourceSize As Long
    Dim ResourceData() As Byte
    
    Set Buffer = New clsBuffer
    
    ResourceSize = LenB(Resource(ResourceNum))
    ReDim ResourceData(ResourceSize - 1)
    CopyMemory ResourceData(0), ByVal VarPtr(Resource(ResourceNum)), ResourceSize
    
    Buffer.WriteLong SUpdateResource
    Buffer.WriteLong ResourceNum
    Buffer.WriteBytes ResourceData

    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendUpdateResourceTo(ByVal index As Long, ByVal ResourceNum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim ResourceSize As Long
    Dim ResourceData() As Byte
    
    Set Buffer = New clsBuffer
    
    ResourceSize = LenB(Resource(ResourceNum))
    ReDim ResourceData(ResourceSize - 1)
    CopyMemory ResourceData(0), ByVal VarPtr(Resource(ResourceNum)), ResourceSize
    
    Buffer.WriteLong SUpdateResource
    Buffer.WriteLong ResourceNum
    Buffer.WriteBytes ResourceData
    
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendShops(ByVal index As Long)
    Dim I As Long

    For I = 1 To MAX_SHOPS

        If LenB(Trim$(Shop(I).Name)) > 0 Then
            Call SendUpdateShopTo(index, I)
        End If

    Next

End Sub

Sub SendUpdateShopToAll(ByVal shopNum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim ShopSize As Long
    Dim ShopData() As Byte
    
    Set Buffer = New clsBuffer
    
    ShopSize = LenB(Shop(shopNum))
    ReDim ShopData(ShopSize - 1)
    CopyMemory ShopData(0), ByVal VarPtr(Shop(shopNum)), ShopSize
    
    Buffer.WriteLong SUpdateShop
    Buffer.WriteLong shopNum
    Buffer.WriteBytes ShopData

    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendUpdateShopTo(ByVal index As Long, ByVal shopNum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim ShopSize As Long
    Dim ShopData() As Byte
    
    Set Buffer = New clsBuffer
    
    ShopSize = LenB(Shop(shopNum))
    ReDim ShopData(ShopSize - 1)
    CopyMemory ShopData(0), ByVal VarPtr(Shop(shopNum)), ShopSize
    
    Buffer.WriteLong SUpdateShop
    Buffer.WriteLong shopNum
    Buffer.WriteBytes ShopData
    
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendSpells(ByVal index As Long)
    Dim I As Long

    For I = 1 To MAX_SPELLS

        If LenB(Trim$(Spell(I).Name)) > 0 Then
            Call SendUpdateSpellTo(index, I)
        End If

    Next

End Sub

Sub SendUpdateSpellToAll(ByVal spellnum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Dim SpellSize As Long
    Dim SpellData() As Byte
    
    Set Buffer = New clsBuffer
    
    SpellSize = LenB(Spell(spellnum))
    ReDim SpellData(SpellSize - 1)
    CopyMemory SpellData(0), ByVal VarPtr(Spell(spellnum)), SpellSize
    
    Buffer.WriteLong SUpdateSpell
    Buffer.WriteLong spellnum
    Buffer.WriteBytes SpellData
    
    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendUpdateSpellTo(ByVal index As Long, ByVal spellnum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim SpellSize As Long
    Dim SpellData() As Byte
    
    Set Buffer = New clsBuffer
    
    SpellSize = LenB(Spell(spellnum))
    ReDim SpellData(SpellSize - 1)
    CopyMemory SpellData(0), ByVal VarPtr(Spell(spellnum)), SpellSize
    
    Buffer.WriteLong SUpdateSpell
    Buffer.WriteLong spellnum
    Buffer.WriteBytes SpellData
    
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPlayerSpells(ByVal index As Long)
    Dim packet As String
    Dim I As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SSpells

    For I = 1 To MAX_PLAYER_SPELLS
        Buffer.WriteLong GetPlayerSpell(index, I)
    Next

    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendResourceCacheTo(ByVal index As Long, ByVal Resource_num As Long)
    Dim Buffer As clsBuffer
    Dim I As Long
    Set Buffer = New clsBuffer
    Buffer.WriteLong SResourceCache
    Buffer.WriteLong ResourceCache(GetPlayerMap(index)).Resource_Count

    If ResourceCache(GetPlayerMap(index)).Resource_Count > 0 Then

        For I = 0 To ResourceCache(GetPlayerMap(index)).Resource_Count
            Buffer.WriteByte ResourceCache(GetPlayerMap(index)).ResourceData(I).ResourceState
            Buffer.WriteLong ResourceCache(GetPlayerMap(index)).ResourceData(I).x
            Buffer.WriteLong ResourceCache(GetPlayerMap(index)).ResourceData(I).y
        Next

    End If

    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendResourceCacheToMap(ByVal mapNum As Long, ByVal Resource_num As Long)
    Dim Buffer As clsBuffer
    Dim I As Long
    Set Buffer = New clsBuffer
    Buffer.WriteLong SResourceCache
    Buffer.WriteLong ResourceCache(mapNum).Resource_Count

    If ResourceCache(mapNum).Resource_Count > 0 Then

        For I = 0 To ResourceCache(mapNum).Resource_Count
            Buffer.WriteByte ResourceCache(mapNum).ResourceData(I).ResourceState
            Buffer.WriteLong ResourceCache(mapNum).ResourceData(I).x
            Buffer.WriteLong ResourceCache(mapNum).ResourceData(I).y
        Next

    End If

    SendDataToMap mapNum, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendDoorAnimation(ByVal mapNum As Long, ByVal x As Long, ByVal y As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SDoorAnimation
    Buffer.WriteLong x
    Buffer.WriteLong y
    
    SendDataToMap mapNum, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendActionMsg(ByVal mapNum As Long, ByVal message As String, ByVal color As Long, ByVal MsgType As Long, ByVal x As Long, ByVal y As Long, Optional PlayerOnlyNum As Long = 0)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SActionMsg
    Buffer.WriteString message
    Buffer.WriteLong color
    Buffer.WriteLong MsgType
    Buffer.WriteLong x
    Buffer.WriteLong y
    
    If PlayerOnlyNum > 0 Then
        SendDataTo PlayerOnlyNum, Buffer.ToArray()
    Else
        SendDataToMap mapNum, Buffer.ToArray()
    End If
    
    Set Buffer = Nothing
End Sub

Sub SendBlood(ByVal mapNum As Long, ByVal x As Long, ByVal y As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SBlood
    Buffer.WriteLong x
    Buffer.WriteLong y
    
    SendDataToMap mapNum, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendAnimation(ByVal mapNum As Long, ByVal Anim As Long, ByVal x As Long, ByVal y As Long, Optional ByVal LockType As Byte = 0, Optional ByVal LockIndex As Long = 0)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SAnimation
    Buffer.WriteLong Anim
    Buffer.WriteLong x
    Buffer.WriteLong y
    Buffer.WriteByte LockType
    Buffer.WriteLong LockIndex
    
    SendDataToMap mapNum, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendCooldown(ByVal index As Long, ByVal Slot As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SCooldown
    Buffer.WriteLong Slot
    
    SendDataTo index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendClearSpellBuffer(ByVal index As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SClearSpellBuffer
    
    SendDataTo index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SayMsg_Map(ByVal mapNum As Long, ByVal index As Long, ByVal message As String, ByVal saycolour As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SSayMsg
    Buffer.WriteString GetPlayerName(index)
    Buffer.WriteLong GetPlayerAccess(index)
    Buffer.WriteLong GetPlayerPK(index)
    Buffer.WriteString message
    Buffer.WriteString "[Map] "
    Buffer.WriteLong saycolour
    
    SendDataToMap mapNum, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SayMsg_Global(ByVal index As Long, ByVal message As String, ByVal saycolour As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SSayMsg
    Buffer.WriteString GetPlayerName(index)
    Buffer.WriteLong GetPlayerAccess(index)
    Buffer.WriteLong GetPlayerPK(index)
    Buffer.WriteString message
    Buffer.WriteString "[Global] "
    Buffer.WriteLong saycolour
    
    SendDataToAll Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub ResetShopAction(ByVal index As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SResetShopAction
    
    SendDataToAll Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendStunned(ByVal index As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SStunned
    Buffer.WriteLong TempPlayer(index).StunDuration
    
    SendDataTo index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendBank(ByVal index As Long)
    Dim Buffer As clsBuffer
    Dim I As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SBank
    
    For I = 1 To MAX_BANK
        Buffer.WriteLong Bank(index).Item(I).Num
        Buffer.WriteLong Bank(index).Item(I).Value
    Next
    
    SendDataTo index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendMapKey(ByVal index As Long, ByVal x As Long, ByVal y As Long, ByVal Value As Byte)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SMapKey
    Buffer.WriteLong x
    Buffer.WriteLong y
    Buffer.WriteByte Value
    SendDataTo index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendMapKeyToMap(ByVal mapNum As Long, ByVal x As Long, ByVal y As Long, ByVal Value As Byte)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SMapKey
    Buffer.WriteLong x
    Buffer.WriteLong y
    Buffer.WriteByte Value
    SendDataToMap mapNum, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendOpenShop(ByVal index As Long, ByVal shopNum As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SOpenShop
    Buffer.WriteLong shopNum
    SendDataTo index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendPlayerMove(ByVal index As Long, ByVal movement As Long, Optional ByVal sendToSelf As Boolean = False)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerMove
    Buffer.WriteLong index
    Buffer.WriteLong GetPlayerX(index)
    Buffer.WriteLong GetPlayerY(index)
    Buffer.WriteLong GetPlayerDir(index)
    Buffer.WriteLong movement
    
    If Not sendToSelf Then
        SendDataToMapBut index, GetPlayerMap(index), Buffer.ToArray()
    Else
        SendDataToMap GetPlayerMap(index), Buffer.ToArray()
    End If
    
    Set Buffer = Nothing
End Sub

Sub SendTrade(ByVal index As Long, ByVal tradeTarget As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong STrade
    Buffer.WriteLong tradeTarget
    Buffer.WriteString Trim$(GetPlayerName(tradeTarget))
    SendDataTo index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendCloseTrade(ByVal index As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SCloseTrade
    SendDataTo index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendTradeUpdate(ByVal index As Long, ByVal dataType As Byte)
Dim Buffer As clsBuffer
Dim I As Long
Dim tradeTarget As Long
Dim totalWorth As Long
    
    tradeTarget = TempPlayer(index).InTrade
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong STradeUpdate
    Buffer.WriteByte dataType
    
    If dataType = 0 Then ' own inventory
        For I = 1 To MAX_INV
            Buffer.WriteLong TempPlayer(index).TradeOffer(I).Num
            Buffer.WriteLong TempPlayer(index).TradeOffer(I).Value
            ' add total worth
            If TempPlayer(index).TradeOffer(I).Num > 0 Then
                ' currency?
                If Item(TempPlayer(index).TradeOffer(I).Num).Type = ITEM_TYPE_CURRENCY Then
                    totalWorth = totalWorth + (Item(GetPlayerInvItemNum(index, TempPlayer(index).TradeOffer(I).Num)).price * TempPlayer(index).TradeOffer(I).Value)
                Else
                    totalWorth = totalWorth + Item(GetPlayerInvItemNum(index, TempPlayer(index).TradeOffer(I).Num)).price
                End If
            End If
        Next
    ElseIf dataType = 1 Then ' other inventory
        For I = 1 To MAX_INV
            Buffer.WriteLong GetPlayerInvItemNum(tradeTarget, TempPlayer(tradeTarget).TradeOffer(I).Num)
            Buffer.WriteLong TempPlayer(tradeTarget).TradeOffer(I).Value
            ' add total worth
            If GetPlayerInvItemNum(tradeTarget, TempPlayer(tradeTarget).TradeOffer(I).Num) > 0 Then
                ' currency?
                If Item(GetPlayerInvItemNum(tradeTarget, TempPlayer(tradeTarget).TradeOffer(I).Num)).Type = ITEM_TYPE_CURRENCY Then
                    totalWorth = totalWorth + (Item(GetPlayerInvItemNum(tradeTarget, TempPlayer(tradeTarget).TradeOffer(I).Num)).price * TempPlayer(tradeTarget).TradeOffer(I).Value)
                Else
                    totalWorth = totalWorth + Item(GetPlayerInvItemNum(tradeTarget, TempPlayer(tradeTarget).TradeOffer(I).Num)).price
                End If
            End If
        Next
    End If
    
    ' send total worth of trade
    Buffer.WriteLong totalWorth
    
    SendDataTo index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendTradeStatus(ByVal index As Long, ByVal Status As Byte)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong STradeStatus
    Buffer.WriteByte Status
    SendDataTo index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendTarget(ByVal index As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong STarget
    Buffer.WriteLong TempPlayer(index).target
    Buffer.WriteLong TempPlayer(index).targetType
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendHotbar(ByVal index As Long)
Dim I As Long
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SHotbar
    For I = 1 To MAX_HOTBAR
        Buffer.WriteLong Player(index).Hotbar(I).Slot
        Buffer.WriteByte Player(index).Hotbar(I).sType
    Next
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendLoginOk(ByVal index As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SLoginOk
    Buffer.WriteLong index
    Buffer.WriteLong Player_HighIndex
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendInGame(ByVal index As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SInGame
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendHighIndex()
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SHighIndex
    Buffer.WriteLong Player_HighIndex
    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPlayerSound(ByVal index As Long, ByVal x As Long, ByVal y As Long, ByVal entityType As Long, ByVal entityNum As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SSound
    Buffer.WriteLong x
    Buffer.WriteLong y
    Buffer.WriteLong entityType
    Buffer.WriteLong entityNum
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendMapSound(ByVal index As Long, ByVal x As Long, ByVal y As Long, ByVal entityType As Long, ByVal entityNum As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SSound
    Buffer.WriteLong x
    Buffer.WriteLong y
    Buffer.WriteLong entityType
    Buffer.WriteLong entityNum
    SendDataToMap GetPlayerMap(index), Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendTradeRequest(ByVal index As Long, ByVal TradeRequest As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong STradeRequest
    Buffer.WriteString Trim$(Player(TradeRequest).Name)
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPartyInvite(ByVal index As Long, ByVal targetPlayer As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SPartyInvite
    Buffer.WriteString Trim$(Player(targetPlayer).Name)
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPartyUpdate(ByVal partyNum As Long)
Dim Buffer As clsBuffer, I As Long

    Set Buffer = New clsBuffer
    Buffer.WriteLong SPartyUpdate
    Buffer.WriteByte 1
    Buffer.WriteLong Party(partyNum).Leader
    For I = 1 To MAX_PARTY_MEMBERS
        Buffer.WriteLong Party(partyNum).Member(I)
    Next
    Buffer.WriteLong Party(partyNum).MemberCount
    SendDataToParty partyNum, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPartyUpdateTo(ByVal index As Long)
Dim Buffer As clsBuffer, I As Long, partyNum As Long

    Set Buffer = New clsBuffer
    Buffer.WriteLong SPartyUpdate
    
    ' check if we're in a party
    partyNum = TempPlayer(index).inParty
    If partyNum > 0 Then
        ' send party data
        Buffer.WriteByte 1
        Buffer.WriteLong Party(partyNum).Leader
        For I = 1 To MAX_PARTY_MEMBERS
            Buffer.WriteLong Party(partyNum).Member(I)
        Next
        Buffer.WriteLong Party(partyNum).MemberCount
    Else
        ' send clear command
        Buffer.WriteByte 0
    End If
    
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPartyVitals(ByVal partyNum As Long, ByVal index As Long)
Dim Buffer As clsBuffer, I As Long

    Set Buffer = New clsBuffer
    Buffer.WriteLong SPartyVitals
    Buffer.WriteLong index
    For I = 1 To Vitals.Vital_Count - 1
        Buffer.WriteLong GetPlayerMaxVital(index, I)
        Buffer.WriteLong Player(index).Vital(I)
    Next
    SendDataToParty partyNum, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendSpawnItemToMap(ByVal mapNum As Long, ByVal index As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SSpawnItem
    Buffer.WriteLong index
    Buffer.WriteString MapItem(mapNum, index).playerName
    Buffer.WriteLong MapItem(mapNum, index).Num
    Buffer.WriteLong MapItem(mapNum, index).Value
    Buffer.WriteLong MapItem(mapNum, index).x
    Buffer.WriteLong MapItem(mapNum, index).y
    SendDataToMap mapNum, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

