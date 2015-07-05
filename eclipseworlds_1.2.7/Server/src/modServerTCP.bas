Attribute VB_Name = "modServerTCP"
Option Explicit

Sub UpdateCaption()
    ' Update the form caption
    frmServer.Caption = Options.Name & " <IP " & frmServer.Socket(0).LocalIP & " Port " & CStr(frmServer.Socket(0).LocalPort) & "> (" & TotalOnlinePlayers & ")"

    ' Update form labels
    frmServer.lblIP = frmServer.Socket(0).LocalIP
    frmServer.lblPort = CStr(frmServer.Socket(0).LocalPort)
    frmServer.lblPlayers = TotalOnlinePlayers & "/" & Trim$(STR$(MAX_PLAYERS))
End Sub

Sub CreateFullMapCache()
    Dim i As Long

    For i = 1 To MAX_MAPS
        Call MapCache_Create(i)
    Next
End Sub

Function IsConnected(ByVal index As Long) As Boolean
    If index < 1 Or index > MAX_PLAYERS Then Exit Function
    
    If frmServer.Socket(index).State = sckConnected Then
        IsConnected = True
    End If
End Function

Function IsPlaying(ByVal index As Long) As Boolean
    If index < 1 Or index > MAX_PLAYERS Then Exit Function
    
    If IsConnected(index) And GetPlayerLogin(index) <> vbNullString Then
        If tempplayer(index).InGame Or tempplayer(index).HasLogged Then
            IsPlaying = True
        End If
    End If
End Function

Function IsLoggedIn(ByVal index As Integer) As Boolean
    If index < 1 Or index > MAX_PLAYERS Then Exit Function

    If IsConnected(index) Then
        If GetPlayerLogin(index) <> vbNullString Then
            IsLoggedIn = True
        End If
    End If
End Function

Function IsMultiAccounts(ByVal Login As String) As Boolean
    Dim i As Long

    For i = 1 To Player_HighIndex
        If IsConnected(i) Then
            If LCase$(Trim$(Account(i).Login)) = LCase$(Login) Then
                IsMultiAccounts = True
                Exit Function
            End If
        End If
    Next
End Function

Function IsMultiIPOnline(ByVal IP As String) As Boolean
    Dim i As Long
    Dim n As Long

    For i = 1 To Player_HighIndex
        If IsConnected(i) Then
            If Trim$(GetPlayerIP(i)) = IP Then
                n = n + 1

                If (n > 1) Then
                    IsMultiIPOnline = True
                    Exit Function
                End If
            End If
        End If
    Next
End Function

Public Function IsBanned(ByVal index As Long, Serial As String) As Boolean
    Dim SendIP As String
    Dim n As Long

    IsBanned = False

    ' Cut off last portion of IP
    SendIP = GetPlayerIP(index)
    For n = Len(SendIP) To 1 Step -1
        If Mid$(SendIP, n, 1) = "." Then Exit For
    Next n
    
    SendIP = Mid$(SendIP, 1, n)

    For n = 1 To MAX_BANS
        If Len(GetPlayerLogin(index)) > 0 Then
            If GetPlayerLogin(index) = Trim$(Ban(n).PlayerLogin) Then
                IsBanned = True
                Exit For
            End If
        End If

        If Len(Trim$(SendIP)) > 0 Then
            If Trim$(SendIP) = Left$(Trim$(Ban(n).IP), Len(SendIP)) Then
                IsBanned = True
                Exit For
            End If
        End If

        If Len(Serial) > 0 Then
            If Serial = Trim$(Ban(n).HDSerial) Then
                IsBanned = True
                Exit For
            End If
        Else
            IsBanned = True
            Exit For
        End If
    Next n
    
    If IsBanned = True Then
        Call AlertMsg(index, "You are banned from " & Options.Name & " and can no longer play.")
    End If
End Function

Sub SendDataTo(ByVal index As Long, ByRef Data() As Byte)
    Dim buffer As clsBuffer
    Dim TempData() As Byte
    
    If IsConnected(index) Then
        Set buffer = New clsBuffer
        TempData = Data
        
        buffer.PreAllocate 4 + (UBound(TempData) - LBound(TempData)) + 1
        buffer.WriteLong (UBound(TempData) - LBound(TempData)) + 1
        buffer.WriteBytes TempData()
        
        If IsConnected(index) Then
            On Error Resume Next
            frmServer.Socket(index).SendData buffer.ToArray()
            DoEvents
        End If
    End If
End Sub

Sub SendDataToAll(ByRef Data() As Byte)
    Dim i As Long

    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            Call SendDataTo(i, Data)
        End If
    Next
End Sub

Sub SendDataToAllBut(ByVal index As Long, ByRef Data() As Byte)
    Dim i As Long

    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            If Not i = index Then
                Call SendDataTo(i, Data)
            End If
        End If
    Next
End Sub

Sub SendDataToMap(ByVal MapNum As Integer, ByRef Data() As Byte)
    Dim i As Long

    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            If GetPlayerMap(i) = MapNum Then
                Call SendDataTo(i, Data)
            End If
        End If
    Next
End Sub

Sub SendDataToMapBut(ByVal index As Long, ByVal MapNum As Integer, ByRef Data() As Byte)
    Dim i As Long

    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            If GetPlayerMap(i) = MapNum Then
                If Not i = index Then
                    Call SendDataTo(i, Data)
                End If
            End If
        End If
    Next
End Sub

Sub SendDataToParty(ByVal PartyNum As Long, ByRef Data() As Byte)
    Dim i As Long

    For i = 1 To Party(PartyNum).MemberCount
        If Party(PartyNum).Member(i) > 0 Then
            Call SendDataTo(Party(PartyNum).Member(i), Data)
        End If
    Next
End Sub

Public Sub GlobalMsg(ByVal Msg As String, ByVal Color As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.WriteLong SGlobalMsg
    buffer.WriteString Msg
    buffer.WriteLong Color
    SendDataToAll buffer.ToArray
    
    Set buffer = Nothing
End Sub

Public Sub AdminMsg(ByVal Msg As String, Color As Long)
    Dim buffer As clsBuffer
    Dim i As Long
    Dim LogMsg As String
    
    Set buffer = New clsBuffer
    
    ' Add server log
    Call AddLog(Msg, "Player")
    
    LogMsg = Msg
    Msg = "[Admin] " & Msg
    
    ' Prevent hacking
    For i = 1 To Len(Msg)

        ' limit the ASCII
        If AscW(Mid$(Msg, i, 1)) < 32 Or AscW(Mid$(Msg, i, 1)) > 126 Then

            ' limit the extended ASCII
            If AscW(Mid$(Msg, i, 1)) < 128 Or AscW(Mid$(Msg, i, 1)) > 168 Then

                ' limit the extended ASCII
                If AscW(Mid$(Msg, i, 1)) < 224 Or AscW(Mid$(Msg, i, 1)) > 253 Then
                    Mid$(Msg, i, 1) = ""
                End If
            End If
        End If

    Next
    
    buffer.WriteLong SAdminMsg
    buffer.WriteString Msg
    buffer.WriteByte Color
    
    For i = 1 To Player_HighIndex
        If IsPlaying(i) And GetPlayerAccess(i) >= STAFF_MODERATOR Then
            SendDataTo i, buffer.ToArray
            Call SendLogs(i, LogMsg, "Admin")
        End If
    Next
    
    Set buffer = Nothing
End Sub

Public Sub PlayerMsg(ByVal index As Long, ByVal Msg As String, ByVal Color As Long, Optional ByVal QuestMsg As Boolean = False, Optional ByVal QuestNum As Long = 0, Optional ByVal Name As String = "")
    Dim buffer As clsBuffer
    Dim i As Long
    
    Set buffer = New clsBuffer
    
    ' Prevent hacking
    For i = 1 To Len(Msg)

        ' limit the ASCII
        If AscW(Mid$(Msg, i, 1)) < 32 Or AscW(Mid$(Msg, i, 1)) > 126 Then

            ' limit the extended ASCII
            If AscW(Mid$(Msg, i, 1)) < 128 Or AscW(Mid$(Msg, i, 1)) > 168 Then

                ' limit the extended ASCII
                If AscW(Mid$(Msg, i, 1)) < 224 Or AscW(Mid$(Msg, i, 1)) > 253 Then
                    Mid$(Msg, i, 1) = ""
                End If
            End If
        End If

    Next
    
    buffer.WriteLong SPlayerMsg
    buffer.WriteString Msg
    buffer.WriteByte Color
    buffer.WriteLong QuestMsg
    buffer.WriteLong QuestNum
    SendDataTo index, buffer.ToArray
    
    Set buffer = Nothing
End Sub

Public Sub MapMsg(ByVal MapNum As Integer, ByVal Msg As String, ByVal Color As Long)
    Dim buffer As clsBuffer
    Dim i As Long

    Set buffer = New clsBuffer

    ' Prevent hacking
    For i = 1 To Len(Msg)

        ' limit the ASCII
        If AscW(Mid$(Msg, i, 1)) < 32 Or AscW(Mid$(Msg, i, 1)) > 126 Then

            ' limit the extended ASCII
            If AscW(Mid$(Msg, i, 1)) < 128 Or AscW(Mid$(Msg, i, 1)) > 168 Then

                ' limit the extended ASCII
                If AscW(Mid$(Msg, i, 1)) < 224 Or AscW(Mid$(Msg, i, 1)) > 253 Then
                    Mid$(Msg, i, 1) = ""
                End If
            End If
        End If

    Next
    
    buffer.WriteLong SMapMsg
    buffer.WriteString Msg
    buffer.WriteByte Color
    SendDataToMap MapNum, buffer.ToArray
    
    Set buffer = Nothing
End Sub

Public Sub GuildMsg(ByVal index As Long, ByVal Msg As String, ByVal Color As Long, Optional HideName As Boolean = False)
    Dim i As Long
    Dim LogMsg As String
    
    ' Prevent hacking
    For i = 1 To Len(Msg)

        ' limit the ASCII
        If AscW(Mid$(Msg, i, 1)) < 32 Or AscW(Mid$(Msg, i, 1)) > 126 Then

            ' limit the extended ASCII
            If AscW(Mid$(Msg, i, 1)) < 128 Or AscW(Mid$(Msg, i, 1)) > 168 Then

                ' limit the extended ASCII
                If AscW(Mid$(Msg, i, 1)) < 224 Or AscW(Mid$(Msg, i, 1)) > 253 Then
                    Mid$(Msg, i, 1) = ""
                End If
            End If
        End If

    Next
    
    ' Add server log
    Call AddLog(Msg, "Player")
    
    ' Set the LogMsg
    If HideName = True Then
        LogMsg = Msg
        Msg = "[Guild] " & Msg
    Else
        LogMsg = GetPlayerName(index) & ": " & Msg
        Msg = "[Guild] " & GetPlayerName(index) & ": " & Msg
    End If

    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            If GetPlayerGuild(i) = GetPlayerGuild(index) Then
                PlayerMsg i, Msg, Color
                Call SendLogs(i, LogMsg, "Guild")
            End If
        End If
    Next
End Sub

Public Sub AlertMsg(ByVal index As Long, ByVal Msg As String)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer

    buffer.WriteLong SAlertMsg
    buffer.WriteString Msg
    SendDataTo index, buffer.ToArray
    DoEvents
    
    Set buffer = Nothing
End Sub

Public Sub PartyMsg(ByVal PartyNum As Long, ByVal Msg As String, ByVal Color As Long)
    Dim i As Long
    Dim LogMsg As String
    
    ' Prevent hacking
    For i = 1 To Len(Msg)

        ' limit the ASCII
        If AscW(Mid$(Msg, i, 1)) < 32 Or AscW(Mid$(Msg, i, 1)) > 126 Then

            ' limit the extended ASCII
            If AscW(Mid$(Msg, i, 1)) < 128 Or AscW(Mid$(Msg, i, 1)) > 168 Then

                ' limit the extended ASCII
                If AscW(Mid$(Msg, i, 1)) < 224 Or AscW(Mid$(Msg, i, 1)) > 253 Then
                    Mid$(Msg, i, 1) = ""
                End If
            End If
        End If

    Next
    
    ' Add server log
    Call AddLog(Msg, "Player")
    
    LogMsg = Msg
    
    Msg = "[Party] " & Msg
    
    ' Send message to all people
    For i = 1 To MAX_PARTY_MEMBERS
        ' Exist?
        If Party(PartyNum).Member(i) > 0 Then
            ' Make sure they're logged on
            If IsPlaying(Party(PartyNum).Member(i)) Then
                PlayerMsg Party(PartyNum).Member(i), Msg, Color
                Call SendLogs(Party(PartyNum).Member(i), LogMsg, "Party")
            End If
        End If
    Next
End Sub

Sub AcceptConnection(ByVal index As Long, ByVal SocketId As Long)
    Dim i As Long

    If (index = 0) Then
        i = FindOpenPlayerSlot

        If Not i = 0 Then
            If Not IsConnected(i) Then
                ' We can connect them
                frmServer.Socket(i).Close
                frmServer.Socket(i).Accept SocketId
                Call SocketConnected(i)
            End If
        End If
    End If
End Sub

Sub SocketConnected(ByVal index As Long)
    Dim i As Long
    
    If index > 0 And index <= MAX_PLAYERS Then
        ' Are they trying to connect more then one connection?
        If Not IsMultiIPOnline(GetPlayerIP(index)) Then
            Call TextAdd("Received connection from " & GetPlayerIP(index) & ".")
        ElseIf Options.MultipleIP = 0 And IsMultiIPOnline(GetPlayerIP(index)) Then
            ' Tried Multiple connections
            Call AlertMsg(index, "Multiple account logins are not authorized.")
            frmServer.Socket(index).Close
        End If
    Else
        Call AlertMsg(index, "The server is full! Try back again later.")
        frmServer.Socket(index).Close
    End If

    ' Re-set the high Index
    Player_HighIndex = 0
    
    For i = MAX_PLAYERS To 1 Step -1
        If IsConnected(i) Then
            Player_HighIndex = i
            Exit For
        End If
    Next
    
    ' Send the new highIndex to all logged in players
    SendPlayer_HighIndex
End Sub

Sub IncomingData(ByVal index As Long, ByVal DataLength As Long)
    Dim buffer() As Byte
    Dim pLength As Long

    If GetPlayerAccess(index) <= 0 Then
        ' Check for data flooding
        If tempplayer(index).DataBytes > 1000 Then Exit Sub
    
        ' Check for Packet flooding
        If tempplayer(index).DataPackets > 105 Then Exit Sub
    End If
            
    ' Check if elapsed time has passed
    tempplayer(index).DataBytes = tempplayer(index).DataBytes + DataLength
    
    If timeGetTime >= tempplayer(index).DataTimer Then
        tempplayer(index).DataTimer = timeGetTime + 1000
        tempplayer(index).DataBytes = 0
        tempplayer(index).DataPackets = 0
    End If
    
    ' Get the data from the socket now
    frmServer.Socket(index).GetData buffer(), vbUnicode, DataLength
    tempplayer(index).buffer.WriteBytes buffer()
    
    If tempplayer(index).buffer.Length >= 4 Then
        pLength = tempplayer(index).buffer.ReadLong(False)
    
        If pLength < 0 Then Exit Sub
    End If
    
    Do While pLength > 0 And pLength <= tempplayer(index).buffer.Length - 4
        If pLength <= tempplayer(index).buffer.Length - 4 Then
            tempplayer(index).DataPackets = tempplayer(index).DataPackets + 1
            tempplayer(index).buffer.ReadLong
            HandleData index, tempplayer(index).buffer.ReadBytes(pLength)
        End If
        
        pLength = 0
        
        If tempplayer(index).buffer.Length >= 4 Then
            pLength = tempplayer(index).buffer.ReadLong(False)
        
            If pLength < 0 Then Exit Sub
        End If
    Loop
            
    tempplayer(index).buffer.Trim
End Sub

Sub CloseSocket(ByVal index As Long)
    Dim i As Long
    
    If index > 0 And index <= Player_HighIndex Then
        If Not tempplayer(index).HasLogged Then
            If Moral(Map(GetPlayerMap(index)).Moral).CanPK = 1 Or Moral(Map(GetPlayerMap(index)).Moral).DropItems = 1 Or Moral(Map(GetPlayerMap(index)).Moral).LoseExp = 1 Or GetPlayerPK(index) = YES Then
                tempplayer(index).PVPTimer = timeGetTime + 30000
            End If
            
            tempplayer(index).HasLogged = True
        End If

        If frmServer.Socket(index).State = sckConnected Then
            frmServer.Socket(index).Close
            Call TextAdd("Connection from " & GetPlayerIP(index) & " has been terminated.")
        End If
        
    End If

End Sub

Public Sub MapCache_Create(ByVal MapNum As Integer)
    Dim MapData As String
    Dim x As Long
    Dim Y As Long
    Dim i As Long, z As Long, w As Long
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    
    buffer.WriteLong MapNum
    buffer.WriteString Trim$(Map(MapNum).Name)
    buffer.WriteString Trim$(Map(MapNum).Music)
    buffer.WriteString Trim$(Map(MapNum).BGS)
    buffer.WriteLong Map(MapNum).Revision
    buffer.WriteByte Map(MapNum).Moral
    buffer.WriteLong Map(MapNum).Up
    buffer.WriteLong Map(MapNum).Down
    buffer.WriteLong Map(MapNum).Left
    buffer.WriteLong Map(MapNum).Right
    buffer.WriteLong Map(MapNum).BootMap
    buffer.WriteByte Map(MapNum).BootX
    buffer.WriteByte Map(MapNum).BootY
    
    buffer.WriteLong Map(MapNum).Weather
    buffer.WriteLong Map(MapNum).WeatherIntensity
    
    buffer.WriteLong Map(MapNum).Fog
    buffer.WriteLong Map(MapNum).FogSpeed
    buffer.WriteLong Map(MapNum).FogOpacity
    
    buffer.WriteLong Map(MapNum).Panorama
    
    buffer.WriteLong Map(MapNum).Red
    buffer.WriteLong Map(MapNum).Green
    buffer.WriteLong Map(MapNum).Blue
    buffer.WriteLong Map(MapNum).Alpha
    
    buffer.WriteByte Map(MapNum).MaxX
    buffer.WriteByte Map(MapNum).MaxY
    
    buffer.WriteByte Map(MapNum).NPC_HighIndex

    For x = 0 To Map(MapNum).MaxX
        For Y = 0 To Map(MapNum).MaxY
            With Map(MapNum).Tile(x, Y)
                For i = 1 To MapLayer.Layer_Count - 1
                    buffer.WriteLong .Layer(i).x
                    buffer.WriteLong .Layer(i).Y
                    buffer.WriteLong .Layer(i).Tileset
                Next
                
                For z = 1 To MapLayer.Layer_Count - 1
                    buffer.WriteLong .Autotile(z)
                Next
                
                buffer.WriteByte .Type
                buffer.WriteLong .Data1
                buffer.WriteLong .Data2
                buffer.WriteLong .Data3
                buffer.WriteString .Data4
                buffer.WriteByte .DirBlock
            End With
        Next
    Next

    For x = 1 To MAX_MAP_NPCS
        buffer.WriteLong Map(MapNum).NPC(x)
        buffer.WriteLong Map(MapNum).NPCSpawnType(x)
    Next

    MapCache(MapNum).Data = buffer.ToArray()
    Set buffer = Nothing
End Sub

' *****************************
' ** Outgoing Server Packets **
' *****************************
Sub SendWhosOnline(ByVal index As Long)
    Dim s As String
    Dim n As Long
    Dim i As Long

    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            If Not i = index Then
                s = s & GetPlayerName(i) & ", "
                n = n + 1
            End If
        End If
    Next

    If n = 0 Then
        s = "There are no other players online."
    ElseIf n = 1 Then
        s = Mid$(s, 1, Len(s) - 2)
        s = "There is " & n & " other player online: " & s & "."
    Else
        s = Mid$(s, 1, Len(s) - 2)
        s = "There are " & n & " other players online: " & s & "."
    End If

    Call PlayerMsg(index, s, WhoColor)
End Sub
'Character Editor
Sub SendPlayersOnline(ByVal index As Long)
    Dim buffer As clsBuffer, i As Long
    Dim list As String

    If index > Player_HighIndex Or index < 1 Then Exit Sub
    Set buffer = New clsBuffer
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
                If i <> Player_HighIndex Then
                    list = list & GetPlayerName(i) & ":" & Account(i).Chars(GetPlayerChar(i)).Access & ":" & Account(i).Chars(GetPlayerChar(i)).Sprite & ", "
                Else
                    list = list & GetPlayerName(i) & ":" & Account(i).Chars(GetPlayerChar(i)).Access & ":" & Account(i).Chars(GetPlayerChar(i)).Sprite
                End If
        End If
    Next
    
    buffer.WriteLong SPlayersOnline
    buffer.WriteString list
 
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

'Character Editor
Sub SendAllCharacters(index As Long, Optional everyone As Boolean = False)
    Dim buffer As clsBuffer, i As Long
    
    Set buffer = New clsBuffer
    buffer.WriteLong SAllCharacters
    
    buffer.WriteString GetCharList
    
    SendDataTo index, buffer.ToArray
    
    Set buffer = Nothing
End Sub

Function PlayerData(ByVal index As Long) As Byte()
    Dim buffer As clsBuffer, i As Long

    If index < 1 Or index > Player_HighIndex Then Exit Function
    
    Set buffer = New clsBuffer
    
    buffer.WriteLong SPlayerData
    buffer.WriteLong index
    buffer.WriteInteger Account(index).Chars(GetPlayerChar(index)).Face
    buffer.WriteString GetPlayerName(index)
    buffer.WriteByte GetPlayerGender(index)
    buffer.WriteByte GetPlayerClass(index)
    buffer.WriteLong GetPlayerExp(index)
    buffer.WriteByte GetPlayerLevel(index)
    buffer.WriteInteger GetPlayerPoints(index)
    buffer.WriteInteger GetPlayerSprite(index)
    buffer.WriteInteger GetPlayerMap(index)
    buffer.WriteByte GetPlayerX(index)
    buffer.WriteByte GetPlayerY(index)
    buffer.WriteByte GetPlayerDir(index)
    buffer.WriteByte GetPlayerAccess(index)
    buffer.WriteByte GetPlayerPK(index)
    buffer.WriteLong tempplayer(index).PVPTimer
    
    If GetPlayerGuild(index) > 0 Then
        buffer.WriteString Guild(GetPlayerGuild(index)).Name
    Else
        buffer.WriteString vbNullString
    End If
    
    buffer.WriteByte GetPlayerGuildAccess(index)
    
    For i = 1 To Stats.Stat_count - 1
        buffer.WriteInteger GetPlayerStat(index, i)
    Next
    
    ' Amount of titles
    buffer.WriteByte Account(index).Chars(GetPlayerChar(index)).AmountOfTitles
    
    ' Send player titles
    For i = 1 To Account(index).Chars(GetPlayerChar(index)).AmountOfTitles
        buffer.WriteByte GetPlayerTitle(index, i)
    Next
    
    ' Send the player's current title
    buffer.WriteByte Account(index).Chars(GetPlayerChar(index)).CurrentTitle
    
    ' Send player status
    buffer.WriteString Account(index).Chars(GetPlayerChar(index)).Status
    
    For i = 1 To Skills.Skill_Count - 1
        buffer.WriteByte GetPlayerSkill(index, i)
        buffer.WriteLong GetPlayerSkillExp(index, i)
    Next

    PlayerData = buffer.ToArray()
    Set buffer = Nothing
End Function

Sub SendJoinMap(ByVal index As Long)
    Dim i As Long, x As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer

    ' Send all players on current map to index
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            If i <> index Then
                If GetPlayerMap(i) = GetPlayerMap(index) Then
                    SendDataTo index, PlayerData(i)
                    
                    For x = 1 To Vital_Count - 1
                        Call SendVitalTo(i, index, x)
                    Next
                End If
            End If
        End If
    Next
    
    ' Send index's player data to everyone on the map including themself
    SendDataToMap GetPlayerMap(index), PlayerData(index)
    
    ' Send the NPC targets to the player
    For i = 1 To Map(GetPlayerMap(index)).NPC_HighIndex
        If MapNPC(GetPlayerMap(index)).NPC(i).Num > 0 Then
            Call SendMapNPCTarget(GetPlayerMap(index), i, MapNPC(GetPlayerMap(index)).NPC(i).target, MapNPC(GetPlayerMap(index)).NPC(i).targetType)
        Else
            ' Send 0 so it uncaches any old data
            Call SendMapNPCTarget(GetPlayerMap(index), i, 0, 0)
        End If
    Next
    
    Set buffer = Nothing
End Sub

Sub SendLeaveMap(ByVal index As Long, ByVal MapNum As Integer)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.WriteLong SLeft
    buffer.WriteLong index
    SendDataToMapBut index, MapNum, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendPlayerData(ByVal index As Long)
    SendDataToMap GetPlayerMap(index), PlayerData(index)
End Sub

Sub SendAccessVerificator(ByVal index As Long, success As Byte, Message As String, currentAccess As Byte)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.WriteLong SAccessVerificator
    buffer.WriteByte success
    buffer.WriteString Message
    buffer.WriteByte currentAccess
    
    SendDataTo index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

'Character Editor
Sub SendExtendedPlayerData(index As Long, playerName As String)
    'Check if He is Online
    Dim i As Long, j As Long, tempPlayer_ As PlayerRec
    For i = 1 To MAX_PLAYERS
        For j = 1 To MAX_CHARS
            If Account(i).Login = "" Then GoTo use_offline_player
            If Trim$(Account(i).Chars(j).Name) = playerName Then
                tempPlayer_ = Account(i).Chars(j)
                GoTo use_online_player
            End If
        Next
    Next
    
use_offline_player:
    'Find associated Account Name
    Dim F As Long
    Dim s As String
    Dim charLogin() As String
    
    F = FreeFile
    
    Open App.path & "\data\accounts\charlist.txt" For Input As #F
        Do While Not EOF(F)
            Input #F, s
            charLogin = Split(s, ":")
            If charLogin(0) = playerName Then Exit Do
        Loop
    Close #F
    
    'Load Character into temp variable - charLogin(0) -> Character Name | charLogin(1) -> Account/Login Name
    Dim tempplayer As AccountRec
    Dim filename As String
    
    filename = App.path & "\data\accounts\" & charLogin(0) & "\data.bin"
    Call ChkDir(App.path & "\data\accounts\", charLogin(0))
    
    F = FreeFile
    
    If Not FileExist(filename, True) Then
        ' Erase that char name
        Call DeleteName(playerName)
        Call PlayerMsg(index, "This character doesn't exist and has been wiped from charlist.txt.", BrightRed)
        Call SendRefreshCharEditor(index)
        Exit Sub
    End If
    
    Open filename For Binary As #F
        Get #F, , tempplayer
    Close #F
    
    'Get Character info, that we are requesting -> playerName
    Dim requestedClientPlayer As PlayerEditableRec
    For i = 1 To MAX_CHARS
        If Trim$(Account(index).Chars(i).Name) = playerName Then
            tempPlayer_ = Account(index).Chars(i)
            Exit For
        End If
    Next
use_online_player:
    'Copy over data that's available...
    requestedClientPlayer.Name = tempPlayer_.Name
    requestedClientPlayer.Level = tempPlayer_.Level
    requestedClientPlayer.Class = tempPlayer_.Class
    requestedClientPlayer.Access = tempPlayer_.Access
    requestedClientPlayer.Exp = tempPlayer_.Exp
    requestedClientPlayer.Gender = tempPlayer_.Gender
    requestedClientPlayer.Login = "XXXX" ' Do we really want to edit it in client? Is it safe to send it?
    requestedClientPlayer.Password = "XXXX" ' Do we really want to edit it in client? Is it safe to send it?
    requestedClientPlayer.Points = tempPlayer_.Points
    requestedClientPlayer.Sprite = tempPlayer_.Sprite
    Dim tempSize As Long
    tempSize = LenB(tempPlayer_.Stat(1)) * UBound(tempPlayer_.Stat)
    CopyMemory ByVal VarPtr(requestedClientPlayer.Stat(1)), ByVal VarPtr(tempPlayer_.Stat(1)), tempSize
    tempSize = LenB(tempPlayer_.Vital(1)) * UBound(tempPlayer_.Vital)
    CopyMemory ByVal VarPtr(requestedClientPlayer.Vital(1)), ByVal VarPtr(tempPlayer_.Vital(1)), tempSize
    
    'Send Data Over Network to Admin
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SExtendedPlayerData
    
    Dim PlayerSize As Long
    Dim PlayerData() As Byte
    
    PlayerSize = LenB(requestedClientPlayer)
    ReDim PlayerData(PlayerSize - 1)
    CopyMemory PlayerData(0), ByVal VarPtr(requestedClientPlayer), PlayerSize
    buffer.WriteBytes PlayerData
    
    SendDataTo index, buffer.ToArray
    
    Set buffer = Nothing
    
    
End Sub

Sub SendMap(ByVal index As Long, ByVal MapNum As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.PreAllocate (UBound(MapCache(MapNum).Data) - LBound(MapCache(MapNum).Data)) + 5
    buffer.WriteLong SMapData
    buffer.WriteBytes MapCache(MapNum).Data()
    SendDataTo index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendMapItemsTo(ByVal index As Long, ByVal MapNum As Integer)
    Dim i As Long
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    
    buffer.WriteLong SMapItemData
    For i = 1 To MAX_MAP_ITEMS
        buffer.WriteString MapItem(MapNum, i).playerName
        buffer.WriteLong MapItem(MapNum, i).Num
        buffer.WriteLong MapItem(MapNum, i).Value
        buffer.WriteInteger MapItem(MapNum, i).Durability
        buffer.WriteByte MapItem(MapNum, i).x
        buffer.WriteByte MapItem(MapNum, i).Y
    Next

    SendDataTo index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendMapItemToMap(ByVal MapNum As Integer, ByVal MapSlotNum As Long)
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    
    buffer.WriteLong SSpawnItem
    buffer.WriteLong MapSlotNum
    buffer.WriteString MapItem(MapNum, MapSlotNum).playerName
    buffer.WriteLong MapItem(MapNum, MapSlotNum).Num
    buffer.WriteLong MapItem(MapNum, MapSlotNum).Value
    buffer.WriteInteger MapItem(MapNum, MapSlotNum).Durability
    buffer.WriteLong MapItem(MapNum, MapSlotNum).x
    buffer.WriteLong MapItem(MapNum, MapSlotNum).Y
    
    SendDataToMap MapNum, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendMapNPCVitals(ByVal MapNum As Integer, ByVal MapNPCNum As Byte)
    Dim buffer As clsBuffer, i As Long
    
    Set buffer = New clsBuffer
    
    buffer.WriteLong SMapNPCVitals
    buffer.WriteByte MapNPCNum
    For i = 1 To Vitals.Vital_Count - 1
        buffer.WriteLong MapNPC(MapNum).NPC(MapNPCNum).Vital(i)
    Next

    SendDataToMap MapNum, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendMapNPCTarget(ByVal MapNum As Integer, ByVal MapNPCNum As Byte, ByVal target As Byte, ByVal targetType As Byte)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    
    buffer.WriteLong SMapNPCTarget
    buffer.WriteByte MapNPCNum
    buffer.WriteByte target
    buffer.WriteByte targetType

    SendDataToMap MapNum, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendMapNPCsTo(ByVal index As Long, ByVal MapNum As Integer)
    Dim i As Long, x As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.WriteLong SMapNPCData

    For i = 1 To MAX_MAP_NPCS
        With MapNPC(MapNum).NPC(i)
            buffer.WriteLong .Num
            buffer.WriteLong .x
            buffer.WriteLong .Y
            buffer.WriteLong .Dir
            For x = 1 To Vitals.Vital_Count - 1
                buffer.WriteLong .Vital(x)
            Next
        End With
    Next

    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendMapNPCsToMap(ByVal MapNum As Integer)
    Dim i As Long, x As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.WriteLong SMapNPCData

    For i = 1 To MAX_MAP_NPCS
        With MapNPC(MapNum).NPC(i)
            buffer.WriteLong .Num
            buffer.WriteLong .x
            buffer.WriteLong .Y
            buffer.WriteLong .Dir
            For x = 1 To Vitals.Vital_Count - 1
                buffer.WriteLong .Vital(x)
            Next
        End With
    Next

    SendDataToMap MapNum, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendMorals(ByVal index As Long)
    Dim i As Long

    For i = 1 To MAX_MORALS
        If Len(Trim$(Moral(i).Name)) > 0 Then
            Call SendUpdateMoralTo(index, i)
        End If
    Next
End Sub

Sub SendClasses(ByVal index As Long)
    Dim i As Long
    
    For i = 1 To MAX_CLASSES
        If Len(Trim$(Class(i).Name)) > 0 Then
            Call SendUpdateClassTo(index, i)
        End If
    Next
End Sub

Sub SendEmoticons(ByVal index As Long)
    Dim i As Long
    
    For i = 1 To MAX_EMOTICONS
        If Len(Trim$(Emoticon(i).Command)) > 0 Then
            Call SendUpdateEmoticonTo(index, i)
        End If
    Next
End Sub

Sub SendItems(ByVal index As Long)
    Dim i As Long

    For i = 1 To MAX_ITEMS
        If Len(Trim$(Item(i).Name)) > 0 Then
            Call SendUpdateItemTo(index, i)
        End If
    Next
End Sub

Sub SendTitles(ByVal index As Long)
    Dim i As Long
    
    For i = 1 To MAX_TITLES
        If Len(Trim$(Title(i).Name)) > 0 Then
            Call SendUpdateTitleTo(index, i)
        End If
    Next
End Sub

Sub SendAnimations(ByVal index As Long)
    Dim i As Long

    For i = 1 To MAX_ANIMATIONS
        If Len(Trim$(Animation(i).Name)) > 0 Then
            Call SendUpdateAnimationTo(index, i)
        End If
    Next
End Sub

Sub SendNPCs(ByVal index As Long)
    Dim i As Long

    For i = 1 To MAX_NPCS
        If Len(Trim$(NPC(i).Name)) > 0 Then
            Call SendUpdateNPCTo(index, i)
        End If
    Next
End Sub

Sub SendResources(ByVal index As Long)
    Dim i As Long

    For i = 1 To MAX_RESOURCES
        If Len(Trim$(Resource(i).Name)) > 0 Then
            Call SendUpdateResourceTo(index, i)
        End If
    Next
End Sub

Sub SendInventory(ByVal index As Long)
    Dim i As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.WriteLong SPlayerInv

    For i = 1 To MAX_INV
        buffer.WriteLong GetPlayerInvItemNum(index, i)
        buffer.WriteLong GetPlayerInvItemValue(index, i)
        buffer.WriteInteger GetPlayerInvItemDur(index, i)
        buffer.WriteByte GetPlayerInvItemBind(index, i)
    Next

    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendInventoryUpdate(ByVal index As Long, ByVal InvSlot As Byte)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.WriteLong SPlayerInvUpdate
    buffer.WriteByte InvSlot
    buffer.WriteLong GetPlayerInvItemNum(index, InvSlot)
    buffer.WriteLong GetPlayerInvItemValue(index, InvSlot)
    buffer.WriteInteger GetPlayerInvItemDur(index, InvSlot)
    buffer.WriteByte GetPlayerInvItemBind(index, InvSlot)
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendWornEquipment(ByVal index As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    Dim i As Byte
    
    buffer.WriteLong SPlayerWornEq
    
    For i = 1 To Equipment.Equipment_Count - 1
        buffer.WriteLong GetPlayerEquipment(index, i)
    Next
    
    For i = 1 To Equipment.Equipment_Count - 1
        buffer.WriteInteger GetPlayerEquipmentDur(index, i)
    Next
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendMapEquipment(ByVal index As Long)
    Dim buffer As clsBuffer
    Dim i As Byte
    
    Set buffer = New clsBuffer
    buffer.WriteLong SMapWornEq
    
    buffer.WriteLong index
    
    For i = 1 To Equipment.Equipment_Count - 1
        buffer.WriteLong GetPlayerEquipment(index, i)
    Next

    SendDataToMap GetPlayerMap(index), buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendMapEquipmentTo(ByVal PlayerNum As Long, ByVal index As Long)
    Dim buffer As clsBuffer
    Dim i As Byte
    
    Set buffer = New clsBuffer
    
    buffer.WriteLong SMapWornEq
    buffer.WriteLong PlayerNum
    
    For i = 1 To Equipment.Equipment_Count - 1
        buffer.WriteLong GetPlayerEquipment(PlayerNum, i)
    Next
    
    SendDataTo index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendVital(ByVal index As Long, ByVal Vital As Vitals)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer

    Select Case Vital
        Case HP
            buffer.WriteLong SPlayerHP
            buffer.WriteLong index
            buffer.WriteLong GetPlayerMaxVital(index, Vitals.HP)
            buffer.WriteLong GetPlayerVital(index, Vitals.HP)
        Case MP
            buffer.WriteLong SPlayerMP
            buffer.WriteLong index
            buffer.WriteLong GetPlayerMaxVital(index, Vitals.MP)
            buffer.WriteLong GetPlayerVital(index, Vitals.MP)
    End Select

    SendDataToAll buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendVitalTo(ByVal index As Long, player As Long, ByVal Vital As Vitals)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer

    Select Case Vital
        Case HP
            buffer.WriteLong SPlayerHP
            buffer.WriteLong player
            buffer.WriteLong GetPlayerMaxVital(player, Vitals.HP)
            buffer.WriteLong GetPlayerVital(player, Vitals.HP)
        Case MP
            buffer.WriteLong SPlayerMP
            buffer.WriteLong player
            buffer.WriteLong GetPlayerMaxVital(player, Vitals.MP)
            buffer.WriteLong GetPlayerVital(player, Vitals.MP)
    End Select

    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendPlayerExp(ByVal index As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.WriteLong SPlayerEXP
    buffer.WriteLong index
    buffer.WriteLong GetPlayerExp(index)
    buffer.WriteLong GetPlayerNextLevel(index)
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendPlayerStats(ByVal index As Long)
    Dim i As Long
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SPlayerStats
    buffer.WriteLong index
    
    For i = 1 To Stats.Stat_count - 1
        buffer.WriteInteger GetPlayerStat(index, i)
    Next
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendPlayerPoints(ByVal index As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    
    buffer.WriteLong SPlayerPoints
    buffer.WriteLong index
    buffer.WriteInteger GetPlayerPoints(index)
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendPlayerLevel(ByVal index As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    
    buffer.WriteLong SPlayerLevel
    buffer.WriteLong index
    buffer.WriteByte GetPlayerLevel(index)
    
    SendDataToMap GetPlayerMap(index), buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendPlayerGuild(ByVal index As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    
    buffer.WriteLong SPlayerGuild
    buffer.WriteLong index
    
    If GetPlayerGuild(index) > 0 Then
        buffer.WriteString Guild(GetPlayerGuild(index)).Name
    Else
        buffer.WriteString vbNullString
    End If
    
    buffer.WriteByte GetPlayerGuildAccess(index)
    
    SendDataToMap GetPlayerMap(index), buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendPlayerSkills(ByVal index As Long)
    Dim i As Long
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SPlayerSkills
    buffer.WriteLong index
    
    For i = 1 To Skills.Skill_Count - 1
        buffer.WriteByte GetPlayerSkill(index, i)
        buffer.WriteLong GetPlayerSkillExp(index, i)
    Next
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendGuildInvite(ByVal index As Long, ByVal OtherPlayer As Long)
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SGuildInvite
    
    buffer.WriteString Trim$(Account(index).Chars(GetPlayerChar(index)).Name)
    buffer.WriteString Trim$(Guild(GetPlayerGuild(index)).Name)
    
    SendDataTo OtherPlayer, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendPlayerGuildMembers(ByVal index As Long, Optional ByVal Ignore As Byte = 0)
    Dim i As Long
    Dim PlayerArray() As String
    Dim PlayerCount As Long
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SGuildMembers
    
    PlayerCount = 0
    
    ' Count members online
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            If Not Ignore = i And Not i = index Then
                If GetPlayerGuild(i) = GetPlayerGuild(index) Then
                    PlayerCount = PlayerCount + 1
                    ReDim Preserve PlayerArray(1 To PlayerCount)
                    PlayerArray(UBound(PlayerArray)) = GetPlayerName(i)
                End If
            End If
        End If
    Next
    
    ' Add to Packet
    buffer.WriteLong PlayerCount
    
    If PlayerCount > 0 Then
        For i = 1 To PlayerCount
            buffer.WriteString PlayerArray(i)
        Next
    End If
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendPlayerSprite(ByVal index As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    
    buffer.WriteLong SPlayerSprite
    buffer.WriteLong index
    buffer.WriteInteger GetPlayerSprite(index)
      
    SendDataToMap GetPlayerMap(index), buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendPlayerTitles(ByVal index As Long)
    Dim buffer As clsBuffer
    Dim i As Long
    
    Set buffer = New clsBuffer
    
    buffer.WriteLong SPlayerTitles
    buffer.WriteLong index
    
    ' Amount of titles
    buffer.WriteByte Account(index).Chars(GetPlayerChar(index)).AmountOfTitles
    
    ' Send player titles
    For i = 1 To Account(index).Chars(GetPlayerChar(index)).AmountOfTitles
        buffer.WriteByte Account(index).Chars(GetPlayerChar(index)).Title(i)
    Next
    
    ' Send the player's current title
    buffer.WriteByte Account(index).Chars(GetPlayerChar(index)).CurrentTitle
    
    SendDataToMap GetPlayerMap(index), buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendPlayerStatus(ByVal index As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    
    buffer.WriteLong SPlayerStatus
    buffer.WriteLong index
    
    buffer.WriteString Account(index).Chars(GetPlayerChar(index)).Status
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendPlayerPK(ByVal index As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    
    buffer.WriteLong SPlayerPK
    buffer.WriteLong index
    
    buffer.WriteByte GetPlayerPK(index)
    buffer.WriteLong tempplayer(index).PVPTimer
    
    SendDataToMap GetPlayerMap(index), buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendWelcome(ByVal index As Long)
    ' Send the MOTD
    If Not Trim$(Options.MOTD) = vbNullString Then
        Call PlayerMsg(index, Options.MOTD, BrightCyan)
    End If
    
    ' Send the SMOTD
    If Not Trim$(Options.SMOTD) = vbNullString Then
        If GetPlayerAccess(index) >= STAFF_MODERATOR Then
            Call PlayerMsg(index, Options.SMOTD, Cyan)
        End If
    End If
    
    ' Send the GMOTD
    If GetPlayerGuild(index) > 0 Then
        If Not Trim$(Guild(GetPlayerGuild(index)).MOTD) = vbNullString Then
            Call PlayerMsg(index, Trim$(Guild(GetPlayerGuild(index)).MOTD), BrightGreen)
        End If
    End If
End Sub

Sub SendLeftGame(ByVal index As Long)
    Dim buffer As clsBuffer, i As Long
    
    Set buffer = New clsBuffer
    
    buffer.WriteLong SLeaveGame
    buffer.WriteLong index
    
    SendDataToAllBut index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendUpdateItemToAll(ByVal ItemNum As Integer)
    Dim buffer As clsBuffer
    Dim ItemSize As Long
    Dim ItemData() As Byte
    Set buffer = New clsBuffer
    ItemSize = LenB(Item(ItemNum))
    
    ReDim ItemData(ItemSize - 1)
    
    CopyMemory ItemData(0), ByVal VarPtr(Item(ItemNum)), ItemSize
    
    buffer.WriteLong SUpdateItem
    buffer.WriteLong ItemNum
    buffer.WriteBytes ItemData
    
    SendDataToAll buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendUpdateQuestTo(ByVal index As Long, ByVal QuestNum As Integer)
    Dim buffer As clsBuffer
    Dim i As Long, II As Long
    
    Set buffer = New clsBuffer
    
        buffer.WriteLong SUpdateQuest
        
        With Quest(QuestNum)
            buffer.WriteLong QuestNum
            buffer.WriteString .Name
            buffer.WriteString .Description
            buffer.WriteLong .CanBeRetaken
            buffer.WriteLong .Max_CLI
            
            For i = 1 To .Max_CLI
                buffer.WriteLong .CLI(i).ItemIndex
                buffer.WriteLong .CLI(i).isNPC
                buffer.WriteLong .CLI(i).Max_Actions
                
                For II = 1 To .CLI(i).Max_Actions
                    buffer.WriteString .CLI(i).Action(II).TextHolder
                    buffer.WriteLong .CLI(i).Action(II).ActionID
                    buffer.WriteLong .CLI(i).Action(II).Amount
                    buffer.WriteLong .CLI(i).Action(II).MainData
                    buffer.WriteLong .CLI(i).Action(II).QuadData
                    buffer.WriteLong .CLI(i).Action(II).SecondaryData
                    buffer.WriteLong .CLI(i).Action(II).TertiaryData
                Next II
            Next i
            
            buffer.WriteLong .Requirements.AccessReq
            buffer.WriteLong .Requirements.ClassReq
            buffer.WriteLong .Requirements.GenderReq
            buffer.WriteLong .Requirements.LevelReq
            buffer.WriteLong .Requirements.SkillLevelReq
            buffer.WriteLong .Requirements.SkillReq
            
            For i = 1 To Stats.Stat_count - 1
                buffer.WriteLong .Requirements.Stat_Req(i)
            Next i
        End With
        
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing

End Sub

Sub SendUpdateItemTo(ByVal index As Long, ByVal ItemNum As Integer)
    Dim buffer As clsBuffer
    Dim ItemSize As Long
    Dim ItemData() As Byte
    
    Set buffer = New clsBuffer
    
    ItemSize = LenB(Item(ItemNum))
    ReDim ItemData(ItemSize - 1)
    CopyMemory ItemData(0), ByVal VarPtr(Item(ItemNum)), ItemSize
    buffer.WriteLong SUpdateItem
    buffer.WriteLong ItemNum
    buffer.WriteBytes ItemData
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendUpdateAnimationToAll(ByVal AnimationNum As Long)
    Dim buffer As clsBuffer
    Dim AnimationSize As Long
    Dim AnimationData() As Byte
    
    Set buffer = New clsBuffer
    
    AnimationSize = LenB(Animation(AnimationNum))
    ReDim AnimationData(AnimationSize - 1)
    CopyMemory AnimationData(0), ByVal VarPtr(Animation(AnimationNum)), AnimationSize
    buffer.WriteLong SUpdateAnimation
    buffer.WriteLong AnimationNum
    buffer.WriteBytes AnimationData
    
    SendDataToAll buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendUpdateAnimationTo(ByVal index As Long, ByVal AnimationNum As Long)
    Dim buffer As clsBuffer
    Dim AnimationSize As Long
    Dim AnimationData() As Byte
    
    Set buffer = New clsBuffer
    
    AnimationSize = LenB(Animation(AnimationNum))
    ReDim AnimationData(AnimationSize - 1)
    CopyMemory AnimationData(0), ByVal VarPtr(Animation(AnimationNum)), AnimationSize
    buffer.WriteLong SUpdateAnimation
    buffer.WriteLong AnimationNum
    buffer.WriteBytes AnimationData
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub
Sub SendAssociatedCharacters()

End Sub
Sub SendUpdateNPCToAll(ByVal NPCNum As Long)
    Dim buffer As clsBuffer
    Dim NPCSize As Long
    Dim NPCData() As Byte
    
    Set buffer = New clsBuffer
    
    NPCSize = LenB(NPC(NPCNum))
    ReDim NPCData(NPCSize - 1)
    CopyMemory NPCData(0), ByVal VarPtr(NPC(NPCNum)), NPCSize
    buffer.WriteLong SUpdateNPC
    buffer.WriteLong NPCNum
    buffer.WriteBytes NPCData
    
    SendDataToAll buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendUpdateNPCTo(ByVal index As Long, ByVal NPCNum As Long)
    Dim buffer As clsBuffer
    Dim NPCSize As Long
    Dim NPCData() As Byte
    
    Set buffer = New clsBuffer
    
    NPCSize = LenB(NPC(NPCNum))
    ReDim NPCData(NPCSize - 1)
    CopyMemory NPCData(0), ByVal VarPtr(NPC(NPCNum)), NPCSize
    buffer.WriteLong SUpdateNPC
    buffer.WriteLong NPCNum
    buffer.WriteBytes NPCData
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendUpdateResourceToAll(ByVal ResourceNum As Long)
    Dim buffer As clsBuffer
    Dim ResourceSize As Long
    Dim ResourceData() As Byte
    
    Set buffer = New clsBuffer
    
    ResourceSize = LenB(Resource(ResourceNum))
    ReDim ResourceData(ResourceSize - 1)
    CopyMemory ResourceData(0), ByVal VarPtr(Resource(ResourceNum)), ResourceSize
    
    buffer.WriteLong SUpdateResource
    buffer.WriteLong ResourceNum
    buffer.WriteBytes ResourceData

    SendDataToAll buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendUpdateResourceTo(ByVal index As Long, ByVal ResourceNum As Long)
    Dim buffer As clsBuffer
    Dim ResourceSize As Long
    Dim ResourceData() As Byte
    
    Set buffer = New clsBuffer
    
    ResourceSize = LenB(Resource(ResourceNum))
    ReDim ResourceData(ResourceSize - 1)
    CopyMemory ResourceData(0), ByVal VarPtr(Resource(ResourceNum)), ResourceSize
    
    buffer.WriteLong SUpdateResource
    buffer.WriteLong ResourceNum
    buffer.WriteBytes ResourceData
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendShops(ByVal index As Long)
    Dim i As Long

    For i = 1 To MAX_SHOPS
        If Len(Trim$(Shop(i).Name)) > 0 Then
            Call SendUpdateShopTo(index, i)
        End If
    Next
End Sub

Sub SendUpdateShopToAll(ByVal ShopNum As Long)
    Dim buffer As clsBuffer
    Dim ShopSize As Long
    Dim ShopData() As Byte
    
    Set buffer = New clsBuffer
    
    ShopSize = LenB(Shop(ShopNum))
    ReDim ShopData(ShopSize - 1)
    CopyMemory ShopData(0), ByVal VarPtr(Shop(ShopNum)), ShopSize
    
    buffer.WriteLong SUpdateShop
    buffer.WriteLong ShopNum
    buffer.WriteBytes ShopData
    
    SendDataToAll buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendUpdateShopTo(ByVal index As Long, ByVal ShopNum As Long)
    Dim buffer As clsBuffer
    Dim ShopSize As Long
    Dim ShopData() As Byte
    
    Set buffer = New clsBuffer
    
    ShopSize = LenB(Shop(ShopNum))
    ReDim ShopData(ShopSize - 1)
    CopyMemory ShopData(0), ByVal VarPtr(Shop(ShopNum)), ShopSize
    
    buffer.WriteLong SUpdateShop
    buffer.WriteLong ShopNum
    buffer.WriteBytes ShopData
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendSpells(ByVal index As Long)
    Dim i As Long

    For i = 1 To MAX_SPELLS
        If Len(Trim$(Spell(i).Name)) > 0 Then
            Call SendUpdateSpellTo(index, i)
        End If
    Next
End Sub

Sub SendUpdateSpellToAll(ByVal SpellNum As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    Dim SpellSize As Long
    Dim SpellData() As Byte
    
    Set buffer = New clsBuffer
    
    SpellSize = LenB(Spell(SpellNum))
    ReDim SpellData(SpellSize - 1)
    CopyMemory SpellData(0), ByVal VarPtr(Spell(SpellNum)), SpellSize
    
    buffer.WriteLong SUpdateSpell
    buffer.WriteLong SpellNum
    buffer.WriteBytes SpellData
    
    SendDataToAll buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendUpdateSpellTo(ByVal index As Long, ByVal SpellNum As Long)
    Dim buffer As clsBuffer
    Dim SpellSize As Long
    Dim SpellData() As Byte
    
    Set buffer = New clsBuffer
    
    SpellSize = LenB(Spell(SpellNum))
    ReDim SpellData(SpellSize - 1)
    CopyMemory SpellData(0), ByVal VarPtr(Spell(SpellNum)), SpellSize
    
    buffer.WriteLong SUpdateSpell
    buffer.WriteLong SpellNum
    buffer.WriteBytes SpellData
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendPlayerSpells(ByVal index As Long)
    Dim i As Long
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    
    buffer.WriteLong SSpells
    For i = 1 To MAX_PLAYER_SPELLS
        buffer.WriteLong GetPlayerSpell(index, i)
    Next
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendPlayerSpell(ByVal index As Long, ByVal SpellSlot As Byte)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    
    buffer.WriteLong SSpell
    buffer.WriteByte SpellSlot
    buffer.WriteLong GetPlayerSpell(index, SpellSlot)
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendResourceCacheTo(ByVal index As Long)
    Dim buffer As clsBuffer
    Dim i As Long
    Set buffer = New clsBuffer
    
    buffer.WriteLong SResourceCache
    buffer.WriteLong ResourceCache(GetPlayerMap(index)).Resource_Count

    If ResourceCache(GetPlayerMap(index)).Resource_Count > 0 Then
        For i = 0 To ResourceCache(GetPlayerMap(index)).Resource_Count
            buffer.WriteByte ResourceCache(GetPlayerMap(index)).ResourceData(i).ResourceState
            buffer.WriteInteger ResourceCache(GetPlayerMap(index)).ResourceData(i).x
            buffer.WriteInteger ResourceCache(GetPlayerMap(index)).ResourceData(i).Y
        Next
    End If

    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendResourceCacheToMap(ByVal MapNum As Integer)
    Dim buffer As clsBuffer
    Dim i As Long
    Set buffer = New clsBuffer
    
    buffer.WriteLong SResourceCache
    buffer.WriteLong ResourceCache(MapNum).Resource_Count

    If ResourceCache(MapNum).Resource_Count > 0 Then
        For i = 0 To ResourceCache(MapNum).Resource_Count
            buffer.WriteByte ResourceCache(MapNum).ResourceData(i).ResourceState
            buffer.WriteInteger ResourceCache(MapNum).ResourceData(i).x
            buffer.WriteInteger ResourceCache(MapNum).ResourceData(i).Y
        Next
    End If

    SendDataToMap MapNum, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendDoorAnimation(ByVal MapNum As Integer, ByVal x As Long, ByVal Y As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SDoorAnimation
    buffer.WriteLong x
    buffer.WriteLong Y
    
    SendDataToMap MapNum, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendActionMsg(ByVal MapNum As Integer, ByVal Message As String, ByVal Color As Long, ByVal MsgType As Long, ByVal x As Long, ByVal Y As Long, Optional PlayerOnlyNum As Long = 0)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SActionMsg
    buffer.WriteString Message
    buffer.WriteLong Color
    buffer.WriteLong MsgType
    buffer.WriteLong x
    buffer.WriteLong Y
    
    If PlayerOnlyNum > 0 Then
        SendDataTo PlayerOnlyNum, buffer.ToArray()
    Else
        SendDataToMap MapNum, buffer.ToArray()
    End If
    
    Set buffer = Nothing
End Sub

Sub SendBlood(ByVal MapNum As Integer, ByVal x As Long, ByVal Y As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SBlood
    buffer.WriteLong x
    buffer.WriteLong Y
    
    SendDataToMap MapNum, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendAnimation(ByVal MapNum As Integer, ByVal Anim As Long, ByVal x As Long, ByVal Y As Long, Optional ByVal LockType As Byte = 0, Optional ByVal LockIndex As Long = 0, Optional ByVal OnlyTo As Long = 0)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SAnimation
    buffer.WriteLong Anim
    buffer.WriteLong x
    buffer.WriteLong Y
    buffer.WriteByte LockType
    buffer.WriteLong LockIndex
    
    If OnlyTo > 0 Then
        SendDataTo OnlyTo, buffer.ToArray
    Else
        SendDataToMap MapNum, buffer.ToArray()
    End If
    Set buffer = Nothing
End Sub

Sub SendSpellCooldown(ByVal index As Long, ByVal Slot As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SSpellCooldown
    buffer.WriteByte Slot
    buffer.WriteLong GetPlayerSpellCD(index, Slot)
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendClearAccountSpellBuffer(ByVal index As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SClearSpellBuffer
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SayMsg_Map(ByVal MapNum As Integer, ByVal index As Long, ByVal Message As String, ByVal SayColor As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SSayMsg
    buffer.WriteString GetPlayerName(index)
    buffer.WriteLong GetPlayerAccess(index)
    buffer.WriteLong GetPlayerPK(index)
    buffer.WriteString Message
    buffer.WriteString "[Map] "
    buffer.WriteLong SayColor
    
    SendDataToMap MapNum, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SayMsg_Global(ByVal index As Long, ByVal Message As String, ByVal SayColor As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SSayMsg
    
    buffer.WriteString GetPlayerName(index)
    buffer.WriteLong GetPlayerAccess(index)
    buffer.WriteLong GetPlayerPK(index)
    buffer.WriteString Message
    buffer.WriteString "[Global] "
    buffer.WriteLong SayColor
    
    SendDataToAll buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub ResetShopAction(ByVal index As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SResetShopAction
    
    SendDataToAll buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendStunned(ByVal index As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SStunned
    
    buffer.WriteLong tempplayer(index).StunDuration
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendBank(ByVal index As Long)
    Dim buffer As clsBuffer
    Dim i As Long
    
    Set buffer = New clsBuffer
    buffer.WriteLong sbank
    
    For i = 1 To MAX_BANK
        buffer.WriteLong Account(index).Bank.Item(i).Num
        buffer.WriteLong Account(index).Bank.Item(i).Value
    Next
    
    SendDataTo index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendOpenShop(ByVal index As Long, ByVal ShopNum As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SOpenShop
    
    buffer.WriteLong ShopNum
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendPlayerMove(ByVal index As Long, ByVal movement As Byte, Optional ByVal SendToSelf As Boolean = False)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SPlayerMove
    buffer.WriteLong index
    buffer.WriteLong GetPlayerX(index)
    buffer.WriteLong GetPlayerY(index)
    buffer.WriteLong GetPlayerDir(index)
    buffer.WriteLong movement

    If Not SendToSelf Then
        SendDataToMap GetPlayerMap(index), buffer.ToArray()
    Else
        SendDataTo index, buffer.ToArray()
    End If
    Set buffer = Nothing
End Sub

Sub SendPlayerXY(ByVal index As Long, Optional ByVal SendToSelf As Boolean = False)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SPlayerWarp
    
    buffer.WriteLong index
    buffer.WriteByte GetPlayerX(index)
    buffer.WriteByte GetPlayerY(index)
    buffer.WriteByte GetPlayerDir(index)
    
    If Not SendToSelf Then
        SendDataToMap GetPlayerMap(index), buffer.ToArray()
    Else
        SendDataTo index, buffer.ToArray()
    End If
    Set buffer = Nothing
End Sub

Sub SendNPCMove(ByVal MapNPCNum As Long, ByVal movement As Byte, ByVal MapNum As Integer)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SNPCMove
    
    buffer.WriteLong MapNPCNum
    buffer.WriteByte MapNPC(MapNum).NPC(MapNPCNum).x
    buffer.WriteByte MapNPC(MapNum).NPC(MapNPCNum).Y
    buffer.WriteByte MapNPC(MapNum).NPC(MapNPCNum).Dir
    buffer.WriteByte movement
    
    SendDataToMap MapNum, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendTrade(ByVal index As Long, ByVal TradeTarget As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong STrade
    
    buffer.WriteLong TradeTarget
    buffer.WriteString Trim$(GetPlayerName(TradeTarget))
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendCloseTrade(ByVal index As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SCloseTrade
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendTradeUpdate(ByVal index As Long, ByVal dataType As Byte)
    Dim buffer As clsBuffer
    Dim i As Long
    Dim TradeTarget As Long
    Dim TotalWorth As Long
    
    TradeTarget = tempplayer(index).InTrade
    
    Set buffer = New clsBuffer
    buffer.WriteLong STradeUpdate
    
    buffer.WriteByte dataType
    If dataType = 0 Then ' own inventory
        For i = 1 To MAX_INV
            buffer.WriteLong tempplayer(index).TradeOffer(i).Num
            buffer.WriteLong tempplayer(index).TradeOffer(i).Value
            ' add total worth
            If tempplayer(index).TradeOffer(i).Num > 0 Then
                 If GetPlayerInvItemNum(index, tempplayer(index).TradeOffer(i).Num) > 0 Then
                    ' currency?
                    If Item(tempplayer(index).TradeOffer(i).Num).stackable = 1 Then
                        TotalWorth = TotalWorth + (Item(GetPlayerInvItemNum(index, tempplayer(index).TradeOffer(i).Num)).Price * tempplayer(index).TradeOffer(i).Value)
                    Else
                        TotalWorth = TotalWorth + Item(GetPlayerInvItemNum(index, tempplayer(index).TradeOffer(i).Num)).Price
                    End If
                End If
            End If
        Next
    ElseIf dataType = 1 Then ' other inventory
        For i = 1 To MAX_INV
            buffer.WriteLong GetPlayerInvItemNum(TradeTarget, tempplayer(TradeTarget).TradeOffer(i).Num)
            buffer.WriteLong tempplayer(TradeTarget).TradeOffer(i).Value
            ' add total worth
            If GetPlayerInvItemNum(TradeTarget, tempplayer(TradeTarget).TradeOffer(i).Num) > 0 Then
                ' currency?
                If Item(GetPlayerInvItemNum(TradeTarget, tempplayer(TradeTarget).TradeOffer(i).Num)).stackable = 1 Then
                    TotalWorth = TotalWorth + (Item(GetPlayerInvItemNum(TradeTarget, tempplayer(TradeTarget).TradeOffer(i).Num)).Price * tempplayer(TradeTarget).TradeOffer(i).Value)
                Else
                    TotalWorth = TotalWorth + Item(GetPlayerInvItemNum(TradeTarget, tempplayer(TradeTarget).TradeOffer(i).Num)).Price
                End If
            End If
        Next
    End If
    ' Send total worth of trade
    buffer.WriteLong TotalWorth
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendTradeStatus(ByVal index As Long, ByVal Status As Byte)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong STradeStatus
    
    buffer.WriteByte Status
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendAttack(ByVal index As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SAttack
    buffer.WriteLong index
    
    SendDataToMapBut index, GetPlayerMap(index), buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendPlayerTarget(ByVal index As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong STarget
    
    buffer.WriteByte tempplayer(index).target
    buffer.WriteByte tempplayer(index).targetType
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendHotbar(ByVal index As Long)
    Dim i As Long
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SHotbar
    
    For i = 1 To MAX_HOTBAR
        buffer.WriteByte Account(index).Chars(GetPlayerChar(index)).Hotbar(i).Slot
        buffer.WriteByte Account(index).Chars(GetPlayerChar(index)).Hotbar(i).SType
    Next
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendNPCSpellBuffer(ByRef MapNum As Integer, ByVal MapNPCNum As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SNPCSpellBuffer
    
    buffer.WriteLong MapNPCNum
    buffer.WriteLong MapNPC(MapNum).NPC(MapNPCNum).SpellBuffer.Spell
    
    Call SendDataToMap(MapNum, buffer.ToArray)
    Set buffer = Nothing
End Sub

Sub SendLogs(ByVal index As Long, Msg As String, Name As String)
    Dim buffer As clsBuffer
    Dim LogSize As Long
    Dim LogData() As Byte
    
    Set buffer = New clsBuffer
    
    Log.Msg = Msg
    Log.File = Name
    LogSize = LenB(Log)
    ReDim LogData(LogSize - 1)
    CopyMemory LogData(0), ByVal VarPtr(Log), LogSize
    buffer.WriteLong SUpdateLogs
    
    buffer.WriteBytes LogData
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub UpdateFriendsList(ByVal index As Integer)
    Dim buffer As clsBuffer
    Dim Name As String
    Dim i As Long, n As Long
    
    Set buffer = New clsBuffer
    buffer.WriteLong SFriendsList
    
    If Account(index).Friends.AmountOfFriends = 0 Then
        buffer.WriteByte Account(index).Friends.AmountOfFriends
        GoTo Finish
    End If
   
    ' Sends the amount of friends in friends list
    buffer.WriteByte Account(index).Friends.AmountOfFriends
   
    ' Check to see if they are online
    For i = 1 To Account(index).Friends.AmountOfFriends
        Name = Trim$(Account(index).Friends.Members(i))
        buffer.WriteString Name
        For n = 1 To Player_HighIndex
            If IsPlaying(FindPlayer(Name)) Then
                buffer.WriteString Name & " Online"
            Else
                buffer.WriteString Name & " Offline"
            End If
        Next
    Next
    
Finish:
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub UpdateFoesList(ByVal index As Integer)
    Dim buffer As clsBuffer
    Dim Name As String
    Dim i As Long, n As Long
    
    Set buffer = New clsBuffer
    buffer.WriteLong SFoesList
    
    If Account(index).Foes.Amount = 0 Then
        buffer.WriteByte Account(index).Foes.Amount
        GoTo Finish
    End If
   
    ' Sends the amount of Foes in Foes list
    buffer.WriteByte Account(index).Foes.Amount
   
    ' Check to see if they are online
    For i = 1 To Account(index).Foes.Amount
        Name = Trim$(Account(index).Foes.Members(i))
        buffer.WriteString Name
        For n = 1 To Player_HighIndex
            If IsPlaying(FindPlayer(Name)) Then
                buffer.WriteString Name & " Online"
            Else
                buffer.WriteString Name & " Offline"
            End If
        Next
    Next
    
Finish:
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendInGame(ByVal index As Long)
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SInGame
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendPlayer_HighIndex()
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SHighIndex
    
    buffer.WriteLong Player_HighIndex
    
    SendDataToAll buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendSoundTo(ByVal index As Integer, Sound As String)
    Dim buffer As clsBuffer
    
    ' Don't send it if there's nothing to send
    If Sound = vbNullString Then Exit Sub
    
    Set buffer = New clsBuffer
    buffer.WriteLong SSound
    
    buffer.WriteString Sound
    SendDataTo index, buffer.ToArray()
End Sub

Sub SendSoundToMap(ByVal MapNum As Integer, Sound As String)
    Dim buffer As clsBuffer
    
    ' Don't send it if there's nothing to send
    If Sound = vbNullString Then Exit Sub
    
    Set buffer = New clsBuffer
    buffer.WriteLong SSound
    
    buffer.WriteString Sound
    SendDataToMap MapNum, buffer.ToArray()
End Sub

Sub SendSoundToAll(ByVal MapNum As Integer, Sound As String)
    Dim buffer As clsBuffer
    
    ' Don't send it if there's nothing to send
    If Sound = vbNullString Then Exit Sub
    
    Set buffer = New clsBuffer
    buffer.WriteLong SSound
    
    buffer.WriteString Sound
    SendDataToAll buffer.ToArray()
End Sub

Sub SendPlayerSound(ByVal index As Long, ByVal x As Long, ByVal Y As Long, ByVal EntityType As Long, ByVal EntityNum As Long)
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SEntitySound
    
    buffer.WriteLong x
    buffer.WriteLong Y
    buffer.WriteLong EntityType
    buffer.WriteLong EntityNum
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendMapSound(ByVal MapNum As Integer, ByVal index As Long, ByVal x As Long, ByVal Y As Long, ByVal EntityType As Long, ByVal EntityNum As Long)
    Dim buffer As clsBuffer

    If EntityNum <= 0 Then Exit Sub
    
    Set buffer = New clsBuffer
    buffer.WriteLong SEntitySound
    
    buffer.WriteLong x
    buffer.WriteLong Y
    buffer.WriteLong EntityType
    buffer.WriteLong EntityNum
    
    SendDataToMap MapNum, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendNPCDeath(ByVal MapNPCNum As Long, MapNum As Integer)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SNPCDead
    
    buffer.WriteLong MapNPCNum
    
    SendDataToMap MapNum, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendNPCAttack(ByVal Attacker As Long, MapNum As Integer)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SNPCAttack
    
    buffer.WriteLong Attacker
    
    SendDataToMap MapNum, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendLogin(ByVal index As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SLoginOk
    
    buffer.WriteLong index
    buffer.WriteLong Player_HighIndex
    buffer.WriteLong Options.MaxLevel
    buffer.WriteLong Options.MaxStat
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendNewCharClasses(ByVal index As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SNewCharClasses
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendGameData(ByVal index As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SGameData
    
    buffer.WriteString Options.Name
    buffer.WriteString Options.Website
    
    buffer.WriteLong MAX_MAPS
    buffer.WriteLong MAX_ITEMS
    buffer.WriteLong MAX_NPCS
    buffer.WriteLong MAX_ANIMATIONS
    buffer.WriteLong MAX_SHOPS
    buffer.WriteLong MAX_SPELLS
    buffer.WriteLong MAX_RESOURCES
    buffer.WriteLong MAX_QUESTS
    buffer.WriteLong MAX_BANS
    buffer.WriteLong MAX_TITLES
    buffer.WriteLong MAX_MORALS
    buffer.WriteLong MAX_CLASSES
    buffer.WriteLong MAX_EMOTICONS
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendNews(ByVal index As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SSendNews
    
    buffer.WriteString Options.News
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendCheckForMap(ByVal index As Long, MapNum As Integer)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SCheckForMap
    
    buffer.WriteInteger MapNum
    buffer.WriteInteger Map(MapNum).Revision
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendTradeRequest(ByVal index As Long, ByVal TradeRequest As Long)
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong STradeRequest
    
    buffer.WriteString Trim$(GetPlayerName(TradeRequest))
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendPartyInvite(ByVal index As Long, ByVal OtherPlayer As Long)
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SPartyInvite
    
    buffer.WriteString Trim$(Account(OtherPlayer).Chars(GetPlayerChar(OtherPlayer)).Name)
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendPartyUpdate(ByVal PartyNum As Long)
    Dim buffer As clsBuffer, i As Long

    Set buffer = New clsBuffer
    buffer.WriteLong SPartyUpdate
    
    buffer.WriteByte 1
    For i = 1 To MAX_PARTY_MEMBERS
        buffer.WriteLong Party(PartyNum).Member(i)
    Next
    buffer.WriteLong Party(PartyNum).MemberCount
    buffer.WriteLong PartyNum
    
    SendDataToParty PartyNum, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendPartyUpdateTo(ByVal index As Long)
    Dim buffer As clsBuffer, i As Long, PartyNum As Long
    
    Set buffer = New clsBuffer
    buffer.WriteLong SPartyUpdate
    
    ' Check if we're in a party
    PartyNum = tempplayer(index).InParty
    
    If PartyNum > 0 Then
        ' Send party data
        buffer.WriteByte 1
        buffer.WriteLong Party(PartyNum).Leader
        For i = 1 To MAX_PARTY_MEMBERS
            buffer.WriteLong Party(PartyNum).Member(i)
        Next
        buffer.WriteLong Party(PartyNum).MemberCount
    Else
        ' Send clear command
        buffer.WriteByte 0
    End If
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendPartyVitals(ByVal PartyNum As Long, ByVal index As Long)
    Dim buffer As clsBuffer, i As Long

    Set buffer = New clsBuffer
    buffer.WriteLong SPartyVitals
    
    buffer.WriteLong index
    For i = 1 To Vitals.Vital_Count - 1
        buffer.WriteLong GetPlayerMaxVital(index, i)
        buffer.WriteLong Account(index).Chars(GetPlayerChar(index)).Vital(i)
    Next
    
    SendDataToParty PartyNum, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendUpdateTitleToAll(ByVal TitleNum As Long)
    Dim buffer As clsBuffer
    Dim TitleSize As Long
    Dim TitleData() As Byte
    
    Set buffer = New clsBuffer
    
    TitleSize = LenB(Title(TitleNum))
    ReDim TitleData(TitleSize - 1)
    CopyMemory TitleData(0), ByVal VarPtr(Title(TitleNum)), TitleSize
    buffer.WriteLong SUpdateTitle
    
    buffer.WriteLong TitleNum
    buffer.WriteBytes TitleData
    
    SendDataToAll buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendUpdateTitleTo(ByVal index As Long, ByVal TitleNum As Long)
    Dim buffer As clsBuffer
    Dim TitleSize As Long
    Dim TitleData() As Byte
    
    Set buffer = New clsBuffer
    
    TitleSize = LenB(Title(TitleNum))
    ReDim TitleData(TitleSize - 1)
    CopyMemory TitleData(0), ByVal VarPtr(Title(TitleNum)), TitleSize
    buffer.WriteLong SUpdateTitle
    
    buffer.WriteLong TitleNum
    buffer.WriteBytes TitleData
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendCloseClient(index As Long)
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SCloseClient
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendUpdateMoralToAll(ByVal MoralNum As Long)
    Dim buffer As clsBuffer
    Dim MoralSize As Long
    Dim MoralData() As Byte
    
    Set buffer = New clsBuffer
    
    MoralSize = LenB(Moral(MoralNum))
    ReDim MoralData(MoralSize - 1)
    CopyMemory MoralData(0), ByVal VarPtr(Moral(MoralNum)), MoralSize
    buffer.WriteLong SUpdateMoral
    
    buffer.WriteLong MoralNum
    buffer.WriteBytes MoralData
    
    SendDataToAll buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendUpdateMoralTo(ByVal index As Long, ByVal MoralNum As Long)
    Dim buffer As clsBuffer
    Dim MoralSize As Long
    Dim MoralData() As Byte
    
    Set buffer = New clsBuffer
    
    MoralSize = LenB(Moral(MoralNum))
    ReDim MoralData(MoralSize - 1)
    CopyMemory MoralData(0), ByVal VarPtr(Moral(MoralNum)), MoralSize
    buffer.WriteLong SUpdateMoral
    
    buffer.WriteLong MoralNum
    buffer.WriteBytes MoralData
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendUpdateClassTo(ByVal index As Long, ByVal ClassNum As Long)
    Dim buffer As clsBuffer
    Dim Classesize As Long
    Dim ClassData() As Byte
    
    Set buffer = New clsBuffer
    
    Classesize = LenB(Class(ClassNum))
    ReDim ClassData(Classesize - 1)
    CopyMemory ClassData(0), ByVal VarPtr(Class(ClassNum)), Classesize
    buffer.WriteLong SUpdateClass
    
    buffer.WriteLong ClassNum
    buffer.WriteBytes ClassData
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendUpdateEmoticonTo(ByVal index As Long, ByVal EmoticonNum As Long)
    Dim buffer As clsBuffer
    Dim EmoticonSize As Long
    Dim EmoticonData() As Byte
    
    Set buffer = New clsBuffer

    EmoticonSize = LenB(Emoticon(EmoticonNum))
    ReDim EmoticonData(EmoticonSize - 1)
    CopyMemory EmoticonData(0), ByVal VarPtr(Emoticon(EmoticonNum)), EmoticonSize
    buffer.WriteLong SUpdateEmoticon
    
    buffer.WriteLong EmoticonNum
    buffer.WriteBytes EmoticonData
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendUpdateEmoticonToAll(ByVal EmoticonNum As Integer)
    Dim buffer As clsBuffer
    Dim EmoticonSize As Long
    Dim EmoticonData() As Byte
    Set buffer = New clsBuffer
    EmoticonSize = LenB(Emoticon(EmoticonNum))
    
    ReDim EmoticonData(EmoticonSize - 1)
    
    CopyMemory EmoticonData(0), ByVal VarPtr(Emoticon(EmoticonNum)), EmoticonSize
    
    buffer.WriteLong SUpdateEmoticon
    buffer.WriteLong EmoticonNum
    buffer.WriteBytes EmoticonData
    
    SendDataToAll buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendCheckEmoticon(ByVal index As Long, ByVal MapNum As Long, ByVal EmoticonNum As Long)
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SCheckEmoticon
    
    buffer.WriteLong index
    buffer.WriteLong EmoticonNum
    
    SendDataToMap MapNum, buffer.ToArray()
End Sub

Sub SendChatBubble(ByVal MapNum As Long, ByVal target As Long, ByVal targetType As Long, ByVal Message As String, ByVal Color As Long)
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SChatBubble
    
    buffer.WriteLong target
    buffer.WriteLong targetType
    buffer.WriteString Message
    buffer.WriteLong Color
    
    SendDataToMap MapNum, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendSpecialEffect(ByVal index As Long, EffectType As Long, Optional Data1 As Long = 0, Optional Data2 As Long = 0, Optional Data3 As Long = 0, Optional Data4 As Long = 0)
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SSpecialEffect
    
    Select Case EffectType
        Case EFFECT_TYPE_FADEIN
            buffer.WriteLong EffectType
        Case EFFECT_TYPE_FADEOUT
            buffer.WriteLong EffectType
        Case EFFECT_TYPE_FLASH
            buffer.WriteLong EffectType
        Case EFFECT_TYPE_FOG
            buffer.WriteLong EffectType
            buffer.WriteLong Data1 ' Fog num
            buffer.WriteLong Data2 ' Fog movement speed
            buffer.WriteLong Data3 ' Opacity
        Case EFFECT_TYPE_WEATHER
            buffer.WriteLong EffectType
            buffer.WriteLong Data1 ' Weather type
            buffer.WriteLong Data2 ' Weather intensity
        Case EFFECT_TYPE_TINT
            buffer.WriteLong EffectType
            buffer.WriteLong Data1 ' Red
            buffer.WriteLong Data2 ' Green
            buffer.WriteLong Data3 ' Blue
            buffer.WriteLong Data4 ' Alpha
    End Select
    
    SendDataTo index, buffer.ToArray
    Set buffer = Nothing
End Sub
