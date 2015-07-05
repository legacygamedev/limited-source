Attribute VB_Name = "modGeneral"
Option Explicit

Public Declare Function GetTickCount Lib "kernel32" () As Long

' Version constants
Public Const CLIENT_MAJOR = 0
Public Const CLIENT_MINOR = 0
Public Const CLIENT_REVISION = 1

' Used for respawning items
Public SpawnSeconds As Long

' Used for weather effects
Public GameWeather As Long
Public WeatherSeconds As Long
Public GameTime As Long
Public TimeSeconds As Long

' Used for closing key doors again
Public KeyTimer As Long

' Used for gradually giving back players and npcs hp
Public GiveHPTimer As Long
Public GiveNPCHPTimer As Long

' Used for logging
Public ServerLog As Boolean

Sub InitServer()
Dim IPMask As String
Dim i As Long
Dim f As Long
    
    Randomize Timer
    
    nid.cbSize = Len(nid)
    nid.hWnd = frmServer.hWnd
    nid.uId = vbNull
    nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    nid.uCallBackMessage = WM_MOUSEMOVE
    nid.hIcon = frmServer.Icon
    nid.szTip = "Mirage Server" & vbNullChar
    ' Add to the sys tray
    Call Shell_NotifyIcon(NIM_ADD, nid)
    
    ' Init atmosphere
    GameWeather = WEATHER_NONE
    WeatherSeconds = 0
    GameTime = TIME_DAY
    TimeSeconds = 0
    
    ' Check if the maps directory is there, if its not make it
    If LCase(Dir(App.Path & "\System\maps", vbDirectory)) <> "maps" Then
        Call MkDir(App.Path & "\System\Maps")
    End If
    
    ' Check if the accounts directory is there, if its not make it
    If LCase(Dir(App.Path & "\accounts", vbDirectory)) <> "accounts" Then
        Call MkDir(App.Path & "\Accounts")
    End If
    
    ' Check if the spells directory is there, if its not make it
    If LCase(Dir(App.Path & "\System\spells", vbDirectory)) <> "spells" Then
        Call MkDir(App.Path & "\System\Spells")
    End If
    
    ' Check if the spells directory is there, if its not make it
    If LCase(Dir(App.Path & "\classes", vbDirectory)) <> "classes" Then
        Call MkDir(App.Path & "\Classes")
    End If
    
    ' Check if the spells directory is there, if its not make it
    If LCase(Dir(App.Path & "\System\Shops", vbDirectory)) <> "shops" Then
        Call MkDir(App.Path & "\System\Shops")
    End If
    
    ' Check if the spells directory is there, if its not make it
    If LCase(Dir(App.Path & "\System\npcs", vbDirectory)) <> "npcs" Then
        Call MkDir(App.Path & "\System\Npcs")
    End If
    
    ' Check if the spells directory is there, if its not make it
    If LCase(Dir(App.Path & "\System\items", vbDirectory)) <> "items" Then
        Call MkDir(App.Path & "\System\Items")
    End If
    
    SEP_CHAR = Chr(0)
    END_CHAR = Chr(237)
    
    ServerLog = False
    
    GAME_NAME = Trim$(GetVar(App.Path & "\System\Config\Data.ini", "GAME", "Name"))
    GAME_PORT = GetVar(App.Path & "\System\Config\Data.ini", "GAME", "Port")
    MAX_PLAYERS = Val(GetVar(App.Path & "\System\Config\Data.ini", "MAX", "Players"))
    MAX_MAPS = Val(GetVar(App.Path & "\System\Config\Data.ini", "MAX", "Maps"))
    MAX_ITEMS = Val(GetVar(App.Path & "\System\Config\Data.ini", "MAX", "Items"))
    MAX_NPCS = Val(GetVar(App.Path & "\System\Config\Data.ini", "MAX", "Npcs"))
    MAX_MAP_ITEMS = Val(GetVar(App.Path & "\System\Config\Data.ini", "MAX", "Map Items"))
    MAX_SHOPS = Val(GetVar(App.Path & "\System\Config\Data.ini", "MAX", "Shops"))
    MAX_SPELLS = Val(GetVar(App.Path & "\System\Config\Data.ini", "MAX", "Spells"))
    MAX_GUILDS = Val(GetVar(App.Path & "\System\Config\Data.ini", "MAX", "Guilds"))
    MAX_GUILD_MEMBERS = Val(GetVar(App.Path & "\System\Config\Data.ini", "MAX", "Guild Members"))
    MAX_LEVEL = Val(GetVar(App.Path & "\System\Config\Data.ini", "MAX", "Level"))
    
    ReDim Player(1 To MAX_PLAYERS) As AccountRec
    ReDim Map(1 To MAX_MAPS) As MapRec
    ReDim Item(0 To MAX_ITEMS) As ItemRec
    ReDim Npc(0 To MAX_NPCS) As NpcRec
    ReDim MapItem(1 To MAX_MAPS, 1 To MAX_MAP_ITEMS) As MapItemRec
    ReDim Shop(1 To MAX_SHOPS) As ShopRec
    ReDim Spell(1 To MAX_SPELLS) As SpellRec
    ReDim Guild(1 To MAX_GUILDS) As GuildRec
    ReDim TempTile(1 To MAX_MAPS) As TempTileRec
    ReDim PlayersOnMap(1 To MAX_MAPS) As Long
    ReDim MapNpc(1 To MAX_MAPS, 1 To MAX_MAP_NPCS) As MapNpcRec
    
    For i = 1 To MAX_GUILDS
        ReDim Guild(i).Member(1 To MAX_GUILD_MEMBERS) As String * NAME_LENGTH
    Next
    
    ' Get the listening socket ready to go
    Set GameServer = New clsServer
        
    ' Init all the player sockets
    For i = 1 To MAX_PLAYERS
        Call ClearPlayer(i)
        
        Call GameServer.Sockets.Add(CStr(i))
    Next i
    
    Call ClearTempTile
    Call ClearMaps
    Call ClearMapItems
    Call ClearMapNpcs
    Call ClearNpcs
    Call ClearItems
    Call ClearShops
    Call ClearSpells
    
    Call LoadClasses
    Call LoadMaps
    Call LoadItems
    Call LoadNpcs
    Call LoadShops
    Call LoadSpells
    Call SpawnAllMapsItems
    Call SpawnAllMapNpcs
    
    If Not FileExist("System\Config\Data.ini") Then
        SpecialPutVar App.Path & "\System\Config\Data.ini", "GAME", "Name", "GameName"
        SpecialPutVar App.Path & "\System\Config\Data.ini", "GAME", "Port", 4000
    End If
        
    ' Check if the master charlist file exists for checking duplicate names, and if it doesnt make it
    If Not FileExist("accounts\charlist.txt") Then
        f = FreeFile
        Open App.Path & "\accounts\charlist.txt" For Output As #f
        Close #f
    End If
    
    Set MyScript = New clsSadScript
    'Set clsScriptCommands = New clsCommands
    MyScript.ReadInCode App.Path & "\Scripts\Main.msl", "Scripts\Main.msl", MyScript.SControl, False
    'MyScript.SControl.AddObject "ScriptHardCode", clsScriptCommands, True
    
    ' Start listening
    GameServer.StartListening
    
    Call UpdateCaption
    
    frmLoad.Visible = False
    frmServer.Show
    
    SpawnSeconds = 0
    frmServer.tmrGameAI.Enabled = True
End Sub

Sub DestroyServer()
Dim i As Long

    nid.cbSize = Len(nid)
    nid.hWnd = frmServer.hWnd
    nid.uId = vbNull
    nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    nid.uCallBackMessage = WM_MOUSEMOVE
    nid.hIcon = frmServer.Icon
    nid.szTip = "Mirage Server" & vbNullChar
    ' Add to the sys tray
    Call Shell_NotifyIcon(NIM_DELETE, nid)
 
    frmLoad.Visible = True
    frmServer.Visible = False
    
    Call SaveAllPlayersOnline
    Call ClearMaps
    Call ClearMapItems
    Call ClearMapNpcs
    Call ClearNpcs
    Call ClearItems
    Call ClearShops
    
    For i = 1 To MAX_PLAYERS
        Call GameServer.Sockets.Remove(CStr(i))
    Next
    Set GameServer = Nothing

    End
End Sub

Sub SetStatus(ByVal Status As String, ByVal Value As String)
    frmLoad.lblStatus.Caption = Status
    frmLoad.prgLoad.Value = Value
End Sub

Sub ServerLogic()
    Call CheckGiveHP
    Call GameAI
End Sub

Sub CheckSpawnMapItems()
Dim X As Long, y As Long

    ' Used for map item respawning
    SpawnSeconds = SpawnSeconds + 1
    
    ' ///////////////////////////////////////////
    ' // This is used for respawning map items //
    ' ///////////////////////////////////////////
    If SpawnSeconds >= 120 Then
        ' 2 minutes have passed
        For y = 1 To MAX_MAPS
            ' Make sure no one is on the map when it respawns
            If PlayersOnMap(y) = False Then
                ' Clear out unnecessary junk
                For X = 1 To MAX_MAP_ITEMS
                    Call ClearMapItem(X, y)
                Next X
                    
                ' Spawn the items
                Call SpawnMapItems(y)
                Call SendMapItemsToAll(y)
            End If
            DoEvents
        Next y
        
        SpawnSeconds = 0
    End If
End Sub

Sub GameAI()
Dim i As Long, X As Long, y As Long, n As Long, x1 As Long, y1 As Long, TickCount As Long
Dim Damage As Long, DistanceX As Long, DistanceY As Long, NpcNum As Long, Target As Long
Dim DidWalk As Boolean
            
    'WeatherSeconds = WeatherSeconds + 1
    'TimeSeconds = TimeSeconds + 1
    
    ' Lets change the weather if its time to
    If WeatherSeconds >= 60 Then
        i = Int(Rnd * 3)
        If i <> GameWeather Then
            GameWeather = i
            Call SendWeatherToAll
        End If
        WeatherSeconds = 0
    End If
    
    ' Check if we need to switch from day to night or night to day
    If TimeSeconds >= 60 Then
        If GameTime = TIME_DAY Then
            GameTime = TIME_NIGHT
        Else
            GameTime = TIME_DAY
        End If
        
        Call SendTimeToAll
        TimeSeconds = 0
    End If
            
    For y = 1 To MAX_MAPS
        If PlayersOnMap(y) = YES Then
            TickCount = GetTickCount
            
            ' ////////////////////////////////////
            ' // This is used for closing doors //
            ' ////////////////////////////////////
            If TickCount > TempTile(y).DoorTimer + 5000 Then
                For y1 = 0 To MAX_MAPY
                    For x1 = 0 To MAX_MAPX
                        If Map(y).Tile(x1, y1).Type = TILE_TYPE_KEY And TempTile(y).DoorOpen(x1, y1) = YES Then
                            TempTile(y).DoorOpen(x1, y1) = NO
                            Call SendDataToMap(y, "MAPKEY" & SEP_CHAR & x1 & SEP_CHAR & y1 & SEP_CHAR & 0 & SEP_CHAR & END_CHAR)
                        End If
                    Next x1
                Next y1
            End If
            
            For X = 1 To MAX_MAP_NPCS
                NpcNum = MapNpc(y, X).num
                
                ' /////////////////////////////////////////
                ' // This is used for ATTACKING ON SIGHT //
                ' /////////////////////////////////////////
                ' Make sure theres a npc with the map
                If Map(y).Npc(X) > 0 And MapNpc(y, X).num > 0 Then
                    ' If the npc is a attack on sight, search for a player on the map
                    If Npc(NpcNum).Behavior = NPC_BEHAVIOR_ATTACKONSIGHT Or Npc(NpcNum).Behavior = NPC_BEHAVIOR_GUARD Then
                        For i = 1 To MAX_PLAYERS
                            If IsPlaying(i) Then
                                If GetPlayerMap(i) = y And MapNpc(y, X).Target = 0 And GetPlayerAccess(i) <= ADMIN_MONITER Then
                                    n = Npc(NpcNum).Range
                                    
                                    DistanceX = MapNpc(y, X).X - GetPlayerX(i)
                                    DistanceY = MapNpc(y, X).y - GetPlayerY(i)
                                    
                                    ' Make sure we get a positive value
                                    If DistanceX < 0 Then DistanceX = DistanceX * -1
                                    If DistanceY < 0 Then DistanceY = DistanceY * -1
                                    
                                    ' Are they in range?  if so GET'M!
                                    If DistanceX <= n And DistanceY <= n Then
                                        If Npc(NpcNum).Behavior = NPC_BEHAVIOR_ATTACKONSIGHT Or GetPlayerPK(i) = YES Then
                                            If Trim(Npc(NpcNum).AttackSay) <> "" Then
                                                Call PlayerMsg(i, "A " & Trim(Npc(NpcNum).Name) & " says, '" & Trim(Npc(NpcNum).AttackSay) & "' to you.", SayColor)
                                            End If
                                            
                                            MapNpc(y, X).Target = i
                                        End If
                                    End If
                                End If
                            End If
                        Next i
                    End If
                End If
                                                                        
                ' /////////////////////////////////////////////
                ' // This is used for NPC walking/targetting //
                ' /////////////////////////////////////////////
                ' Make sure theres a npc with the map
                If Map(y).Npc(X) > 0 And MapNpc(y, X).num > 0 Then
                    Target = MapNpc(y, X).Target
                    
                    ' Check to see if its time for the npc to walk
                    If Npc(NpcNum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER Then
                        ' Check to see if we are following a player or not
                        If Target > 0 Then
                            ' Check if the player is even playing, if so follow'm
                            If IsPlaying(Target) And GetPlayerMap(Target) = y Then
                                DidWalk = False
                                
                                i = Int(Rnd * 5)
                                
                                ' Lets move the npc
                                Select Case i
                                    Case 0
                                        ' Up
                                        If MapNpc(y, X).y > GetPlayerY(Target) And DidWalk = False Then
                                            If CanNpcMove(y, X, DIR_UP) Then
                                                Call NpcMove(y, X, DIR_UP, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                        ' Down
                                        If MapNpc(y, X).y < GetPlayerY(Target) And DidWalk = False Then
                                            If CanNpcMove(y, X, DIR_DOWN) Then
                                                Call NpcMove(y, X, DIR_DOWN, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                        ' Left
                                        If MapNpc(y, X).X > GetPlayerX(Target) And DidWalk = False Then
                                            If CanNpcMove(y, X, DIR_LEFT) Then
                                                Call NpcMove(y, X, DIR_LEFT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                        ' Right
                                        If MapNpc(y, X).X < GetPlayerX(Target) And DidWalk = False Then
                                            If CanNpcMove(y, X, DIR_RIGHT) Then
                                                Call NpcMove(y, X, DIR_RIGHT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                    
                                    Case 1
                                        ' Right
                                        If MapNpc(y, X).X < GetPlayerX(Target) And DidWalk = False Then
                                            If CanNpcMove(y, X, DIR_RIGHT) Then
                                                Call NpcMove(y, X, DIR_RIGHT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                        ' Left
                                        If MapNpc(y, X).X > GetPlayerX(Target) And DidWalk = False Then
                                            If CanNpcMove(y, X, DIR_LEFT) Then
                                                Call NpcMove(y, X, DIR_LEFT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                        ' Down
                                        If MapNpc(y, X).y < GetPlayerY(Target) And DidWalk = False Then
                                            If CanNpcMove(y, X, DIR_DOWN) Then
                                                Call NpcMove(y, X, DIR_DOWN, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                        ' Up
                                        If MapNpc(y, X).y > GetPlayerY(Target) And DidWalk = False Then
                                            If CanNpcMove(y, X, DIR_UP) Then
                                                Call NpcMove(y, X, DIR_UP, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                        
                                    Case 2
                                        ' Down
                                        If MapNpc(y, X).y < GetPlayerY(Target) And DidWalk = False Then
                                            If CanNpcMove(y, X, DIR_DOWN) Then
                                                Call NpcMove(y, X, DIR_DOWN, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                        ' Up
                                        If MapNpc(y, X).y > GetPlayerY(Target) And DidWalk = False Then
                                            If CanNpcMove(y, X, DIR_UP) Then
                                                Call NpcMove(y, X, DIR_UP, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                        ' Right
                                        If MapNpc(y, X).X < GetPlayerX(Target) And DidWalk = False Then
                                            If CanNpcMove(y, X, DIR_RIGHT) Then
                                                Call NpcMove(y, X, DIR_RIGHT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                        ' Left
                                        If MapNpc(y, X).X > GetPlayerX(Target) And DidWalk = False Then
                                            If CanNpcMove(y, X, DIR_LEFT) Then
                                                Call NpcMove(y, X, DIR_LEFT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                    
                                    Case 3
                                        ' Left
                                        If MapNpc(y, X).X > GetPlayerX(Target) And DidWalk = False Then
                                            If CanNpcMove(y, X, DIR_LEFT) Then
                                                Call NpcMove(y, X, DIR_LEFT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                        ' Right
                                        If MapNpc(y, X).X < GetPlayerX(Target) And DidWalk = False Then
                                            If CanNpcMove(y, X, DIR_RIGHT) Then
                                                Call NpcMove(y, X, DIR_RIGHT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                        ' Up
                                        If MapNpc(y, X).y > GetPlayerY(Target) And DidWalk = False Then
                                            If CanNpcMove(y, X, DIR_UP) Then
                                                Call NpcMove(y, X, DIR_UP, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                        ' Down
                                        If MapNpc(y, X).y < GetPlayerY(Target) And DidWalk = False Then
                                            If CanNpcMove(y, X, DIR_DOWN) Then
                                                Call NpcMove(y, X, DIR_DOWN, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                End Select
                                
                                
                            
                                ' Check if we can't move and if player is behind something and if we can just switch dirs
                                If Not DidWalk Then
                                    If MapNpc(y, X).X - 1 = GetPlayerX(Target) And MapNpc(y, X).y = GetPlayerY(Target) Then
                                        If MapNpc(y, X).Dir <> DIR_LEFT Then
                                            Call NpcDir(y, X, DIR_LEFT)
                                        End If
                                        DidWalk = True
                                    End If
                                    If MapNpc(y, X).X + 1 = GetPlayerX(Target) And MapNpc(y, X).y = GetPlayerY(Target) Then
                                        If MapNpc(y, X).Dir <> DIR_RIGHT Then
                                            Call NpcDir(y, X, DIR_RIGHT)
                                        End If
                                        DidWalk = True
                                    End If
                                    If MapNpc(y, X).X = GetPlayerX(Target) And MapNpc(y, X).y - 1 = GetPlayerY(Target) Then
                                        If MapNpc(y, X).Dir <> DIR_UP Then
                                            Call NpcDir(y, X, DIR_UP)
                                        End If
                                        DidWalk = True
                                    End If
                                    If MapNpc(y, X).X = GetPlayerX(Target) And MapNpc(y, X).y + 1 = GetPlayerY(Target) Then
                                        If MapNpc(y, X).Dir <> DIR_DOWN Then
                                            Call NpcDir(y, X, DIR_DOWN)
                                        End If
                                        DidWalk = True
                                    End If
                                    
                                    ' We could not move so player must be behind something, walk randomly.
                                    If Not DidWalk Then
                                        i = Int(Rnd * 2)
                                        If i = 1 Then
                                            i = Int(Rnd * 4)
                                            If CanNpcMove(y, X, i) Then
                                                Call NpcMove(y, X, i, MOVING_WALKING)
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                                MapNpc(y, X).Target = 0
                            End If
                        Else
                            i = Int(Rnd * 4)
                            If i = 1 Then
                                i = Int(Rnd * 4)
                                If CanNpcMove(y, X, i) Then
                                    Call NpcMove(y, X, i, MOVING_WALKING)
                                End If
                            End If
                        End If
                    End If
                End If
                
                ' /////////////////////////////////////////////
                ' // This is used for npcs to attack players //
                ' /////////////////////////////////////////////
                ' Make sure theres a npc with the map
                If Map(y).Npc(X) > 0 And MapNpc(y, X).num > 0 Then
                    Target = MapNpc(y, X).Target
                    
                    ' Check if the npc can attack the targeted player player
                    If Target > 0 Then
                        ' Is the target playing and on the same map?
                        If IsPlaying(Target) And GetPlayerMap(Target) = y Then
                            ' Can the npc attack the player?
                            If CanNpcAttackPlayer(X, Target) Then
                                If Not CanPlayerBlockHit(Target) Then
                                    Damage = Npc(NpcNum).STR - GetPlayerProtection(Target)
                                    If Damage > 0 Then
                                        Call NpcAttackPlayer(X, Target, Damage)
                                    Else
                                        Call PlayerMsg(Target, "The " & Trim(Npc(NpcNum).Name) & "'s hit didn't even phase you!", BrightBlue)
                                    End If
                                Else
                                    Call PlayerMsg(Target, "Your " & Trim(Item(GetPlayerInvItemNum(Target, GetPlayerShieldSlot(Target))).Name) & " blocks the " & Trim(Npc(NpcNum).Name) & "'s hit!", BrightCyan)
                                End If
                            End If
                        Else
                            ' Player left map or game, set target to 0
                            MapNpc(y, X).Target = 0
                        End If
                    End If
                End If
                
                ' ////////////////////////////////////////////
                ' // This is used for regenerating NPC's HP //
                ' ////////////////////////////////////////////
                ' Check to see if we want to regen some of the npc's hp
                If MapNpc(y, X).num > 0 And TickCount > GiveNPCHPTimer + 10000 Then
                    If MapNpc(y, X).HP > 0 Then
                        MapNpc(y, X).HP = MapNpc(y, X).HP + GetNpcHPRegen(NpcNum)
                    
                        ' Check if they have more then they should and if so just set it to max
                        If MapNpc(y, X).HP > GetNpcMaxHP(NpcNum) Then
                            MapNpc(y, X).HP = GetNpcMaxHP(NpcNum)
                        End If
                    End If
                End If
                    
                ' ////////////////////////////////////////////////////////
                ' // This is used for checking if an NPC is dead or not //
                ' ////////////////////////////////////////////////////////
                ' Check if the npc is dead or not
                'If MapNpc(y, x).Num > 0 Then
                '    If MapNpc(y, x).HP <= 0 And Npc(MapNpc(y, x).Num).STR > 0 And Npc(MapNpc(y, x).Num).DEF > 0 Then
                '        MapNpc(y, x).Num = 0
                '        MapNpc(y, x).SpawnWait = TickCount
                '   End If
                'End If
                
                ' //////////////////////////////////////
                ' // This is used for spawning an NPC //
                ' //////////////////////////////////////
                ' Check if we are supposed to spawn an npc or not
                If MapNpc(y, X).num = 0 And Map(y).Npc(X) > 0 Then
                    If TickCount > MapNpc(y, X).SpawnWait + (Npc(Map(y).Npc(X)).SpawnSecs * 1000) Then
                        Call SpawnNpc(X, y)
                    End If
                End If
            Next X
        End If
        DoEvents
    Next y
    
    ' Make sure we reset the timer for npc hp regeneration
    If GetTickCount > GiveNPCHPTimer + 10000 Then
        GiveNPCHPTimer = GetTickCount
    End If

    ' Make sure we reset the timer for door closing
    If GetTickCount > KeyTimer + 15000 Then
        KeyTimer = GetTickCount
    End If
End Sub

Sub CheckGiveHP()
Dim i As Long, n As Long

    If GetTickCount > GiveHPTimer + 10000 Then
        For i = 1 To MAX_PLAYERS
            If IsPlaying(i) Then
                Call SetPlayerHP(i, GetPlayerHP(i) + GetPlayerHPRegen(i))
                Call SendHP(i)
                Call SetPlayerMP(i, GetPlayerMP(i) + GetPlayerMPRegen(i))
                Call SendMP(i)
                Call SetPlayerSP(i, GetPlayerSP(i) + GetPlayerSPRegen(i))
                Call SendSP(i)
            End If
            DoEvents
        Next i
        
        GiveHPTimer = GetTickCount
    End If
End Sub

Sub PlayerSaveTimer()
Static MinPassed As Long
Dim i As Long

    MinPassed = MinPassed + 1
    If MinPassed >= 10 Then
        If TotalOnlinePlayers > 0 Then
            Call TextAdd(frmServer.txtText, "Saving all online players...", True)
            Call GlobalMsg("Saving all online players...", Pink)
            For i = 1 To MAX_PLAYERS
                If IsPlaying(i) Then
                    Call SavePlayer(i)
                End If
                DoEvents
            Next i
        End If
        
        MinPassed = 0
    End If
End Sub

Function GetIP() As String
On Error Resume Next
Dim IP As String

IP = frmServer.Inet1.OpenURL("http://www.whatismyip.org/")

    If IsNumeric(Left$(IP, 2)) = False Then
        IP = Split(frmServer.Inet1.OpenURL("http://whatismyip.com/"), "<h1>")(1)
        IP = Split(IP, " ")(3)
        IP = Trim$(Split(IP, "</h1>")(0))

        If Not IP <> vbNullString Then IP = "Localhost(127.0.0.1)"
    End If

    GetIP = Trim$(IP)
    IP = ""

End Function
