Attribute VB_Name = "modGeneral"
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

Sub InitServer()
Dim IPMask As String
Dim i As Long
Dim f As Long
    
    Randomize Timer
    
    ' Check for server configuration file
    If Not FileExist("Data\Data.ini") Then
        frmConfig.Show vbModal
    End If
    
    ' Check if server needs IP configuration
    If GetVar(App.Path & "\Data\Data.ini", "CONFIG", "Always") = NO Then
        frmIPConfig.Show vbModal
    End If
    
    '' Init atmosphere
    'GameWeather = WEATHER_NONE
    'WeatherSeconds = 0
    'GameTime = TIME_DAY
    'TimeSeconds = 0
    
    ' Check if the maps directory is there, if its not make it
    If LCase(Dir(App.Path & "\maps", vbDirectory)) <> "maps" Then
        Call MkDir(App.Path & "\maps")
    End If
    
    ' Check if the accounts directory is there, if its not make it
    If LCase(Dir(App.Path & "\accounts", vbDirectory)) <> "accounts" Then
        Call MkDir(App.Path & "\accounts")
    End If
    
    ' Check if the logs directory is there, if its not make it
    If LCase(Dir(App.Path & "\logs", vbDirectory)) <> "logs" Then
        Call MkDir(App.Path & "\logs")
    End If
    
    SEP_CHAR = Chr(0)
    END_CHAR = Chr(237)
    
    ServerLog = True
    
    Call SetStatus("Loading settings...")
    
    GAME_NAME = Trim(GetVar(App.Path & "\Data\Data.ini", "CONFIG", "GameName"))
    GAME_WEBSITE = Trim(GetVar(App.Path & "\Data\Data.ini", "CONFIG", "WebSite"))
    GAME_IP = Trim(GetVar(App.Path & "\Data\Data.ini", "CONFIG", "IP"))
    GAME_PORT = Trim(GetVar(App.Path & "\Data\Data.ini", "CONFIG", "Port"))
    MAX_PLAYERS = GetVar(App.Path & "\Data\Data.ini", "MAX", "MAX_PLAYERS")
    MAX_ITEMS = GetVar(App.Path & "\Data\Data.ini", "MAX", "MAX_ITEMS")
    MAX_NPCS = GetVar(App.Path & "\Data\Data.ini", "MAX", "MAX_NPCS")
    MAX_SHOPS = GetVar(App.Path & "\Data\Data.ini", "MAX", "MAX_SHOPS")
    MAX_SPELLS = GetVar(App.Path & "\Data\Data.ini", "MAX", "MAX_SPELLS")
    MAX_SKILLS = GetVar(App.Path & "\Data\Data.ini", "MAX", "MAX_SKILLS")
    MAX_MAPS = GetVar(App.Path & "\Data\Data.ini", "MAX", "MAX_MAPS")
    MAX_MAP_ITEMS = GetVar(App.Path & "\Data\Data.ini", "MAX", "MAX_MAP_ITEMS")
    MAX_GUILDS = GetVar(App.Path & "\Data\Data.ini", "MAX", "MAX_GUILDS")
    MAX_GUILD_MEMBERS = GetVar(App.Path & "\Data\Data.ini", "MAX", "MAX_GUILD_MEMBERS")
    'MAX_EMOTICONS = GetVar(App.Path & "\Data\Data.ini", "MAX", "MAX_EMOTICONS")
    MAX_QUESTS = GetVar(App.Path & "\Data\Data.ini", "MAX", "MAX_QUESTS")
    'MAX_LEVEL = GetVar(App.Path & "\Data\Data.ini", "MAX", "MAX_LEVEL")
    'Scripting = GetVar(App.Path & "\Data\Data.ini", "CONFIG", "Scripting")

    'MAX_MAPX = 30
    'MAX_MAPY = 30
    'If GetVar(App.Path & "\Data\Data.ini", "CONFIG", "Scrolling") = 0 Then
        'MAX_MAPX = 19
        'MAX_MAPY = 14
    'ElseIf GetVar(App.Path & "\Data\Data.ini", "CONFIG", "Scrolling") = 1 Then
        'MAX_MAPX = 30
        'MAX_MAPY = 30
    'End If
        
    ReDim Map(1 To MAX_MAPS) As MapRec
    ReDim TempTile(1 To MAX_MAPS) As TempTileRec
    ReDim PushTile(1 To MAX_MAPS) As PushTileRec
    ReDim PlayersOnMap(1 To MAX_MAPS) As Long
    ReDim Player(1 To MAX_PLAYERS) As AccountRec
    ReDim Item(0 To MAX_ITEMS) As ItemRec
    ReDim Npc(0 To MAX_NPCS) As NpcRec
    ReDim MapItem(1 To MAX_MAPS, 1 To MAX_MAP_ITEMS) As MapItemRec
    ReDim MapNpc(1 To MAX_MAPS, 1 To MAX_MAP_NPCS) As MapNpcRec
    ReDim MapResource(1 To MAX_MAPS, 1 To MAX_MAP_RESOURCES) As MapResourceRec
    ReDim Shop(1 To MAX_SHOPS) As ShopRec
    ReDim Spell(1 To MAX_SPELLS) As SpellRec
    ReDim Skill(1 To MAX_SKILLS) As SkillRec
    ReDim Guild(1 To MAX_GUILDS) As GuildRec
    'ReDim Emoticons(0 To MAX_EMOTICONS) As EmoRec
    ReDim Quest(1 To MAX_QUESTS) As QuestRec
    For i = 1 To MAX_GUILDS
        ReDim Guild(i).Member(1 To MAX_GUILD_MEMBERS) As String * NAME_LENGTH
    Next i
    For i = 1 To MAX_MAPS
        'ReDim Map(i).Tile(0 To MAX_MAPX, 0 To MAX_MAPY) As TileRec
        ReDim TempTile(i).DoorOpen(0 To MAX_MAPX, 0 To MAX_MAPY) As Byte
        ReDim PushTile(i).Pushed(0 To MAX_MAPX, 0 To MAX_MAPY) As Byte
    Next i
    'ReDim Experience(1 To MAX_LEVEL) As Long
    
    ' Get the listening socket ready To go
    Set GameServer = New clsServer
        
    ' Init all the player sockets
    For i = 1 To MAX_PLAYERS
        Call SetStatus("Initializing player array...")
        Call ClearPlayer(i)
        
        Call GameServer.Sockets.Add(CStr(i))
    Next i
    
    Call SetStatus("Clearing GUI's...")
    Call ClearGUIS
    Call SetStatus("Clearing temp tile fields...")
    Call ClearTempTile
    'Call SetStatus("Clearing maps...")
    'Call ClearMaps
    Call SetStatus("Clearing map items...")
    Call ClearMapItems
    Call SetStatus("Clearing map npcs...")
    Call ClearMapNpcs
    Call SetStatus("Clearing map resources...")
    Call ClearMapResources
    Call SetStatus("Clearing npcs...")
    Call ClearNpcs
    Call SetStatus("Clearing items...")
    Call ClearItems
    Call SetStatus("Clearing shops...")
    Call ClearShops
    Call SetStatus("Clearing skills...")
    Call ClearSkills
    Call SetStatus("Clearing spells...")
    Call ClearSpells
    Call SetStatus("Clearing spells...")
    Call ClearQuests
    Call SetStatus("Loading GUI's...")
    Call LoadGUIS
    Call SetStatus("Loading quests...")
    Call LoadClasses
    Call SetStatus("Loading maps...")
    Call LoadMaps
    Call SetStatus("Loading items...")
    Call LoadItems
    Call SetStatus("Loading npcs...")
    Call LoadNpcs
    Call SetStatus("Loading shops...")
    Call LoadShops
    Call SetStatus("Loading skills...")
    Call LoadSkills
    Call SetStatus("Loading spells...")
    Call LoadSpells
    Call SetStatus("Loading quests...")
    Call LoadQuests
    Call SetStatus("Spawning map items...")
    Call SpawnAllMapsItems
    Call SetStatus("Spawning map npcs...")
    Call SpawnAllMapNpcs
    Call SetStatus("Spawning map resources...")
    Call SpawnAllMapResources
        
    ' Check if the master charlist file exists for checking duplicate names, and if it doesnt make it
    If Not FileExist("accounts\charlist.txt") Then
        f = FreeFile
        Open App.Path & "\accounts\charlist.txt" For Output As #f
        Close #f
    End If
    
    ' Check if the packetlog file exists, and if it doesn't make it
    If Not FileExist("packet.log") Then
        f = FreeFile
        Open App.Path & "\packet.log" For Output As #f
        Close #f
    End If
    
    ' Start listening
    GameServer.StartListening
    
    Call UpdateCaption
    
    frmLoad.Visible = False
    frmCServer.Show
    
    SpawnSeconds = 0
    frmCServer.tmrGameAI.Enabled = True
End Sub

Sub SetStatus(ByVal Status As String)
    frmLoad.lblStatus.Caption = Status
End Sub

Sub DestroyServer()
Dim i As Long
 
    frmLoad.Visible = True
    frmCServer.Visible = False
    
    Call SetStatus("Saving players online...")
    Call SaveAllPlayersOnline
    Call SetStatus("Clearing maps...")
    Call ClearMaps
    Call SetStatus("Clearing map items...")
    Call ClearMapItems
    Call SetStatus("Clearing map npcs...")
    Call ClearMapNpcs
    Call SetStatus("Clearing map resources...")
    Call ClearMapResources
    Call SetStatus("Clearing npcs...")
    Call ClearNpcs
    Call SetStatus("Clearing items...")
    Call ClearItems
    Call SetStatus("Clearings skills...")
    Call ClearSkills
    Call SetStatus("Clearings Spells...")
    Call ClearSpells
    Call SetStatus("Clearings Quests...")
    Call ClearQuests
    Call SetStatus("Clearing shops...")
    Call ClearShops
    For i = 1 To MAX_PLAYERS
        Call SetStatus("Unloading sockets And timers... " & i & "/" & MAX_PLAYERS)
        DoEvents

        Call GameServer.Sockets.Remove(CStr(i))
    Next i
    Set GameServer = Nothing

    End
End Sub

Sub ServerLogic()
    'Call CheckGiveHP
    Call GameAI
End Sub

Sub CheckSpawnMapItems()
Dim x As Long, y As Long

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
                For x = 1 To MAX_MAP_ITEMS
                    Call ClearMapItem(x, y)
                Next x
                    
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
Dim i As Long, x As Long, y As Long, n As Long, x1 As Long, y1 As Long, TickCount As Long
Dim Damage As Long, DistanceX As Long, DistanceY As Long, NpcNum As Long, Target As Long
Dim DidWalk As Boolean
            
    ''WeatherSeconds = WeatherSeconds + 1
    ''TimeSeconds = TimeSeconds + 1
    
    '' Lets change the weather if its time to
    'If WeatherSeconds >= 60 Then
        'i = Int(Rnd * 3)
        'If i <> GameWeather Then
            'GameWeather = i
            'Call SendWeatherToAll
        'End If
        'WeatherSeconds = 0
    'End If
    
    '' Check if we need to switch from day to night or night to day
    'If TimeSeconds >= 60 Then
        'If GameTime = TIME_DAY Then
            'GameTime = TIME_NIGHT
        'Else
            'GameTime = TIME_DAY
        'End If
        
        'Call SendTimeToAll
        'TimeSeconds = 0
    'End If
            
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
            
            ' /////////////////////////////////////////
            ' // This is used for closing pushblocks //
            ' /////////////////////////////////////////
            If TickCount > PushTile(y).PushedTimer + 5000 Then
                For y1 = 0 To MAX_MAPY
                    For x1 = 0 To MAX_MAPX
                        If Map(y).Tile(x1, y1).Type = TILE_TYPE_PUSHBLOCK And PushTile(y).Pushed(x1, y1) = YES Then
                            If PushBlockBlocked(y, x1, y1) = False Then
                                PushTile(y).Pushed(x1, y1) = NO
                                Call SendDataToMap(y, "PUSHBLOCK" & SEP_CHAR & x1 & SEP_CHAR & y1 & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & 1 & SEP_CHAR & END_CHAR)
                            Else
                                PushTile(y).PushedTimer = GetTickCount
                            End If
                        End If
                    Next x1
                Next y1
            End If
            
            For x = 1 To MAX_MAP_NPCS
                NpcNum = MapNpc(y, x).Num
                
                ' /////////////////////////////////////////
                ' // This is used for ATTACKING ON SIGHT //
                ' /////////////////////////////////////////
                ' Make sure theres a npc with the map
                If Map(y).Npc(x) > 0 And MapNpc(y, x).Num > 0 Then
                    ' If the npc is a attack on sight, search for a player on the map
                    If Npc(NpcNum).Behavior = NPC_BEHAVIOR_ATTACKONSIGHT Or Npc(NpcNum).Behavior = NPC_BEHAVIOR_GUARD Then
                        For i = 1 To HighIndex
                            If IsPlaying(i) Then
                                ' temp remove admin block
                                If GetPlayerMap(i) = y And MapNpc(y, x).Target = 0 Then 'And GetPlayerAccess(i) <= ADMIN_MONITER Then
                                    n = Npc(NpcNum).Range
                                    
                                    DistanceX = MapNpc(y, x).x - GetPlayerX(i)
                                    DistanceY = MapNpc(y, x).y - GetPlayerY(i)
                                    
                                    ' Make sure we get a positive value
                                    If DistanceX < 0 Then DistanceX = DistanceX * -1
                                    If DistanceY < 0 Then DistanceY = DistanceY * -1
                                    
                                    ' Are they in range?  if so GET'M!
                                    If DistanceX <= n And DistanceY <= n Then
                                        If Npc(NpcNum).Behavior = NPC_BEHAVIOR_ATTACKONSIGHT Or GetPlayerPK(i) = YES Then
                                            
                                            MapNpc(y, x).Target = i
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
                If Map(y).Npc(x) > 0 And MapNpc(y, x).Num > 0 Then
                    Target = MapNpc(y, x).Target
                    
                    ' Check to see if its time for the npc to walk
                    If Npc(NpcNum).Behavior <> NPC_BEHAVIOR_RESOURCE Then
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
                                        If MapNpc(y, x).y > GetPlayerY(Target) And DidWalk = False Then
                                            If CanNpcMove(y, x, DIR_UP) Then
                                                Call NpcMove(y, x, DIR_UP, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                        ' Down
                                        If MapNpc(y, x).y < GetPlayerY(Target) And DidWalk = False Then
                                            If CanNpcMove(y, x, DIR_DOWN) Then
                                                Call NpcMove(y, x, DIR_DOWN, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                        ' Left
                                        If MapNpc(y, x).x > GetPlayerX(Target) And DidWalk = False Then
                                            If CanNpcMove(y, x, DIR_LEFT) Then
                                                Call NpcMove(y, x, DIR_LEFT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                        ' Right
                                        If MapNpc(y, x).x < GetPlayerX(Target) And DidWalk = False Then
                                            If CanNpcMove(y, x, DIR_RIGHT) Then
                                                Call NpcMove(y, x, DIR_RIGHT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                    
                                    Case 1
                                        ' Right
                                        If MapNpc(y, x).x < GetPlayerX(Target) And DidWalk = False Then
                                            If CanNpcMove(y, x, DIR_RIGHT) Then
                                                Call NpcMove(y, x, DIR_RIGHT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                        ' Left
                                        If MapNpc(y, x).x > GetPlayerX(Target) And DidWalk = False Then
                                            If CanNpcMove(y, x, DIR_LEFT) Then
                                                Call NpcMove(y, x, DIR_LEFT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                        ' Down
                                        If MapNpc(y, x).y < GetPlayerY(Target) And DidWalk = False Then
                                            If CanNpcMove(y, x, DIR_DOWN) Then
                                                Call NpcMove(y, x, DIR_DOWN, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                        ' Up
                                        If MapNpc(y, x).y > GetPlayerY(Target) And DidWalk = False Then
                                            If CanNpcMove(y, x, DIR_UP) Then
                                                Call NpcMove(y, x, DIR_UP, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                        
                                    Case 2
                                        ' Down
                                        If MapNpc(y, x).y < GetPlayerY(Target) And DidWalk = False Then
                                            If CanNpcMove(y, x, DIR_DOWN) Then
                                                Call NpcMove(y, x, DIR_DOWN, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                        ' Up
                                        If MapNpc(y, x).y > GetPlayerY(Target) And DidWalk = False Then
                                            If CanNpcMove(y, x, DIR_UP) Then
                                                Call NpcMove(y, x, DIR_UP, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                        ' Right
                                        If MapNpc(y, x).x < GetPlayerX(Target) And DidWalk = False Then
                                            If CanNpcMove(y, x, DIR_RIGHT) Then
                                                Call NpcMove(y, x, DIR_RIGHT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                        ' Left
                                        If MapNpc(y, x).x > GetPlayerX(Target) And DidWalk = False Then
                                            If CanNpcMove(y, x, DIR_LEFT) Then
                                                Call NpcMove(y, x, DIR_LEFT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                    
                                    Case 3
                                        ' Left
                                        If MapNpc(y, x).x > GetPlayerX(Target) And DidWalk = False Then
                                            If CanNpcMove(y, x, DIR_LEFT) Then
                                                Call NpcMove(y, x, DIR_LEFT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                        ' Right
                                        If MapNpc(y, x).x < GetPlayerX(Target) And DidWalk = False Then
                                            If CanNpcMove(y, x, DIR_RIGHT) Then
                                                Call NpcMove(y, x, DIR_RIGHT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                        ' Up
                                        If MapNpc(y, x).y > GetPlayerY(Target) And DidWalk = False Then
                                            If CanNpcMove(y, x, DIR_UP) Then
                                                Call NpcMove(y, x, DIR_UP, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                        ' Down
                                        If MapNpc(y, x).y < GetPlayerY(Target) And DidWalk = False Then
                                            If CanNpcMove(y, x, DIR_DOWN) Then
                                                Call NpcMove(y, x, DIR_DOWN, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                End Select
                                
                                
                            
                                ' Check if we can't move and if player is behind something and if we can just switch dirs
                                If Not DidWalk Then
                                    If MapNpc(y, x).x - 1 = GetPlayerX(Target) And MapNpc(y, x).y = GetPlayerY(Target) Then
                                        If MapNpc(y, x).Dir <> DIR_LEFT Then
                                            Call NpcDir(y, x, DIR_LEFT)
                                        End If
                                        DidWalk = True
                                    End If
                                    If MapNpc(y, x).x + 1 = GetPlayerX(Target) And MapNpc(y, x).y = GetPlayerY(Target) Then
                                        If MapNpc(y, x).Dir <> DIR_RIGHT Then
                                            Call NpcDir(y, x, DIR_RIGHT)
                                        End If
                                        DidWalk = True
                                    End If
                                    If MapNpc(y, x).x = GetPlayerX(Target) And MapNpc(y, x).y - 1 = GetPlayerY(Target) Then
                                        If MapNpc(y, x).Dir <> DIR_UP Then
                                            Call NpcDir(y, x, DIR_UP)
                                        End If
                                        DidWalk = True
                                    End If
                                    If MapNpc(y, x).x = GetPlayerX(Target) And MapNpc(y, x).y + 1 = GetPlayerY(Target) Then
                                        If MapNpc(y, x).Dir <> DIR_DOWN Then
                                            Call NpcDir(y, x, DIR_DOWN)
                                        End If
                                        DidWalk = True
                                    End If
                                    
                                    ' We could not move so player must be behind something, walk randomly.
                                    If Not DidWalk Then
                                        i = Int(Rnd * 2)
                                        If i = 1 Then
                                            i = Int(Rnd * 4)
                                            If CanNpcMove(y, x, i) Then
                                                Call NpcMove(y, x, i, MOVING_WALKING)
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                                MapNpc(y, x).Target = 0
                            End If
                        Else
                            i = Int(Rnd * 3)
                            If i = 1 Then
                                i = Int(Rnd * 4)
                                If CanNpcMove(y, x, i) Then
                                    Call NpcMove(y, x, i, MOVING_WALKING)
                                End If
                            End If
                        End If
                    End If
                End If
                
                ' /////////////////////////////////////////////
                ' // This is used for npcs to attack players //
                ' /////////////////////////////////////////////
                ' Make sure theres a npc with the map
                If Map(y).Npc(x) > 0 And MapNpc(y, x).Num > 0 Then
                    Target = MapNpc(y, x).Target
                    
                    ' Check if the npc can attack the targeted player player
                    If Target > 0 Then
                        ' Is the target playing and on the same map?
                        If IsPlaying(Target) And GetPlayerMap(Target) = y Then
                            ' Can the npc attack the player?
                            If CanNpcAttackPlayer(x, Target) Then
                                If Not CanPlayerBlockHit(Target) Then
                                    Damage = Npc(NpcNum).STR - GetPlayerProtection(Target)
                                    If Damage > 0 Then
                                        Call NpcAttackPlayer(x, Target, Damage)
                                    Else
                                        Call SendDataTo(Target, "BLITPLAYERMSG" & SEP_CHAR & "  Miss" & SEP_CHAR & Magenta & SEP_CHAR & END_CHAR)
                                    End If
                                Else
                                    Call SendDataTo(Target, "BLITPLAYERMSG" & SEP_CHAR & "  Block" & SEP_CHAR & Yellow & SEP_CHAR & END_CHAR)
                                End If
                            End If
                        Else
                            ' Player left map or game, set target to 0
                            MapNpc(y, x).Target = 0
                        End If
                    End If
                End If
                
                '' ////////////////////////////////////////////
                '' // This is used for regenerating NPC's HP //
                '' ////////////////////////////////////////////
                '' Check to see if we want to regen some of the npc's hp
                'If MapNpc(y, x).Num > 0 And TickCount > GiveNPCHPTimer + 10000 Then
                    'If MapNpc(y, x).HP > 0 Then
                        'MapNpc(y, x).HP = MapNpc(y, x).HP + GetNpcHPRegen(NpcNum)
                    
                        '' Check if they have more then they should and if so just set it to max
                        'If MapNpc(y, x).HP > GetNpcMaxHP(NpcNum) Then
                            'MapNpc(y, x).HP = GetNpcMaxHP(NpcNum)
                        'End If
                    'End If
                'End If
                    
                '' ////////////////////////////////////////////////////////
                '' // This is used for checking if an NPC is dead or not //
                '' ////////////////////////////////////////////////////////
                '' Check if the npc is dead or not
                ''If MapNpc(y, x).Num > 0 Then
                ''    If MapNpc(y, x).HP <= 0 And Npc(MapNpc(y, x).Num).STR > 0 And Npc(MapNpc(y, x).Num).DEF > 0 Then
                ''        MapNpc(y, x).Num = 0
                ''        MapNpc(y, x).SpawnWait = TickCount
                ''   End If
                ''End If
                
                ' //////////////////////////////////////
                ' // This is used for spawning an NPC //
                ' //////////////////////////////////////
                ' Check if we are supposed to spawn an npc or not
                If MapNpc(y, x).Num = 0 And Map(y).Npc(x) > 0 Then
                    If TickCount > MapNpc(y, x).SpawnWait + (Npc(Map(y).Npc(x)).SpawnSecs * 1000) Then
                        Call SpawnNpc(x, y)
                    End If
                End If
            Next x
            ' //////////////////////////////////////////
            ' // This is used for spawning a Resource //
            ' //////////////////////////////////////////
            For x = 1 To MAX_MAP_RESOURCES
                ' Check if we are supposed to spawn a resource or not
                If MapResource(y, x).Num = 0 And Map(y).Resource(x) > 0 Then
                    If TickCount > MapResource(y, x).SpawnWait + (Npc(Map(y).Resource(x)).SpawnSecs * 1000) Then
                        Call SpawnResource(x, y)
                    End If
                End If
            Next x
        End If
        DoEvents
    Next y
    
    '' Make sure we reset the timer for npc hp regeneration
    'If GetTickCount > GiveNPCHPTimer + 10000 Then
        'GiveNPCHPTimer = GetTickCount
    'End If

    ' Make sure we reset the timer for door closing
    If GetTickCount > KeyTimer + 15000 Then
        KeyTimer = GetTickCount
    End If
End Sub

