Attribute VB_Name = "modGeneral"
Option Explicit

Sub InitServer()
'On Error GoTo errorhandler:
Dim IPMask As String
Dim i As Long
Dim f As Long
    
    Randomize Timer
    'Set Game Name
    'G_NAME = GetVar(App.Path & "\Data\data.ini", "Info", "Name")
    
    f = FreeFile
    
    ' Init atmosphere
    GameWeather = WEATHER_NONE
    WeatherSeconds = 0
    GameTime = TIME_DAY
    TimeSeconds = 0
    
    'Set server time variables
    Server_Second = 1
    Server_Minute = 1
    Server_Hour = 1
    
    ' Check if the accounts directory is there, if its not make it
    If LCase(Dir(App.Path & "\accounts", vbDirectory)) <> "accounts" Then
        Call MkDir(App.Path & "\accounts")
    End If
    
    ' Check if the data directory is there, if its not make it
    If LCase$(Dir(App.Path & "\data", vbDirectory)) <> "data" Then
        Call MkDir(App.Path & "\Data")
    Open App.Path & "\Data\classes.ini" For Output As #f
        Print #f, "; Distribute 20 points"
        Print #f, "; Make sure that STR DEF SPEED are all at least 1, MAGI can be 0"
        Print #f, "; Keep in mind that the number of Max Classes and Max Visible classes starts with 0 and goes to the specified"
        Print #f, "; The Map, X, and Y choices determine where the class starting location will be"
        Print #f, ""
        Print #f, "[INIT]"
        Print #f, "MaxClasses = 0"
        Print #f, "MaxVisibleClasses = 0"
        Print #f, ""
        Print #f, "[CLASS0]"
        Print #f, "Name = Knight"
        Print #f, "Sprite = 0"
        Print #f, "HP = 20"
        Print #f, "MP = 0"
        Print #f, "SP = 10"
        Print #f, "STR = 8"
        Print #f, "DEF = 7"
        Print #f, "SPEED = 3"
        Print #f, "MAGI = 0"
        Print #f, "Map = 1"
        Print #f, "x = 3"
        Print #f, "y = 4"
    Close #f
    
    Open App.Path & "\Data\data.ini" For Output As #f
        Print #f, "[Address]"
        Print #f, "Port = 2000"
        Print #f, "[Info]"
        Print #f, "Name = ""Game Name"""
        Print #f, "MOTD = ""Welcome to the Dual Solace Engine, RealFeel!"""
        Print #f, "Msg = ""This is a test of the server message system!"""
        Print #f, "[Settings]"
        Print #f, "MAX_PLAYERS = 70"
        Print #f, "MAX_MAPS = 200"
        Print #f, "MAX_ITEMS = 255"
        Print #f, "MAX_NPCS = 255"
        Print #f, "MAX_SHOPS = 255"
        Print #f, "MAX_SPELLS = 255"
        Print #f, "MAX_GUILDS = 20"
        Print #f, "MAX_EXPERIENCE = 100"
    Close #f
    End If
    
    ' Check if the DLLs directory is there, if its not make it
    If LCase$(Dir(App.Path & "\dlls", vbDirectory)) <> "dlls" Then
        Call MkDir(App.Path & "\DLLs")
    End If
    
    ' Check if the log directory is there, if its not make it
    If LCase$(Dir(App.Path & "\Logs", vbDirectory)) <> "logs" Then
        Call MkDir(App.Path & "\Logs")
    End If
    
    ' Check if the items directory is there, if its not make it
    If LCase$(Dir(App.Path & "\items", vbDirectory)) <> "items" Then
        Call MkDir(App.Path & "\items")
    End If
    
    ' Check if the maps directory is there, if its not make it
    If LCase$(Dir(App.Path & "\maps", vbDirectory)) <> "maps" Then
        Call MkDir(App.Path & "\maps")
    End If
    
    ' Check if the npcs directory is there, if its not make it
    If LCase$(Dir(App.Path & "\npcs", vbDirectory)) <> "npcs" Then
        Call MkDir(App.Path & "\npcs")
    End If
    
    ' Check if the scripts directory is there, if its not make it
    If LCase$(Dir(App.Path & "\scripts", vbDirectory)) <> "scripts" Then
        Call MkDir(App.Path & "\scripts")
    End If
    
    ' Check if the shops directory is there, if its not make it
    If LCase$(Dir(App.Path & "\shops", vbDirectory)) <> "shops" Then
        Call MkDir(App.Path & "\shops")
    End If
    
    ' Check if the spells directory is there, if its not make it
    If LCase$(Dir(App.Path & "\spells", vbDirectory)) <> "spells" Then
        Call MkDir(App.Path & "\spells")
    End If
    
    
    SEP_CHAR = Chr(0)
    END_CHAR = Chr(237)
    
    ServerLog = False
    
    'set the server port
    If GetVar((App.Path & "\data\data.ini"), "Address", "Port") = "" Then
        Game_Port = 7234
    Else
        Game_Port = GetVar(App.Path & "\data\data.ini", "Address", "Port")
    End If
    
    'set the maximum players
    If GetVar((App.Path & "\data\data.ini"), "Info", "Name") = "" Then
        GAME_NAME = "Dual Solace's Real-Feel Engine"
    Else
        GAME_NAME = GetVar(App.Path & "\data\data.ini", "Info", "Name")
    End If
    
    'set the maximum players
    If GetVar((App.Path & "\data\data.ini"), "Settings", "MAX_PLAYERS") = "" Then
        MAX_PLAYERS = 70
    Else
        MAX_PLAYERS = CLng(GetVar(App.Path & "\data\data.ini", "Settings", "MAX_PLAYERS"))
    End If
    
    'set the maximum players
    If GetVar((App.Path & "\data\data.ini"), "Settings", "MAX_MAPS") = "" Then
        MAX_MAPS = 1000
    Else
        MAX_MAPS = CLng(GetVar(App.Path & "\data\data.ini", "Settings", "MAX_MAPS"))
    End If
    
    'set the maximum items
    If GetVar((App.Path & "\data\data.ini"), "Settings", "MAX_ITEMS") = "" Then
        MAX_ITEMS = 255
    Else
        MAX_ITEMS = CLng(GetVar(App.Path & "\data\data.ini", "Settings", "MAX_ITEMS"))
    End If
    
    'set the maximum npcs
    If GetVar((App.Path & "\data\data.ini"), "Settings", "MAX_NPCS") = "" Then
        MAX_NPCS = 255
    Else
        MAX_NPCS = CLng(GetVar(App.Path & "\data\data.ini", "Settings", "MAX_NPCS"))
    End If
    
    'set the maximum shops
    If GetVar((App.Path & "\data\data.ini"), "Settings", "MAX_SHOPS") = "" Then
        MAX_SHOPS = 255
    Else
        MAX_SHOPS = CLng(GetVar(App.Path & "\data\data.ini", "Settings", "MAX_SHOPS"))
    End If
    
    'set the maximum spells
    If GetVar((App.Path & "\data\data.ini"), "Settings", "MAX_SPELLS") = "" Then
        MAX_SPELLS = 255
    Else
        MAX_SPELLS = CLng(GetVar(App.Path & "\data\data.ini", "Settings", "MAX_SPELLS"))
    End If
    
    'set the maximum guilds
    If GetVar((App.Path & "\data\data.ini"), "Settings", "MAX_GUILDS") = "" Then
        MAX_GUILDS = 20
    Else
        MAX_GUILDS = CLng(GetVar(App.Path & "\data\data.ini", "Settings", "MAX_GUILDS"))
    End If
    
    'set the maximum experience level
    If GetVar((App.Path & "\data\data.ini"), "Settings", "MAX_EXPERIENCE") = "" Then
        MAX_EXPERIENCE = 100
    Else
        MAX_EXPERIENCE = CLng(GetVar(App.Path & "\data\data.ini", "Settings", "MAX_EXPERIENCE"))
    End If
    
    ReDim Map(1 To MAX_MAPS) As MapRec
    ReDim TempTile(1 To MAX_MAPS) As TempTileRec
    ReDim PlayersOnMap(1 To MAX_MAPS) As Long
    ReDim Player(1 To MAX_PLAYERS) As AccountRec
    ReDim Item(0 To MAX_ITEMS) As ItemRec
    ReDim Npc(0 To MAX_NPCS) As NpcRec
    ReDim MapItem(1 To MAX_MAPS, 1 To MAX_MAP_ITEMS) As MapItemRec
    ReDim MapNpc(1 To MAX_MAPS, 1 To MAX_MAP_NPCS) As MapNpcRec
    ReDim Shop(1 To MAX_SHOPS) As ShopRec
    ReDim Spell(1 To MAX_SPELLS) As SpellRec
    ReDim Guild(1 To MAX_GUILDS) As GuildRec
    ReDim Experience(1 To MAX_EXPERIENCE) As Long
    
    ' Get the listening socket ready to go
    frmServer.Socket(0).RemoteHost = frmServer.Socket(0).LocalIP
    frmServer.Socket(0).LocalPort = Game_Port
        
    ' Init all the player sockets
    For i = 1 To MAX_PLAYERS
        Call SetStatus("Initializing player array...")
        Call ClearPlayer(i)
        
        Load frmServer.Socket(i)
    Next i
    
    'Set names
    'frmLoad.Caption = G_NAME
    'frmServer.Caption = G_NAME & " Server"
    
    Call SetStatus("Clearing temp tile fields...")
    Call ClearTempTile
    Call SetStatus("Clearing maps...")
    Call ClearMaps
    Call SetStatus("Clearing map items...")
    Call ClearMapItems
    Call SetStatus("Clearing map npcs...")
    Call ClearMapNpcs
    Call SetStatus("Clearing npcs...")
    Call ClearNpcs
    Call SetStatus("Clearing items...")
    Call ClearItems
    Call SetStatus("Clearing shops...")
    Call ClearShops
    Call SetStatus("Clearing spells...")
    Call ClearSpells
    Call SetStatus("Loading classes...")
    Call LoadClasses
    Call SetStatus("Loading maps...")
    Call LoadMaps
    Call SetStatus("Loading items...")
    Call LoadItems
    Call SetStatus("Loading npcs...")
    Call LoadNpcs
    Call SetStatus("Loading shops...")
    Call LoadShops
    Call SetStatus("Loading spells...")
    Call LoadSpells
    Call SetStatus("Spawning map items...")
    Call SpawnAllMapsItems
    Call SetStatus("Spawning map npcs...")
    Call SpawnAllMapNpcs
    Call SetStatus("Clearing exp...")
    Call ClearExps
    Call SetStatus("Loading exp...")
    Call LoadExps
    Call SetStatus("Loading main.txt")
    Call InitSADScript
    
    frmLoad.lblScript.ForeColor = &H8000&
    frmLoad.lblScript.Caption = "DONE"
    frmLoad.lblStatus.ForeColor = &H8000&
    frmLoad.lblStatus.Caption = "COMPLETE"
    DoEvents
        
    ' Check if the master charlist file exists for checking duplicate names, and if it doesnt make it
    If Not FileExist("accounts\charlist.txt") Then
        f = FreeFile
        Open App.Path & "\accounts\charlist.txt" For Output As #f
        Close #f
    End If
    
    ' Start listening
    frmServer.Socket(0).Listen
    
    Call UpdateCaption
    
    frmLoad.Visible = False
    frmServer.Show
    
    SpawnSeconds = 0
    frmServer.tmrGameAI.Enabled = True
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modGeneral.bas", "InitServer", Err.Number, Err.Description)
End Sub

Sub DestroyServer()
'On Error GoTo errorhandler:
Dim i As Long
    
    ' Add to the sys tray
    Call Shell_NotifyIcon(NIM_DELETE, nid)
    frmLoad.Visible = True
    frmServer.Visible = False
    
    Call SetStatus("Saving players online...")
    Call SaveAllPlayersOnline
    Call SetStatus("Clearing maps...")
    Call ClearMaps
    Call SetStatus("Clearing map items...")
    Call ClearMapItems
    Call SetStatus("Clearing map npcs...")
    Call ClearMapNpcs
    Call SetStatus("Clearing npcs...")
    Call ClearNpcs
    Call SetStatus("Clearing items...")
    Call ClearItems
    Call SetStatus("Clearing shops...")
    Call ClearShops
    Call SetStatus("Unloading sockets and timers...")
    For i = 1 To MAX_PLAYERS
        Unload frmServer.Socket(i)
    Next i
    
    Set MyScript = Nothing
    Set Commands = Nothing
    End
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modGeneral.bas", "DestroyServer", Err.Number, Err.Description)
End Sub

Sub SetStatus(ByVal Status As String)
'On Error GoTo errorhandler:
    frmLoad.lblStatus.Caption = Status
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modGeneral.bas", "SetStatus", Err.Number, Err.Description)
End Sub

Sub ServerLogic()
'On Error GoTo errorhandler:
Dim i As Long

    ' Check for disconnections
    For i = 1 To MAX_PLAYERS
        If frmServer.Socket(i).State > 7 Then
            Call CloseSocket(i)
        End If
    Next i
        
    Call CheckGiveHP
    Call GameAI
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modGeneral.bas", "ServerLogic", Err.Number, Err.Description)
End Sub

Sub CheckSpawnMapItems()
'On Error GoTo errorhandler:
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
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modGeneral.bas", "CheckSpawnMapItems", Err.Number, Err.Description)
End Sub

Sub GameAI()
'On Error GoTo errorhandler:
Dim i As Long, x As Long, y As Long, n As Long, x1 As Long, y1 As Long, TickCount As Long
Dim Damage As Long, DistanceX As Long, DistanceY As Long, NpcNum As Long, Target As Long
Dim Chance As Byte, WalkPath As Byte
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
                        If Map(y).Tile(x1, y1).Key = True And TempTile(y).DoorOpen(x1, y1) = YES Then
                            TempTile(y).DoorOpen(x1, y1) = NO
                            Call SendDataToMap(y, "MAPKEY" & SEP_CHAR & x1 & SEP_CHAR & y1 & SEP_CHAR & 0 & SEP_CHAR & END_CHAR)
                        End If
                    Next x1
                Next y1
            End If
            
            For x = 1 To MAX_MAP_NPCS
                NpcNum = MapNpc(y, x).Num
                'NPC FEAR -smchronos
                If Npc(NpcNum).Fear = True And (MapNpc(y, x).HP * 5) < MapNpc(y, x).MaxHP Then
                    Target = MapNpc(y, x).Target
                    If Target > 0 Then
                    Chance = 1
                    Do While DidWalk = False
                        If Chance = 5 Then
                            DidWalk = True
                            Exit Do
                        End If
                        WalkPath = Int(Rnd * 4) + 1
                        Select Case WalkPath
                        Case 1:
                        ' Up
                        If MapNpc(y, x).y > GetPlayerY(Target) And DidWalk = False Then
                            If CanNpcMove(y, x, DIR_DOWN) Then
                                Call NpcMove(y, x, DIR_DOWN, MOVING_WALKING)
                                DidWalk = True
                            ElseIf CanNpcMove(y, x, DIR_UP) And DidWalk = False Then
                                Call NpcMove(y, x, DIR_UP, MOVING_WALKING)
                                DidWalk = True
                            Else
                                Chance = Chance + 1
                            End If
                        End If
                        Case 2:
                        ' Down
                        If MapNpc(y, x).y < GetPlayerY(Target) And DidWalk = False Then
                            If CanNpcMove(y, x, DIR_UP) Then
                                Call NpcMove(y, x, DIR_UP, MOVING_WALKING)
                                DidWalk = True
                            ElseIf CanNpcMove(y, x, DIR_DOWN) And DidWalk = False Then
                                Call NpcMove(y, x, DIR_DOWN, MOVING_WALKING)
                                DidWalk = True
                            Else
                                Chance = Chance + 1
                            End If
                        End If
                        Case 3:
                        ' Left
                        If MapNpc(y, x).x > GetPlayerX(Target) And DidWalk = False Then
                            If CanNpcMove(y, x, DIR_RIGHT) Then
                                Call NpcMove(y, x, DIR_RIGHT, MOVING_WALKING)
                                DidWalk = True
                            ElseIf CanNpcMove(y, x, DIR_LEFT) And DidWalk = False Then
                                Call NpcMove(y, x, DIR_LEFT, MOVING_WALKING)
                                DidWalk = True
                            Else
                                Chance = Chance + 1
                            End If
                        End If
                        Case 4:
                        ' Right
                        If MapNpc(y, x).x < GetPlayerX(Target) And DidWalk = False Then
                            If CanNpcMove(y, x, DIR_LEFT) Then
                                Call NpcMove(y, x, DIR_LEFT, MOVING_WALKING)
                                DidWalk = True
                            ElseIf CanNpcMove(y, x, DIR_RIGHT) And DidWalk = False Then
                                Call NpcMove(y, x, DIR_RIGHT, MOVING_WALKING)
                                DidWalk = True
                            Else
                                Chance = Chance + 1
                            End If
                        End If
                        End Select
                    Loop
                    Else
                    For i = 1 To MAX_PLAYERS
                        If IsPlaying(i) Then
                        n = Npc(NpcNum).Range
                                    
                        DistanceX = MapNpc(y, x).x - GetPlayerX(i)
                        DistanceY = MapNpc(y, x).y - GetPlayerY(i)
                                    
                        ' Make sure we get a positive value
                        If DistanceX < 0 Then DistanceX = DistanceX * -1
                        If DistanceY < 0 Then DistanceY = DistanceY * -1
                                    
                        ' Are they in range?  if so GET'M!
                        If DistanceX <= n And DistanceY <= n Then
                            MapNpc(y, x).Target = i
                        End If
                        End If
                    Next i
                    End If
                Else
                ' ////////////////////////////////////////////////////
                ' // This is used for ATTACKING ON SIGHT or TRAITOR //
                ' ////////////////////////////////////////////////////
                ' Make sure theres a npc with the map
                If Map(y).Npc(x) > 0 And MapNpc(y, x).Num > 0 Then
                    ' If the npc is a attack on sight, search for a player on the map
                    If MapNpc(y, x).Behavior = NPC_BEHAVIOR_ATTACKONSIGHT Or MapNpc(y, x).Behavior = NPC_BEHAVIOR_GUARD Or MapNpc(y, x).Behavior = NPC_BEHAVIOR_FOLLOW Then
                        For i = 1 To MAX_PLAYERS
                            If IsPlaying(i) Then
                                If GetPlayerMap(i) = y And MapNpc(y, x).Target = 0 And GetPlayerAccess(i) <= ADMIN_MONITER Then
                                    n = Npc(NpcNum).Range
                                    
                                    DistanceX = MapNpc(y, x).x - GetPlayerX(i)
                                    DistanceY = MapNpc(y, x).y - GetPlayerY(i)
                                    
                                    ' Make sure we get a positive value
                                    If DistanceX < 0 Then DistanceX = DistanceX * -1
                                    If DistanceY < 0 Then DistanceY = DistanceY * -1
                                    
                                    ' Are they in range?  if so GET'M!
                                    If DistanceX <= n And DistanceY <= n Then
                                        If MapNpc(y, x).Behavior = NPC_BEHAVIOR_ATTACKONSIGHT Or GetPlayerPK(i) = YES And MapNpc(y, x).Behavior = NPC_BEHAVIOR_FOLLOW Then
                                            If Trim$(Npc(NpcNum).AttackSay) <> "" Then
                                                Call PlayerMsg(i, "A " & Trim$(Npc(NpcNum).Name) & " says, '" & Trim$(Npc(NpcNum).AttackSay) & "' to you.", SayColor)
                                            End If
                                            
                                            MapNpc(y, x).Target = i
                                        End If
                                    End If
                                End If
                            End If
                        Next i
                    ElseIf MapNpc(y, x).Behavior = NPC_BEHAVIOR_TRAITOR Or MapNpc(y, x).Behavior = NPC_BEHAVIOR_FOLLOW Then
                    For i = 1 To MAX_MAP_NPCS
                        If i <> x Then
                            If MapNpc(y, x).Target = 0 And MapNpc(y, i).Num > 0 Then
                                n = Npc(NpcNum).Range
                                    
                                DistanceX = MapNpc(y, x).x - CLng(MapNpc(y, i).x)
                                DistanceY = MapNpc(y, x).y - CLng(MapNpc(y, i).y)
                                    
                                ' Make sure we get a positive value
                                If DistanceX < 0 Then DistanceX = DistanceX * -1
                                If DistanceY < 0 Then DistanceY = DistanceY * -1
                                    
                                ' Are they in range?  if so GET'M!
                                If DistanceX <= n And DistanceY <= n Then
                                    If MapNpc(y, x).Behavior = NPC_BEHAVIOR_TRAITOR Or MapNpc(y, x).Behavior = NPC_BEHAVIOR_FOLLOW Then
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
                    If MapNpc(y, x).Behavior <> NPC_BEHAVIOR_SHOPKEEPER Then
                        ' Check to see if we are following a player or not
                        If Target > 0 Then
                        If MapNpc(y, x).Behavior <> NPC_BEHAVIOR_TRAITOR Then
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
                        Else '-smchronos
                            If Target <= MAX_MAP_NPCS And Target > 0 And MapNpc(y, Target).HP > 0 Then
                                DidWalk = False
                                
                                i = Int(Rnd * 5)
                                
                                ' Lets move the npc
                                Select Case i
                                    Case 0
                                        ' Up
                                        If MapNpc(y, x).y > MapNpc(y, Target).y And DidWalk = False Then
                                            If CanNpcMove(y, x, DIR_UP) Then
                                                Call NpcMove(y, x, DIR_UP, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                        ' Down
                                        If MapNpc(y, x).y < MapNpc(y, Target).y And DidWalk = False Then
                                            If CanNpcMove(y, x, DIR_DOWN) Then
                                                Call NpcMove(y, x, DIR_DOWN, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                        ' Left
                                        If MapNpc(y, x).x > MapNpc(y, Target).x And DidWalk = False Then
                                            If CanNpcMove(y, x, DIR_LEFT) Then
                                                Call NpcMove(y, x, DIR_LEFT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                        ' Right
                                        If MapNpc(y, x).x < MapNpc(y, Target).x And DidWalk = False Then
                                            If CanNpcMove(y, x, DIR_RIGHT) Then
                                                Call NpcMove(y, x, DIR_RIGHT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                    
                                    Case 1
                                        ' Right
                                        If MapNpc(y, x).x < MapNpc(y, Target).x And DidWalk = False Then
                                            If CanNpcMove(y, x, DIR_RIGHT) Then
                                                Call NpcMove(y, x, DIR_RIGHT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                        ' Left
                                        If MapNpc(y, x).x > MapNpc(y, Target).x And DidWalk = False Then
                                            If CanNpcMove(y, x, DIR_LEFT) Then
                                                Call NpcMove(y, x, DIR_LEFT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                        ' Down
                                        If MapNpc(y, x).y < MapNpc(y, Target).y And DidWalk = False Then
                                            If CanNpcMove(y, x, DIR_DOWN) Then
                                                Call NpcMove(y, x, DIR_DOWN, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                        ' Up
                                        If MapNpc(y, x).y > MapNpc(y, Target).y And DidWalk = False Then
                                            If CanNpcMove(y, x, DIR_UP) Then
                                                Call NpcMove(y, x, DIR_UP, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                        
                                    Case 2
                                        ' Down
                                        If MapNpc(y, x).y < MapNpc(y, Target).y And DidWalk = False Then
                                            If CanNpcMove(y, x, DIR_DOWN) Then
                                                Call NpcMove(y, x, DIR_DOWN, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                        ' Up
                                        If MapNpc(y, x).y > MapNpc(y, Target).y And DidWalk = False Then
                                            If CanNpcMove(y, x, DIR_UP) Then
                                                Call NpcMove(y, x, DIR_UP, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                        ' Right
                                        If MapNpc(y, x).x < MapNpc(y, Target).x And DidWalk = False Then
                                            If CanNpcMove(y, x, DIR_RIGHT) Then
                                                Call NpcMove(y, x, DIR_RIGHT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                        ' Left
                                        If MapNpc(y, x).x > MapNpc(y, Target).x And DidWalk = False Then
                                            If CanNpcMove(y, x, DIR_LEFT) Then
                                                Call NpcMove(y, x, DIR_LEFT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                    
                                    Case 3
                                        ' Left
                                        If MapNpc(y, x).x > MapNpc(y, Target).x And DidWalk = False Then
                                            If CanNpcMove(y, x, DIR_LEFT) Then
                                                Call NpcMove(y, x, DIR_LEFT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                        ' Right
                                        If MapNpc(y, x).x < MapNpc(y, Target).x And DidWalk = False Then
                                            If CanNpcMove(y, x, DIR_RIGHT) Then
                                                Call NpcMove(y, x, DIR_RIGHT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                        ' Up
                                        If MapNpc(y, x).y > MapNpc(y, Target).y And DidWalk = False Then
                                            If CanNpcMove(y, x, DIR_UP) Then
                                                Call NpcMove(y, x, DIR_UP, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                        ' Down
                                        If MapNpc(y, x).y < MapNpc(y, Target).y And DidWalk = False Then
                                            If CanNpcMove(y, x, DIR_DOWN) Then
                                                Call NpcMove(y, x, DIR_DOWN, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
                                End Select

                                ' Check if we can't move and if player is behind something and if we can just switch dirs
                                If Not DidWalk Then
                                    If MapNpc(y, x).x - 1 = MapNpc(y, Target).x And MapNpc(y, x).y = MapNpc(y, Target).y Then
                                        If MapNpc(y, x).Dir <> DIR_LEFT Then
                                            Call NpcDir(y, x, DIR_LEFT)
                                        End If
                                        DidWalk = True
                                    End If
                                    If MapNpc(y, x).x + 1 = MapNpc(y, Target).x And MapNpc(y, x).y = MapNpc(y, Target).y Then
                                        If MapNpc(y, x).Dir <> DIR_RIGHT Then
                                            Call NpcDir(y, x, DIR_RIGHT)
                                        End If
                                        DidWalk = True
                                    End If
                                    If MapNpc(y, x).x = MapNpc(y, Target).x And MapNpc(y, x).y - 1 = MapNpc(y, Target).y Then
                                        If MapNpc(y, x).Dir <> DIR_UP Then
                                            Call NpcDir(y, x, DIR_UP)
                                        End If
                                        DidWalk = True
                                    End If
                                    If MapNpc(y, x).x = MapNpc(y, Target).x And MapNpc(y, x).y + 1 = MapNpc(y, Target).y Then
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
                        End If
                        Else
                            i = Int(Rnd * 4)
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
            If MapNpc(y, x).Behavior <> NPC_BEHAVIOR_TRAITOR Then '-smchronos
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
                                        Call PlayerMsg(Target, "The " & Trim$(Npc(NpcNum).Name) & "'s hit didn't even phase you!", BrightBlue)
                                    End If
                                Else
                                    Call PlayerMsg(Target, "Your " & Trim$(Item(GetPlayerInvItemNum(Target, GetPlayerShieldSlot(Target))).Name) & " blocks the " & Trim$(Npc(NpcNum).Name) & "'s hit!", BrightCyan)
                                    ' Send message to play sound
                                    Call SendDataToMap(GetPlayerMap(Target), "PLAYSOUND" & SEP_CHAR & Trim$(Item(GetPlayerInvItemNum(Target, GetPlayerShieldSlot(Target))).Sound) & SEP_CHAR & END_CHAR)
                                End If
                            End If
                        Else
                            ' Player left map or game, set target to 0
                            MapNpc(y, x).Target = 0
                        End If
                    End If
                End If
                
                ' /////////////////////////////////////////////
                ' // This is used for npcs to attack npcs ///// -smchronos
                ' /////////////////////////////////////////////
            ElseIf MapNpc(y, x).Behavior = NPC_BEHAVIOR_TRAITOR Then
                ' Make sure theres a npc with the map
                If Map(y).Npc(x) > 0 And MapNpc(y, x).Num > 0 Then
                    Target = MapNpc(y, x).Target
                If Target > 0 Then
                    'Check to see the targeted npc is still on the map, not dead
                    If Map(y).Npc(Target) > 0 And MapNpc(y, Target).Num > 0 Then
                        ' Check if the npc can attack the targeted npc
                        If MapNpc(y, Target).HP > 0 Then
                                ' Can the npc attack the target?
                                If CanNpcAttackNpc(x, Target, y) Then
                                'Attack the npc if the damage is over 0!
                                'Npc(Target).DEF
                                    Damage = CLng(Npc(NpcNum).STR) - CLng(Npc(Target).DEF)
                                    If Damage > 0 Then
                                        Call NpcAttackNpc(x, Target, Damage, y)
                                    End If
                                End If
                        End If
                    End If
                End If
                End If
            End If
                
                ' ////////////////////////////////////////////
                ' // This is used for regenerating NPC's HP //
                ' ////////////////////////////////////////////
                ' Check to see if we want to regen some of the npc's hp
                If MapNpc(y, x).Num > 0 And TickCount > GiveNPCHPTimer + 10000 Then
                    If MapNpc(y, x).HP > 0 Then
                        MapNpc(y, x).HP = MapNpc(y, x).HP + GetNpcHPRegen(NpcNum)
                    
                        ' Check if they have more then they should and if so just set it to max
                        If MapNpc(y, x).HP > GetNpcMaxHP(NpcNum) Then
                            MapNpc(y, x).HP = GetNpcMaxHP(NpcNum)
                            'Call SendDataToMap(y, "NPCHP" & SEP_CHAR & x & SEP_CHAR & MapNpc(y, x).HP & SEP_CHAR & GetNpcMaxHP(NpcNum) & SEP_CHAR & END_CHAR)
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
                If MapNpc(y, x).Num = 0 And Map(y).Npc(x) > 0 Then
                    If TickCount > MapNpc(y, x).SpawnWait + (Npc(Map(y).Npc(x)).SpawnSecs * 1000) Then
                        Call SpawnNpc(x, y)
                    End If
                End If
            End If
            Next x
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
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modGeneral.bas", "GameAI", Err.Number, Err.Description)
End Sub

Sub CheckGiveHP()
'On Error GoTo errorhandler:
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
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modGeneral.bas", "CheckGiveHP", Err.Number, Err.Description)
End Sub

Sub PlayerSaveTimer()
'On Error GoTo errorhandler:
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
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modGeneral.bas", "PlayerSaveTimer", Err.Number, Err.Description)
End Sub

Sub ResetShopStock(ByVal ShopNum As Long)
'On Error GoTo errorhandler:
Dim x As Byte
    For x = 1 To MAX_TRADES
        Shop(ShopNum).TradeItem(x).Stock = Shop(ShopNum).TradeItem(x).MaxStock
    Next x
ErrorHandlerExit:
  Exit Sub
errorhandler:
  Call ReportError("modGeneral.bas", "ResetShopStock", Err.Number, Err.Description)
End Sub
