Attribute VB_Name = "modGameLogic"
Option Explicit

' ------------------------------------------
' --               Asphodel               --
' ------------------------------------------

Public Sub GameLoop()
Dim Tick As Currency
Dim TickFPS As Currency
Dim FPS As Long
Dim WalkTimer As Currency
Dim ScrollTextTimer As Currency
Dim tmr2000 As Currency

    ' Used for calculating fps
    TickFPS = GetTickCountNew
    FPS = 0
    
    ' Start the main game loop
    Do While InGame
    
        Tick = GetTickCountNew
        
        ' Set their ingame status to their TCP
        ' connection status, if they lose connection
        ' with the server, close the game this way
        InGame = IsConnected
        
        ' Check all of the movement keys to make
        ' sure they aren't trying to auto move
        CheckKeys
        
        ' Lets process the sound buffers to make sure
        ' there aren't any memory leaks
        ProcessSoundBuffers
        
        ' Draw the spell cool down bars
        If frmMainGame.picSpells.Visible Then DrawSpellsWaiting
        
        If LenB(ScrollText.Text) > 0 Then
            If frmMainGame.lblSignText.Caption <> ScrollText.Text Then
                If ScrollTextTimer < GetTickCountNew Then
                    HandleScrollText
                    ScrollTextTimer = GetTickCountNew + 50
                End If
            Else
                ScrollText.Running = False
                If ScrollTextTimer < GetTickCountNew Then
                    frmMainGame.lblPressEnter.Visible = Not frmMainGame.lblPressEnter.Visible
                    ScrollTextTimer = GetTickCountNew + 500
                End If
            End If
        End If
        
        ' Cycle the map animation on and off
        ' every quarter of a second
        If MapAnimTimer < Tick Then
            MapAnim = Not MapAnim
            MapAnimTimer = Tick + 250
        End If
        
        ' If they are on a warp tile, then they can't move
        If Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex)).Type = Tile_Type.Warp_ Then CanMoveNow = False
        
        ' Process the movement before you start
        ' drawing, or else it will be behind a frame
        If Not GettingMap Then
            If WalkTimer < Tick Then
                ' Process player movements
                ProcessMovement
                
                ' Process NPC movements
                ProcessNpcMovement
                
                WalkTimer = Tick + 15
            End If
            
            If CanMoveNow Then
            
                If GetPlayerAccess(MyIndex) > 0 Then
                    If frmAdmin.chkItemPick.Value = 1 Then
                        CheckMapGetItem
                    End If
                End If
                
                ' Check if player is trying to move
                CheckMovement
                
                ' Check to see if player is trying to attack
                CheckAttack
                
            End If
        End If
        
        ' Handle all of the drawing so players
        ' will be able to see stuff!
        Render_Graphics
        
        ' Calculate the ping
        If PingEnabled Then
            If Not WaitingonPing Then
                If tmr2000 < GetTickCountNew Then
                    DeterminePing
                    tmr2000 = GetTickCountNew + 2000
                End If
            End If
        End If
        
        ' Calculate the FPS
        If TickFPS > Tick Then
            FPS = FPS + 1
        Else
            GameFPS = FPS
            TickFPS = Tick + 1000
            FPS = 0
        End If
        
        DoEvents
        Sleep 1
        
        ' FPS cap
        If FPS_Lock > 0 Then
            If (GetTickCountNew - Tick) < FPS_Lock Then
                Sleep FPS_Lock - (GetTickCountNew - Tick)
            End If
        End If
        
    Loop
    
    frmMainGame.Visible = False
    
    If isLogging Then
        frmMainGame.picInventoryList.Cls
        frmMainGame.picSpellList.Cls
        GetRidOfShop
        ZeroMemory ByVal VarPtr(ShopTrade), LenB(ShopTrade)
        If frmAdmin.Visible Then Unload frmAdmin
        CurrentWindow = Window_State.Main_Menu
        frmMainGame.txtChat.Text = vbNullString
        Load_GameConfig
        frmMainMenu.Caption = Game_Name & " [Main Menu]"
        frmMainMenu.Visible = True
        GettingMap = True
    Else
        ' Shutdown the game
        SetStatus "Destroying game data..."
        DestroyGame
    End If
    
End Sub

Public Sub GetRidOfShop()

    With frmMainGame
        .lblWelcome.Caption = vbNullString
        .lblShopItem.Caption = "None"
        .lblShopCost.Caption = "Nothing"
        .lblShopDesc.Caption = "No item selected."
        .picShopList.Cls
        .picShop.Visible = False
    End With
    
    ReadyToRepair = False
    ReadyToSell = False
    
    ZeroMemory ByVal VarPtr(ShopTrade), LenB(ShopTrade)
    
End Sub

Private Sub DeterminePing()
    WaitingonPing = True
    PingCounter = GetTickCountNew
    SendData CPing & END_CHAR
End Sub

Private Sub HandleScrollText()
    If ScrollText.CurLetter = Len(ScrollText.Text) + 1 Then Exit Sub
    With frmMainGame.lblSignText
        .Caption = .Caption & Mid$(ScrollText.Text, ScrollText.CurLetter, 1)
    End With
    ScrollText.CurLetter = ScrollText.CurLetter + 1
End Sub

Public Sub Handle_PlayerAnims()
Dim i As Long
Dim ii As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            If GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                With Player(i)
                
                    If .Moving = 0 Then
                        If .Attacking = 0 Then
                            GoTo Skipper
                        End If
                    End If
                    
                    If .Moving > 0 Then
                    
                        If .WalkAnim < 0 Then .WalkAnim = 0
                        
                        If .WalkTimer < GetTickCountNew Then
                            .WalkAnim = .WalkAnim + 1
                            If .WalkAnim > UBound(GameConfig.WalkFrame) Then .WalkAnim = 1
                            .WalkTimer = GetTickCountNew + GameConfig.WalkAnim_Speed
                        End If
                        
                        GoTo Skipper
                        
                    ElseIf .Attacking Then
                    
                        If GameConfig.Total_AttackFrames > 0 Then
                        
                            If .WalkAnim < 0 Then .WalkAnim = 0
                            
                            If .WalkTimer < GetTickCountNew Then
                                .WalkAnim = .WalkAnim + 1
                                If .WalkAnim > UBound(GameConfig.AttackFrame) Then .WalkAnim = 1
                                .WalkTimer = GetTickCountNew + (UBound(GameConfig.AttackFrame) / 1000)
                            End If
                            
                        Else
                            .WalkAnim = GameConfig.StandFrame
                        End If
                        
                        GoTo Skipper
                        
                    End If
Skipper:
                End With
            End If
        End If
    Next
    
End Sub

Public Sub Handle_NpcAnims()
Dim i As Long
Dim ii As Long

    For i = 1 To UBound(MapNpc)
    
        With MapNpc(i)
            If .Num > 0 Then
                
                If .Moving = 0 Then
                    If .Attacking = 0 Then
                        GoTo Skipper
                    End If
                End If
                
                If .Moving > 0 Then
                
                    If .WalkAnim < 0 Then .WalkAnim = 0
                    
                    If .WalkTimer < GetTickCountNew Then
                        .WalkAnim = .WalkAnim + 1
                        If .WalkAnim > GameConfig.Total_WalkFrames Then .WalkAnim = 1
                        .WalkTimer = GetTickCountNew + GameConfig.WalkAnim_Speed
                    End If
                    
                    GoTo Skipper
                    
                ElseIf .Attacking Then
                
                    If GameConfig.Total_AttackFrames > 0 Then
                        If .WalkAnim < 0 Then .WalkAnim = 0
                        If .WalkTimer < GetTickCountNew Then
                            .WalkAnim = .WalkAnim + 1
                            If .WalkAnim > UBound(GameConfig.AttackFrame) Then .WalkAnim = 1
                            .WalkTimer = GetTickCountNew + (1000 / UBound(GameConfig.AttackFrame))
                        End If
                    Else
                        .WalkAnim = GameConfig.StandFrame
                    End If
                    
                    GoTo Skipper
                    
                End If
            End If
        End With
Skipper:
    Next
    
End Sub

Public Sub ProcessMovement()
Dim MovementSpeed As Long
Dim Index As Long
Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            If GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                Index = i
                
                If Player(Index).Moving > 0 Then
                
                    ' Check if player is walking, and if so process moving them over
                    If Player(Index).Moving = MovementType.Walking Then
                        MovementSpeed = WALK_SPEED
                    ElseIf Player(Index).Moving = MovementType.Running Then
                        MovementSpeed = RUN_SPEED
                    End If
                    
                    Select Case GetPlayerDir(Index)
                        Case E_Direction.Up_
                            Player(Index).YOffset = Player(Index).YOffset - MovementSpeed
                            If Player(Index).YOffset < 0 Then Player(Index).YOffset = 0
                        Case E_Direction.Down_
                            Player(Index).YOffset = Player(Index).YOffset + MovementSpeed
                            If Player(Index).YOffset > 0 Then Player(Index).YOffset = 0
                        Case E_Direction.Left_
                            Player(Index).XOffset = Player(Index).XOffset - MovementSpeed
                            If Player(Index).XOffset < 0 Then Player(Index).XOffset = 0
                        Case E_Direction.Right_
                            Player(Index).XOffset = Player(Index).XOffset + MovementSpeed
                            If Player(Index).XOffset > 0 Then Player(Index).XOffset = 0
                    End Select
                    
                    ' Check if completed walking over to the next tile
                    If Player(Index).XOffset = 0 Then
                        If Player(Index).YOffset = 0 Then
                            Player(Index).Moving = 0
                        End If
                    End If
                    
                End If
            End If
        End If
    Next
    
End Sub

Public Sub ProcessNpcMovement()
Dim MapNpcNum As Long

    For MapNpcNum = 1 To UBound(MapNpc)
    
        If MapNpc(MapNpcNum).Num > 0 Then
            ' Check if player is walking, and if so process moving them over
            If MapNpc(MapNpcNum).Moving = MovementType.Walking Then
                Select Case GetMapNpcDir(MapNpcNum)
                    Case E_Direction.Up_
                        MapNpc(MapNpcNum).YOffset = MapNpc(MapNpcNum).YOffset - WALK_SPEED
                        If MapNpc(MapNpcNum).YOffset < 0 Then MapNpc(MapNpcNum).YOffset = 0
                    Case E_Direction.Down_
                        MapNpc(MapNpcNum).YOffset = MapNpc(MapNpcNum).YOffset + WALK_SPEED
                        If MapNpc(MapNpcNum).YOffset > 0 Then MapNpc(MapNpcNum).YOffset = 0
                    Case E_Direction.Left_
                        MapNpc(MapNpcNum).XOffset = MapNpc(MapNpcNum).XOffset - WALK_SPEED
                        If MapNpc(MapNpcNum).XOffset < 0 Then MapNpc(MapNpcNum).XOffset = 0
                    Case E_Direction.Right_
                        MapNpc(MapNpcNum).XOffset = MapNpc(MapNpcNum).XOffset + WALK_SPEED
                        If MapNpc(MapNpcNum).XOffset > 0 Then MapNpc(MapNpcNum).XOffset = 0
                End Select
                
                ' Check if completed walking over to the next tile
                If MapNpc(MapNpcNum).XOffset = 0 Then
                    If MapNpc(MapNpcNum).YOffset = 0 Then
                        MapNpc(MapNpcNum).Moving = 0
                    End If
                End If
                
            End If
        End If
        
    Next
    
End Sub

Public Sub HandleKeypresses(ByVal KeyAscii As Integer)
Dim ChatText As String
Dim i As Long
Dim n As Long
Dim Command() As String

    If frmMainGame.picSign.Visible Then Exit Sub
    
    ChatText = Trim$(MyText)
    
    If LenB(ChatText) = 0 Then
        If KeyAscii = vbKeyReturn Then
            SendData CPressReturn & END_CHAR
            Exit Sub
        End If
        Exit Sub
    End If
    
    MyText = LCase$(ChatText)
    
    ' Handle when the player presses the return key
    If KeyAscii = vbKeyReturn Then
    
        If LenB(ChatText) = 0 Then
            SendData CPressReturn & END_CHAR
            Exit Sub
        End If
        
        ' Broadcast message
        If Left$(ChatText, 1) = "'" Then
            ChatText = Mid$(ChatText, 2, Len(ChatText) - 1)
            If Len(ChatText) > 0 Then
                SendMessage ChatType.BroadcastMsg, ChatText
            End If
            MyText = vbNullString
            frmMainGame.txtMyChat.Text = vbNullString
            Exit Sub
        End If
        
        ' Emote message
        If Left$(ChatText, 1) = "-" Then
            MyText = Mid$(ChatText, 2, Len(ChatText) - 1)
            If Len(ChatText) > 0 Then
                SendMessage ChatType.EmoteMsg, ChatText
            End If
            MyText = vbNullString
            frmMainGame.txtMyChat.Text = vbNullString
            Exit Sub
        End If
        
        ' Player message
        If Left$(ChatText, 1) = "!" Then
            ChatText = Mid$(ChatText, 2, Len(ChatText) - 1)
            
            ' Make sure they are actually sending something
            If Len(ChatText) - i > 0 Then
                MyText = Mid$(ChatText, i + 1, Len(ChatText) - i)
                    
                ' Send the message to the player
                SendMessage ChatType.PrivateMsg, ChatText
            Else
                AddText "Usage: !playername (message)", AlertColor
            End If
            MyText = vbNullString
            frmMainGame.txtMyChat.Text = vbNullString
            Exit Sub
        End If
        
        ' Global Message
        If Left$(ChatText, 1) = vbQuote Then
            If GetPlayerAccess(MyIndex) >= StaffType.Mapper Then
                ChatText = Mid$(ChatText, 2, Len(ChatText) - 1)
                If Len(ChatText) > 0 Then
                    SendMessage ChatType.GlobalMsg, ChatText
                End If
                MyText = vbNullString
                frmMainGame.txtMyChat.Text = vbNullString
                Exit Sub
            End If
        End If
        
        ' Admin Message
        If Left$(ChatText, 1) = "=" Then
            If GetPlayerAccess(MyIndex) >= StaffType.Mapper Then
                ChatText = Mid$(ChatText, 2, Len(ChatText) - 1)
                If Len(ChatText) > 0 Then
                    SendMessage ChatType.AdminMsg, ChatText
                End If
                MyText = vbNullString
                frmMainGame.txtMyChat.Text = vbNullString
                Exit Sub
            End If
        End If
        
        If Left$(MyText, 1) = "/" Then
            Command = Split(MyText, " ")
            
            Select Case Command(0)
                Case "/guild"
                    If LenB(Trim$(Player(MyIndex).GuildName)) < 1 Then
                        AddText "You aren't in a guild!", AlertColor
                        GoTo Continue
                    End If
                    
                    frmMainGame.picGuildCP.Visible = Not frmMainGame.picGuildCP.Visible
                    
                Case "/help"
                    AddText "Social Commands:", HelpColor
                    AddText "'msghere = Broadcast Message", HelpColor
                    AddText "-msghere = Emote Message", HelpColor
                    AddText "!namehere msghere = Player Message", HelpColor
                    AddText "Available Commands: /help, /info, /who, /fps, /ping, /inv, /stats, /party, /join, /leave", HelpColor
                    
                Case "/info"
                    ' Checks to make sure we have more than one string in the array
                    If UBound(Command) < 1 Then
                        AddText "Usage: /info (name)", AlertColor
                        GoTo Continue
                    End If
                    
                    If IsNumeric(Command(1)) Then
                        AddText "Usage: /info (name)", AlertColor
                        GoTo Continue
                    End If
                    
                    SendData CPlayerInfoRequest & SEP_CHAR & Command(1) & END_CHAR
                    
                ' Whos Online
                Case "/who"
                    SendWhosOnline
                                
                ' Checking fps
                Case "/fps"
                    BFPS = Not BFPS
                    
                    If BFPS Then
                        frmMainGame.chkFPS.Value = 1
                    Else
                        frmMainGame.chkFPS.Value = 0
                    End If
                    
                ' Checking ping
                Case "/ping"
                    If PingEnabled = 0 Then PingEnabled = 1 Else PingEnabled = 0
                    PutVar App.Path & "/info.ini", "BASIC", "Ping", CStr(PingEnabled)
                    
                    frmMainGame.chkPing.Value = PingEnabled
                    
                ' Show inventory
                Case "/inv"
                    frmMainGame.picInv.Visible = True
                    UpdateInventory
                    
                ' Request stats
                Case "/stats"
                    SendData CGetStats & END_CHAR
                    
                ' Party request
                Case "/party"
                    ' Make sure they are actually sending something
                    If UBound(Command) < 1 Then
                        AddText "Usage: /party (name)", AlertColor
                        GoTo Continue
                    End If
                    
                    If IsNumeric(Command(1)) Then
                        AddText "Usage: /party (name)", AlertColor
                        GoTo Continue
                    End If
                        
                    SendPartyRequest (Command(1))
                    
                ' Join party
                Case "/join"
                    SendJoinParty
                
                ' Leave party
                Case "/leave"
                    SendLeaveParty
                
                ' // Monitor Admin Commands //
                Case "/acp"
                    If GetPlayerAccess(MyIndex) < StaffType.Mapper Then
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                        GoTo Continue
                    End If
                    
                    frmAdmin.Show
                    
                ' Admin Help
                Case "/admin"
                    If GetPlayerAccess(MyIndex) < StaffType.Mapper Then
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                        GoTo Continue
                    End If
                    
                    AddText "Social Commands:", HelpColor
                    AddText """msghere = Global Admin Message", HelpColor
                    AddText "=msghere = Private Admin Message", HelpColor
                    AddText "Available Commands: /acp, /admin, /loc, /editmap, /warpmeto, /warptome, /warpto, /setsprite, /mapreport, /kick, /ban, /edititem, /respawn, /editnpc, /motd, /editshop, /editspell", HelpColor
                    
                ' Kicking a player
                Case "/kick"
                    If GetPlayerAccess(MyIndex) < StaffType.Mapper Then
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                        GoTo Continue
                    End If
                    
                    If UBound(Command) < 1 Then
                        AddText "Usage: /kick (name)", AlertColor
                        GoTo Continue
                    End If
                    
                    If IsNumeric(Command(1)) Then
                        AddText "Usage: /kick (name)", AlertColor
                        GoTo Continue
                    End If

                    SendKick Command(1)
                            
                ' // Mapper Admin Commands //
                ' Location
                Case "/loc"
                
                    If GetPlayerAccess(MyIndex) < StaffType.Mapper Then
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                        GoTo Continue
                    End If
                    
                    BLoc = Not BLoc
                    
                    If BLoc Then
                        frmAdmin.chkDisplayCurrent.Value = 1
                    Else
                        frmAdmin.chkDisplayCurrent.Value = 0
                    End If
                    
                ' Warping to a player
                Case "/warpmeto"
                    If GetPlayerAccess(MyIndex) < StaffType.Mapper Then
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                        GoTo Continue
                    End If
                    
                    If UBound(Command) < 1 Then
                        AddText "Usage: /warpmeto (name)", AlertColor
                        GoTo Continue
                    End If
                    
                    If IsNumeric(Command(1)) Then
                        AddText "Usage: /warpmeto (name)", AlertColor
                        GoTo Continue
                    End If
                    
                    WarpMeTo Command(1)
                            
                ' Warping a player to you
                Case "/warptome"
                    If GetPlayerAccess(MyIndex) < StaffType.Mapper Then
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                        GoTo Continue
                    End If
                    
                    If UBound(Command) < 1 Then
                        AddText "Usage: /warptome (name)", AlertColor
                        GoTo Continue
                    End If
                    
                    If IsNumeric(Command(1)) Then
                        AddText "Usage: /warptome (name)", AlertColor
                        GoTo Continue
                    End If
                    
                    WarpToMe Command(1)
                            
                ' Warping to a map
                Case "/warpto"
                    If GetPlayerAccess(MyIndex) < StaffType.Mapper Then
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                        GoTo Continue
                    End If
                    
                    If UBound(Command) < 1 Then
                        AddText "Usage: /warpto (map #)", AlertColor
                        GoTo Continue
                    End If
                    
                    If Not IsNumeric(Command(1)) Then
                        AddText "Usage: /warpto (map #)", AlertColor
                        GoTo Continue
                    End If
                    
                    n = CLng(Command(1))
                
                    ' Check to make sure its a valid map #
                    If n > 0 And n <= MAX_MAPS Then
                        WarpTo (n)
                    Else
                        AddText "Invalid map number.", Color.red
                    End If
                    
                ' Setting sprite
                Case "/setsprite"
                    If GetPlayerAccess(MyIndex) < StaffType.Mapper Then
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                        GoTo Continue
                    End If
                    
                    If UBound(Command) < 1 Then
                        AddText "Usage: /setsprite (sprite #)", AlertColor
                        GoTo Continue
                    End If
                    
                    If Not IsNumeric(Command(1)) Then
                        AddText "Usage: /setsprite (sprite #)", AlertColor
                        GoTo Continue
                    End If
                    
                    SendSetSprite CLng(Command(1))
                
                ' Map report
                Case "/mapreport"
                    If GetPlayerAccess(MyIndex) < StaffType.Mapper Then
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                        GoTo Continue
                    End If
                        
                    SendData CMapReport & END_CHAR
            
                ' Respawn request
                Case "/respawn"
                    If GetPlayerAccess(MyIndex) < StaffType.Mapper Then
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                        GoTo Continue
                    End If
                    
                    SendMapRespawn
            
                ' MOTD change
                Case "/motd"
                    If GetPlayerAccess(MyIndex) < StaffType.Mapper Then
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                        GoTo Continue
                    End If
                    
                    If UBound(Command) < 1 Then
                        AddText "Usage: /motd (new motd)", AlertColor
                        GoTo Continue
                    End If

                    SendMOTDChange Right$(ChatText, Len(ChatText) - 5)
                
                ' Check the ban list
                Case "/banlist"
                    If GetPlayerAccess(MyIndex) < StaffType.Mapper Then
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                        GoTo Continue
                    End If
                    
                    SendBanList
                
                ' Banning a player
                Case "/ban"
                    If GetPlayerAccess(MyIndex) < StaffType.Mapper Then
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                        GoTo Continue
                    End If
                    
                    If UBound(Command) < 1 Then
                        AddText "Usage: /ban (name)", AlertColor
                        GoTo Continue
                    End If
                    
                    SendBan Command(1)
                
                ' // Developer Admin Commands //
                ' Editing item request
                Case "/edititem"
                    If GetPlayerAccess(MyIndex) < StaffType.Developer Then
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                        GoTo Continue
                    End If
                    
                    SendRequestEditItem
                    
                ' Editing npc request
                Case "/editnpc"
                    If GetPlayerAccess(MyIndex) < StaffType.Developer Then
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                        GoTo Continue
                    End If
                    
                    SendRequestEditNpc
                    
                ' Editing map request
                Case "/editmap"
editmap:
                    If GetPlayerAccess(MyIndex) < StaffType.Mapper Then
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                        GoTo Continue
                    End If
                    
                    SendRequestEditMap
                    
                Case "/mapeditor"
                    GoTo editmap
                    
                ' Editing shop request
                Case "/editshop"
                    If GetPlayerAccess(MyIndex) < StaffType.Developer Then
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                        GoTo Continue
                    End If
                    
                    SendRequestEditShop
                    
                ' Editing spell request
                Case "/editspell"
                    If GetPlayerAccess(MyIndex) < StaffType.Developer Then
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                        GoTo Continue
                    End If
                    
                    SendRequestEditSpell
                    
                ' Editing anim request
                Case "/editanim"
                    If GetPlayerAccess(MyIndex) < StaffType.Developer Then
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                        GoTo Continue
                    End If
                    
                    SendRequestEditAnim
                    
                ' Editing sign request
                Case "/editsign"
                    If GetPlayerAccess(MyIndex) < StaffType.Developer Then
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                        GoTo Continue
                    End If
                    
                    SendRequestEditSign
                    
                ' // Creator Admin Commands //
                ' Giving another player access
                Case "/setaccess"
                    If GetPlayerAccess(MyIndex) < StaffType.Creator Then
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                        GoTo Continue
                    End If
                    
                    If UBound(Command) < 2 Then
                        AddText "Usage: /setaccess (name) (access)", AlertColor
                        GoTo Continue
                    End If
                    
                    If IsNumeric(Command(1)) Or Not IsNumeric(Command(2)) Then
                        AddText "Usage: /setaccess (name) (access)", AlertColor
                        GoTo Continue
                    End If
                    
                    SendSetAccess Command(1), CLng(Command(2))
                
                ' Ban destroy
                Case "/destroybanlist"
                    If GetPlayerAccess(MyIndex) < StaffType.Creator Then
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                        GoTo Continue
                    End If
                    
                    SendBanDestroy
                
                Case Else
                    AddText "Not a valid command!", HelpColor
                    
            End Select
            
'continue label where we go instead of exiting the sub
Continue:
            MyText = vbNullString
            frmMainGame.txtMyChat.Text = vbNullString
            Exit Sub
        End If
        
        ' Say message
        If Len(ChatText) > 0 Then
            SendMessage ChatType.MapMsg, ChatText
        End If
        
        MyText = vbNullString
        frmMainGame.txtMyChat.Text = vbNullString
        Exit Sub
    End If
    
    ' Handle when the user presses the backspace key
    If (KeyAscii = vbKeyBack) Then
        If LenB(MyText) > 0 Then MyText = Mid$(MyText, 1, Len(MyText) - 1)
    End If
    
    ' And if neither, then add the character to the user's text buffer
    If (KeyAscii <> vbKeyReturn) Then
        If (KeyAscii <> vbKeyBack) Then
            
            ' Make sure the character is on standard English keyboard
            If KeyAscii >= 32 Then
                If KeyAscii <= 126 Then
                    MyText = MyText & ChrW$(KeyAscii)
                End If
            End If
            
        End If
    End If
    
End Sub

Public Sub CheckMapGetItem()
    If GetPlayerAccess(MyIndex) > 0 Then
        If frmAdmin.chkItemPick.Value = 1 Then
            GoTo SkipIfStatement
        End If
    End If
    If Player(MyIndex).MapGetTimer < GetTickCountNew Then
        If LenB(Trim$(MyText)) = 0 Then
SkipIfStatement:
            Player(MyIndex).MapGetTimer = GetTickCountNew + 250
            SendData CMapGetItem & END_CHAR
        End If
    End If
End Sub

Public Sub CheckAttack()
    If ControlDown Then
        If Player(MyIndex).AttackTimer + 1000 < GetTickCountNew Then
            If Player(MyIndex).Attacking = 0 Then
                With Player(MyIndex)
                    .Attacking = 1
                    .AttackTimer = GetTickCountNew
                End With
                SendData CAttack & END_CHAR
            End If
        End If
    End If
End Sub

Public Sub CheckInput(ByVal KeyState As Byte, ByVal KeyCode As Integer)
    If Not GettingMap Then
        If KeyState = 1 Then
            Select Case KeyCode
            
                Case vbKeyReturn
                    CheckMapGetItem
                    
                Case vbKeyInsert
                    If Selected_Spell > 0 Then
                        If Player(MyIndex).CastTimer(Selected_Spell) < GetTickCountNew Then
                            If Player(MyIndex).Moving = 0 Then
                                SendData CCast & SEP_CHAR & Selected_Spell & END_CHAR
                            Else
                                AddText "Cannot cast while walking!", Color.BrightRed
                            End If
                        End If
                    Else
                        AddText "No spell memorized! Click a spell to memorize it.", Color.BrightRed
                    End If
                    
            End Select
        Else
            Select Case KeyCode
            
                Case vbKeyUp
                    DirUp = False
                    
                Case vbKeyDown
                    DirDown = False
                    
                Case vbKeyLeft
                    DirLeft = False
                    
                Case vbKeyRight
                    DirRight = False
                    
                Case vbKeyShift
                    ShiftDown = False
                    
                Case vbKeyControl
                    ControlDown = False

            End Select
        End If
    End If
End Sub

Function IsTryingToMove() As Boolean
    IsTryingToMove = (DirUp Or DirDown Or DirLeft Or DirRight)
End Function

Function CanMove() As Boolean
Dim d As Long

    CanMove = True
   
    ' Make sure they aren't trying to move when they are already moving
    If Player(MyIndex).Moving <> 0 Then
        CanMove = False
        Exit Function
    End If
    
    ' Make sure they haven't just casted a spell
    If Player(MyIndex).CastedSpell = YES Then
        If GetTickCountNew > Player(MyIndex).AttackTimer + 1000 Then
            Player(MyIndex).CastedSpell = NO
        Else
            CanMove = False
            Exit Function
        End If
    End If
    
    If frmMainGame.picSign.Visible Or frmMainGame.picShop.Visible Then
        CanMove = False
        Exit Function
    End If
    
    d = GetPlayerDir(MyIndex)
    If DirUp Then
        SetPlayerDir MyIndex, E_Direction.Up_
       
        ' Check to see if they are trying to go out of bounds
        If GetPlayerY(MyIndex) > 0 Then
            If CheckDirection(E_Direction.Up_) Then
                CanMove = False
               
                ' Set the new direction if they weren't facing that direction
                If d <> E_Direction.Up_ Then
                    SendPlayerDir
                End If
                Exit Function
            End If
        Else
            ' Check if they can warp to a new map
            If Map.Up > 0 Then
                MapEditorLeaveMap
                SendPlayerRequestNewMap
                GettingMap = True
                CanMoveNow = False
            End If
            CanMove = False
            Exit Function
        End If
    End If
           
    If DirDown Then
        SetPlayerDir MyIndex, E_Direction.Down_
        
        ' Check to see if they are trying to go out of bounds
        If GetPlayerY(MyIndex) < MAX_MAPY Then
            If CheckDirection(E_Direction.Down_) Then
                CanMove = False
                
                ' Set the new direction if they weren't facing that direction
                If d <> E_Direction.Down_ Then
                    SendPlayerDir
                End If
                Exit Function
            End If
        Else
            ' Check if they can warp to a new map
            If Map.Down > 0 Then
                MapEditorLeaveMap
                SendPlayerRequestNewMap
                GettingMap = True
                CanMoveNow = False
            End If
            CanMove = False
            Exit Function
        End If
    End If
               
    If DirLeft Then
        SetPlayerDir MyIndex, E_Direction.Left_
       
        ' Check to see if they are trying to go out of bounds
        If GetPlayerX(MyIndex) > 0 Then
            If CheckDirection(E_Direction.Left_) Then
                CanMove = False
               
                ' Set the new direction if they weren't facing that direction
                If d <> E_Direction.Left_ Then
                    SendPlayerDir
                End If
                Exit Function
            End If
        Else
            ' Check if they can warp to a new map
            If Map.Left > 0 Then
                MapEditorLeaveMap
                SendPlayerRequestNewMap
                GettingMap = True
                CanMoveNow = False
            End If
            CanMove = False
            Exit Function
        End If
    End If
       
    If DirRight Then
        SetPlayerDir MyIndex, E_Direction.Right_
       
        ' Check to see if they are trying to go out of bounds
        If GetPlayerX(MyIndex) < MAX_MAPX Then
            If CheckDirection(E_Direction.Right_) Then
                CanMove = False
                ' Set the new direction if they weren't facing that direction
                If d <> E_Direction.Right_ Then
                    SendPlayerDir
                End If
                Exit Function
            End If
        Else
            ' Check if they can warp to a new map
            If Map.Right > 0 Then
                MapEditorLeaveMap
                SendPlayerRequestNewMap
                GettingMap = True
                CanMoveNow = False
            End If
            CanMove = False
            Exit Function
        End If
    End If
End Function

Function CheckDirection(ByVal Direction As Byte) As Boolean
Dim X As Long
Dim Y As Long
Dim i As Long
   
    Select Case Direction
        Case E_Direction.Up_
            X = GetPlayerX(MyIndex)
            Y = GetPlayerY(MyIndex) - 1
        Case E_Direction.Down_
            X = GetPlayerX(MyIndex)
            Y = GetPlayerY(MyIndex) + 1
        Case E_Direction.Left_
            X = GetPlayerX(MyIndex) - 1
            Y = GetPlayerY(MyIndex)
        Case E_Direction.Right_
            X = GetPlayerX(MyIndex) + 1
            Y = GetPlayerY(MyIndex)
    End Select
   
    ' Check to see if the map tile is blocked or not
    If Map.Tile(X, Y).Type = Tile_Type.Blocked_ Then
        CheckDirection = True
        Exit Function
    End If
    
    ' Check to see if the map tile is a sign or not
    If Map.Tile(X, Y).Type = Tile_Type.Sign_ Then
        CheckDirection = True
        Exit Function
    End If
    
    ' Check to see if a sign is open or not
    If frmMainGame.picSign.Visible Or frmMainGame.picSign.Visible Then
        CheckDirection = True
        Exit Function
    End If
    
    ' Check to see if the key door is open or not
    If Map.Tile(X, Y).Type = Tile_Type.Key_ Then
        ' This actually checks if its open or not
        If TempTile(X, Y).DoorOpen = NO Then
            CheckDirection = True
            Exit Function
        End If
    End If
    
    ' Check to see if a player is already on that tile
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            If GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                If GetPlayerX(i) = X Then
                    If GetPlayerY(i) = Y Then
                        CheckDirection = True
                        Exit Function
                    End If
                End If
            End If
        End If
    Next
    
    ' Check to see if a npc is already on that tile
    For i = 1 To MAX_MAP_NPCS
        If MapNpc(i).Num > 0 Then
            If MapNpc(i).X = X Then
                If MapNpc(i).Y = Y Then
                    CheckDirection = True
                    Exit Function
                End If
            End If
        End If
    Next
    
End Function

Public Sub CheckMovement()
    If IsTryingToMove Then
        If CanMove Then
        
            ' Check if player has the shift key down for running
            If ShiftDown Then
                Player(MyIndex).Moving = MovementType.Running
            Else
                Player(MyIndex).Moving = MovementType.Walking
            End If
        
            Select Case GetPlayerDir(MyIndex)
                Case E_Direction.Up_
                    SendPlayerMove
                    Player(MyIndex).YOffset = PIC_Y
                    SetPlayerY MyIndex, GetPlayerY(MyIndex) - 1
            
                Case E_Direction.Down_
                    SendPlayerMove
                    Player(MyIndex).YOffset = -PIC_Y
                    SetPlayerY MyIndex, GetPlayerY(MyIndex) + 1
            
                Case E_Direction.Left_
                    SendPlayerMove
                    Player(MyIndex).XOffset = PIC_X
                    SetPlayerX MyIndex, GetPlayerX(MyIndex) - 1
            
                Case E_Direction.Right_
                    SendPlayerMove
                    Player(MyIndex).XOffset = -PIC_X
                    SetPlayerX MyIndex, GetPlayerX(MyIndex) + 1
            End Select
            
            If Player(MyIndex).XOffset = 0 Then
                If Player(MyIndex).YOffset = 0 Then
                    If Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex)).Type = Tile_Type.Warp_ Then
                        GettingMap = True
                    End If
                End If
            End If
            
        End If
    End If
End Sub

Public Sub UpdateInventory()
Dim i As Long

    For i = 1 To Equipment.Equipment_Count - 1
        SetPlayerEquipmentSlot MyIndex, GetPlayerEquipmentSlot(MyIndex, i), i
    Next
    
    DrawInventoryList
    
End Sub

Public Sub PlayerSearch()
    If isInBounds Then SendData CSearch & SEP_CHAR & CurX & SEP_CHAR & CurY & END_CHAR
End Sub

Public Function isInBounds()
    isInBounds = ((CurX >= 0) And (CurX <= MAX_MAPX) And (CurY >= 0) And (CurY <= MAX_MAPY))
End Function

Public Sub MapEditorSetSpawn()
Dim X As Long
Dim Y As Long

    X = CurX
    Y = CurY
    
    If Map.Tile(X, Y).Type <> Tile_Type.None_ Then
        AddText "The tile must not have an attribute!", Color.BrightRed
        Exit Sub
    End If
    
    If X > -1 And X <= MAX_MAPX And Y > -1 And Y <= MAX_MAPY Then
        MapSpawnY(frmMapProperties.lstUseNpcs.ListIndex + 1) = Y
        MapSpawnX(frmMapProperties.lstUseNpcs.ListIndex + 1) = X
    Else
        MapSpawnY(frmMapProperties.lstUseNpcs.ListIndex + 1) = -1
        MapSpawnX(frmMapProperties.lstUseNpcs.ListIndex + 1) = -1
    End If
    
    If MapSpawnY(frmMapProperties.lstUseNpcs.ListIndex + 1) = -1 Then
        frmMapProperties.lblSpawnY.Caption = "Y: None"
    Else
        frmMapProperties.lblSpawnY.Caption = "Y: " & MapSpawnY(frmMapProperties.lstUseNpcs.ListIndex + 1)
    End If
    
    If MapSpawnX(frmMapProperties.lstUseNpcs.ListIndex + 1) = -1 Then
        frmMapProperties.lblSpawnX.Caption = "X: None"
    Else
        frmMapProperties.lblSpawnX.Caption = "X: " & MapSpawnX(frmMapProperties.lstUseNpcs.ListIndex + 1)
    End If
    
    SettingSpawn = False
    
    frmMainGame.Hide
    frmMapProperties.Visible = True
    
End Sub

Public Sub UpdateDrawMapName()

    DrawMapNameX = (((MAX_MAPX + 1) * PIC_X) / 2) - ((Len(Trim$(Map.Name)) * FONT_WIDTH) / 2)
    DrawMapNameY = 0
    
    Select Case Map.Moral
        Case MAP_MORAL_NONE
            DrawMapNameColor = ColorTable(Color.BrightRed)
            
        Case MAP_MORAL_SAFE
            DrawMapNameColor = ColorTable(Color.White)
            
        Case Else
            DrawMapNameColor = ColorTable(Color.White)
    End Select
    
End Sub

Private Sub CheckKeys()

    If GetForegroundWindow <> frmMainGame.hWnd Then Exit Sub
    If Not CanMoveNow Then Exit Sub
    
    ControlDown = GetAsyncKeyState(VK_CONTROL) < 0
    ShiftDown = GetAsyncKeyState(VK_SHIFT) < 0
    
    If Not ControlDown Then
        DirUp = GetAsyncKeyState(VK_UP) < 0
        DirDown = GetAsyncKeyState(VK_DOWN) < 0
        DirLeft = GetAsyncKeyState(VK_LEFT) < 0
        DirRight = GetAsyncKeyState(VK_RIGHT) < 0
    Else
        DirUp = False
        DirDown = False
        DirLeft = False
        DirRight = False
    End If
    
End Sub

Public Function IsItem(ByVal X As Single, ByVal Y As Single) As Long
Dim tempRec As RECT
Dim ItemNum As Long

    For ItemNum = 1 To MAX_INV
        If Player(MyIndex).Inv(ItemNum).Num > 0 Then
            tempRec = Get_RECT(ItemIconY + ((ItemOffsetY + PIC_Y) * ((ItemNum - 1) \ ItemsInRow)), ItemIconY + ((ItemOffsetX + PIC_X) * (((ItemNum - 1) Mod ItemsInRow))))
            
            If X >= tempRec.Left Then
                If X <= tempRec.Right Then
                    If Y >= tempRec.Top Then
                        If Y <= tempRec.Bottom Then
                            IsItem = ItemNum
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
    Next
    
End Function

Public Function IsShopItem(ByVal X As Single, ByVal Y As Single) As Long
Dim tempRec As RECT
Dim ItemNum As Long

    For ItemNum = 1 To MAX_TRADES
        If ShopTrade.TradeItem(ItemNum).GetItem > 0 Then
            tempRec = Get_RECT(ShopIconY + ((ShopOffsetY + PIC_Y) * ((ItemNum - 1) \ ShopIconsInRow)), ShopIconY + ((ShopOffsetX + PIC_X) * (((ItemNum - 1) Mod ShopIconsInRow))))
            
            If X >= tempRec.Left Then
                If X <= tempRec.Right Then
                    If Y >= tempRec.Top Then
                        If Y <= tempRec.Bottom Then
                            IsShopItem = ItemNum
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
    Next
    
End Function

Public Sub SendACPAction(ByVal ActionID As ACP_Action)
Dim Packet As String

    If GetPlayerAccess(MyIndex) < StaffType.Monitor Then
        AddText "You need to be a high enough staff member to do this (" & StaffType.Mapper & ")!", AlertColor
        Exit Sub
    End If
    
    Packet = CACPAction & SEP_CHAR
    
    Select Case ActionID
    
        Case ACP_Action.LevelSelf
            Packet = Packet & ActionID
            
        Case ACP_Action.LevelTarget
            If LenB(Trim$(frmAdmin.txtTargetPlayer.Text)) < 1 Then
                AddText "You don't have a target player's name typed in the box!", AlertColor
                Exit Sub
            End If
            Packet = Packet & ActionID & SEP_CHAR & Trim$(frmAdmin.txtTargetPlayer.Text)
            
        Case ACP_Action.SetTargetSprite
            If LenB(Trim$(frmAdmin.txtTargetPlayer.Text)) < 1 Then
                AddText "You don't have a target player's name typed in the box!", AlertColor
                Exit Sub
            End If
            Packet = Packet & ActionID & SEP_CHAR & Trim$(frmAdmin.txtTargetPlayer.Text) & SEP_CHAR & frmAdmin.scrlSpriteNum.Value
            
        Case ACP_Action.CheckAccount
            If LenB(Trim$(frmAdmin.txtAccount.Text)) < 1 Then
                AddText "You don't have a target account's name typed in the box!", AlertColor
                Exit Sub
            End If
            Packet = Packet & ActionID & SEP_CHAR & Trim$(frmAdmin.txtAccount.Text)
            
        Case ACP_Action.CheckInventory
            If LenB(Trim$(frmAdmin.txtTargetPlayer.Text)) < 1 Then
                AddText "You don't have a target player's name typed in the box!", AlertColor
                Exit Sub
            End If
            Packet = Packet & ActionID & SEP_CHAR & Trim$(frmAdmin.txtTargetPlayer.Text)
            
        Case ACP_Action.GiveSelfPK
            Packet = Packet & ActionID
            
        Case ACP_Action.GiveTargetPK
            If LenB(Trim$(frmAdmin.txtTargetPlayer.Text)) < 1 Then
                AddText "You don't have a target player's name typed in the box!", AlertColor
                Exit Sub
            End If
            Packet = Packet & ActionID & SEP_CHAR & Trim$(frmAdmin.txtTargetPlayer.Text)
            
        Case ACP_Action.MutePlayer
            If LenB(Trim$(frmAdmin.txtTargetPlayer.Text)) < 1 Then
                AddText "You don't have a target player's name typed in the box!", AlertColor
                Exit Sub
            End If
            Packet = Packet & ActionID & SEP_CHAR & Trim$(frmAdmin.txtTargetPlayer.Text) & SEP_CHAR & frmAdmin.scrlTime.Value
            
    End Select
    
    SendData Packet & END_CHAR
    
End Sub

Public Sub ErrorReport(ByVal ErrorTag As String)

    AddLog ErrorTag, "\error.txt"
    AddText "An error has occured. Please send your error.txt to an administrator!", Color.BrightRed
    
End Sub

Public Sub LoadWindows()

    Set Windows(Window_State.Main_Menu) = frmMainMenu
    Set Windows(Window_State.Login) = frmMainMenu
    Set Windows(Window_State.New_Account) = frmMainMenu
    Set Windows(Window_State.New_Char) = frmNewChar
    Set Windows(Window_State.Chars) = frmChars
    Set Windows(Window_State.Main_Game) = frmMainGame
    Set Windows(Window_State.Credits) = frmMainMenu
    
End Sub

Public Sub ClearMapAttribs()
Dim LoopI As Long

    MapAttribType = Tile_Type.None_
    
    For LoopI = 1 To 3
        MapAttribName(LoopI) = vbNullString
        MapAttribMin(LoopI) = 0
        MapAttribMax(LoopI) = 0
        'ZeroMemory ByVal VarPtr(MapAttribName(LoopI)), LenB(MapAttribName(LoopI))
        'ZeroMemory ByVal VarPtr(MapAttribMin(LoopI)), LenB(MapAttribMin(LoopI))
        'ZeroMemory ByVal VarPtr(MapAttribMax(LoopI)), LenB(MapAttribMax(LoopI))
    Next
    
End Sub

Public Sub SetMapAttrib(ByVal Index As Long, ByVal Caption As String, ByVal Min As Long, ByVal Max As Long)

    MapAttribName(Index) = Caption
    MapAttribMin(Index) = Min
    MapAttribMax(Index) = Max
    
End Sub

Public Sub Build_Lookups()
Dim LoopI As Long
Dim LoopI2 As Long

    For LoopI = LBound(MultiplyPicX) To UBound(MultiplyPicX)
        MultiplyPicX(LoopI) = LoopI * 32
    Next
    
    For LoopI = LBound(ColorTable) To UBound(ColorTable)
        ColorTable(LoopI) = QBColor(LoopI)
    Next
    
    For LoopI = 1 To MAX_INTEGER
        For LoopI2 = 1 To MAX_BYTE
            ModularTable(LoopI, LoopI2) = LoopI Mod LoopI2
        Next
    Next
    
End Sub

Public Sub UpdateItemDescription(ByVal ItemNum As Long, ByVal ItemVal As Long)
Dim BuildString As String
Dim PotionString As String

    Select Case Item(ItemNum).Type
        Case ItemType.Currency_
        
            BuildString = "Amount: " & FormatNumber2(ItemVal) & vbNewLine
            
        Case ItemType.Key
        
            BuildString = "[ this is a key ]" & vbNewLine
            
        Case ItemType.Spell_
        
            BuildString = "This item will teach you:" & vbNewLine
            
            If Item(ItemNum).Data1 > 0 Then
                BuildString = BuildString & "Spell: " & Trim$(Spell(Item(ItemNum).Data1).Name) & vbNewLine
            Else
                BuildString = BuildString & "Nothing" & vbNewLine
            End If
            
            BuildString = BuildString & vbNewLine
            
            BuildString = BuildString & "[ Requires ]" & vbNewLine
            If Item(ItemNum).Required(Item_Requires.Access_) > 0 Then BuildString = BuildString & Item(ItemNum).Required(Item_Requires.Access_) & " Access" & vbNewLine
            
            If Item(ItemNum).Required(Item_Requires.Class_) > 0 Then BuildString = BuildString & "Class: " & Trim$(Class(Item(ItemNum).Required(Item_Requires.Class_)).Name) & vbNewLine
            
            If Item(ItemNum).Required(Item_Requires.Level_) > 1 Then BuildString = BuildString & "Level: " & Item(ItemNum).Required(Item_Requires.Level_) & vbNewLine
            If Item(ItemNum).Required(Item_Requires.Strength_) > 0 Then BuildString = BuildString & "Strength: " & Item(ItemNum).Required(Item_Requires.Strength_) & vbNewLine
            If Item(ItemNum).Required(Item_Requires.Defense_) > 0 Then BuildString = BuildString & "Defense: " & Item(ItemNum).Required(Item_Requires.Defense_) & vbNewLine
            If Item(ItemNum).Required(Item_Requires.Speed_) > 0 Then BuildString = BuildString & "Speed: " & Item(ItemNum).Required(Item_Requires.Speed_) & vbNewLine
            If Item(ItemNum).Required(Item_Requires.Magic_) > 0 Then BuildString = BuildString & "Magic: " & Item(ItemNum).Required(Item_Requires.Magic_) & vbNewLine
            
            If Item(ItemNum).Data1 > 0 Then
                If BuildString = "This item will teach you:" & vbNewLine & _
                                 "Spell: " & Trim$(Spell(Item(ItemNum).Data1).Name) & vbNewLine & vbNewLine & _
                                 "[ Requires ]" & vbNewLine Then BuildString = BuildString & "Nothing" & vbNewLine
            Else
                If BuildString = "This item will teach you:" & vbNewLine & _
                                 "Nothing" & vbNewLine & vbNewLine & _
                                 "[ Requires ]" & vbNewLine Then BuildString = BuildString & "Nothing" & vbNewLine
            End If
            
        Case Is < ItemType.Shield_ + 1
        
            If Item(ItemNum).Type > 0 Then
                BuildString = "[ Requires ]" & vbNewLine
                If Item(ItemNum).Required(Item_Requires.Access_) > 0 Then BuildString = BuildString & Item(ItemNum).Required(Item_Requires.Access_) & " Access" & vbNewLine
                
                If Item(ItemNum).Required(Item_Requires.Class_) > 0 Then BuildString = BuildString & "Class: " & Trim$(Class(Item(ItemNum).Required(Item_Requires.Class_)).Name) & vbNewLine
                
                If Item(ItemNum).Required(Item_Requires.Level_) > 1 Then BuildString = BuildString & "Level: " & Item(ItemNum).Required(Item_Requires.Level_) & vbNewLine
                If Item(ItemNum).Required(Item_Requires.Strength_) > 0 Then BuildString = BuildString & "Strength: " & Item(ItemNum).Required(Item_Requires.Strength_) & vbNewLine
                If Item(ItemNum).Required(Item_Requires.Defense_) > 0 Then BuildString = BuildString & "Defense: " & Item(ItemNum).Required(Item_Requires.Defense_) & vbNewLine
                If Item(ItemNum).Required(Item_Requires.Speed_) > 0 Then BuildString = BuildString & "Speed: " & Item(ItemNum).Required(Item_Requires.Speed_) & vbNewLine
                If Item(ItemNum).Required(Item_Requires.Magic_) > 0 Then BuildString = BuildString & "Magic: " & Item(ItemNum).Required(Item_Requires.Magic_) & vbNewLine
                
                If BuildString = "[ Requires ]" & vbNewLine Then BuildString = BuildString & "Nothing" & vbNewLine
                
                BuildString = BuildString & vbNewLine
                
                BuildString = BuildString & "[ Bonuses ]" & vbNewLine
                If Item(ItemNum).BuffStats(Stats.Strength) > 0 Then BuildString = BuildString & "Strength +" & Item(ItemNum).BuffStats(Stats.Strength) & vbNewLine
                If Item(ItemNum).BuffStats(Stats.Defense) > 0 Then BuildString = BuildString & "Defense +" & Item(ItemNum).BuffStats(Stats.Defense) & vbNewLine
                If Item(ItemNum).BuffStats(Stats.Speed) > 0 Then BuildString = BuildString & "Speed +" & Item(ItemNum).BuffStats(Stats.Speed) & vbNewLine
                If Item(ItemNum).BuffStats(Stats.Magic) > 0 Then BuildString = BuildString & "Magic +" & Item(ItemNum).BuffStats(Stats.Magic) & vbNewLine
                If Item(ItemNum).BuffVitals(Vitals.HP) > 0 Then BuildString = BuildString & "HP +" & Item(ItemNum).BuffVitals(Vitals.HP) & vbNewLine
                If Item(ItemNum).BuffVitals(Vitals.MP) > 0 Then BuildString = BuildString & "MP +" & Item(ItemNum).BuffVitals(Vitals.MP) & vbNewLine
                If Item(ItemNum).BuffVitals(Vitals.SP) > 0 Then BuildString = BuildString & "SP +" & Item(ItemNum).BuffVitals(Vitals.SP) & vbNewLine
            End If
            
        Case Else
        
            If Item(ItemNum).Type = ItemType.Potion Then
                BuildString = "[ Bonuses ]" & vbNewLine
                If Item(ItemNum).Data1 <> 0 Then
                    If Item(ItemNum).Data1 > 0 Then
                        BuildString = BuildString & "HP +" & Item(ItemNum).Data1 & vbNewLine
                    Else
                        BuildString = BuildString & "HP " & Item(ItemNum).Data1 & vbNewLine
                    End If
                End If
                
                If Item(ItemNum).Data2 <> 0 Then
                    If Item(ItemNum).Data2 > 0 Then
                        BuildString = BuildString & "MP +" & Item(ItemNum).Data2 & vbNewLine
                    Else
                        BuildString = BuildString & "MP " & Item(ItemNum).Data2 & vbNewLine
                    End If
                End If
                
                If Item(ItemNum).Data3 <> 0 Then
                    If Item(ItemNum).Data3 > 0 Then
                        BuildString = BuildString & "SP +" & Item(ItemNum).Data3 & vbNewLine
                    Else
                        BuildString = BuildString & "SP " & Item(ItemNum).Data3 & vbNewLine
                    End If
                End If
                
                BuildString = BuildString & vbNewLine
                
                BuildString = BuildString & "[ Requires ]" & vbNewLine
                If Item(ItemNum).Required(Item_Requires.Access_) > 0 Then PotionString = Item(ItemNum).Required(Item_Requires.Access_) & " Access" & vbNewLine
                
                If Item(ItemNum).Required(Item_Requires.Class_) > 0 Then PotionString = PotionString & "Class: " & Trim$(Class(Item(ItemNum).Required(Item_Requires.Class_)).Name) & vbNewLine
                
                If Item(ItemNum).Required(Item_Requires.Level_) > 1 Then PotionString = PotionString & "Level: " & Item(ItemNum).Required(Item_Requires.Level_) & vbNewLine
                If Item(ItemNum).Required(Item_Requires.Strength_) > 0 Then PotionString = PotionString & "Strength: " & Item(ItemNum).Required(Item_Requires.Strength_) & vbNewLine
                If Item(ItemNum).Required(Item_Requires.Defense_) > 0 Then PotionString = PotionString & "Defense: " & Item(ItemNum).Required(Item_Requires.Defense_) & vbNewLine
                If Item(ItemNum).Required(Item_Requires.Speed_) > 0 Then PotionString = PotionString & "Speed: " & Item(ItemNum).Required(Item_Requires.Speed_) & vbNewLine
                If Item(ItemNum).Required(Item_Requires.Magic_) > 0 Then PotionString = PotionString & "Magic: " & Item(ItemNum).Required(Item_Requires.Magic_) & vbNewLine
                
                If LenB(PotionString) < 1 Then
                    BuildString = BuildString & "Nothing" & vbNewLine
                Else
                    BuildString = BuildString & PotionString & vbNewLine
                End If
                
            End If
            
    End Select
    
    If LenB(BuildString) < 1 Then BuildString = "None" & vbNewLine
    
    If Item(ItemNum).Type <> ItemType.Currency_ Then
        If Item(ItemNum).CostItem > 0 And Item(ItemNum).CostAmount > 0 Then
            BuildString = "[ Worth: " & Item(ItemNum).CostAmount & " " & Trim$(Item(Item(ItemNum).CostItem).Name) & " ]" & vbNewLine & vbNewLine & BuildString
        Else
            BuildString = "[ Worth: Nothing ]" & vbNewLine & vbNewLine & BuildString
        End If
    End If
    
    With frmMainGame
        .lblDescription.Caption = BuildString
        .picItemDesc.Height = (.lblDescription.Top + .lblDescription.Height) + .picItemDescBottom.Height
        .picItemDescBottom.Top = .picItemDesc.Height - .picItemDescBottom.Height
    End With
    
End Sub

Public Sub UpdateSpellDescription(ByVal SpellNum As Long)
Dim BuildString As String
Dim AOEString As String

    If Spell(SpellNum).AOE = 1 Then AOEString = "*Area of Effect*" & vbNewLine & vbNewLine
    
    If Spell(SpellNum).Timer > 0 Then
        BuildString = BuildString & vbNewLine & Round(Spell(SpellNum).Timer / 1000, 1) & " second cool-down"
    Else
        BuildString = BuildString & vbNewLine & "No cool-down"
    End If
    
    Select Case Spell(SpellNum).Type
        Case Spell_Type.AddHP
            BuildString = BuildString & "Power +" & Spell(SpellNum).Data1 & " HP"
        Case Spell_Type.AddMP
            BuildString = BuildString & "Power +" & Spell(SpellNum).Data1 & " MP"
        Case Spell_Type.AddSP
            BuildString = BuildString & "Power +" & Spell(SpellNum).Data1 & " SP"
        Case Spell_Type.GiveItem
            BuildString = BuildString & "Gives Item: " & Trim$(Item(Spell(SpellNum).Data1).Name)
        Case Spell_Type.SubHP
            BuildString = BuildString & "Power -" & Spell(SpellNum).Data1 & " HP"
        Case Spell_Type.SubMP
            BuildString = BuildString & "Power -" & Spell(SpellNum).Data1 & " MP"
        Case Spell_Type.SubSP
            BuildString = BuildString & "Power -" & Spell(SpellNum).Data1 & " SP"
        Case Else
            GoTo Skipppy
    End Select
    
    If Spell(SpellNum).Range > 0 Then
        BuildString = BuildString & vbNewLine & "Range: " & Spell(SpellNum).Range & " tiles"
    Else
        BuildString = BuildString & vbNewLine & "Range: Full Map"
    End If
    
    If LenB(BuildString) > 0 Then
        If Spell(SpellNum).MPReq > 0 Then
            BuildString = AOEString & "[Requires]" & vbNewLine & Spell(SpellNum).MPReq & " MP" & vbNewLine & vbNewLine & "[Misc Info]" & vbNewLine & BuildString
        Else
            BuildString = AOEString & "[Requires]" & vbNewLine & "Nothing" & vbNewLine & vbNewLine & "[Misc Info]" & vbNewLine & BuildString
        End If
    End If
    
Skipppy:
    
    If LenB(BuildString) < 1 Then BuildString = "None"
    
    BuildString = BuildString & vbNewLine
    
    With frmMainGame
        .lblSpell.Caption = Trim$(Spell(SpellNum).Name)
        .lblSpellDescription.Caption = BuildString
        .picSpellDesc.Height = (.lblSpellDescription.Top + .lblSpellDescription.Height) + .picItemDescBottom.Height
        .picSpellDescBottom.Top = .picSpellDesc.Height - .picSpellDescBottom.Height
    End With
    
End Sub

Public Sub HandleNews(ByVal NewsText As String)

    frmMainMenu.txtNews.Text = vbNullString
    
    If LenB(Trim$(NewsText)) < 1 Then NewsText = "[center]No news!"
    
    NewsText = Replace$(NewsText, "[br]", vbNewLine, , , vbTextCompare)
    NewsText = Replace$(NewsText, "[tab]", Space$(2), , , vbTextCompare)
    
    If InStr(1, NewsText, "[center]", vbTextCompare) Then
        frmMainMenu.txtNews.SelStart = 0
        frmMainMenu.txtNews.SelAlignment = vbCenter
        NewsText = Replace$(NewsText, "[center]", vbNullString, , , vbTextCompare)
        frmMainMenu.txtNews.SelLength = Len(NewsText)
    End If
    
    frmMainMenu.txtNews.SelText = NewsText
    
End Sub
