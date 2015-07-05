Attribute VB_Name = "modGameLogic"
Option Explicit

Sub Main()
Dim i As Long
        
    Randomize Timer
    
    ' Check if the maps directory is there, if its not make it
    If LCase$(Dir$(App.Path & "\maps", vbDirectory)) <> "maps" Then
        Call MkDir(App.Path & "\maps")
    End If
    
    LoadConfiguration
    
    ' Make sure we set that we aren't in the game
    InGame = False
    GettingMap = True
    InEditor = False
    InItemsEditor = False
    InNpcEditor = False
    InShopEditor = False
    InEmoticonEditor = False
    
    ' Now we set all our UDT sizes
    AnimationSize = LenB(Animation(1))
    EmoticonSize = LenB(Emoticons(i))
    ItemSize = LenB(Item(1))
    NpcSize = LenB(Npc(1))
    ShopSize = LenB(Shop(1))
    SpellSize = LenB(Spell(1))
    QuestSize = LenB(Quest(1))
    
    ' Clear out players
    For i = 1 To MAX_PLAYERS
        Call ClearPlayer(i)
    Next
    
    ' Set the size of picScreen
    frmMainGame.picScreen.Width = ScreenX
    frmMainGame.picScreen.Height = ScreenY
    
    InitTcp
    InitDirectX
    
    frmMainMenu.Visible = True
    frmSendGetData.Visible = False
End Sub

Public Sub LoadConfiguration()

    '****************
    '** Login Info **
    '****************
    frmLogin.chkRemember.Value = Val(ReadIniValue(App.Path & "\Core Files\Configuration.ini", "Account Information", "Remember"))
    frmLogin.txtName.Text = ReadIniValue(App.Path & "\Core Files\Configuration.ini", "Account Information", "Account")
    frmLogin.txtPassword.Text = ReadIniValue(App.Path & "\Core Files\Configuration.ini", "Account Information", "Password")
    
    ShowItemLinks = Val(ReadIniValue(App.Path & "\Core Files\Configuration.ini", "Config", "ShowItemLinks"))
End Sub

Public Sub SaveConfiguration()
    '****************
    '** Login Info **
    '****************
    If frmLogin.chkRemember.Value Then
        WriteIniValue App.Path & "\Core Files\Configuration.ini", "Account Information", "Remember", "1"
        WriteIniValue App.Path & "\Core Files\Configuration.ini", "Account Information", "Account", frmLogin.txtName.Text
        WriteIniValue App.Path & "\Core Files\Configuration.ini", "Account Information", "Password", frmLogin.txtPassword.Text
    Else
        WriteIniValue App.Path & "\Core Files\Configuration.ini", "Account Information", "Remember", "0"
        WriteIniValue App.Path & "\Core Files\Configuration.ini", "Account Information", "Account", ""
        WriteIniValue App.Path & "\Core Files\Configuration.ini", "Account Information", "Password", ""
    End If
End Sub

Sub SetStatus(ByVal Caption As String)
    frmSendGetData.lblStatus.Caption = Caption
    DoEvents
End Sub

Sub MenuState(ByVal State As Long)
    
    frmEvent.Visible = False
    frmSendGetData.Visible = True
    Call SetStatus("Loading...")
    
    Select Case CurrentState
        Case MenuStates.MainMenu
            frmMainMenu.Visible = True
            frmEvent.Visible = False
            
        Case MenuStates.Login
            frmLogin.Visible = True
            frmSendGetData.Visible = False
            
        Case MenuStates.NewAccount
            frmNewAccount.Visible = True
            frmEvent.Visible = False
        
        Case MenuStates.NewChar
            frmNewChar.Visible = True
            frmEvent.Visible = False
            
        Case MenuStates.Chars
            frmChars.Visible = True
            frmEvent.Visible = False
           
        Case MenuStates.Shutdown
            InGame = False
            TcpDestroy
            GameDestroy
            
    End Select

    If Not IsConnected Then
        'frmMainMenu.Visible = True
        frmSendGetData.Visible = False
        frmEvent.Visible = True
        frmEvent.lblInformation.Caption = "Connection Error. (Err: #1)"
    End If
End Sub

Sub GameInit()
    frmMainGame.Visible = True
    frmSendGetData.Visible = False
    'InitDirectX
End Sub

Sub GameLoop()
Dim startTick As Long
Dim TickFPS As Long
Dim FPS As Long
Dim X As Long
Dim Y As Long
Dim i As Long
Dim rec As RECT
Dim rec_pos As RECT
Dim rec_clear As RECT

    frmMainGame.picScreen.SetFocus      ' Set the focus

    ' TODO : GET RID OF THIS ?
    UpdateInventory                     ' Update the inventory for the first time
    UpdateEquipment                     ' Update the equipment for the first time

    With rec
        .Bottom = (MAX_MAPY + 3) * PIC_Y
        .Right = (MAX_MAPX + 3) * PIC_X
    End With
    rec_clear = rec
    
    With rec_pos
        .Bottom = ((MAX_MAPY + 1) * PIC_Y)
        .Right = ((MAX_MAPX + 1) * PIC_X)
    End With
    
    TickFPS = GetTickCount

    Do While InGame
        startTick = GetTickCount

        ' Check to make sure we are still connected
        InGame = IsConnected

        ' Check to make sure they aren't trying to auto do anything
        If GetAsyncKeyState(VK_UP) >= 0 And DirUp = True Then DirUp = False
        If GetAsyncKeyState(VK_DOWN) >= 0 And DirDown = True Then DirDown = False
        If GetAsyncKeyState(VK_LEFT) >= 0 And DirLeft = True Then DirLeft = False
        If GetAsyncKeyState(VK_RIGHT) >= 0 And DirRight = True Then DirRight = False
        If GetAsyncKeyState(VK_CONTROL) >= 0 And ControlDown = True Then ControlDown = False
        If GetAsyncKeyState(VK_SHIFT) >= 0 And ShiftDown = True Then ShiftDown = False
        If GetAsyncKeyState(VK_ALT) >= 0 And AltDown = True Then AltDown = False

        ' Check if player is trying to move
        Call CheckMovement

        ' Check to see if player is trying to attack
        Call CheckAttack

        ' Process player movements (actually move them)
        For i = 1 To MapPlayersCount
            If Player(MapPlayers(i)).Moving > 0 Then
                ProcessMovement MapPlayers(i)
            End If
        Next

        ' Process npc movements (actually move them)
        For i = 1 To MapNpcCount
            If MapNpc(i).Num > 0 Then
                ProcessNpcMovement i
            End If
        Next

        ' Change map animation every 250 milliseconds
        If startTick > MapAnimTimer Then
            MapAnim = Not MapAnim
            MapAnimTimer = GetTickCount + 250
        End If

'        ////////////////
'        //  Graphics  //
'        ////////////////

        If NeedToRestoreSurfaces Then                               ' Check if we need to restore surfaces
            InitSurfaces
            DD.RestoreAllSurfaces
        End If
        
        DD_BackBuffer.BltColorFill rec_clear, 0                           ' Clear out the backbuffer
                
        If Not GettingMap Then
            UpdateCamera
            BltTiles
            BltItems
            BltAnimations 0
            BltPlayers
            BltNpcs
            BltAnimations 1
            BltFringeTiles
            BltTargets
            BltCastBar
        End If
        
        ' Lock the backbuffer so we can draw text and names
        TexthDC = DD_BackBuffer.GetDC
            If Not GettingMap Then
                'Draw the names and target
                CheckNames
        
                ' Draw player names
                For i = MapPlayersCount To 1 Step -1
                    If InViewPort(Current_X(MapPlayers(i)), Current_Y(MapPlayers(i))) Then
                        BltPlayerName MapPlayers(i)
                    End If
                Next
        
                ' Draws ActionMsg Messages
                For i = 1 To MAX_BYTE
                    If ActionMsg(i).Message <> vbNullString Then
                        BltActionMsg i
                    End If
                Next
        
                ' Draw the attributes if in the map editor
                BltAttributes
        
                ' DRAW SPELLNAME IF CASTING
                If CastingSpell Then
                     DrawText TexthDC, Camera.Left + HalfX - (getTextWidth(Trim$(Spell(Player(MyIndex).Spell(CastingSpell).SpellNum).Name)) / 2), Camera.Top + 327, Trim$(Spell(Player(MyIndex).Spell(CastingSpell).SpellNum).Name), QBColor(White)
                End If
        
                ' Draw map name
                If LenB(Trim$(Map.Name)) > 0 Then
                    If Map.Moral = MAP_MORAL_NONE Then
                        Call DrawText(TexthDC, Camera.Left + HalfX - (getTextWidth(Trim$(Map.Name)) / 2), Camera.Top + 1, Trim$(Map.Name), QBColor(BrightRed))
                    Else
                        Call DrawText(TexthDC, Camera.Left + HalfX - (getTextWidth(Trim$(Map.Name)) / 2), Camera.Top + 1, Trim$(Map.Name), QBColor(White))
                    End If
                End If
                
                ' Draw FPS
                If ShowFPS Then DrawText TexthDC, Camera.Left + 5, Camera.Top + 5, GameFPS, QBColor(White)
                
                ' Draw death timer
                If Current_IsDead(MyIndex) Then
                    DrawText TexthDC, Camera.Right - getTextWidth(DeathTimer) - 5, Camera.Top + 5, DeathTimer, QBColor(White)
                End If
            Else
                Call DrawText(TexthDC, Camera.Left + 10, Camera.Top + 10, "Receiving Map...", QBColor(BrightCyan))
            End If
            
            'frmMainGame.txtMyTextBox.Text = MyText
        DD_BackBuffer.ReleaseDC TexthDC                             ' Release DC
        
        ' More graphical things that should go above text
        BltNpcIcons
        BltEmoticons
        
        With rec
            .Top = Camera.Top
            .Bottom = .Top + ScreenY
            .Left = Camera.Left
            .Right = .Left + ScreenX
        End With
        
        DX.GetWindowRect frmMainGame.picScreen.hwnd, rec_pos        ' Get the rect to blit to
        DD_PrimarySurf.Blt rec_pos, DD_BackBuffer, rec, DDBLT_WAIT  ' Blit the backbuffer
        
        ' Calculate fps
        If GetTickCount > TickFPS Then
            GameFPS = FPS
            FPS = 0
            TickFPS = GetTickCount + 1000
        End If
        FPS = FPS + 1
        
        MakeMidiLoop
        
        ' Lock fps
        If Not LockFPS Then
            Do While GetTickCount < startTick + 30
                DoEvents
                Sleep 1
            Loop
        Else
            DoEvents
        End If
    Loop

    frmMainGame.Visible = False
    If Not frmEvent.Visible Then
        frmSendGetData.Visible = True
        Call SetStatus("Destroying game data...")

        'Report disconnection if server disconnects
        If IsConnected = False Then
            frmEvent.Visible = True
            frmEvent.lblInformation.Caption = "The server has disconnected. (Err: #2)"
        End If
    End If

    'Shutdown the game
    GameDestroy
End Sub

Sub GameDestroy()
    
    Set PlayerBuffer = Nothing
    Set frmMainGame.Hypertext = Nothing
    
    StopMidi
    DestroyDirectX
    
    End
End Sub

Public Sub CheckNames()
Dim i As Long

    If InEditor = False Then
    
        If MyTargetType = TARGET_TYPE_NPC Then BltMapNPCName MyTarget
        
        For i = MapNpcCount To 1 Step -1
            If MapNpc(i).Num > 0 Then
                If MapNpc(i).X = CurX Then
                    If MapNpc(i).Y = CurY Then
                        BltMapNPCName i
                        Exit Sub
                    End If
                End If
            End If
        Next
        
        For i = MapPlayersCount To 1 Step -1
            If Player(MapPlayers(i)).X = CurX Then
                If Player(MapPlayers(i)).Y = CurY Then
                    Exit Sub
                End If
            End If
        Next
        
        For i = MAX_MAP_ITEMS To 1 Step -1
            If MapItem(i).Num > 0 Then
                If MapItem(i).X = CurX Then
                    If MapItem(i).Y = CurY Then
                        BltMapItemName i
                        Exit Sub
                    End If
                End If
            End If
        Next
    End If
End Sub

Sub ProcessMovement(ByVal Index As Long)
    
    ' Check if player is walking, and if so process moving them over
'    If Player(Index).Moving = MOVING_WALKING Then
        Select Case Current_Dir(Index)
            Case DIR_UP
                Player(Index).YOffset = Player(Index).YOffset - WALK_Speed
            Case DIR_DOWN
                Player(Index).YOffset = Player(Index).YOffset + WALK_Speed
            Case DIR_LEFT
                Player(Index).XOffset = Player(Index).XOffset - WALK_Speed
            Case DIR_RIGHT
                Player(Index).XOffset = Player(Index).XOffset + WALK_Speed
        End Select
        
        ' Check if completed walking over to the next tile
        If (Player(Index).XOffset = 0) Then
            If (Player(Index).YOffset = 0) Then
                Player(Index).Moving = 0
            End If
        End If
'    End If

    ' Check if player is running, and if so process moving them over
    'If Player(Index).Moving = MOVING_RUNNING Then
    '    Select Case Current_Dir(Index)
    '        Case DIR_UP
    '            Player(Index).YOffset = Player(Index).YOffset - RUN_Speed
    '        Case DIR_DOWN
    '            Player(Index).YOffset = Player(Index).YOffset + RUN_Speed
    '        Case DIR_LEFT
    '            Player(Index).XOffset = Player(Index).XOffset - RUN_Speed
    '        Case DIR_RIGHT
    '            Player(Index).XOffset = Player(Index).XOffset + RUN_Speed
    '    End Select
        
        ' Check if completed walking over to the next tile
'        If (Player(Index).XOffset = 0) And (Player(Index).YOffset = 0) Then
'            Player(Index).Moving = 0
'        End If
    'End If
End Sub

Sub ProcessNpcMovement(ByVal MapNpcNum As Long)

    ' Check which movement Speed the npc is
    Select Case MapNpc(MapNpcNum).Moving
        Case 1
            Select Case MapNpc(MapNpcNum).Dir
                Case DIR_UP
                    MapNpc(MapNpcNum).YOffset = MapNpc(MapNpcNum).YOffset - NPC_Speed
                Case DIR_DOWN
                    MapNpc(MapNpcNum).YOffset = MapNpc(MapNpcNum).YOffset + NPC_Speed
                Case DIR_LEFT
                    MapNpc(MapNpcNum).XOffset = MapNpc(MapNpcNum).XOffset - NPC_Speed
                Case DIR_RIGHT
                    MapNpc(MapNpcNum).XOffset = MapNpc(MapNpcNum).XOffset + NPC_Speed
            End Select

            ' Check if completed walking over to the next tile
            If (MapNpc(MapNpcNum).XOffset = 0) And (MapNpc(MapNpcNum).YOffset = 0) Then
                MapNpc(MapNpcNum).Moving = 0
            End If

        Case 2
            Select Case MapNpc(MapNpcNum).Dir
                Case DIR_UP
                    MapNpc(MapNpcNum).YOffset = MapNpc(MapNpcNum).YOffset - NPC_Speed_FAST
                Case DIR_DOWN
                    MapNpc(MapNpcNum).YOffset = MapNpc(MapNpcNum).YOffset + NPC_Speed_FAST
                Case DIR_LEFT
                    MapNpc(MapNpcNum).XOffset = MapNpc(MapNpcNum).XOffset - NPC_Speed_FAST
                Case DIR_RIGHT
                    MapNpc(MapNpcNum).XOffset = MapNpc(MapNpcNum).XOffset + NPC_Speed_FAST
            End Select

            ' Check if completed walking over to the next tile
            If (MapNpc(MapNpcNum).XOffset = 0) And (MapNpc(MapNpcNum).YOffset = 0) Then
                MapNpc(MapNpcNum).Moving = 0
            End If

        Case 3
            Select Case MapNpc(MapNpcNum).Dir
                Case DIR_UP
                    MapNpc(MapNpcNum).YOffset = MapNpc(MapNpcNum).YOffset - NPC_Speed_FASTEST
                Case DIR_DOWN
                    MapNpc(MapNpcNum).YOffset = MapNpc(MapNpcNum).YOffset + NPC_Speed_FASTEST
                Case DIR_LEFT
                    MapNpc(MapNpcNum).XOffset = MapNpc(MapNpcNum).XOffset - NPC_Speed_FASTEST
                Case DIR_RIGHT
                    MapNpc(MapNpcNum).XOffset = MapNpc(MapNpcNum).XOffset + NPC_Speed_FASTEST
            End Select

            ' Check if completed walking over to the next tile
            If (MapNpc(MapNpcNum).XOffset = 0) And (MapNpc(MapNpcNum).YOffset = 0) Then
                MapNpc(MapNpcNum).Moving = 0
            End If
    End Select
End Sub

Sub HandleKeypresses(ByVal KeyAscii As Integer)
Dim ChatText As String
Dim Name As String
Dim i As Long
Dim n As Long
Dim Command() As String
    
    ' Handle when the player presses the return key
    If KeyAscii = vbKeyReturn Then
        If Left$(MyText, 1) = "/" Then
            Command = Split(Trim$(MyText), " ")
        
            Select Case LCase$(Command(0))
                Case "/help"
                    AddText "Communication hotkeys:", AlertColor
                    AddText "F1 = Set chat: Immediate chat (current map).", AlertColor
                    AddText "F2 = Set chat: Realm-wide chat (everyone).", AlertColor
                    AddText "F3 = Set chat: Guild chat.", AlertColor
                    AddText "F4 = Set chat: Party chat.", AlertColor
                    AddText "/s 'Message' or /say 'Message' = Immediate chat (current map).", AlertColor
                    AddText "/r 'Message' or /realm 'Message' = Realm-wide chat (everyone).", AlertColor
                    AddText "/w 'player' 'Message' or /whisper 'player' 'Message' = Whisper to adventurer.", AlertColor
                    AddText "/g 'Message' or /guild 'Message' = Guild chat.", AlertColor
                    AddText "/p 'Message' or /party 'Message' = Guild chat.", AlertColor
                    
                    AddText vbNullString, AlertColor
                    AddText "Other:", AlertColor
                    AddText "/em 'Message' or /emote 'Message' = Emote Message (current map).", AlertColor
                    
                    AddText vbNullString, AlertColor
                    AddText "Available Commands: /help, /fix, /clear, /online, /fps, /stats, /emotes.", AlertColor
                    
                    AddText vbNullString, AlertColor
                    AddText "Party Commands: /invite 'adventurer', /join, /leave.", AlertColor
                    
                    AddText vbNullString, AlertColor
                    AddText "Guild Commands: /gcreate '3 Chars(Guild Abbreviation)' '20 Chars(Guild Name)', /gmotd 'Message', /gquit, /gdelete, /gkick 'guild member', /gpromote 'guild member', /gdemote 'guild member', /ginvite 'adventurer', /gjoin, /gdecline.", AlertColor
                    
                    If Current_Access(MyIndex) >= ADMIN_MONITOR Then
                        AddText vbNullString, AlertColor
                        AddText "Monitor Commands:", AlertColor
                        AddText "/admin (Admin Message) or /a (Admin Message)", AlertColor
                        AddText "/kick 'adventurer', /kill 'adventurer', /info 'adventurer'", AlertColor
                        
                        If Current_Access(MyIndex) >= ADMIN_MAPPER Then
                            AddText vbNullString, AlertColor
                            AddText "Mapper Commands:", AlertColor
                            AddText "/loc, /emoticoneditor, /mapeditor, /warptome, /warpmeto, /warpto, /setsprite 'adventurer' '#', /mapreport, /respawn, /motd, /banlist, /ban", AlertColor
                            
                            If Current_Access(MyIndex) >= ADMIN_DEVELOPER Then
                                AddText vbNullString, AlertColor
                                AddText "Developer Commands:", AlertColor
                                AddText "/edititem, /editnpc, /editshop, /editspell, /editanim, /editquest", AlertColor
                                
                                If Current_Access(MyIndex) >= ADMIN_CREATOR Then
                                    AddText vbNullString, AlertColor
                                    AddText "Owner Commands:", AlertColor
                                    AddText "/setaccess, /destroybanlist", AlertColor
                                End If
                            End If
                        End If
                    End If
                    
                Case "/inv", "/inventory"
                    InvVisible = Not InvVisible
                    
                Case "/s", "/say"
                    If UBound(Command) >= 1 Then
                        SayMsg Right$(MyText, Len(MyText) - Len(Command(0)) - 1)
                    Else
                        AddText "Usage: /s 'Message' or /say 'Message'", AlertColor
                    End If
                    
                Case "/r", "/realm"
                    If UBound(Command) >= 1 Then
                        RealmMsg Right$(MyText, Len(MyText) - Len(Command(0)) - 1)
                    Else
                        AddText "Usage: /r 'Message' or /realm 'Message'", AlertColor
                    End If
                    
                Case "/w", "/whisper"
                    If UBound(Command) >= 2 Then
                        If IsAlpha(Command(1)) Then
                            PlayerMsg Right$(MyText, Len(MyText) - (Len(Command(0)) + Len(Command(1))) - 1), Command(1)
                        Else
                            AddText "Usage: /w 'player' 'Message' or /whisper 'player' 'Message'", AlertColor
                        End If
                    Else
                        AddText "Usage: /w 'player' 'Message' or /whisper 'player' 'Message'", AlertColor
                    End If
                    
                Case "/g", "/guild"
                    If UBound(Command) >= 1 Then
                        GuildMsg Right$(MyText, Len(MyText) - Len(Command(0)) - 1)
                    Else
                        AddText "Usage: /g 'Message' or /guild 'Message'", AlertColor
                    End If
                    
                Case "/p", "/party"
                    If UBound(Command) >= 1 Then
                        PartyMsg Right$(MyText, Len(MyText) - Len(Command(0)) - 1)
                    Else
                        AddText "Usage: /p 'Message' or /party 'Message'", AlertColor
                    End If
                    
                Case "/em", "/emote"
                    If UBound(Command) >= 1 Then
                        EmoteMsg Right$(MyText, Len(MyText) - Len(Command(0)))
                    Else
                        AddText "Usage: /me 'Message'", AlertColor
                    End If
                        
                Case "/online"
                    SendWhosOnline
                    
                Case "/clear"
                    frmMainGame.txtChat.Text = vbNullString
                    
                Case "/fps"
                    ShowFPS = Not ShowFPS
                
                Case "/stats"
                    SendGetStats
                    
                Case "/fix"
                    SendFix
                
                Case "/emotes"
                    For i = 1 To MAX_EMOTICONS
                        If Trim$(Emoticons(i).Command) <> vbNullString Then
                            ChatText = ChatText & Trim$(Emoticons(i).Command) & ", "
                        End If
                    Next
                    AddText "Available Emotes: " & ChatText, White
                    
                Case "/invite"
                    If UBound(Command) >= 1 Then
                        If IsAlpha(Command(1)) Then
                            SendPartyRequest Command(1)
                        Else
                            AddText "Usage: /party 'adventurer'", AlertColor
                        End If
                    Else
                        AddText "Usage: /party 'adventurer'", AlertColor
                    End If
                    
                Case "/join"
                    SendJoinParty
                    
                Case "/leave"
                    SendLeaveParty
                
                Case "/gmotd"
                    If UBound(Command) >= 1 Then
                        SendSetGMOTD Right$(MyText, Len(MyText) - Len(Command(0)) - 1)
                    Else
                        AddText "/gmotd 'guild Message'", AlertColor
                    End If
                        
                Case "/gcreate"
                    ' command(1) = guild abbreviation - 3 chars or less
                    ' command(2) - command(?) = guild name - 20 chars or less
                    If UBound(Command) >= 2 Then
                        Name = Right$(MyText, Len(MyText) - Len(Command(0)) - Len(Command(1)) - 2)
                        SendGCreate Name, Command(1)
                    Else
                        AddText "Usage: /ggcreate '3 Chars(Guild Abbreviation)' '20 Chars(Guild Name)'", AlertColor
                    End If
                
                Case "/gquit"
                    SendGQuit
                    
                Case "/gdelete"
                    SendGDelete
                
                Case "/gkick"
                    If UBound(Command) >= 1 Then
                        If IsAlpha(Command(1)) Then
                            SendGKick Command(1)
                        Else
                            AddText "Usage: /gkick 'adventurer'", AlertColor
                        End If
                    Else
                        AddText "Usage: /gkick 'adventurer'", AlertColor
                    End If
                    
                Case "/gpromote"
                    If UBound(Command) >= 1 Then
                        If IsAlpha(Command(1)) Then
                            SendGPromote Command(1)
                        Else
                            AddText "Usage: /gpromote 'adventurer'", AlertColor
                        End If
                    Else
                        AddText "Usage: /gpromote 'adventurer'", AlertColor
                    End If
                
                Case "/gdemote"
                    If UBound(Command) >= 1 Then
                        If IsAlpha(Command(1)) Then
                            SendGDemote Command(1)
                        Else
                            AddText "Usage: /gdemote 'adventurer'", AlertColor
                        End If
                    Else
                        AddText "Usage: /gdemote 'adventurer'", AlertColor
                    End If
                    
                Case "/ginvite"
                    If UBound(Command) >= 1 Then
                        If IsAlpha(Command(1)) Then
                            SendGInvite Command(1)
                        Else
                            AddText "Usage: /ginvite 'adventurer'", AlertColor
                        End If
                    Else
                        AddText "Usage: /ginvite 'adventurer'", AlertColor
                    End If
                    
                Case "/gjoin"
                    SendGJoin
                
                Case "/gdecline"
                    SendGDecline
                
                Case "/release"
                    SendRelease
                    
                Case "/revive"
                    SendRevive
                    
                ' // Monitor Admin Commands //
                Case "/a", "/admin"
                    If Current_Access(MyIndex) >= ADMIN_MONITOR Then
                         If UBound(Command) >= 1 Then
                            AdminMsg Right$(MyText, Len(MyText) - Len(Command(0)) - 1)
                         Else
                            AddText "Usage: /admin 'Message' or /a 'Message'", AlertColor
                        End If
                    Else
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                    End If
                    
                Case "/kick"
                    If Current_Access(MyIndex) >= ADMIN_MONITOR Then
                         If UBound(Command) >= 1 Then
                            If IsAlpha(Command(1)) Then
                                SendKick Command(1)
                            Else
                                AddText "Usage: /kick 'adventurer'", AlertColor
                            End If
                         Else
                            AddText "Usage: /kick 'adventurer'", AlertColor
                        End If
                    Else
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                    End If
                    
                Case "/kill"
                    If Current_Access(MyIndex) >= ADMIN_MONITOR Then
                         If UBound(Command) >= 1 Then
                            If IsAlpha(Command(1)) Then
                                SendKill Command(1)
                            Else
                                AddText "Usage: /kill 'adventurer'", AlertColor
                            End If
                         Else
                            AddText "Usage: /kill 'adventurer'", AlertColor
                        End If
                    Else
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                    End If
                    
                Case "/info"
                    If Current_Access(MyIndex) >= ADMIN_MONITOR Then
                        If UBound(Command) >= 1 Then
                            If IsAlpha(Command(1)) Then
                                SendPlayerInfoRequest Command(1)
                            Else
                                AddText "Usage: /info 'adventurer'", AlertColor
                            End If
                        Else
                            AddText "Usage: /info 'adventurer'", AlertColor
                        End If
                    Else
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                    End If
                    
                ' // Mapper Admin Commands //
                Case "/loc"
                    If Current_Access(MyIndex) >= ADMIN_MAPPER Then
                        SendRequestLocation
                    Else
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                    End If
                    
                Case "/editemoticon"
                    If Current_Access(MyIndex) >= ADMIN_MAPPER Then
                        SendRequestEditEmoticon
                    Else
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                    End If
                    
                Case "/emoticoneditor"
                    If Current_Access(MyIndex) >= ADMIN_MAPPER Then
                        SendRequestEditEmoticon
                    Else
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                    End If
                
                Case "/mapeditor"
                    If Current_Access(MyIndex) >= ADMIN_MAPPER Then
                        SendRequestEditMap
                    Else
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                    End If
                    
                Case "/warpmeto"
                    If Current_Access(MyIndex) >= ADMIN_MAPPER Then
                        If UBound(Command) >= 1 Then
                            If IsAlpha(Command(1)) Then
                                WarpMeTo Command(1)
                            Else
                                AddText "Usage: /warpmeto 'adventurer'", AlertColor
                            End If
                        Else
                            AddText "Usage: /warpmeto 'adventurer'", AlertColor
                        End If
                    Else
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                    End If
                                                           
                Case "/warptome"
                    If Current_Access(MyIndex) >= ADMIN_MAPPER Then
                        If UBound(Command) >= 1 Then
                            If IsAlpha(Command(1)) Then
                                WarpToMe Command(1)
                            Else
                                AddText "Usage: /warptome 'adventurer'", AlertColor
                            End If
                        Else
                            AddText "Usage: /warptome 'adventurer'", AlertColor
                        End If
                    Else
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                    End If
                           
                Case "/warpto"
                    If Current_Access(MyIndex) >= ADMIN_MAPPER Then
                        If UBound(Command) >= 1 Then
                            If IsNumeric(Command(1)) Then
                                WarpTo CLng(Command(1))
                            Else
                                AddText "Usage: /warpto '#'", AlertColor
                            End If
                        Else
                            AddText "Usage: /warpto '#'", AlertColor
                        End If
                    Else
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                    End If
                
                Case "/setsprite"
                    ' command(1) = adventurer
                    ' command(2) = sprite num
                    If Current_Access(MyIndex) >= ADMIN_MAPPER Then
                        If UBound(Command) >= 2 Then
                            If IsNumeric(Command(2)) Then
                                If IsAlpha(Command(1)) Then
                                    SendSetSprite Command(1), CLng(Command(2))
                                Else
                                    AddText "Usage: /setsprite 'adventurer' '#'", AlertColor
                                End If
                            Else
                                AddText "Usage: /setsprite 'adventurer' '#'", AlertColor
                            End If
                        Else
                            AddText "Usage: /setsprite 'adventurer' '#'", AlertColor
                        End If
                    Else
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                    End If
                
                Case "/mapreport"
                    If Current_Access(MyIndex) >= ADMIN_MAPPER Then
                        SendMapReport
                    Else
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                    End If
                
                Case "/respawn"
                    If Current_Access(MyIndex) >= ADMIN_MAPPER Then
                        SendMapRespawn
                    Else
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                    End If
                
                Case "/motd"
                    If Current_Access(MyIndex) >= ADMIN_MAPPER Then
                        If UBound(Command) >= 1 Then
                            SendMOTDChange Right$(MyText, Len(MyText) - Len(Command(0)) - 1)
                        Else
                            AddText "Usage: /motd 'Message'", AlertColor
                        End If
                    Else
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                    End If
                
                Case "/banlist"
                    If Current_Access(MyIndex) >= ADMIN_MAPPER Then
                        SendBanList
                    Else
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                    End If
                        
                Case "/ban"
                    If Current_Access(MyIndex) >= ADMIN_MAPPER Then
                        If UBound(Command) >= 1 Then
                            If IsAlpha(Command(1)) Then
                                SendBan Command(1)
                            Else
                                AddText "Usage: /ban 'adventurer'", AlertColor
                            End If
                        Else
                            AddText "Usage: /ban 'adventurer'", AlertColor
                        End If
                    Else
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                    End If
                        
                ' // Developer Admin Commands //
                Case "/edititem"
                    If Current_Access(MyIndex) >= ADMIN_DEVELOPER Then
                        SendRequestEditItem
                    Else
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                    End If
                
                Case "/editnpc"
                    If Current_Access(MyIndex) >= ADMIN_DEVELOPER Then
                        SendRequestEditNpc
                    Else
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                    End If
                
                Case "/editshop"
                    If Current_Access(MyIndex) >= ADMIN_DEVELOPER Then
                        SendRequestEditShop
                    Else
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                    End If
               
               Case "/editspell"
                    If Current_Access(MyIndex) >= ADMIN_DEVELOPER Then
                        SendRequestEditSpell
                    Else
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                    End If
                
                Case "/editanim"
                    If Current_Access(MyIndex) >= ADMIN_DEVELOPER Then
                        SendRequestEditAnimation
                    Else
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                    End If
                
                Case "/editquest"
                    If Current_Access(MyIndex) >= ADMIN_DEVELOPER Then
                        SendRequestEditQuest
                    Else
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                    End If
                    
                ' // Creator Admin Commands //
                Case "/setaccess"
                    If Current_Access(MyIndex) >= ADMIN_CREATOR Then
                        If UBound(Command) >= 2 Then
                            If IsNumeric(Command(2)) Then
                                If IsAlpha(Command(1)) Then
                                    SendSetAccess Command(1), CLng(Command(2))
                                Else
                                    AddText "Usage: /setaccess 'adventurer' '#'", AlertColor
                                End If
                            Else
                                AddText "Usage: /setaccess 'adventurer' '#'", AlertColor
                            End If
                        Else
                            AddText "Usage: /setaccess 'adventurer' '#'", AlertColor
                        End If
                    Else
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                    End If
                
                Case "/destroybanlist"
                    If Current_Access(MyIndex) >= ADMIN_CREATOR Then
                        SendBanDestroy
                    Else
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                    End If
                
                Case "/lock"
                    If Current_Access(MyIndex) >= ADMIN_CREATOR Then
                        LockFPS = Not LockFPS
                    Else
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                    End If
                
                Case Else
                    ' Check for emoticons
                    For i = 1 To MAX_EMOTICONS
                        If Trim$(Emoticons(i).Command) <> vbNullString Then
                            If Trim$(Emoticons(i).Command) = Command(0) Then
                                SendCheckEmoticon i
                                n = n + 1
                                Exit For
                            End If
                        End If
                    Next
                    ' If we don't find a emoticon, then it's an invalid command
                    If n = 0 Then AddText "Invalid command!", AlertColor
                    
            End Select
            
            MyText = vbNullString
            Exit Sub
        End If
        
        ' Global Message
        If LenB(MyText) > 0 Then
            If frmMainGame.optglobal.Value Then
                RealmMsg MyText
            End If
            
            ' Guild Message
            If frmMainGame.optGuild.Value Then
                GuildMsg MyText
            End If
            
            ' Guild Message
            If frmMainGame.optParty.Value Then
                PartyMsg MyText
            End If
            
            ' Say Message
            If frmMainGame.optMap.Value Then
                SayMsg MyText
            End If
        End If
        
        MyText = vbNullString
        Exit Sub
    End If
    
    ' Handle when the user presses the backspace key
    If KeyAscii = vbKeyBack Then
        If LenB(MyText) > 0 Then
            MyText = Left$(MyText, Len(MyText) - 1)
            Exit Sub
        End If
    End If
    
    ' And if neither, then add the character to the user's text buffer
    If KeyAscii <> vbKeyReturn Then
        If KeyAscii <> vbKeyBack Then
            ' Make sure they just use standard keys, no gay shitty macro keys
            If KeyAscii >= 32 Then
                If KeyAscii <= 126 Then
                    MyText = MyText & ChrW$(KeyAscii)
                End If
            End If
        End If
    End If
End Sub

Sub CheckMapGetItem()
    If GetTickCount > Player(MyIndex).MapGetTimer Then
        If Trim$(MyText) = vbNullString Then
            Player(MyIndex).MapGetTimer = GetTickCount + 250
            SendMapGetItem
        End If
    End If
End Sub

Sub CheckAttack()
    If ControlDown Then
        ' Doesnt' matter if they are dead
        If Player(MyIndex).IsDead Then Exit Sub
        
        ' cancel spellcasting if they are casting
        CheckCasting
        
        If GetTickCount > Player(MyIndex).AttackTimer + 1000 Then
            If Player(MyIndex).Attacking = 0 Then
                Player(MyIndex).Attacking = 1
                Player(MyIndex).AttackTimer = GetTickCount
                SendAttack
            End If
        End If
    End If
End Sub

Public Sub CheckCasting()
    If CastingSpell Then
        CastingSpell = 0
        CastTime = 0
        SendCancelSpell
    End If
End Sub

Sub CheckInput(ByVal KeyState As Byte, ByVal KeyCode As Integer, ByVal Shift As Integer)
    If GettingMap = False Then
        If KeyState = 1 Then
            If KeyCode = vbKeyReturn Then
                CheckMapGetItem
            End If
            If KeyCode = vbKeyControl Then
                ControlDown = True
            End If
            If KeyCode = vbKeyUp Then
                DirUp = True
                DirDown = False
                DirLeft = False
                DirRight = False
            End If
            If KeyCode = vbKeyDown Then
                DirUp = False
                DirDown = True
                DirLeft = False
                DirRight = False
            End If
            If KeyCode = vbKeyLeft Then
                DirUp = False
                DirDown = False
                DirLeft = True
                DirRight = False
            End If
            If KeyCode = vbKeyRight Then
                DirUp = False
                DirDown = False
                DirLeft = False
                DirRight = True
            End If
            If KeyCode = vbKeyShift Then
                ShiftDown = True
            End If
            If KeyCode = vbKeyMenu Then
                AltDown = True
            End If
            If KeyCode = vbKeyEscape Then
                SendClearTarget
            End If
        Else
            If KeyCode = vbKeyUp Then DirUp = False
            If KeyCode = vbKeyDown Then DirDown = False
            If KeyCode = vbKeyLeft Then DirLeft = False
            If KeyCode = vbKeyRight Then DirRight = False
            If KeyCode = vbKeyShift Then ShiftDown = False
            If KeyCode = vbKeyMenu Then AltDown = False
            If KeyCode = vbKeyControl Then ControlDown = False
        End If
    End If
End Sub

Function IsTryingToMove() As Boolean

    IsTryingToMove = False
    
    If (DirUp = True) Or (DirDown = True) Or (DirLeft = True) Or (DirRight = True) Then
        IsTryingToMove = True
    End If
End Function

Function CheckDirection(ByVal Direction As Byte) As Boolean
Dim X As Long, Y As Long
Dim i As Long

    CheckDirection = True
    
    Select Case Direction
        Case DIR_UP
            X = Current_X(MyIndex)
            Y = Current_Y(MyIndex) - 1
        Case DIR_DOWN
            X = Current_X(MyIndex)
            Y = Current_Y(MyIndex) + 1
        Case DIR_LEFT
            X = Current_X(MyIndex) - 1
            Y = Current_Y(MyIndex)
        Case DIR_RIGHT
            X = Current_X(MyIndex) + 1
            Y = Current_Y(MyIndex)
    End Select
    
    If Not IsValidMapPoint(X, Y) Then Exit Function
    
    ' Check to see if the map tile is blocked or not
    If Map.Tile(X, Y).Type = TILE_TYPE_BLOCKED Then Exit Function
    
    ' Check for item block
    If Map.Tile(X, Y).Type = TILE_TYPE_ITEM Then
        If Map.Tile(X, Y).Data3 Then
            Exit Function
        End If
    End If
                                
    ' Check to see if the key door is open or not
    If Map.Tile(X, Y).Type = TILE_TYPE_KEY Then
        ' This actually checks if its open or not
        If Not TempTile(X, Y).Open Then
            Exit Function
        End If
    End If
    
    ' Check to see if a player is already on that tile
    For i = 1 To MapPlayersCount
        If Current_X(MapPlayers(i)) = X Then
            If Current_Y(MapPlayers(i)) = Y Then
                Exit Function
            End If
        End If
    Next
'    For i = 1 To MAX_PLAYERS
'        If IsPlaying(i) Then
'            If Current_Map(i) = Current_Map(MyIndex) Then
'                If Current_X(i) = X Then
'                    If Current_Y(i) = Y Then
'                        Exit Function
'                    End If
'                End If
'            End If
'        End If
'    Next

    ' Check to see if a npc is already on that tile
    For i = 1 To MapNpcCount
        If MapNpc(i).Num Then
            If MapNpc(i).X = X Then
                If MapNpc(i).Y = Y Then
                    Exit Function
                End If
            End If
        End If
    Next
    
    CheckDirection = False
    
End Function

Function CanMove() As Boolean
Dim d As Long
Dim OldDir As Byte

    CanMove = True
    
    ' Doesnt' matter if they are dead
    If Player(MyIndex).IsDead Then
        CanMove = False
        Exit Function
    End If
        
    ' Make sure they aren't trying to move when they are already moving
    If Player(MyIndex).Moving <> 0 Then
        CanMove = False
        Exit Function
    End If
    
    ' Check if they just want to change dir
    If AltDown Then
        OldDir = Current_Dir(MyIndex)
        If DirUp Then Update_Dir MyIndex, DIR_UP
        If DirDown Then Update_Dir MyIndex, DIR_DOWN
        If DirLeft Then Update_Dir MyIndex, DIR_LEFT
        If DirRight Then Update_Dir MyIndex, DIR_RIGHT
        If OldDir <> Current_Dir(MyIndex) Then SendPlayerDir
        CanMove = False
        Exit Function
    End If
    
    ' Make sure they haven't just casted a spell
    If Player(MyIndex).CastedSpell = YES Then
        If GetTickCount > Player(MyIndex).AttackTimer + 1000 Then
            Player(MyIndex).CastedSpell = NO
        Else
            CanMove = False
            Exit Function
        End If
    End If
    
    d = Current_Dir(MyIndex)
    If DirUp Then
        Call Update_Dir(MyIndex, DIR_UP)
        
        ' Check to see if they are trying to go out of bounds
        If Current_Y(MyIndex) > 0 Then
            If CheckDirection(DIR_UP) Then
                CanMove = False
                
                ' Set the new direction if they weren't facing that direction
                If d <> DIR_UP Then
                    Call SendPlayerDir
                End If
            End If
        Else
            ' Check if they can warp to a new map
            If Map.Up > 0 Then
                Call SendPlayerRequestNewMap
                GettingMap = True
            End If
            
            CanMove = False
        End If
        
        Exit Function
    End If
            
    If DirDown Then
        Call Update_Dir(MyIndex, DIR_DOWN)
        
        ' Check to see if they are trying to go out of bounds
        If Current_Y(MyIndex) < Map.MaxY Then
            If CheckDirection(DIR_DOWN) Then
                CanMove = False
                
                ' Set the new direction if they weren't facing that direction
                If d <> DIR_DOWN Then
                    Call SendPlayerDir
                End If
            End If
        Else
            ' Check if they can warp to a new map
            If Map.Down > 0 Then
                Call SendPlayerRequestNewMap
                GettingMap = True
            End If
            CanMove = False
        End If
        
        Exit Function
    End If
                
    If DirLeft Then
        Call Update_Dir(MyIndex, DIR_LEFT)
        
        ' Check to see if they are trying to go out of bounds
        If Current_X(MyIndex) > 0 Then
            If CheckDirection(DIR_LEFT) Then
                CanMove = False
                
                ' Set the new direction if they weren't facing that direction
                If d <> DIR_LEFT Then
                    Call SendPlayerDir
                End If
            End If
        Else
            ' Check if they can warp to a new map
            If Map.Left > 0 Then
                Call SendPlayerRequestNewMap
                GettingMap = True
            End If
            CanMove = False
        End If
        
        Exit Function
    End If
        
    If DirRight Then
        Call Update_Dir(MyIndex, DIR_RIGHT)
        
        ' Check to see if they are trying to go out of bounds
        If Current_X(MyIndex) < Map.MaxX Then
            If CheckDirection(DIR_RIGHT) Then
                CanMove = False
                
                ' Set the new direction if they weren't facing that direction
                If d <> DIR_RIGHT Then
                    Call SendPlayerDir
                End If
            End If
        Else
            ' Check if they can warp to a new map
            If Map.Right > 0 Then
                Call SendPlayerRequestNewMap
                GettingMap = True
            End If
            CanMove = False
        End If
        
        Exit Function
    End If
End Function

Public Sub CheckMovement()

    If Not GettingMap Then
        If IsTryingToMove Then
            ' cancel spellcasting if they are casting
            CheckCasting
            
            If CanMove Then
                ' Check if player has the shift key down for running
                'If ShiftDown Then
                '    Player(MyIndex).Moving = MOVING_RUNNING
                'Else
'                    Player(MyIndex).Moving = MOVING_WALKING
                'End If
                
                Player(MyIndex).Moving = MOVING_WALKING
                
                Select Case Current_Dir(MyIndex)
                    Case DIR_UP
                        'Call SendPlayerMove
                        Player(MyIndex).YOffset = PIC_Y
                        Update_Y MyIndex, Current_Y(MyIndex) - 1
                
                    Case DIR_DOWN
                        'Call SendPlayerMove
                        Player(MyIndex).YOffset = -PIC_Y
                        Update_Y MyIndex, Current_Y(MyIndex) + 1
                
                    Case DIR_LEFT
                        'Call SendPlayerMove
                        Player(MyIndex).XOffset = PIC_X
                        Update_X MyIndex, Current_X(MyIndex) - 1
                
                    Case DIR_RIGHT
                        'Call SendPlayerMove
                        Player(MyIndex).XOffset = -PIC_X
                        Update_X MyIndex, Current_X(MyIndex) + 1
                End Select
            
                SendPlayerMove
                
                ' Gotta check :)
                If Map.Tile(Current_X(MyIndex), Current_Y(MyIndex)).Type = TILE_TYPE_WARP Then
                    GettingMap = True
                End If
            End If
        End If
    End If
End Sub

Public Sub UpdateInventory()
    If InGame Then Call BltInventory
End Sub

Public Sub UpdateEquipment()
Dim rec As RECT
Dim rec_pos As RECT
Dim i As Long
    
    If InGame Then
        With rec_pos
            .Top = 0
            .Bottom = PIC_Y
            .Left = 0
            .Right = PIC_X
        End With
        
        For i = 1 To Slots.Slot_Count
            frmMainGame.picEquipment(i - 1).Visible = False
            If Current_EquipmentSlot(MyIndex, i) Then
                With rec
                    .Top = Item(Current_EquipmentSlot(MyIndex, i)).Pic * PIC_Y
                    .Bottom = .Top + PIC_Y
                    .Left = 0
                    .Right = PIC_X
                End With
                        
                frmMainGame.picEquipment(i - 1).Visible = True
                DD_ItemSurf.BltToDC frmMainGame.picEquipment(i - 1).hdc, rec, rec_pos
            End If
        Next
    End If
End Sub

Public Function StatName(ByVal Stat As Stats) As String
    Select Case Stat
        Case Stats.Strength:        StatName = "Strength"
        Case Stats.Dexterity:       StatName = "Dexterity"
        Case Stats.Vitality:        StatName = "Vitality"
        Case Stats.Intelligence:    StatName = "Intelligence"
        Case Stats.Wisdom:          StatName = "Wisdom"
    End Select
End Function

Public Function StatAbbreviation(ByVal Stat As Stats) As String
    Select Case Stat
        Case Stats.Strength:        StatAbbreviation = "STR"
        Case Stats.Dexterity:       StatAbbreviation = "DEX"
        Case Stats.Vitality:        StatAbbreviation = "VIT"
        Case Stats.Intelligence:    StatAbbreviation = "INT"
        Case Stats.Wisdom:          StatAbbreviation = "WIS"
    End Select
End Function

Public Function VitalName(ByVal Vital As Vitals) As String
    Select Case Vital
        Case Vitals.HP: VitalName = "HP"
        Case Vitals.MP: VitalName = "MP"
        Case Vitals.SP: VitalName = "SP"
    End Select
End Function

Public Function EquipmentName(ByVal EquipmentSlot As Slots) As String
    Select Case EquipmentSlot
        Case Slots.Armor:  EquipmentName = "Armor"
        Case Slots.Weapon: EquipmentName = "Weapon"
        Case Slots.Helmet: EquipmentName = "Helmet"
        Case Slots.Shield: EquipmentName = "Shield"
    End Select
End Function

Public Function TargetName(ByVal Target As Targets) As String
    Select Case Target
        Case Targets.Target_None: TargetName = "None"
        Case Targets.Target_SelfOnly: TargetName = "Self Only"
        Case Targets.Target_PlayerHostile: TargetName = "Player - Hostile"
        Case Targets.Target_PlayerBeneficial: TargetName = "Player - Beneficial"
        Case Targets.Target_Npc: TargetName = "Npc"
        Case Targets.Target_PlayerParty: TargetName = "Player - Party Member"
    End Select
End Function

Public Function BindName(ByVal Bind As ItemBind) As String
    Select Case Bind
        Case ItemBind.None: BindName = "None"
        Case ItemBind.BindOnEquip: BindName = "Bind on Equip"
        Case ItemBind.BindOnPickUp: BindName = "Bind on Pickup"
    End Select
End Function

Public Function IsAlpha(ByRef s As String) As Boolean
Dim i As Long
Dim n As Long

    IsAlpha = True
    For i = 1 To Len(s)
        n = Asc(Mid$(s, i, 1))
        If (n < 64) Or (n > 90 And n < 97) Or (n > 122) Then
            IsAlpha = False
            Exit Function
        End If
    Next
End Function

Public Function Clamp(ByVal Value As Long, ByVal Min As Long, ByVal Max As Long) As Long
    Clamp = Value
    If Value < Min Then Clamp = Min
    If Value > Max Then Clamp = Max
End Function

Public Function MapNpcLimit() As Long
    MapNpcLimit = ((CLng(Map.MaxX) \ MAX_MAPX) + (CLng(Map.MaxY) \ MAX_MAPY)) * 5
End Function

Public Sub UpdateMapNpcCount()
Dim i As Long
    
    MapNpcCount = 0
    For i = 1 To MAX_MOBS
        MapNpcCount = MapNpcCount + Map.Mobs(i).NpcCount
    Next
    
    ReDim Preserve MapNpc(MapNpcCount)
End Sub

Public Sub Update_StatsWindow()
Dim i As Long

    frmMainGame.lblName.Caption = Current_Name(MyIndex) & " the level " & Current_Level(MyIndex) & " " & Trim$(Class(Current_Class(MyIndex)).Name)
    frmMainGame.lblPoints.Caption = "Stat points: " & Current_Points(MyIndex)
    For i = 1 To Stats.Stat_Count
        frmMainGame.lblStat(i - 1).Caption = StatName(i) & ": " & Current_Stat(MyIndex, i) & " (" & Current_BaseStat(MyIndex, i) & " +" & Current_ModStat(MyIndex, i) & ")"
        ' If the arrows should be visible
        If Current_Points(MyIndex) Then
            frmMainGame.lblStatUp(i - 1).Visible = True
        Else
            frmMainGame.lblStatUp(i - 1).Visible = False
        End If
    Next
    For i = 1 To Vitals.Vital_Count
        frmMainGame.lblVital(i - 1).Caption = VitalName(i) & ": " & Current_Vital(MyIndex, i) & " / " & Current_MaxVital(MyIndex, i)
    Next
    frmMainGame.lblDamage = "Damage: " & Int(Current_Damage(MyIndex) * 0.9) + 1 & " - " & Int(Current_Damage(MyIndex) * 1.1) + 1
    frmMainGame.lblProtection = "Protection: " & Int(Current_Protection(MyIndex) * 0.9) + 1 & " - " & Int(Current_Protection(MyIndex) * 1.1) + 1
    frmMainGame.lblMagicDamage = "Magic Bonus: " & Current_MagicDamage(MyIndex)
    frmMainGame.lblMagicProtection = "Magic Protection: " & Current_MagicProtection(MyIndex)
End Sub

Public Sub CastSpell(ByVal SpellSlot As Long)
Dim i As Long
Dim SpellNum As Long

    ' If dead , doesn't matter
    If Current_IsDead(MyIndex) Then Exit Sub
    
    If SpellSlot <= 0 Then
        ' add error Message
        Exit Sub
    End If
    
    If SpellSlot > MAX_PLAYER_SPELLS Then
        ' add error Message
        Exit Sub
    End If
    
    ' Set the spellnum
    SpellNum = Player(MyIndex).Spell(SpellSlot).SpellNum
    
    ' Check if you actually have a spell here
    If SpellNum = 0 Then
        AddText "No spell here.", BrightRed
        Exit Sub
    End If
    
    If GetTickCount < Player(MyIndex).AttackTimer + 1000 Then
        ' add error Message
        Exit Sub
    End If
    
    If Player(MyIndex).Moving > 0 Then
        ' add error Message
        Exit Sub
    End If
    
    If CastingSpell > 0 Then
        AddText "Already casting a spell.", BrightRed
        Exit Sub
    End If
    
    ' first check if it's not on cooldown
    If Player(MyIndex).Spell(SpellSlot).Cooldown > 0 Then
        AddText "Spell on cooldown.", BrightRed
        Exit Sub
    End If
    
    ' Prelim check for vital required to cast
    For i = 1 To Vitals.Vital_Count
        If Spell(SpellNum).VitalReq(i) > Current_Vital(MyIndex, i) Then
            AddText CStr(Spell(SpellNum).VitalReq(i)) & " " & VitalName(i) & " required.", BrightRed
            Exit Sub
        End If
    Next
    
    ' Prelim check for target
    ' Will check the target flags on the spell and make sure you have the appropriate target
    ' If Self cast - doesn't matter if you have a target or not - overrides other flags
    If Not Spell(SpellNum).TargetFlags And Targets.Target_SelfOnly Then
        ' If you have a target
        If MyTarget > 0 Then
            ' If your target is a player check if the spell can be cast on players
            If MyTargetType = TARGET_TYPE_PLAYER Then
                ' Check if it's hostile
                If (Spell(SpellNum).TargetFlags And Targets.Target_PlayerHostile) = Not Targets.Target_PlayerHostile Then
                    AddText "Can not cast this spell on players.", BrightRed
                    Exit Sub
                End If
                
                ' Check if it's Target_PlayerParty
                ' We do this because Target_PlayerParty overrides the other player target flags
                If Spell(SpellNum).TargetFlags And Targets.Target_PlayerParty Then
                    ' All these checks will have to be on the server for right now
                    ' TODO: Finish - Will need to send party shit
'                    ' Check if you're in a party
'                    If Player(Index).InParty Then
'                        ' Check if the target is in your party
'                        If Player(Index).PartyIndex <> Player(Player(Index).Target).PartyIndex Then
'                            SendPlayerMsg Index, "Can only cast this spell on party members.", BrightRed
'                            Exit Sub
'                        End If
'                    ' Not in a party, can only cast on self then
'                    Else
'                        If Player(Index).Target <> Index Then
'                            SendPlayerMsg Index, "Can only cast this spell on party members.", BrightRed
'                            Exit Sub
'                        End If
'                    End If
                ' Since it's not a PlayerParty spell then we check the other player target flags
                Else
                    ' Check if it's beneficial
                    If (Spell(SpellNum).TargetFlags And Targets.Target_PlayerBeneficial) = Not Targets.Target_PlayerBeneficial Then
                        AddText "Can not cast this spell on players.", BrightRed
                        Exit Sub
                    End If
                    
                    ' If hostile - check if you can actually attack them
                    If (Spell(SpellNum).TargetFlags And Targets.Target_PlayerHostile) Then
                        If Not CheckAttackPlayer(MyTarget) Then
                            'sendplayermsg Index, "Can not cast this spell on players.", BrightRed
                            Exit Sub
                        End If
                    End If
                End If
                
                ' Should mean they can cast on player - check their range now
                If Not PlayerInRange(Current_X(MyTarget), Current_Y(MyTarget), Spell(SpellNum).Range) Then
                    AddText "Target not in range.", BrightRed
                    Exit Sub
                End If
                                
                ' Now we check the spell type and if the player needs to be alive
                Select Case Spell(SpellNum).Type
                    ' Revive is the only spell type the target must be dead
                    Case SPELL_TYPE_REVIVE
                        ' Check if they are dead
                        If Not Current_IsDead(MyTarget) Then
                            AddText "Target not dead.", BrightRed
                            Exit Sub
                        End If
                    ' All other spell types the player must be alive
                    Case Else
                        ' Can't cast on a dead player
                        If Current_IsDead(MyTarget) Then
                            AddText "Target is dead.", BrightRed
                            Exit Sub
                        End If
                        
                End Select
                
            ' If your target is a npc check if the spell can be cast on npcs
            ElseIf MyTargetType = TARGET_TYPE_NPC Then
                If Not (Spell(SpellNum).TargetFlags And Targets.Target_Npc) = Targets.Target_Npc Then
                    AddText "Can not cast this spell on npcs.", BrightRed
                    Exit Sub
                End If
                
                ' Check if it's a npc you can attack
                If Npc(MapNpc(MyTarget).Num).Behavior = NPC_BEHAVIOR_FRIENDLY Then
                    AddText "Can not cast this spell on friendly npcs.", BrightRed
                    Exit Sub
                End If
                
                If Npc(MapNpc(MyTarget).Num).Behavior = NPC_BEHAVIOR_SHOPKEEPER Then
                    AddText "Can not cast this spell on friendly npcs.", BrightRed
                    Exit Sub
                End If
                
                ' Should mean they can cast on npc - check their range now
                If Not NpcInRange(MyTarget, Spell(SpellNum).Range) Then
                    AddText "Target not in range.", BrightRed
                    Exit Sub
                End If
                
            End If
        ' No target
        Else
            ' For now we can't cast if we don't have a target
            AddText "You need a target.", BrightRed
            Exit Sub
        End If
    End If
    
    ' You can cast
    SendCast SpellSlot
                        
    Player(MyIndex).Attacking = 1
    Player(MyIndex).AttackTimer = GetTickCount
    Player(MyIndex).CastedSpell = YES
    
    CastingSpell = SpellSlot
    CastTime = GetTickCount + (Spell(Player(MyIndex).Spell(SpellSlot).SpellNum).CastTime * 1000)
    
End Sub

Function DeathTimer() As String
Dim Minutes As Long
Dim Seconds As Long
Dim Timer As Long
    
    Timer = (Player(MyIndex).IsDeadTimer - GetTickCount) \ 1000
    Minutes = Timer \ 60
    Seconds = Timer Mod 60
    DeathTimer = "Release in " & Format$(Minutes, "00") & ":" & Format$(Seconds, "00")
End Function

Public Function Rand(ByVal Low As Long, ByVal High As Long) As Long
    Rand = Int((High - Low + 1) * Rnd) + Low
End Function

Public Function CheckAttackPlayer(ByVal Victim As Long) As Boolean

    CheckAttackPlayer = False
    
    If MyIndex = Victim Then
        AddText "As much as you would like to, you can't attack yourself...", AlertColor
        Exit Function
    End If
    
    ' Check if map is attackable
    If Map.Moral = MAP_MORAL_SAFE Then
        If Current_PK(Victim) = 0 Then
            AddText "This is a haven.", AlertColor
            Exit Function
        End If
    End If
        
    ' Check if they are dead
    If Current_IsDead(Victim) Then
        AddText "That player is currently dead.", AlertColor
        Exit Function
    End If
    
    ' Check to make sure that they dont have access
    If Current_Access(MyIndex) > ADMIN_MONITOR Then
        AddText "Realm Masters may not murder.", AlertColor
        Exit Function
    End If
        
    ' Check to make sure the victim isn't an admin
    If Current_Access(Victim) > ADMIN_MONITOR Then
        AddText "You cannot attack a Realm Master.", AlertColor
        Exit Function
    End If
          
    ' Make sure they are high enough level
    If Current_Level(MyIndex) < 10 Then
        AddText "You must be level 10+ to fight.", AlertColor
        Exit Function
    End If
          
    ' Make sure they are high enough level
    If Current_Level(Victim) < 10 Then
        AddText "They are under level 10.", AlertColor
        Exit Function
    End If
    
    If LenB(Player(MyIndex).GuildName) > 0 Then
        If Player(MyIndex).GuildName = Player(Victim).GuildName Then
            AddText "Cannot attack guild members.", AlertColor
            Exit Function
        End If
    End If
    
    ' if you get through the checks, you're golden
    CheckAttackPlayer = True
End Function

Function PlayerInRange(ByVal X As Long, ByVal Y As Long, ByVal Distance As Byte) As Boolean
Dim DistanceX As Long, DistanceY As Long

    PlayerInRange = False
    
    DistanceX = X - Current_X(MyIndex)
    DistanceY = Y - Current_Y(MyIndex)
    
    ' Make sure we get a positive value
    If DistanceX < 0 Then DistanceX = -DistanceX
    If DistanceY < 0 Then DistanceY = -DistanceY
    
    ' Are they in range?
    If DistanceX <= Distance Then
        If DistanceY <= Distance Then
            PlayerInRange = True
        End If
    End If
End Function

Function NpcInRange(ByVal MapNpcNum As Byte, ByVal Distance As Byte) As Boolean
Dim DistanceX As Long, DistanceY As Long

    NpcInRange = False
    
    If MapNpc(MapNpcNum).Num <= 0 Then Exit Function
    
    DistanceX = MapNpc(MapNpcNum).X - Current_X(MyIndex)
    DistanceY = MapNpc(MapNpcNum).Y - Current_Y(MyIndex)
    
    ' Make sure we get a positive value
    If DistanceX < 0 Then DistanceX = -DistanceX
    If DistanceY < 0 Then DistanceY = -DistanceY
    
    ' Are they in range?
    If DistanceX <= Distance Then
        If DistanceY <= Distance Then
            NpcInRange = True
        End If
    End If
End Function

Sub SwitchInvSlots(ByVal OldSlot As Long, ByVal NewSlot As Long)
Dim OldNum As Long
Dim OldValue As Long
Dim OldBound As Boolean
Dim NewNum As Long
Dim NewValue As Long
Dim NewBound As Boolean
Dim OverFlow As Long

    If OldSlot <= 0 Then Exit Sub
    If OldSlot > MAX_INV Then Exit Sub
    If NewSlot <= 0 Then Exit Sub
    If NewSlot > MAX_INV Then Exit Sub
    
    OldNum = Current_InvItemNum(MyIndex, OldSlot)
    OldValue = Current_InvItemValue(MyIndex, OldSlot)
    OldBound = Current_InvItemBound(MyIndex, OldSlot)
    
    NewNum = Current_InvItemNum(MyIndex, NewSlot)
    NewValue = Current_InvItemValue(MyIndex, NewSlot)
    NewBound = Current_InvItemBound(MyIndex, NewSlot)
    
    ' Combine item values if same
    If OldNum > 0 Then
        If NewNum > 0 Then
            If OldNum = NewNum Then
                ' Check to see if stackable
                If Item(OldNum).Stack = 1 And Item(NewNum).Stack = 1 Then
                    ' If the newvalue is at max value it wasn't switching inv slots
                    ' Added below check to try to fix it
                    If NewValue <> Item(NewNum).StackMax Then
                        ' Check if the item values will overflow
                        If OldValue + NewValue > Item(NewNum).StackMax Then
                            OverFlow = Item(NewNum).StackMax - OldValue
                            OldValue = OldValue + OverFlow
                            NewValue = NewValue - OverFlow
                            If OldValue <= 0 Then
                                NewNum = 0
                                NewValue = 0
                                NewBound = False
                            End If
                        Else
                            OldValue = OldValue + NewValue
                            NewNum = 0
                            NewValue = 0
                            NewBound = False
                        End If
                    End If
                End If
            End If
        End If
    End If
    
    Update_InvItemNum MyIndex, NewSlot, OldNum
    Update_InvItemValue MyIndex, NewSlot, OldValue
    Update_InvItemBound MyIndex, NewSlot, OldBound
    
    Update_InvItemNum MyIndex, OldSlot, NewNum
    Update_InvItemValue MyIndex, OldSlot, NewValue
    Update_InvItemBound MyIndex, OldSlot, NewBound
    
    SendChangeInvSlots OldSlot, NewSlot
End Sub

Public Sub UpdateMapPlayers()
Dim i As Long
Dim ii As Long

    ' Clear the map player count so we can recalculate it
    MapPlayersCount = 0
    
    For i = 1 To MAX_PLAYERS
        If Current_Map(i) = Current_Map(MyIndex) Then
            MapPlayersCount = MapPlayersCount + 1
        End If
    Next
    
    If MapPlayersCount = 0 Then Exit Sub
    
    ' Clear the map players array
    ReDim MapPlayers(1 To MapPlayersCount)
    
    ' Loop the OnlinePlayersCount checking for players on this map
    For i = 1 To MAX_PLAYERS
        If Current_Map(i) = Current_Map(MyIndex) Then
            ii = ii + 1
            MapPlayers(ii) = i
        End If
    Next
End Sub

Public Sub AlertMessage(ByVal Message As String, Optional ByRef callBack As Long, Optional ByVal OkayOnly As Boolean = True)
Dim myForm As New frmAlert
    
    myForm.sMessage = Message
    myForm.OkayOnly = OkayOnly
    myForm.callBack = callBack
    myForm.Show
End Sub

Public Sub Release_Click(ByVal YesNo As Long, ByRef Data As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    If YesNo = YES Then
        SendRelease
    End If
End Sub

Public Sub Revive_Click(ByVal YesNo As Long, ByRef Data As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    If YesNo = YES Then
        SendRevive
    End If
End Sub
