Attribute VB_Name = "modGameLogic"
Option Explicit

Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Const SRCAND = &H8800C6
Public Const SRCCOPY = &HCC0020
Public Const SRCPAINT = &HEE0086

Public Const VK_UP = &H26
Public Const VK_DOWN = &H28
Public Const VK_LEFT = &H25
Public Const VK_RIGHT = &H27
Public Const VK_SHIFT = &H10
Public Const VK_RETURN = &HD
Public Const VK_CONTROL = &H11

' Menu states
Public Const MENU_STATE_NEWACCOUNT = 0
Public Const MENU_STATE_DELACCOUNT = 1
Public Const MENU_STATE_LOGIN = 2
Public Const MENU_STATE_GETCHARS = 3
Public Const MENU_STATE_NEWCHAR = 4
Public Const MENU_STATE_ADDCHAR = 5
Public Const MENU_STATE_DELCHAR = 6
Public Const MENU_STATE_USECHAR = 7
Public Const MENU_STATE_INIT = 8

' Speed moving vars
Public Const WALK_SPEED = 4
Public Const RUN_SPEED = 8
Public Const GM_WALK_SPEED = 4
Public Const GM_RUN_SPEED = 8
'Set the variable to your desire,
'32 is a safe and recommended setting

' Game direction vars
Public DirUp As Boolean
Public DirDown As Boolean
Public DirLeft As Boolean
Public DirRight As Boolean
Public ShiftDown As Boolean
Public ControlDown As Boolean

' Game text buffer
Public MyText As String

' Index of actual player
Public MyIndex As Long

' Map animation #, used to keep track of what map animation is currently on
Public MapAnim As Byte
Public MapAnimTimer As Long

' PK Label Switch #
Public PKSet As Byte
Public PKTimer As Long

' Used to freeze controls when getting a new map
Public GettingMap As Boolean

' Used to check if in editor or not and variables for use in editor
Public InEditor As Boolean
Public EditorTileX As Long
Public EditorTileY As Long
Public EditorWarpMap As Long
Public EditorWarpX As Long
Public EditorWarpY As Long

' Used for map item editor
Public ItemEditorNum As Long
Public ItemEditorValue As Long

' Used for map key editor
Public KeyEditorNum As Long
Public KeyEditorTake As Long

' Used for map key opene ditor
Public KeyOpenEditorX As Long
Public KeyOpenEditorY As Long
Public KeyOpenEditorMsg As String

' Map for local use
Public SaveMap As MapRec
'Public SaveMapItem() As MapItemRec
Public SaveMapNpc(1 To MAX_MAP_NPCS) As MapNpcRec

' Used for index based editors
Public InItemsEditor As Boolean
Public InNpcEditor As Boolean
Public InShopEditor As Boolean
Public InSpellEditor As Boolean
Public InEmoticonEditor As Boolean
Public InArrowEditor As Boolean
Public EditorIndex As Long
Public InScriptEditor As Boolean
Public InEffectEditor As Boolean

' Game fps
Public GameFPS As Long

' For BankSystem
Public BankSelect As Integer
Public Inventory As Long

' Used for atmosphere
Public GameWeather As Long
Public GameTime As Long

' Scrolling Variables
Public NewPlayerX As Long
Public NewPlayerY As Long
Public NewXOffset As Long
Public NewYOffset As Long
Public NewX As Long
Public NewY As Long

' Damage Variables
Public DmgDamage As Long
Public DmgTime As Long
Public NPCDmgDamage As Long
Public NPCDmgTime As Long
Public NPCWho As Long

Public EditorItemX As Long
Public EditorItemY As Long

Public EditorShopNum As Long

Public EditorItemNum1 As Byte
Public EditorItemNum2 As Byte
Public EditorItemNum3 As Byte

Public Arena1 As Byte
Public Arena2 As Byte
Public Arena3 As Byte

Public ii As Long, iii As Long
Public sx As Long

Public MouseDownX As Long
Public MouseDownY As Long

Public SpritePic As Long
Public SpriteItem As Long
Public SpritePrice As Long

Public SoundFileName As String

Public ScreenMode As Byte

Public SignLine1 As String
Public SignLine2 As String
Public SignLine3 As String

Public ClassChange As Long
Public ClassChangeReq As Long

Public NoticeTitle As String
Public NoticeText As String
Public NoticeSound As String

Public ScriptNum As Long

Public Connucted As Boolean
                    
Sub Main()
Dim i As Long
    ScreenMode = 0

    frmSendGetData.Visible = True
    Call SetStatus("Checking folders...")
    DoEvents
    
    ' Check if the maps directory is there, if its not make it
    If LCase$(Dir(App.Path & "\Maps", vbDirectory)) <> "maps" Then
        Call MkDir(App.Path & "\Maps")
    End If
    If UCase(Dir(App.Path & "\graphics", vbDirectory)) <> "GRAPHICS" Then
        Call MkDir(App.Path & "\Graphics")
    End If
    If UCase(Dir(App.Path & "\Music", vbDirectory)) <> "MUSIC" Then
        Call MkDir(App.Path & "\Music")
    End If
    If UCase(Dir(App.Path & "\Sounds", vbDirectory)) <> "SOUNDS" Then
        Call MkDir(App.Path & "\Sounds")
    End If
    If UCase(Dir(App.Path & "\Animations", vbDirectory)) <> "ANIMATIONS" Then
        Call MkDir(App.Path & "\Animations")
    End If
    
    DoEvents
        
    ' Make sure we set that we aren't in the game
    InGame = False
    GettingMap = True
    InEditor = False
    InItemsEditor = False
    InNpcEditor = False
    InShopEditor = False
    InArrowEditor = False
    InEmoticonEditor = False
    InScriptEditor = False
    
    frmEndieko.picItems.Picture = LoadPicture(App.Path & "\Graphics\items.bmp")
    frmEndieko.VisSpellTimer.Enabled = True
    frmEndieko.picSpells.Picture = LoadPicture(App.Path & "\Graphics\spellicons.bmp")
    frmSpriteChange.picSprites.Picture = LoadPicture(App.Path & "\Graphics\sprites.bmp")
    
    Call SetStatus("Initializing TCP Settings...")
    DoEvents
    
    Call TcpInit
    frmMainMenu.Visible = True
    frmSendGetData.Visible = False
End Sub

Sub SetStatus(ByVal Caption As String)
    frmSendGetData.lblStatus.Caption = Caption
End Sub

Sub MenuState(ByVal State As Long)
    Connucted = True
    frmSendGetData.Visible = True
    Call SetStatus("Connecting to server...")
    Select Case State
        Case MENU_STATE_NEWACCOUNT
            frmNewAccount.Visible = False
            If ConnectToServer = True Then
                Call SetStatus("Connected, sending new account information...")
                Call SendNewAccount(frmNewAccount.txtName.Text, frmNewAccount.txtPassword.Text)
            End If
            
        Case MENU_STATE_DELACCOUNT
            frmDeleteAccount.Visible = False
            If ConnectToServer = True Then
                Call SetStatus("Connected, sending account deletion request ...")
                Call SendDelAccount(frmDeleteAccount.txtName.Text, frmDeleteAccount.txtPassword.Text)
            End If
        
        Case MENU_STATE_LOGIN
            frmLogin.Visible = False
            If ConnectToServer = True Then
                Call SetStatus("Connected, sending login information...")
                Call SendLogin(frmLogin.txtName.Text, frmLogin.txtPassword.Text)
            End If
        
        Case MENU_STATE_NEWCHAR
            frmChars.Visible = False
            Call SetStatus("Connected, getting available classes...")
            Call SendGetClasses
            
        Case MENU_STATE_ADDCHAR
            frmNewChar.Visible = False
            If ConnectToServer = True Then
                Call SetStatus("Connected, sending character addition data...")
                If frmNewChar.optMale.Value = True Then
                    Call SendAddChar(frmNewChar.txtName, 0, 0, 1)
                Else
                    Call SendAddChar(frmNewChar.txtName, 1, 0, 1)
                End If
            End If
        
        Case MENU_STATE_DELCHAR
            frmChars.Visible = False
            If ConnectToServer = True Then
                Call SetStatus("Connected, sending character deletion request...")
                Call SendDelChar(1)
            End If
            
        Case MENU_STATE_USECHAR
            frmChars.Visible = False
            If ConnectToServer = True Then
                Call SetStatus("Connected, sending char data...")
                Call SendUseChar(1)
            End If
    End Select

    If Not IsConnected And Connucted = True Then
        frmMainMenu.Visible = True
        frmSendGetData.Visible = False
        Call MsgBox("Sorry, the server seems to be down.  Please try to reconnect in a few minutes or visit " & WEBSITE, vbOKOnly, GAME_NAME)
    End If
End Sub
Sub GameInit()
    frmEndieko.Visible = True
    frmSendGetData.Visible = False
    Call InitDirectX
End Sub

Sub GameLoop()
Dim Tick As Long
Dim TickFPS As Long
Dim FPS As Long
Dim X As Long
Dim Y As Long
Dim i As Long
Dim rec_back As RECT
    
    ' Set the focus
    frmEndieko.picScreen.SetFocus
    
    ' Set font
    Call SetFont("Fixedsys", 18)
    'Call SetFont("Franklin Gothic Demi", 20)
                
    ' Used for calculating fps
    TickFPS = GetTickCount
    FPS = 0
    
    Do While InGame
        Tick = GetTickCount
        
        ' Check to make sure they aren't trying to auto do anything
        If GetAsyncKeyState(VK_UP) >= 0 And DirUp = True Then DirUp = False
        If GetAsyncKeyState(VK_DOWN) >= 0 And DirDown = True Then DirDown = False
        If GetAsyncKeyState(VK_LEFT) >= 0 And DirLeft = True Then DirLeft = False
        If GetAsyncKeyState(VK_RIGHT) >= 0 And DirRight = True Then DirRight = False
        If GetAsyncKeyState(VK_CONTROL) >= 0 And ControlDown = True Then ControlDown = False
        If GetAsyncKeyState(VK_SHIFT) >= 0 And ShiftDown = True Then ShiftDown = False
        
        ' Check to make sure we are still connected
        If Not IsConnected Then InGame = False
        
        ' Check if we need to restore surfaces
        If NeedToRestoreSurfaces Then
            DD.RestoreAllSurfaces
            Call InitSurfaces
        End If
                
        If GettingMap = False Then
        ' Visual Inventory
        Dim Q As Long
        Dim Qq As Long
        Dim IT As Long
               
        If GetTickCount > IT + 500 And frmEndieko.picInv3.Visible = True Then
            For Q = 0 To MAX_INV - 1
                Qq = Player(MyIndex).Inv(Q + 1).Num

                If frmEndieko.picInv(Q).Picture <> LoadPicture() Then
                    frmEndieko.picInv(Q).Picture = LoadPicture()
                Else
                    If Qq = 0 Then
                        frmEndieko.picInv(Q).Picture = LoadPicture()
                    Else
                        Call BitBlt(frmEndieko.picInv(Q).hDC, 0, 0, PIC_X, PIC_Y, frmEndieko.picItems.hDC, (Item(Qq).Pic - Int(Item(Qq).Pic / 6) * 6) * PIC_X, Int(Item(Qq).Pic / 6) * PIC_Y, SRCCOPY)
                    End If
                End If
            Next Q
        End If
                       
        If GetTickCount > IT + 500 And frmEndieko.picPlayerSpells.Visible = True Then
            For Q = 1 To MAX_PLAYER_SPELLS - 1
                Qq = Player(MyIndex).Spell(Q + 1)
               
                If frmEndieko.picSpell(Q).Picture <> LoadPicture() Then
                    frmEndieko.picSpell(Q).Picture = LoadPicture()
                Else
                    If Qq = 0 Then
                        frmEndieko.picSpell(Q).Picture = LoadPicture()
                    Else
                        Call BitBlt(frmEndieko.picSpell(Q).hDC, 0, 0, PIC_X, PIC_Y, frmEndieko.picSpells.hDC, (Spell(Qq).Pic - Int(Spell(Qq).Pic / 6) * 6) * PIC_X, Int(Spell(Qq).Pic / 6) * PIC_Y, SRCCOPY)
                    End If
                End If
            Next Q
        End If
                        
        NewX = 10
        NewY = 7

        NewPlayerY = Player(MyIndex).Y - NewY
        NewPlayerX = Player(MyIndex).X - NewX

        NewX = NewX * PIC_X
        NewY = NewY * PIC_Y

        NewXOffset = Player(MyIndex).XOffset
        NewYOffset = Player(MyIndex).YOffset

        If Player(MyIndex).Y - 7 < 1 Then
            NewY = Player(MyIndex).Y * PIC_Y + Player(MyIndex).YOffset
            NewYOffset = 0
            NewPlayerY = 0
            
            If Player(MyIndex).Y = 7 And Player(MyIndex).Dir = DIR_UP Then
                NewPlayerY = Player(MyIndex).Y - 7
                NewY = 7 * PIC_Y
                NewYOffset = Player(MyIndex).YOffset
            End If
        ElseIf Player(MyIndex).Y + 9 > MAX_MAPY + 1 Then
            NewY = (Player(MyIndex).Y - 16) * PIC_Y + Player(MyIndex).YOffset
            NewYOffset = 0
            NewPlayerY = MAX_MAPY - 14
            
            If Player(MyIndex).Y = 23 And Player(MyIndex).Dir = DIR_DOWN Then
                NewPlayerY = Player(MyIndex).Y - 7
                NewY = 7 * PIC_Y
                NewYOffset = Player(MyIndex).YOffset
            End If
        End If

        If Player(MyIndex).X - 10 < 1 Then
            NewX = Player(MyIndex).X * PIC_X + Player(MyIndex).XOffset
            NewXOffset = 0
            NewPlayerX = 0
            
            If Player(MyIndex).X = 10 And Player(MyIndex).Dir = DIR_LEFT Then
                NewPlayerX = Player(MyIndex).X - 10
                NewX = 10 * PIC_X
                NewXOffset = Player(MyIndex).XOffset
            End If
        ElseIf Player(MyIndex).X + 11 > MAX_MAPX + 1 Then
            NewX = (Player(MyIndex).X - 11) * PIC_X + Player(MyIndex).XOffset
            NewXOffset = 0
            NewPlayerX = MAX_MAPX - 19
            
            If Player(MyIndex).X = 21 And Player(MyIndex).Dir = DIR_RIGHT Then
                NewPlayerX = Player(MyIndex).X - 10
                NewX = 10 * PIC_X
                NewXOffset = Player(MyIndex).XOffset
            End If
        End If
        
        sx = 32
        If MAX_MAPX = 19 Then
            NewX = Player(MyIndex).X * PIC_X + Player(MyIndex).XOffset
            NewXOffset = 0
            NewPlayerX = 0
            NewY = Player(MyIndex).Y * PIC_Y + Player(MyIndex).YOffset
            NewYOffset = 0
            NewPlayerY = 0
            sx = 0
        End If
        
        ' Blit out tiles layers ground/anim1/anim2
        For Y = 0 To MAX_MAPY
            For X = 0 To MAX_MAPX
                Call BltTile(X, Y)
            Next X
        Next Y
       
        ' Blit out the items
        For i = 1 To MAX_MAP_ITEMS
            If MapItem(i).Num > 0 Then
                Call BltItem(i)
            End If
        Next i
        
    If ScreenMode = 0 Then
        ' Blit players bar
        If ReadINI("CONFIG", "PlayerBar", App.Path & "\config.ini") = 1 Then
            Call BltPlayerBars(MyIndex)
        End If

        
        ' Blit out the sprite change attribute
        For Y = 0 To MAX_MAPY
            For X = 0 To MAX_MAPX
                Call BltSpriteChange(X, Y)
            Next X
        Next Y

        ' Blit out players
'        For i = 1 To MAX_PLAYERS
'            If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
'                Call BltPlayer(i)
'            End If
'        Next i

        ' Blit out players
        For i = 1 To MAX_PLAYERS
            If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                Call BltPlayer(i)
            End If
        Next i
        
        ' Blit out the npcs
        For i = 1 To MAX_MAP_NPCS
            Call BltNpc(i)
        Next i
        
        ' Blit out the npcs
        For i = 1 To MAX_MAP_NPCS
            Call BltNpcTop(i)
        Next i
        
        ' Blit out arrows
        For i = 1 To MAX_PLAYERS
             If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                 Call BltArrow(i)
             End If
        Next i
        
        ' Blit out the sprite change attribute
        For Y = 0 To MAX_MAPY
            For X = 0 To MAX_MAPX
                Call BltSpriteChange2(X, Y)
            Next X
        Next Y
    End If
                
        ' Blit out tile layer fringe
        For Y = 0 To MAX_MAPY
            For X = 0 To MAX_MAPX
                Call BltFringeTile(X, Y)
            Next X
        Next Y
        
        ' Blit Out Night
        If GettingMap = False Then
            If GameTime = TIME_NIGHT And InEditor = False Then
                Call Night
            End If
        End If

        If InEditor = True And ReadINI("CONFIG", "MapGrid", App.Path & "\config.ini") = 1 Then
            For Y = 0 To MAX_MAPY
                For X = 0 To MAX_MAPX
                    Call BltTile2(X * 32, Y * 32, 1)
                Next X
            Next Y
        End If
    End If

        ' Lock the backbuffer so we can draw text and names
        TexthDC = DD_BackBuffer.GetDC
        
    If GettingMap = False Then
    If ScreenMode = 0 Then
    
        ' Blt Overhead Damage Done to Players
        If ReadINI("CONFIG", "PlayerName", App.Path & "\config.ini") = 0 Then
            If GetTickCount < NPCDmgTime + 2000 Then
                Call DrawText(TexthDC, (Int(Len(NPCDmgDamage)) / 2) * 3 + NewX + sx, NewY - 22 - ii + sx, NPCDmgDamage, QBColor(White))
            End If
        Else
            If GetPlayerGuild(MyIndex) <> vbNullString Then
                If GetTickCount < NPCDmgTime + 2000 Then
                    Call DrawText(TexthDC, (Int(Len(NPCDmgDamage)) / 2) * 3 + NewX + sx, NewY - 42 - ii + sx, NPCDmgDamage, QBColor(White))
                End If
            Else
                If GetTickCount < NPCDmgTime + 2000 Then
                    Call DrawText(TexthDC, (Int(Len(NPCDmgDamage)) / 2) * 3 + NewX + sx, NewY - 22 - ii + sx, NPCDmgDamage, QBColor(White))
                End If
            End If
        End If
        ii = ii + 1
        
        ' Blt Overhead Damage Done to NPCs
        If NPCWho > 0 Then
            If MapNpc(NPCWho).Num > 0 Then
                If ReadINI("CONFIG", "NPCName", App.Path & "\config.ini") = 0 Then
                    If Npc(MapNpc(NPCWho).Num).Big = 0 Then
                        If GetTickCount < DmgTime + 2000 Then
                            Call DrawText(TexthDC, (MapNpc(NPCWho).X - NewPlayerX) * PIC_X + sx + (Int(Len(DmgDamage)) / 2) * 3 + MapNpc(NPCWho).XOffset - NewXOffset, (MapNpc(NPCWho).Y - NewPlayerY) * PIC_Y + sx - 20 + MapNpc(NPCWho).YOffset - NewYOffset - iii, DmgDamage, QBColor(BrightRed))
                        End If
                    Else
                        If GetTickCount < DmgTime + 2000 Then
                            Call DrawText(TexthDC, (MapNpc(NPCWho).X - NewPlayerX) * PIC_X + sx + (Int(Len(DmgDamage)) / 2) * 3 + MapNpc(NPCWho).XOffset - NewXOffset, (MapNpc(NPCWho).Y - NewPlayerY) * PIC_Y + sx - 47 + MapNpc(NPCWho).YOffset - NewYOffset - iii, DmgDamage, QBColor(BrightRed))
                        End If
                    End If
                Else
                    If Npc(MapNpc(NPCWho).Num).Big = 0 Then
                        If GetTickCount < DmgTime + 2000 Then
                            Call DrawText(TexthDC, (MapNpc(NPCWho).X - NewPlayerX) * PIC_X + sx + (Int(Len(DmgDamage)) / 2) * 3 + MapNpc(NPCWho).XOffset - NewXOffset, (MapNpc(NPCWho).Y - NewPlayerY) * PIC_Y + sx - 30 + MapNpc(NPCWho).YOffset - NewYOffset - iii, DmgDamage, QBColor(BrightRed))
                        End If
                    Else
                        If GetTickCount < DmgTime + 2000 Then
                            Call DrawText(TexthDC, (MapNpc(NPCWho).X - NewPlayerX) * PIC_X + sx + (Int(Len(DmgDamage)) / 2) * 3 + MapNpc(NPCWho).XOffset - NewXOffset, (MapNpc(NPCWho).Y - NewPlayerY) * PIC_Y + sx - 57 + MapNpc(NPCWho).YOffset - NewYOffset - iii, DmgDamage, QBColor(BrightRed))
                        End If
                    End If
                End If
                iii = iii + 1
            End If
        End If
             
        ' Blt Overhead Msgs over Self
        If ReadINI("CONFIG", "PlayerName", App.Path & "\config.ini") = 0 Then
            If GetTickCount < Overhead.Time + 2000 Then
                Call DrawText(TexthDC, (Int(Len(Overhead.Msg)) / 2) * 3 + NewX + sx, NewY - 22 - Overhead.ii + sx, Overhead.Msg, QBColor(Overhead.Color))
            End If
        Else
            If GetPlayerGuild(MyIndex) <> vbNullString Then
                If GetTickCount < Overhead.Time + 2000 Then
                    Call DrawText(TexthDC, (Int(Len(Overhead.Msg)) / 2) * 3 + NewX + sx, NewY - 42 - Overhead.ii + sx, Overhead.Msg, QBColor(Overhead.Color))
                End If
            Else
                If GetTickCount < Overhead.Time + 2000 Then
                    Call DrawText(TexthDC, (Int(Len(Overhead.Msg)) / 2) * 3 + NewX + sx, NewY - 22 - Overhead.ii + sx, Overhead.Msg, QBColor(Overhead.Color))
                End If
            End If
        End If
        Overhead.ii = Overhead.ii + 1
            
        If ReadINI("CONFIG", "PlayerName", App.Path & "\config.ini") = 1 Then
            For i = 1 To MAX_PLAYERS
                 If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                    If GetPlayerSprite(i) <> 1000 Then
                        Call BltPlayerName(i)
                    End If
                 End If
            Next i
        End If

'        'Draw NPC Names
'        If ReadINI("CONFIG", "NPCName", App.Path & "\config.ini") = 1 Then
'            For i = LBound(MapNpc) To UBound(MapNpc)
'                If MapNpc(i).Num > 0 Then
'                    Call BltMapNPCName(i)
'                End If
'            Next i
'        End If
                
        ' Blit out attribs if in editor
        If InEditor Then
            For Y = 0 To MAX_MAPY
                For X = 0 To MAX_MAPX
                    With Map.Tile(X, Y)
                        If .Type = TILE_TYPE_BLOCKED Then Call DrawText(TexthDC, X * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, Y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "B", QBColor(BrightRed))
                        If .Type = TILE_TYPE_WARP Then Call DrawText(TexthDC, X * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, Y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "W", QBColor(BrightBlue))
                        If .Type = TILE_TYPE_ITEM Then Call DrawText(TexthDC, X * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, Y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "I", QBColor(White))
                        If .Type = TILE_TYPE_NPCAVOID Then Call DrawText(TexthDC, X * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, Y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "N", QBColor(White))
                        If .Type = TILE_TYPE_KEY Then Call DrawText(TexthDC, X * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, Y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "K", QBColor(White))
                        If .Type = TILE_TYPE_KEYOPEN Then Call DrawText(TexthDC, X * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, Y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "O", QBColor(White))
                        If .Type = TILE_TYPE_HEAL Then Call DrawText(TexthDC, X * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, Y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "H", QBColor(BrightGreen))
                        If .Type = TILE_TYPE_KILL Then Call DrawText(TexthDC, X * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, Y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "K", QBColor(BrightRed))
                        If .Type = TILE_TYPE_SHOP Then Call DrawText(TexthDC, X * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, Y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "S", QBColor(Yellow))
                        If .Type = TILE_TYPE_CBLOCK Then Call DrawText(TexthDC, X * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, Y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "CB", QBColor(Black))
                        If .Type = TILE_TYPE_ARENA Then Call DrawText(TexthDC, X * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, Y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "A", QBColor(BrightGreen))
                        If .Type = TILE_TYPE_SOUND Then Call DrawText(TexthDC, X * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, Y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "PS", QBColor(Yellow))
                        If .Type = TILE_TYPE_SPRITE_CHANGE Then Call DrawText(TexthDC, X * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, Y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "SC", QBColor(Grey))
                        If .Type = TILE_TYPE_SIGN Then Call DrawText(TexthDC, X * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, Y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "SI", QBColor(Yellow))
                        If .Type = TILE_TYPE_DOOR Then Call DrawText(TexthDC, X * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, Y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "D", QBColor(Black))
                        If .Type = TILE_TYPE_NOTICE Then Call DrawText(TexthDC, X * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, Y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "N", QBColor(BrightGreen))
                        If .Type = TILE_TYPE_CHEST Then Call DrawText(TexthDC, X * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, Y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "C", QBColor(Brown))
                        If .Type = TILE_TYPE_CLASS_CHANGE Then Call DrawText(TexthDC, X * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, Y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "CG", QBColor(White))
                        If .Type = TILE_TYPE_SCRIPTED Then Call DrawText(TexthDC, X * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, Y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "SC", QBColor(Yellow))
                        If .Type = TILE_TYPE_BANK Then Call DrawText(TexthDC, X * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, Y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "Bank", QBColor(White))
                    End With
                Next X
            Next Y
        End If

        ' Blit the text they are putting in
        frmEndieko.txtMyTextBox.Text = MyText
        
        If Len(MyText) > 4 Then
            frmEndieko.txtMyTextBox.SelStart = Len(frmEndieko.txtMyTextBox.Text) + 1
        End If
                
        ' Draw map name
        If Map.Moral = MAP_MORAL_NONE Then
            'Call DrawText(TexthDC, Int((20.5) * PIC_X / 2) - (Int(Len(Trim$(Map.Name)) / 2) * 8) + sx, 2 + sx, Trim$(Map.Name), QBColor(BrightRed))
            frmEndieko.lblMapName.Caption = "[" & Trim$(Map.Name) & "]"
            frmEndieko.lblMapName.ForeColor = QBColor(BrightRed)
        ElseIf Map.Moral = MAP_MORAL_SAFE Then
            'Call DrawText(TexthDC, Int((20.5) * PIC_X / 2) - (Int(Len(Trim$(Map.Name)) / 2) * 8) + sx, 2 + sx, Trim$(Map.Name), QBColor(BrightCyan))
            frmEndieko.lblMapName.Caption = "[" & Trim$(Map.Name) & "]"
            frmEndieko.lblMapName.ForeColor = QBColor(BrightCyan)
        ElseIf Map.Moral = MAP_MORAL_NO_PENALTY Then
            'Call DrawText(TexthDC, Int((20.5) * PIC_X / 2) - (Int(Len(Trim$(Map.Name)) / 2) * 8) + sx, 2 + sx, Trim$(Map.Name), QBColor(Black))
            frmEndieko.lblMapName.Caption = "[" & Trim$(Map.Name) & "]"
            frmEndieko.lblMapName.ForeColor = QBColor(Black)
        End If
    End If
    End If

        ' Check if we are getting a map, and if we are tell them so
        If GettingMap = True Then
            Call DrawText(TexthDC, 36, 36, "Receiving Map...", QBColor(BrightCyan))
        End If
                        
        ' Release DC
        Call DD_BackBuffer.ReleaseDC(TexthDC)
        
        ' Blit out emoticons
        For i = 1 To MAX_PLAYERS
            If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                Call BltEmoticons(i)
            End If
        Next i
        
        ' Get the rect for the back buffer to blit from
        rec.Top = 0
        rec.Bottom = (MAX_MAPY + 1) * PIC_Y
        rec.Left = 0
        rec.Right = (MAX_MAPX + 1) * PIC_X
        
        ' Get the rect to blit to
        Call DX.GetWindowRect(frmEndieko.picScreen.hWnd, rec_pos)
        rec_pos.Bottom = rec_pos.Top - sx + ((MAX_MAPY + 1) * PIC_Y)
        rec_pos.Right = rec_pos.Left - sx + ((MAX_MAPX + 1) * PIC_X)
        rec_pos.Top = rec_pos.Bottom - ((MAX_MAPY + 1) * PIC_Y)
        rec_pos.Left = rec_pos.Right - ((MAX_MAPX + 1) * PIC_X)
        
        ' Blit the backbuffer
        Call DD_PrimarySurf.Blt(rec_pos, DD_BackBuffer, rec, DDBLT_WAIT)
        
        ' Check if player is trying to move
        Call CheckMovement
        
        ' Check to see if player is trying to attack
        Call CheckAttack
        
        ' Process player movements (actually move them)
        For i = 1 To MAX_PLAYERS
            If IsPlaying(i) Then
                Call ProcessMovement(i)
            End If
        Next i
        
        ' Process npc movements (actually move them)
        For i = 1 To MAX_MAP_NPCS
            If Map.Npc(i) > 0 Then
                Call ProcessNpcMovement(i)
            End If
        Next i
  
        ' Change map animation every 250 milliseconds
        If GetTickCount > MapAnimTimer + 250 Then
            If MapAnim = 0 Then
                MapAnim = 1
            Else
                MapAnim = 0
            End If
            MapAnimTimer = GetTickCount
        End If
        
        ' Change pk label every 150 milliseconds
        If GetTickCount > PKTimer + 150 Then
            If PKSet = 0 Then
                PKSet = 1
            Else
                PKSet = 0
            End If
            PKTimer = GetTickCount
        End If
                
        ' Lock fps
        Do While GetTickCount < Tick + 30
            DoEvents
        Loop
        
        ' Calculate fps
        If GetTickCount > TickFPS + 1000 Then
            GameFPS = FPS
            TickFPS = GetTickCount
            FPS = 0
        Else
            FPS = FPS + 1
        End If
        
        Call MakeMidiLoop
        
        DoEvents
    Loop
    
    frmEndieko.Visible = False
    frmSendGetData.Visible = True
    Call SetStatus("Destroying game data...")
    
    ' Shutdown the game
    Call GameDestroy
    
    ' Report disconnection if server disconnects
    If IsConnected = False Then
        Call MsgBox("Thank you for playing " & GAME_NAME & "!", vbOKOnly, GAME_NAME)
    End If
End Sub

Sub GameDestroy()
    Call DestroyDirectX
    Call StopMidi
    End
End Sub

Sub BltTile(ByVal X As Long, ByVal Y As Long)
Dim Ground As Long
Dim Anim1 As Long
Dim Anim2 As Long
Dim Mask2 As Long
Dim M2Anim As Long

    Ground = Map.Tile(X, Y).Ground
    Anim1 = Map.Tile(X, Y).Mask
    Anim2 = Map.Tile(X, Y).Anim
    Mask2 = Map.Tile(X, Y).Mask2
    M2Anim = Map.Tile(X, Y).M2Anim
    
    ' Only used if ever want to switch to blt rather then bltfast
    With rec_pos
        .Top = (Y - NewPlayerY) * PIC_Y + sx - NewYOffset
        .Bottom = .Top + PIC_Y
        .Left = (X - NewPlayerX) * PIC_X + sx - NewXOffset
        .Right = .Left + PIC_X
    End With
    
    rec.Top = Int(Ground / 14) * PIC_Y
    rec.Bottom = rec.Top + PIC_Y
    rec.Left = (Ground - Int(Ground / 14) * 14) * PIC_X
    rec.Right = rec.Left + PIC_X
    Call DD_BackBuffer.Blt(rec_pos, DD_TileSurf, rec, DDBLT_WAIT)
    'Call DD_BackBuffer.BltFast((X - NewPlayerX) * PIC_X - NewXOffset, (Y - NewPlayerY) * PIC_Y - NewYOffset, DD_TileSurf, rec, DDBLTFAST_WAIT)
    
    If (MapAnim = 0) Or (Anim2 <= 0) Then
        ' Is there an animation tile to plot?
        If Anim1 > 0 And TempTile(X, Y).DoorOpen = NO Then
            rec.Top = Int(Anim1 / 14) * PIC_Y
            rec.Bottom = rec.Top + PIC_Y
            rec.Left = (Anim1 - Int(Anim1 / 14) * 14) * PIC_X
            rec.Right = rec.Left + PIC_X
            Call DD_BackBuffer.Blt(rec_pos, DD_TileSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
            'Call DD_BackBuffer.BltFast((X - NewPlayerX) * PIC_X - NewXOffset, (Y - NewPlayerY) * PIC_Y - NewYOffset, DD_TileSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    Else
        ' Is there a second animation tile to plot?
        If Anim2 > 0 Then
            rec.Top = Int(Anim2 / 14) * PIC_Y
            rec.Bottom = rec.Top + PIC_Y
            rec.Left = (Anim2 - Int(Anim2 / 14) * 14) * PIC_X
            rec.Right = rec.Left + PIC_X
            Call DD_BackBuffer.Blt(rec_pos, DD_TileSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
            'Call DD_BackBuffer.BltFast((X - NewPlayerX) * PIC_X - NewXOffset, (Y - NewPlayerY) * PIC_Y - NewYOffset, DD_TileSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    End If
    
    If (MapAnim = 0) Or (M2Anim <= 0) Then
        ' Is there an animation tile to plot?
        If Mask2 > 0 Then
            rec.Top = Int(Mask2 / 14) * PIC_Y
            rec.Bottom = rec.Top + PIC_Y
            rec.Left = (Mask2 - Int(Mask2 / 14) * 14) * PIC_X
            rec.Right = rec.Left + PIC_X
            Call DD_BackBuffer.Blt(rec_pos, DD_TileSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
            'Call DD_BackBuffer.BltFast((X - NewPlayerX) * PIC_X - NewXOffset, (Y - NewPlayerY) * PIC_Y - NewYOffset, DD_TileSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    Else
        ' Is there a second animation tile to plot?
        If M2Anim > 0 Then
            rec.Top = Int(M2Anim / 14) * PIC_Y
            rec.Bottom = rec.Top + PIC_Y
            rec.Left = (M2Anim - Int(M2Anim / 14) * 14) * PIC_X
            rec.Right = rec.Left + PIC_X
            Call DD_BackBuffer.Blt(rec_pos, DD_TileSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
            'Call DD_BackBuffer.BltFast((X - NewPlayerX) * PIC_X - NewXOffset, (Y - NewPlayerY) * PIC_Y - NewYOffset, DD_TileSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    End If
End Sub

Sub BltItem(ByVal ItemNum As Long)
    ' Only used if ever want to switch to blt rather then bltfast
    With rec_pos
        .Top = MapItem(ItemNum).Y * PIC_Y
        .Bottom = .Top + PIC_Y
        .Left = MapItem(ItemNum).X * PIC_X
        .Right = .Left + PIC_X
    End With
    
    rec.Top = Int(Item(MapItem(ItemNum).Num).Pic / 6) * PIC_Y
    rec.Bottom = rec.Top + PIC_Y
    rec.Left = (Item(MapItem(ItemNum).Num).Pic - Int(Item(MapItem(ItemNum).Num).Pic / 6) * 6) * PIC_X
    rec.Right = rec.Left + PIC_X
    
    'Call DD_BackBuffer.Blt(rec_pos, DD_ItemSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
    Call DD_BackBuffer.BltFast((MapItem(ItemNum).X - NewPlayerX) * PIC_X + sx - NewXOffset, (MapItem(ItemNum).Y - NewPlayerY) * PIC_Y + sx - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End Sub

Sub BltFringeTile(ByVal X As Long, ByVal Y As Long)
Dim Fringe As Long
Dim FAnim As Long
Dim Fringe2 As Long
Dim F2Anim As Long

    ' Only used if ever want to switch to blt rather then bltfast
    With rec_pos
        .Top = Y * PIC_Y
        .Bottom = .Top + PIC_Y
        .Left = X * PIC_X
        .Right = .Left + PIC_X
    End With
    
    Fringe = Map.Tile(X, Y).Fringe
    FAnim = Map.Tile(X, Y).FAnim
    Fringe2 = Map.Tile(X, Y).Fringe2
    F2Anim = Map.Tile(X, Y).F2Anim
        
    If (MapAnim = 0) Or (FAnim <= 0) Then
        ' Is there an animation tile to plot?
        
        If Fringe > 0 Then
        rec.Top = Int(Fringe / 14) * PIC_Y
        rec.Bottom = rec.Top + PIC_Y
        rec.Left = (Fringe - Int(Fringe / 14) * 14) * PIC_X
        rec.Right = rec.Left + PIC_X
        'Call DD_BackBuffer.Blt(rec_pos, DD_TileSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
        Call DD_BackBuffer.BltFast((X - NewPlayerX) * PIC_X + sx - NewXOffset, (Y - NewPlayerY) * PIC_Y + sx - NewYOffset, DD_TileSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    
    Else
    
        If FAnim > 0 Then
        rec.Top = Int(FAnim / 14) * PIC_Y
        rec.Bottom = rec.Top + PIC_Y
        rec.Left = (FAnim - Int(FAnim / 14) * 14) * PIC_X
        rec.Right = rec.Left + PIC_X
        'Call DD_BackBuffer.Blt(rec_pos, DD_TileSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
        Call DD_BackBuffer.BltFast((X - NewPlayerX) * PIC_X + sx - NewXOffset, (Y - NewPlayerY) * PIC_Y + sx - NewYOffset, DD_TileSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
        
    End If

    If (MapAnim = 0) Or (F2Anim <= 0) Then
        ' Is there an animation tile to plot?
        
        If Fringe2 > 0 Then
        rec.Top = Int(Fringe2 / 14) * PIC_Y
        rec.Bottom = rec.Top + PIC_Y
        rec.Left = (Fringe2 - Int(Fringe2 / 14) * 14) * PIC_X
        rec.Right = rec.Left + PIC_X
        'Call DD_BackBuffer.Blt(rec_pos, DD_TileSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
        Call DD_BackBuffer.BltFast((X - NewPlayerX) * PIC_X + sx - NewXOffset, (Y - NewPlayerY) * PIC_Y + sx - NewYOffset, DD_TileSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    
    Else
    
        If F2Anim > 0 Then
        rec.Top = Int(F2Anim / 14) * PIC_Y
        rec.Bottom = rec.Top + PIC_Y
        rec.Left = (F2Anim - Int(F2Anim / 14) * 14) * PIC_X
        rec.Right = rec.Left + PIC_X
        'Call DD_BackBuffer.Blt(rec_pos, DD_TileSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
        Call DD_BackBuffer.BltFast((X - NewPlayerX) * PIC_X + sx - NewXOffset, (Y - NewPlayerY) * PIC_Y + sx - NewYOffset, DD_TileSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
        
    End If
End Sub

Sub BltPlayer(ByVal Index As Long)
Dim Anim As Byte
Dim X As Long, Y As Long

If Index = MyIndex Then

    ' Check for animation
    Anim = 0
    If Player(MyIndex).Attacking = 0 Then
        Select Case GetPlayerDir(MyIndex)
            Case DIR_UP
                If (Player(MyIndex).YOffset < PIC_Y / 2) Then Anim = 1
            Case DIR_DOWN
                If (Player(MyIndex).YOffset < PIC_Y / 2 * -1) Then Anim = 1
            Case DIR_LEFT
                If (Player(MyIndex).XOffset < PIC_Y / 2) Then Anim = 1
            Case DIR_RIGHT
                If (Player(MyIndex).XOffset < PIC_Y / 2 * -1) Then Anim = 1
        End Select
    Else
        If Player(MyIndex).AttackTimer + 500 > GetTickCount Then
            Anim = 2
        End If
    End If

    ' Check to see if we want to stop making him attack
    If Player(MyIndex).AttackTimer + 1000 < GetTickCount Then
        Player(MyIndex).Attacking = 0
        Player(MyIndex).AttackTimer = 0
    End If

    rec.Top = GetPlayerSprite(MyIndex) * PIC_Y
    rec.Bottom = rec.Top + PIC_Y
    rec.Left = (GetPlayerDir(MyIndex) * 3 + Anim) * PIC_X
    rec.Right = rec.Left + PIC_X

    X = NewX + sx
    Y = NewY + sx

    Call DD_BackBuffer.BltFast(X, Y, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
Else
    ' Only used if ever want to switch to blt rather then bltfast
    With rec_pos
        .Top = GetPlayerY(Index) * PIC_Y + Player(Index).YOffset
        .Bottom = .Top + PIC_Y
        .Left = GetPlayerX(Index) * PIC_X + Player(Index).XOffset
        .Right = .Left + PIC_X
    End With

    ' Check for animation
    Anim = 0
    If Player(Index).Attacking = 0 Then
        Select Case GetPlayerDir(Index)
            Case DIR_UP
                If (Player(Index).YOffset < PIC_Y / 2) Then Anim = 1
            Case DIR_DOWN
                If (Player(Index).YOffset < PIC_Y / 2 * -1) Then Anim = 1
            Case DIR_LEFT
                If (Player(Index).XOffset < PIC_Y / 2) Then Anim = 1
            Case DIR_RIGHT
                If (Player(Index).XOffset < PIC_Y / 2 * -1) Then Anim = 1
        End Select
    Else
        If Player(Index).AttackTimer + 500 > GetTickCount Then
            Anim = 2
        End If
    End If

    ' Check to see if we want to stop making him attack
    If Player(Index).AttackTimer + 1000 < GetTickCount Then
        Player(Index).Attacking = 0
        Player(Index).AttackTimer = 0
    End If

    rec.Top = GetPlayerSprite(Index) * PIC_Y
    rec.Bottom = rec.Top + PIC_Y
    rec.Left = (GetPlayerDir(Index) * 3 + Anim) * PIC_X
    rec.Right = rec.Left + PIC_X

    X = GetPlayerX(Index) * PIC_X + sx + Player(Index).XOffset
    Y = GetPlayerY(Index) * PIC_Y + sx + Player(Index).YOffset '- 4

    ' Check if its out of bounds because of the offset
    If Y < 0 Then
        Y = 0
        rec.Top = rec.Top + (Y * -1)
    End If

    'Call DD_BackBuffer.Blt(rec_pos, DD_SpriteSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
    Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End If
End Sub

Sub BltMapNPCName(ByVal Index As Long)
Dim TextX As Long
Dim TextY As Long

If Npc(MapNpc(Index).Num).Big = 0 Then
    With Npc(MapNpc(Index).Num)
    'Draw name
        TextX = MapNpc(Index).X * PIC_X + sx + MapNpc(Index).XOffset + CLng(PIC_X / 2) - ((Len(Trim$(.Name)) / 2) * 8)
        TextY = MapNpc(Index).Y * PIC_Y + sx + MapNpc(Index).YOffset - CLng(PIC_Y / 2) - 4
        DrawPlayerNameText TexthDC, TextX - (NewPlayerX * PIC_X) - NewXOffset, TextY - (NewPlayerY * PIC_Y) - NewYOffset, Trim$(.Name), vbWhite
    End With
Else
    With Npc(MapNpc(Index).Num)
    'Draw name
        TextX = MapNpc(Index).X * PIC_X + sx + MapNpc(Index).XOffset + CLng(PIC_X / 2) - ((Len(Trim$(.Name)) / 2) * 8)
        TextY = MapNpc(Index).Y * PIC_Y + sx + MapNpc(Index).YOffset - CLng(PIC_Y / 2) - 32
        DrawPlayerNameText TexthDC, TextX - (NewPlayerX * PIC_X) - NewXOffset, TextY - (NewPlayerY * PIC_Y) - NewYOffset, Trim$(.Name), vbWhite
    End With
End If
End Sub

Sub BltNpc(ByVal MapNpcNum As Long)
Dim Anim As Byte
Dim X As Long, Y As Long

    ' Make sure that theres an npc there, and if not exit the sub
    If MapNpc(MapNpcNum).Num <= 0 Then
        Exit Sub
    End If
    
    ' Only used if ever want to switch to blt rather then bltfast
    With rec_pos
        .Top = MapNpc(MapNpcNum).Y * PIC_Y + MapNpc(MapNpcNum).YOffset
        .Bottom = .Top + PIC_Y
        .Left = MapNpc(MapNpcNum).X * PIC_X + MapNpc(MapNpcNum).XOffset
        .Right = .Left + PIC_X
    End With
    
    ' Check for animation
    Anim = 0
    If MapNpc(MapNpcNum).Attacking = 0 Then
        Select Case MapNpc(MapNpcNum).Dir
            Case DIR_UP
                If (MapNpc(MapNpcNum).YOffset < PIC_Y / 2) Then Anim = 1
            Case DIR_DOWN
                If (MapNpc(MapNpcNum).YOffset < PIC_Y / 2 * -1) Then Anim = 1
            Case DIR_LEFT
                If (MapNpc(MapNpcNum).XOffset < PIC_Y / 2) Then Anim = 1
            Case DIR_RIGHT
                If (MapNpc(MapNpcNum).XOffset < PIC_Y / 2 * -1) Then Anim = 1
        End Select
    Else
        If MapNpc(MapNpcNum).AttackTimer + 500 > GetTickCount Then
            Anim = 2
        End If
    End If
    
    ' Check to see if we want to stop making him attack
    If MapNpc(MapNpcNum).AttackTimer + 1000 < GetTickCount Then
        MapNpc(MapNpcNum).Attacking = 0
        MapNpc(MapNpcNum).AttackTimer = 0
    End If
        
    If Npc(MapNpc(MapNpcNum).Num).Big = 0 Then
        rec.Top = Npc(MapNpc(MapNpcNum).Num).Sprite * 64 + PIC_Y
        rec.Bottom = rec.Top + PIC_Y
        rec.Left = (MapNpc(MapNpcNum).Dir * 3 + Anim) * PIC_X
        rec.Right = rec.Left + PIC_X
        
        X = MapNpc(MapNpcNum).X * PIC_X + sx + MapNpc(MapNpcNum).XOffset
        Y = MapNpc(MapNpcNum).Y * PIC_Y + sx + MapNpc(MapNpcNum).YOffset
        
        ' Check if its out of bounds because of the offset
        If Y < 0 Then
            Y = 0
            rec.Top = rec.Top + (Y * -1)
        End If
            
        'Call DD_BackBuffer.Blt(rec_pos, DD_SpriteSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
        Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    Else
        rec.Top = Npc(MapNpc(MapNpcNum).Num).Sprite * 64 + 32
        rec.Bottom = rec.Top + 32
        rec.Left = (MapNpc(MapNpcNum).Dir * 3 + Anim) * 64
        rec.Right = rec.Left + 64
    
        X = MapNpc(MapNpcNum).X * 32 + sx - 16 + MapNpc(MapNpcNum).XOffset
        Y = MapNpc(MapNpcNum).Y * 32 + sx + MapNpc(MapNpcNum).YOffset
  
        If Y < 0 Then
            rec.Top = Npc(MapNpc(MapNpcNum).Num).Sprite * 64 + 32
            rec.Bottom = rec.Top + 32
            Y = MapNpc(MapNpcNum).YOffset + sx
        End If
        
        If X < 0 Then
            rec.Left = (MapNpc(MapNpcNum).Dir * 3 + Anim) * 64 + 16
            rec.Right = rec.Left + 48
            X = MapNpc(MapNpcNum).XOffset + sx
        End If
        
        If X > MAX_MAPX * 32 Then
            rec.Left = (MapNpc(MapNpcNum).Dir * 3 + Anim) * 64
            rec.Right = rec.Left + 48
            X = MAX_MAPX * 32 + sx - 16 + MapNpc(MapNpcNum).XOffset
        End If

        Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_BigSpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    End If
End Sub

Sub BltNpcTop(ByVal MapNpcNum As Long)
Dim Anim As Byte
Dim X As Long, Y As Long

    ' Make sure that theres an npc there, and if not exit the sub
    If MapNpc(MapNpcNum).Num <= 0 Then
        Exit Sub
    End If
    
    If Npc(MapNpc(MapNpcNum).Num).Big = 0 Then Exit Sub
    
    ' Only used if ever want to switch to blt rather then bltfast
    With rec_pos
        .Top = MapNpc(MapNpcNum).Y * PIC_Y + MapNpc(MapNpcNum).YOffset
        .Bottom = .Top + PIC_Y
        .Left = MapNpc(MapNpcNum).X * PIC_X + MapNpc(MapNpcNum).XOffset
        .Right = .Left + PIC_X
    End With
    
    ' Check for animation
    Anim = 0
    If MapNpc(MapNpcNum).Attacking = 0 Then
        Select Case MapNpc(MapNpcNum).Dir
            Case DIR_UP
                If (MapNpc(MapNpcNum).YOffset < PIC_Y / 2) Then Anim = 1
            Case DIR_DOWN
                If (MapNpc(MapNpcNum).YOffset < PIC_Y / 2 * -1) Then Anim = 1
            Case DIR_LEFT
                If (MapNpc(MapNpcNum).XOffset < PIC_Y / 2) Then Anim = 1
            Case DIR_RIGHT
                If (MapNpc(MapNpcNum).XOffset < PIC_Y / 2 * -1) Then Anim = 1
        End Select
    Else
        If MapNpc(MapNpcNum).AttackTimer + 500 > GetTickCount Then
            Anim = 2
        End If
    End If
    
    ' Check to see if we want to stop making him attack
    If MapNpc(MapNpcNum).AttackTimer + 1000 < GetTickCount Then
        MapNpc(MapNpcNum).Attacking = 0
        MapNpc(MapNpcNum).AttackTimer = 0
    End If
    
    rec.Top = Npc(MapNpc(MapNpcNum).Num).Sprite * PIC_Y
        
     rec.Top = Npc(MapNpc(MapNpcNum).Num).Sprite * 64
     rec.Bottom = rec.Top + 32
     rec.Left = (MapNpc(MapNpcNum).Dir * 3 + Anim) * 64
     rec.Right = rec.Left + 64
 
     X = MapNpc(MapNpcNum).X * 32 + sx - 16 + MapNpc(MapNpcNum).XOffset
     Y = MapNpc(MapNpcNum).Y * 32 + sx - 32 + MapNpc(MapNpcNum).YOffset

     If Y < 0 Then
         rec.Top = Npc(MapNpc(MapNpcNum).Num).Sprite * 64 + 32
         rec.Bottom = rec.Top
         Y = MapNpc(MapNpcNum).YOffset + sx
     End If
     
     If X < 0 Then
         rec.Left = (MapNpc(MapNpcNum).Dir * 3 + Anim) * 64 + 16
         rec.Right = rec.Left + 48
         X = MapNpc(MapNpcNum).XOffset + sx
     End If
     
     If X > MAX_MAPX * 32 Then
         rec.Left = (MapNpc(MapNpcNum).Dir * 3 + Anim) * 64
         rec.Right = rec.Left + 48
         X = MAX_MAPX * 32 + sx - 16 + MapNpc(MapNpcNum).XOffset
     End If

     Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_BigSpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End Sub

Sub BltPlayerName(ByVal Index As Long)
Dim TextX As Long
Dim TextY As Long
Dim Color As Long
    
    ' Check access level
    If GetPlayerPK(Index) = NO Then
        Select Case GetPlayerAccess(Index)
            Case 0
                If GetPlayerGuild(Index) <> vbNullString Then
                    Color = QBColor(White)
                Else
                    Color = QBColor(Grey)
                End If
            Case 1
                Color = QBColor(BrightBlue)
            Case 2
                Color = QBColor(BrightBlue)
            Case 3
                Color = QBColor(BrightBlue)
            Case 4
                Color = QBColor(BrightBlue)
        End Select
    Else
        If PKSet = 0 Then
            Color = QBColor(Grey)
        Else
            Color = QBColor(Red)
        End If
    End If
        
    
If Index = MyIndex Then
    TextX = NewX + sx + Int(PIC_X / 2) - ((Len(GetPlayerName(MyIndex)) / 2) * 8)
    TextY = NewY + sx - Int(PIC_Y / 2)
    
    Call DrawText(TexthDC, TextX, TextY, GetPlayerName(MyIndex), Color)
Else
    ' Draw name
    TextX = GetPlayerX(Index) * PIC_X + sx + Player(Index).XOffset + Int(PIC_X / 2) - ((Len(GetPlayerName(Index)) / 2) * 8)
    TextY = GetPlayerY(Index) * PIC_Y + sx + Player(Index).YOffset - Int(PIC_Y / 2)
    Call DrawText(TexthDC, TextX - (NewPlayerX * PIC_X) - NewXOffset, TextY - (NewPlayerY * PIC_Y) - NewYOffset, GetPlayerName(Index), Color)
End If
End Sub


Sub ProcessMovement(ByVal Index As Long)
    ' Check if player is walking, and if so process moving them over
If Player(Index).Moving = MOVING_WALKING Then
        If Player(Index).Access > 0 Then
            Select Case GetPlayerDir(Index)
                Case DIR_UP
                    Player(Index).YOffset = Player(Index).YOffset - GM_WALK_SPEED
                Case DIR_DOWN
                    Player(Index).YOffset = Player(Index).YOffset + GM_WALK_SPEED
                Case DIR_LEFT
                    Player(Index).XOffset = Player(Index).XOffset - GM_WALK_SPEED
                Case DIR_RIGHT
                    Player(Index).XOffset = Player(Index).XOffset + GM_WALK_SPEED
            End Select
        Else
            Select Case GetPlayerDir(Index)
                Case DIR_UP
                    Player(Index).YOffset = Player(Index).YOffset - WALK_SPEED
                Case DIR_DOWN
                    Player(Index).YOffset = Player(Index).YOffset + WALK_SPEED
                Case DIR_LEFT
                    Player(Index).XOffset = Player(Index).XOffset - WALK_SPEED
                Case DIR_RIGHT
                    Player(Index).XOffset = Player(Index).XOffset + WALK_SPEED
            End Select
        End If
        
        ' Check if completed walking over to the next tile
        If (Player(Index).XOffset = 0) And (Player(Index).YOffset = 0) Then
            Player(Index).Moving = 0
        End If
    End If

    ' Check if player is running, and if so process moving them over
If Player(Index).Moving = MOVING_RUNNING Then
            If Player(Index).Access > 0 Then
            Select Case GetPlayerDir(Index)
                Case DIR_UP
                    Player(Index).YOffset = Player(Index).YOffset - GM_RUN_SPEED
                Case DIR_DOWN
                    Player(Index).YOffset = Player(Index).YOffset + GM_RUN_SPEED
                Case DIR_LEFT
                    Player(Index).XOffset = Player(Index).XOffset - GM_RUN_SPEED
                Case DIR_RIGHT
                    Player(Index).XOffset = Player(Index).XOffset + GM_RUN_SPEED
            End Select
        Else
            Select Case GetPlayerDir(Index)
                Case DIR_UP
                    Player(Index).YOffset = Player(Index).YOffset - RUN_SPEED
                Case DIR_DOWN
                    Player(Index).YOffset = Player(Index).YOffset + RUN_SPEED
                Case DIR_LEFT
                    Player(Index).XOffset = Player(Index).XOffset - RUN_SPEED
                Case DIR_RIGHT
                    Player(Index).XOffset = Player(Index).XOffset + RUN_SPEED
            End Select
        End If
        
        ' Check if completed walking over to the next tile
        If (Player(Index).XOffset = 0) And (Player(Index).YOffset = 0) Then
            Player(Index).Moving = 0
        End If
    End If
End Sub

Sub ProcessNpcMovement(ByVal MapNpcNum As Long)
    ' Check if npc is walking, and if so process moving them over
    If MapNpc(MapNpcNum).Moving = MOVING_WALKING Then
        Select Case MapNpc(MapNpcNum).Dir
            Case DIR_UP
                MapNpc(MapNpcNum).YOffset = MapNpc(MapNpcNum).YOffset - WALK_SPEED
            Case DIR_DOWN
                MapNpc(MapNpcNum).YOffset = MapNpc(MapNpcNum).YOffset + WALK_SPEED
            Case DIR_LEFT
                MapNpc(MapNpcNum).XOffset = MapNpc(MapNpcNum).XOffset - WALK_SPEED
            Case DIR_RIGHT
                MapNpc(MapNpcNum).XOffset = MapNpc(MapNpcNum).XOffset + WALK_SPEED
        End Select
        
        ' Check if completed walking over to the next tile
        If (MapNpc(MapNpcNum).XOffset = 0) And (MapNpc(MapNpcNum).YOffset = 0) Then
            MapNpc(MapNpcNum).Moving = 0
        End If
    End If
End Sub

Sub HandleKeypresses(ByVal KeyAscii As Integer)
Dim ChatText As String
Dim Name As String
Dim i As Long
Dim n As Long

    ' Handle when the player presses the return key
    If (KeyAscii = vbKeyReturn) Then
        If Player(MyIndex).Y - 1 > -1 Then
            If Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).Type = TILE_TYPE_SIGN And Player(MyIndex).Dir = DIR_UP Then
                Call AddText("The Sign Reads:", Black)
                If Trim$(Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).String1) <> "" Then
                    Call AddText(Trim$(Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).String1), Grey)
                End If
                If Trim$(Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).String2) <> "" Then
                    Call AddText(Trim$(Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).String2), Grey)
                End If
                If Trim$(Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).String3) <> "" Then
                    Call AddText(Trim$(Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).String3), Grey)
                End If
            Exit Sub
            End If
        End If
        ' Broadcast message
        If Mid$(MyText, 1, 2) = "/b" Then
            ChatText = Mid$(MyText, 3, Len(MyText) - 1)
            If Len(Trim$(ChatText)) > 0 Then
                Call BroadcastMsg(ChatText)
            End If
            MyText = ""
            Exit Sub
        End If
        
        ' Emote message
        If Mid$(MyText, 1, 2) = "//" Then
            ChatText = Mid$(MyText, 3, Len(MyText) - 1)
            If Len(Trim$(ChatText)) > 0 Then
                Call EmoteMsg(ChatText)
            End If
            MyText = ""
            Exit Sub
        End If
        
        ' Guild message
        If Mid$(MyText, 1, 1) = "@" Then
            ChatText = Mid$(MyText, 2, Len(MyText) - 1)
            If Len(Trim(ChatText)) > 0 Then
                Call SendData("guildchat" & SEP_CHAR & ChatText & SEP_CHAR & END_CHAR)
            End If
            MyText = ""
            Exit Sub
        End If
        
        ' Guild message
        If Mid$(MyText, 1, 6) = "/gchat" Then
            ChatText = Mid$(MyText, 7, Len(MyText) - 6)
            If Len(Trim(ChatText)) > 0 Then
                Call SendData("guildchat" & SEP_CHAR & ChatText & SEP_CHAR & END_CHAR)
            End If
            MyText = ""
            Exit Sub
        End If
        
        ' Party message
        If Mid$(MyText, 1, 1) = "#" Then
            ChatText = Mid$(MyText, 2, Len(MyText) - 1)
            If Len(Trim(ChatText)) > 0 Then
                Call SendData("partychat" & SEP_CHAR & ChatText & SEP_CHAR & END_CHAR)
            End If
            MyText = ""
            Exit Sub
        End If
        
        ' Party message
        If Mid$(MyText, 1, 6) = "/pchat" Then
            ChatText = Mid$(MyText, 7, Len(MyText) - 6)
            If Len(Trim(ChatText)) > 0 Then
                Call SendData("partychat" & SEP_CHAR & ChatText & SEP_CHAR & END_CHAR)
            End If
            MyText = ""
            Exit Sub
        End If
        
        ' Player message
        If Mid$(MyText, 1, 2) = "/t" Then
            ChatText = Mid$(MyText, 3, Len(MyText) - 2)
            Name = ""
                    
            ' Get the desired player from the user text
            For i = 1 To Len(ChatText)
                If Mid$(ChatText, i, 2) <> " " Then
                    Name = Name & Mid$(ChatText, i, 2)
                Else
                    Exit For
                End If
            Next i
                    
            ' Make sure they are actually sending something
            If Len(ChatText) - i > 0 Then
                ChatText = Mid$(ChatText, i + 2, Len(ChatText) - i)
                    
                ' Send the message to the player
                Call PlayerMsg(ChatText, Name)
            Else
                Call AddText("Usage: /playername msghere", AlertColor)
            End If
            MyText = ""
            Exit Sub
        End If
        
        ' // Commands //
        ' Help
        If LCase$(Mid$(MyText, 1, 5)) = "/help" Then
            If GetPlayerAccess(MyIndex) < 1 Then
                Call AddText("Social Commands:", HelpColor)
                Call AddText("/b msghere = Broadcast Message", HelpColor)
                Call AddText("/e msghere = Emote Message", HelpColor)
                Call AddText("!namehere msghere = Player Message", HelpColor)
                Call AddText("Available Commands: /help, /info, /who, /fps, /inv, /stats, /train, /party, /join, /leave, /refresh", HelpColor)
                MyText = ""
                Exit Sub
            Else
                Call AddText("Social Commands:", HelpColor)
                Call AddText("""msghere = Global Admin Message", HelpColor)
                Call AddText("=msghere = Private Admin Message", HelpColor)
                Call AddText("Available Commands: :admin, :loc, :mapeditor, :warpmeto, :warptome, :warpto, :setsprite, :mapreport, :kick, :ban, :unban, :destroybanlist, :mute, :unmute, :edititem, :respawn, :editnpc, :motd, :editshop, :editspell, :map *playername*, :broadcast *playername*, :private *playername*, :global *playername*, :emot *playername*, :admin *playername*", HelpColor)
                MyText = ""
                Exit Sub
            End If
        End If
        
        ' Call upon admins
        If LCase$(Mid$(MyText, 1, 11)) = "/calladmins" Then
            Call SendData("calladmins" & SEP_CHAR & END_CHAR)
            MyText = ""
            Exit Sub
        End If
        
        ' Verification User
        If LCase$(Mid$(MyText, 1, 5)) = "/info" Then
            ChatText = Mid$(MyText, 6, Len(MyText) - 5)
            Call SendData("playerinforequest" & SEP_CHAR & ChatText & SEP_CHAR & END_CHAR)
            MyText = ""
            Exit Sub
        End If
        
        ' Whos Online
        If LCase$(Mid$(MyText, 1, 4)) = "/who" Then
            Call SendWhosOnline
            MyText = ""
            Exit Sub
        End If
                        
        ' Checking fps
        If LCase$(Mid$(MyText, 1, 4)) = "/fps" Then
            Call AddText("FPS: " & GameFPS, Pink)
            MyText = ""
            Exit Sub
        End If
                
        ' Show inventory
        If LCase$(Mid$(MyText, 1, 4)) = "/inv" Then
            Call UpdateInventory
            frmEndieko.picInv3.Visible = True
            MyText = ""
            Exit Sub
        End If
        
        ' Request stats
        If LCase$(Mid$(MyText, 1, 6)) = "/stats" Then
            Call SendData("getstats" & SEP_CHAR & END_CHAR)
            MyText = ""
            Exit Sub
        End If
    
        ' Show training
        If LCase$(Mid$(MyText, 1, 6)) = "/train" Then
            frmTraining.Show vbModal
            MyText = ""
            Exit Sub
        End If
     
        ' Refresh Player
        If LCase$(Mid$(MyText, 1, 8)) = "/refresh" Then
            Call SendData("refresh" & SEP_CHAR & END_CHAR)
            MyText = ""
            Exit Sub
        End If
        
        ' Decline Chat
        If LCase$(Mid$(MyText, 1, 12)) = "/chatdecline" Then
            Call SendData("dchat" & SEP_CHAR & END_CHAR)
            MyText = ""
            Exit Sub
        End If
        
        ' Accept Chat
        If LCase$(Mid$(MyText, 1, 5)) = "/chat" Then
            Call SendData("achat" & SEP_CHAR & END_CHAR)
            MyText = ""
            Exit Sub
        End If
        
        If LCase$(Mid$(MyText, 1, 6)) = "/trade" Then
            ' Make sure they are actually sending something
            If Len(MyText) > 7 Then
                ChatText = Mid$(MyText, 8, Len(MyText) - 7)
                Call SendTradeRequest(ChatText)
            Else
                Call AddText("Usage: /trade playernamehere", AlertColor)
            End If
            MyText = ""
            Exit Sub
        End If
        
        ' Accept Trade
        If LCase$(Mid$(MyText, 1, 7)) = "/accept" Then
            Call SendAcceptTrade
            MyText = ""
            Exit Sub
        End If
        
        ' Decline Trade
        If LCase$(Mid$(MyText, 1, 8)) = "/decline" Then
            Call SendDeclineTrade
            MyText = ""
            Exit Sub
        End If
        
        ' Party request
        If LCase$(Mid$(MyText, 1, 6)) = "/party" Then
            ' Make sure they are actually sending something
            If Len(MyText) > 7 Then
                ChatText = Mid$(MyText, 8, Len(MyText) - 7)
                Call SendPartyRequest(ChatText)
            Else
                Call AddText("Usage: /party playernamehere", AlertColor)
            End If
            MyText = ""
            Exit Sub
        End If
        
        ' Join party
        If LCase$(Mid$(MyText, 1, 5)) = "/join" Then
            Call SendJoinParty
            MyText = ""
            Exit Sub
        End If
        
        ' Leave party
        If LCase$(Mid$(MyText, 1, 6)) = "/leave" Then
            Call SendLeaveParty
            MyText = ""
            Exit Sub
        End If
        
        ' Kick player from guild
        If LCase$(Mid$(MyText, 1, 10)) = "/guildkick" Then
            If Len(MyText) > 10 Then
                ChatText = Mid$(MyText, 11, Len(MyText) - 10)
                Call SendData("kickfromguild" & SEP_CHAR & ChatText & SEP_CHAR & END_CHAR)
            Else
                Call AddText("Usage: /guildkick playernamehere", AlertColor)
            End If
            MyText = ""
            Exit Sub
        End If
        
        If LCase$(Mid$(MyText, 1, 9)) = "/guildwho" Then
            Call SendData("guildwho" & SEP_CHAR & END_CHAR)
            MyText = ""
            Exit Sub
        End If
        
        If LCase$(Mid$(MyText, 1, 11)) = "/guildleave" Then
            Call SendData("guildleave" & SEP_CHAR & END_CHAR)
            MyText = ""
            Exit Sub
        End If
        
        ' Invite player to guild
        If LCase$(Mid$(MyText, 1, 12)) = "/guildinvite" Then
            If Len(MyText) > 12 Then
                ChatText = Mid$(MyText, 13, Len(MyText) - 12)
                Call SendData("invitetoguild" & SEP_CHAR & ChatText & SEP_CHAR & END_CHAR)
            Else
                Call AddText("Usage: /guildinvite playernamehere", AlertColor)
            End If
            MyText = ""
            Exit Sub
        End If
        
        If LCase$(Mid$(MyText, 1, 12)) = "/guildaccept" Then
            Call SendData("guildinvite" & SEP_CHAR & 0 & SEP_CHAR & END_CHAR)
            MyText = ""
            Exit Sub
        End If
        
        If LCase$(Mid$(MyText, 1, 13)) = "/guilddecline" Then
            Call SendData("guildinvite" & SEP_CHAR & 1 & SEP_CHAR & END_CHAR)
            MyText = ""
            Exit Sub
        End If
  
        ' Buy guild
        If LCase$(Mid$(MyText, 1, 9)) = "/guildnew" Then
            If Len(MyText) > 9 Then
                ChatText = Mid$(MyText, 10, Len(MyText) - 9)
                Call SendData("buyguild" & SEP_CHAR & ChatText & SEP_CHAR & END_CHAR)
            Else
                Call AddText("Usage: /createguild guildnamehere", AlertColor)
            End If
            MyText = ""
            Exit Sub
        End If
        
        ' Destroy Guild
        If LCase$(Mid$(MyText, 1, 13)) = "/guilddisband" Then
            Call SendData("destroyguild" & SEP_CHAR & END_CHAR)
            MyText = ""
            Exit Sub
        End If
        
        ' // Moniter Admin Commands //
        If GetPlayerAccess(MyIndex) > 0 Then
            ' Kicking a player
            If LCase$(Mid$(MyText, 1, 5)) = "/kick" Then
                If Len(MyText) > 6 Then
                    MyText = Mid$(MyText, 7, Len(MyText) - 6)
                    Call SendKick(MyText)
                End If
                MyText = ""
                Exit Sub
            End If
        
            ' Global Message
            If Mid$(MyText, 1, 1) = """" Then
                ChatText = Mid$(MyText, 2, Len(MyText) - 1)
                If Len(Trim$(ChatText)) > 0 Then
                    Call GlobalMsg(ChatText)
                End If
                MyText = ""
                Exit Sub
            End If
        
            ' Admin Message
            If Mid$(MyText, 1, 1) = "=" Then
                ChatText = Mid$(MyText, 2, Len(MyText) - 1)
                If Len(Trim$(ChatText)) > 0 Then
                    Call AdminMsg(ChatText)
                End If
                MyText = ""
                Exit Sub
            End If
        End If
        
        ' // Mapper Admin Commands //
        If GetPlayerAccess(MyIndex) >= ADMIN_MAPPER Then
            ' Location
            If LCase$(Mid$(MyText, 1, 4)) = "/loc" Then
                Call SendRequestLocation
                MyText = ""
                Exit Sub
            End If
            
            ' Map Editor
            If LCase$(Mid$(MyText, 1, 8)) = "/editmap" Then
                Call SendRequestEditMap
                MyText = ""
                Exit Sub
            End If
            
            ' Map report
            If LCase$(Mid$(MyText, 1, 10)) = "/mapreport" Then
                Call SendData("mapreport" & SEP_CHAR & END_CHAR)
                MyText = ""
                Exit Sub
            End If
            
            ' Warping to a player
            If LCase$(Mid$(MyText, 1, 9)) = "/warpmeto" Then
                If Len(MyText) > 10 Then
                    MyText = Mid$(MyText, 10, Len(MyText) - 9)
                    Call WarpMeTo(MyText)
                End If
                MyText = ""
                Exit Sub
            End If
                        
            ' Warping a player to you
            If LCase$(Mid$(MyText, 1, 9)) = "/warptome" Then
                If Len(MyText) > 10 Then
                    MyText = Mid$(MyText, 10, Len(MyText) - 9)
                    Call WarpToMe(MyText)
                End If
                MyText = ""
                Exit Sub
            End If
                        
            ' Warping to a map
            If LCase$(Mid$(MyText, 1, 7)) = "/warpto" Then
                If Len(MyText) > 8 Then
                    MyText = Mid$(MyText, 8, Len(MyText) - 7)
                    n = Val(MyText)
                
                    ' Check to make sure its a valid map #
                    If n > 0 And n <= MAX_MAPS Then
                        Call WarpTo(n)
                    Else
                        Call AddText("Invalid map number.", Red)
                    End If
                End If
                MyText = ""
                Exit Sub
            End If
            
            ' Setting sprite
            If LCase$(Mid$(MyText, 1, 10)) = "/setsprite" Then
                If Len(MyText) > 11 Then
                    ' Get sprite #
                    MyText = Mid$(MyText, 12, Len(MyText) - 11)
                
                    Call SendSetSprite(Val(MyText))
                End If
                MyText = ""
                Exit Sub
            End If
            
            ' Setting player sprite
            If LCase$(Mid$(MyText, 1, 16)) = "/setplayersprite" Then
                If Len(MyText) > 19 Then
                    i = Val(Mid$(MyText, 17, 1))
                
                    MyText = Mid$(MyText, 18, Len(MyText) - 17)
                    Call SendSetPlayerSprite(i, Val(MyText))
                End If
                MyText = ""
                Exit Sub
            End If
        
            ' Respawn request
            If Mid$(MyText, 1, 8) = "/respawn" Then
                Call SendMapRespawn
                MyText = ""
                Exit Sub
            End If
        
            ' MOTD change
            If Mid$(MyText, 1, 5) = "/motd" Then
                If Len(MyText) > 6 Then
                    MyText = Mid$(MyText, 7, Len(MyText) - 6)
                    If Trim$(MyText) <> "" Then
                        Call SendMOTDChange(MyText)
                    End If
                End If
                MyText = ""
                Exit Sub
            End If
            
            ' Check the ban list
             If LCase$(Mid$(MyText, 1, 8)) = "/banlist" Then
                 Call SendBanList
                 MyText = ""
                 Exit Sub
             End If
            
            ' Banning a player
            If LCase$(Mid$(MyText, 1, 4)) = "/ban" Then
                If Len(MyText) > 5 Then
                    MyText = Mid$(MyText, 6, Len(MyText) - 5)
                    Call SendBan(MyText)
                    MyText = ""
                End If
                Exit Sub
            End If
            
            ' unBanning a player
             If LCase$(Mid$(MyText, 1, 6)) = "/unban" Then
                 If Len(MyText) > 7 Then
                     MyText = Mid$(MyText, 8, Len(MyText) - 7)
                     Call SendUnBan(MyText)
                     MyText = ""
                 End If
                 Exit Sub
             End If
        End If
            
        ' // Developer Admin Commands //
        If GetPlayerAccess(MyIndex) >= ADMIN_DEVELOPER Then
            ' Turn Invisible to Monitor Players
            If LCase$(Mid$(MyText, 1, 8)) = "/monitor" Then
                Call SetInvisiblity
                MyText = ""
                Exit Sub
            End If
        
            ' Stop Player Broadcast Messages
            If LCase$(Mid$(MyText, 1, 10)) = "/broadcast" Then
                If Len(MyText) > 11 Then
                    MyText = Mid$(MyText, 12, Len(MyText) - 11)
                    Call PlayerBroadcast(MyText)
                    MyText = ""
                End If
                Exit Sub
            End If
            
            ' Stop Player Map Messages
            If LCase$(Mid$(MyText, 1, 4)) = "/map" Then
                If Len(MyText) > 5 Then
                    MyText = Mid$(MyText, 6, Len(MyText) - 5)
                    Call PlayerMap(MyText)
                    MyText = ""
                End If
                Exit Sub
            End If
            
            ' Stop Player Emote Messages
            If LCase$(Mid$(MyText, 1, 6)) = "/emote" Then
                If Len(MyText) > 7 Then
                    MyText = Mid$(MyText, 8, Len(MyText) - 7)
                    Call PlayerEmote(MyText)
                    MyText = ""
                End If
                Exit Sub
            End If
            
            ' Stop Player Private Messages
            If LCase$(Mid$(MyText, 1, 8)) = ":private" Then
                If Len(MyText) > 9 Then
                    MyText = Mid$(MyText, 10, Len(MyText) - 9)
                    Call PlayerEmote(MyText)
                    MyText = ""
                End If
                Exit Sub
            End If
            
            ' Stop Player Guild Messages
            If LCase$(Mid$(MyText, 1, 6)) = "/guild" Then
                If Len(MyText) > 7 Then
                    MyText = Mid$(MyText, 8, Len(MyText) - 7)
                    Call PlayerGuild(MyText)
                    MyText = ""
                End If
                Exit Sub
            End If
            
            ' Stop Player Admin Messages
            If LCase$(Mid$(MyText, 1, 6)) = "/admin" Then
                If Len(MyText) > 7 Then
                    MyText = Mid$(MyText, 8, Len(MyText) - 7)
                    Call PlayerAdmin(MyText)
                    MyText = ""
                End If
                Exit Sub
            End If
            
            ' Stop Player Global Messages
            If LCase$(Mid$(MyText, 1, 7)) = "/global" Then
                If Len(MyText) > 8 Then
                    MyText = Mid$(MyText, 9, Len(MyText) - 8)
                    Call PlayerGlobal(MyText)
                    MyText = ""
                End If
                Exit Sub
            End If
            
            ' Stop Player Party Messages
            If LCase$(Mid$(MyText, 1, 6)) = "/party" Then
                If Len(MyText) > 7 Then
                    MyText = Mid$(MyText, 8, Len(MyText) - 7)
                    Call PlayerParty(MyText)
                    MyText = ""
                End If
                Exit Sub
            End If
            
            ' Jail
            If LCase$(Mid$(MyText, 1, 5)) = "/jail" Then
                If Len(MyText) > 6 Then
                    MyText = Mid$(MyText, 7, Len(MyText) - 6)
                    Call JailPlayer(MyText)
                    MyText = ""
                End If
                Exit Sub
            End If
            
            ' Mute
            If LCase$(Mid$(MyText, 1, 5)) = "/mute" Then
                If Len(MyText) > 6 Then
                    MyText = Mid$(MyText, 7, Len(MyText) - 6)
                    Call MutePlayer(MyText)
                    MyText = ""
                End If
                Exit Sub
            End If
            
            ' Un-Mute
            If LCase$(Mid$(MyText, 1, 7)) = "/unmute" Then
                If Len(MyText) > 8 Then
                    MyText = Mid$(MyText, 9, Len(MyText) - 8)
                    Call UnMutePlayer(MyText)
                    MyText = ""
                End If
                Exit Sub
            End If
            
            ' Day/Night
            If LCase$(Mid$(MyText, 1, 9)) = "/daynight" Then
                Call SendGameTime
                Call SendData("daynight" & SEP_CHAR & END_CHAR)
                MyText = ""
                Exit Sub
            End If
            
            ' Editing item request
            If Mid$(MyText, 1, 9) = "/edititem" Then
                Call SendRequestEditItem
                MyText = ""
                Exit Sub
            End If
            
            ' Editing emoticon request
            If Mid$(MyText, 1, 13) = "/editemoticon" Then
                Call SendRequestEditEmoticon
                MyText = ""
                Exit Sub
            End If
            
            ' Editing npc request
            If Mid$(MyText, 1, 8) = "/editnpc" Then
                Call SendRequestEditNpc
                MyText = ""
                Exit Sub
            End If
            
            ' Editing arrow request
             If Mid$(MyText, 1, 13) = "/editarrow" Then
                 Call SendRequestEditArrow
                 MyText = ""
                 Exit Sub
             End If
            
            ' Editing shop request
            If Mid$(MyText, 1, 9) = "/editshop" Then
                Call SendRequestEditShop
                MyText = ""
                Exit Sub
            End If
        
            ' Editing spell request
            If LCase$(Trim$(MyText)) = "/editspell" Then
                Call SendRequestEditSpell
                MyText = ""
                Exit Sub
            End If
            
            ' Editing Effect request
            If LCase$(Trim$(MyText)) = "/editeffect" Then
                Call SendRequestEditEffect
                MyText = ""
                Exit Sub
            End If
        End If
        
        ' // Creator Admin Commands //
        If GetPlayerAccess(MyIndex) >= ADMIN_CREATOR Then
            ' Giving another player access
            If LCase$(Mid$(MyText, 1, 10)) = "/setaccess" Then
                ' Get access #
                i = Val(Mid$(MyText, 12, 1))
                
                MyText = Mid$(MyText, 14, Len(MyText) - 13)
                
                Call SendSetAccess(MyText, i)
                MyText = ""
                Exit Sub
            End If
            
            ' Ban destroy
            If LCase$(Mid$(MyText, 1, 15)) = "/destroybanlist" Then
                Call SendBanDestroy
                MyText = ""
                Exit Sub
            End If
            
            ' Editing script request
            If Mid$(MyText, 1, 11) = "/editscript" Then
                Call SendRequestEditScript
                MyText = ""
                Exit Sub
            End If
        End If
        
        ' Tell them its not a valid command
        If Left$(Trim$(MyText), 1) = "/" Then
            For i = 0 To MAX_EMOTICONS
                If Trim$(Emoticons(i).Command) = Trim$(MyText) And Trim$(Emoticons(i).Command) <> "/" Then
                    Call SendData("checkemoticons" & SEP_CHAR & i & SEP_CHAR & END_CHAR)
                    MyText = ""
                Exit Sub
                End If
            Next i
            Call SendData("checkcommands" & SEP_CHAR & MyText & SEP_CHAR & END_CHAR)
            MyText = ""
        Exit Sub
        End If
            
        ' Say message
        If Len(Trim$(MyText)) > 0 Then
            Call SayMsg(MyText)
        End If
        MyText = ""
        Exit Sub
    End If
                
    ' Handle when the user presses the backspace key
    If (KeyAscii = vbKeyBack) Then
        If Len(MyText) > 0 Then
            MyText = Mid$(MyText, 1, Len(MyText) - 1)
        End If
    End If
    
    ' And if neither, then add the character to the user's text buffer
    If (KeyAscii <> vbKeyReturn) And (KeyAscii <> vbKeyBack) Then
        ' Make sure they just use standard keys, no gay shitty macro keys
        If KeyAscii >= 32 And KeyAscii <= 126 Then
            MyText = MyText & Chr(KeyAscii)
        End If
    End If
End Sub

Sub CheckMapGetItem()
    If GetTickCount > Player(MyIndex).MapGetTimer + 250 And Trim$(MyText) = "" Then
        Player(MyIndex).MapGetTimer = GetTickCount
        Call SendData("mapgetitem" & SEP_CHAR & END_CHAR)
    End If
End Sub

Sub CheckAttack()
    If ControlDown = True And Player(MyIndex).AttackTimer + 1000 < GetTickCount And Player(MyIndex).Attacking = 0 Then
        Player(MyIndex).Attacking = 1
        Player(MyIndex).AttackTimer = GetTickCount
        Call SendData("attack" & SEP_CHAR & END_CHAR)
    End If
End Sub

Sub CheckInput2()
    If GettingMap = False Then
        If GetKeyState(VK_RETURN) < 0 Then
            Call CheckMapGetItem
        End If
        If GetKeyState(VK_CONTROL) < 0 Then
            ControlDown = True
        Else
            ControlDown = False
        End If
        If GetKeyState(VK_UP) < 0 Then
            DirUp = True
            DirDown = False
            DirLeft = False
            DirRight = False
        Else
            DirUp = False
        End If
        If GetKeyState(VK_DOWN) < 0 Then
            DirUp = False
            DirDown = True
            DirLeft = False
            DirRight = False
        Else
            DirDown = False
        End If
        If GetKeyState(VK_LEFT) < 0 Then
            DirUp = False
            DirDown = False
            DirLeft = True
            DirRight = False
        Else
            DirLeft = False
        End If
        If GetKeyState(VK_RIGHT) < 0 Then
            DirUp = False
            DirDown = False
            DirLeft = False
            DirRight = True
        Else
            DirRight = False
        End If
        If GetKeyState(VK_SHIFT) < 0 Then
            ShiftDown = True
        Else
            ShiftDown = False
        End If
    End If
End Sub

Sub CheckInput(ByVal KeyState As Byte, ByVal KeyCode As Integer, ByVal Shift As Integer)
    If GettingMap = False Then
        If KeyState = 1 Then
            If KeyCode = vbKeyReturn Then
                Call CheckMapGetItem
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
        Else
            If KeyCode = vbKeyUp Then DirUp = False
            If KeyCode = vbKeyDown Then DirDown = False
            If KeyCode = vbKeyLeft Then DirLeft = False
            If KeyCode = vbKeyRight Then DirRight = False
            If KeyCode = vbKeyShift Then ShiftDown = False
            If KeyCode = vbKeyControl Then ControlDown = False
        End If
    End If
End Sub

Function IsTryingToMove() As Boolean
    If (DirUp = True) Or (DirDown = True) Or (DirLeft = True) Or (DirRight = True) Then
        IsTryingToMove = True
    Else
        IsTryingToMove = False
    End If
End Function

Function CanMove() As Boolean
Dim i As Long, d As Long

    CanMove = True
    
    ' Make sure they aren't trying to move when they are already moving
    If Player(MyIndex).Moving <> 0 Then
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
    
    d = GetPlayerDir(MyIndex)
    If DirUp Then
        Call SetPlayerDir(MyIndex, DIR_UP)
        
        ' Check to see if they are trying to go out of bounds
        If GetPlayerY(MyIndex) > 0 Then
            ' Check to see if the map tile is blocked or not
            If Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).Type = TILE_TYPE_BLOCKED Or Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).Type = TILE_TYPE_SIGN Then
                CanMove = False
                
                ' Set the new direction if they weren't facing that direction
                If d <> DIR_UP Then
                    Call SendPlayerDir
                End If
                Exit Function
            End If
            
            If Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).Type = TILE_TYPE_CBLOCK Then
                If Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).Data1 = Player(MyIndex).Class Then Exit Function
                If Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).Data2 = Player(MyIndex).Class Then Exit Function
                If Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).Data3 = Player(MyIndex).Class Then Exit Function
                CanMove = False
                
                ' Set the new direction if they weren't facing that direction
                If d <> DIR_DOWN Then
                    Call SendPlayerDir
                End If
            End If
                                                    
            ' Check to see if the key door is open or not
            If Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).Type = TILE_TYPE_KEY Or Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).Type = TILE_TYPE_DOOR Then
                ' This actually checks if its open or not
                If TempTile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).DoorOpen = NO Then
                    CanMove = False
                
                    ' Set the new direction if they weren't facing that direction
                    If d <> DIR_UP Then
                        Call SendPlayerDir
                    End If
                    Exit Function
                End If
            End If
                        
            ' Check to see if a player is already on that tile
            For i = 1 To MAX_PLAYERS
                If IsPlaying(i) Then
                    If GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                        If (GetPlayerX(i) = GetPlayerX(MyIndex)) And (GetPlayerY(i) = GetPlayerY(MyIndex) - 1) Then
                            CanMove = False
                        
                            ' Set the new direction if they weren't facing that direction
                            If d <> DIR_UP Then
                                Call SendPlayerDir
                            End If
                            Exit Function
                        End If
                    End If
                End If
            Next i
        
            ' Check to see if a npc is already on that tile
            For i = 1 To MAX_MAP_NPCS
                If MapNpc(i).Num > 0 Then
                    If (MapNpc(i).X = GetPlayerX(MyIndex)) And (MapNpc(i).Y = GetPlayerY(MyIndex) - 1) Then
                        CanMove = False
                        
                        ' Set the new direction if they weren't facing that direction
                        If d <> DIR_UP Then
                            Call SendPlayerDir
                        End If
                        Exit Function
                    End If
                End If
            Next i
        Else
            ' Check if they can warp to a new map
            If Map.Up > 0 Then
                Call SendPlayerRequestNewMap
                GettingMap = True
            End If
            CanMove = False
            Exit Function
        End If
    End If
            
    If DirDown Then
        Call SetPlayerDir(MyIndex, DIR_DOWN)
        
        ' Check to see if they are trying to go out of bounds
        If GetPlayerY(MyIndex) < MAX_MAPY Then
            ' Check to see if the map tile is blocked or not
            If Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) + 1).Type = TILE_TYPE_BLOCKED Or Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) + 1).Type = TILE_TYPE_SIGN Then
                CanMove = False
                
                ' Set the new direction if they weren't facing that direction
                If d <> DIR_DOWN Then
                    Call SendPlayerDir
                End If
                Exit Function
            End If
                        
            If Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) + 1).Type = TILE_TYPE_CBLOCK Then
                If Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) + 1).Data1 = Player(MyIndex).Class Then Exit Function
                If Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) + 1).Data2 = Player(MyIndex).Class Then Exit Function
                If Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) + 1).Data3 = Player(MyIndex).Class Then Exit Function
                CanMove = False
                
                ' Set the new direction if they weren't facing that direction
                If d <> DIR_DOWN Then
                    Call SendPlayerDir
                End If
                Exit Function
            End If
                                        
            ' Check to see if the key door is open or not
            If Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) + 1).Type = TILE_TYPE_KEY Or Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) + 1).Type = TILE_TYPE_DOOR Then
                ' This actually checks if its open or not
                If TempTile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) + 1).DoorOpen = NO Then
                    CanMove = False
                
                    ' Set the new direction if they weren't facing that direction
                    If d <> DIR_DOWN Then
                        Call SendPlayerDir
                    End If
                    Exit Function
                End If
            End If
                        
            ' Check to see if a player is already on that tile
            For i = 1 To MAX_PLAYERS
                If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                    If (GetPlayerX(i) = GetPlayerX(MyIndex)) And (GetPlayerY(i) = GetPlayerY(MyIndex) + 1) Then
                        CanMove = False
                        
                        ' Set the new direction if they weren't facing that direction
                        If d <> DIR_DOWN Then
                            Call SendPlayerDir
                        End If
                        Exit Function
                    End If
                End If
            Next i
            
            ' Check to see if a npc is already on that tile
            For i = 1 To MAX_MAP_NPCS
                If MapNpc(i).Num > 0 Then
                    If (MapNpc(i).X = GetPlayerX(MyIndex)) And (MapNpc(i).Y = GetPlayerY(MyIndex) + 1) Then
                        CanMove = False
                        
                        ' Set the new direction if they weren't facing that direction
                        If d <> DIR_DOWN Then
                            Call SendPlayerDir
                        End If
                        Exit Function
                    End If
                End If
            Next i
        Else
            ' Check if they can warp to a new map
            If Map.Down > 0 Then
                Call SendPlayerRequestNewMap
                GettingMap = True
            End If
            CanMove = False
            Exit Function
        End If
    End If
                
    If DirLeft Then
        Call SetPlayerDir(MyIndex, DIR_LEFT)
        
        ' Check to see if they are trying to go out of bounds
        If GetPlayerX(MyIndex) > 0 Then
            ' Check to see if the map tile is blocked or not
            If Map.Tile(GetPlayerX(MyIndex) - 1, GetPlayerY(MyIndex)).Type = TILE_TYPE_BLOCKED Or Map.Tile(GetPlayerX(MyIndex) - 1, GetPlayerY(MyIndex)).Type = TILE_TYPE_SIGN Then
                CanMove = False
                
                ' Set the new direction if they weren't facing that direction
                If d <> DIR_LEFT Then
                    Call SendPlayerDir
                End If
                Exit Function
            End If
                        
            If Map.Tile(GetPlayerX(MyIndex) - 1, GetPlayerY(MyIndex)).Type = TILE_TYPE_CBLOCK Then
                If Map.Tile(GetPlayerX(MyIndex) - 1, GetPlayerY(MyIndex)).Data1 = Player(MyIndex).Class Then Exit Function
                If Map.Tile(GetPlayerX(MyIndex) - 1, GetPlayerY(MyIndex)).Data2 = Player(MyIndex).Class Then Exit Function
                If Map.Tile(GetPlayerX(MyIndex) - 1, GetPlayerY(MyIndex)).Data3 = Player(MyIndex).Class Then Exit Function
                CanMove = False
                
                ' Set the new direction if they weren't facing that direction
                If d <> DIR_LEFT Then
                    Call SendPlayerDir
                End If
                Exit Function
            End If
                                        
            ' Check to see if the key door is open or not
            If Map.Tile(GetPlayerX(MyIndex) - 1, GetPlayerY(MyIndex)).Type = TILE_TYPE_KEY Or Map.Tile(GetPlayerX(MyIndex) - 1, GetPlayerY(MyIndex)).Type = TILE_TYPE_DOOR Then
                ' This actually checks if its open or not
                If TempTile(GetPlayerX(MyIndex) - 1, GetPlayerY(MyIndex)).DoorOpen = NO Then
                    CanMove = False
                    
                    ' Set the new direction if they weren't facing that direction
                    If d <> DIR_LEFT Then
                        Call SendPlayerDir
                    End If
                    Exit Function
                End If
            End If
                        
            ' Check to see if a player is already on that tile
            For i = 1 To MAX_PLAYERS
                If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                    If (GetPlayerX(i) = GetPlayerX(MyIndex) - 1) And (GetPlayerY(i) = GetPlayerY(MyIndex)) Then
                        CanMove = False
                        
                        ' Set the new direction if they weren't facing that direction
                        If d <> DIR_LEFT Then
                            Call SendPlayerDir
                        End If
                        Exit Function
                    End If
                End If
            Next i
        
            ' Check to see if a npc is already on that tile
            For i = 1 To MAX_MAP_NPCS
                If MapNpc(i).Num > 0 Then
                    If (MapNpc(i).X = GetPlayerX(MyIndex) - 1) And (MapNpc(i).Y = GetPlayerY(MyIndex)) Then
                        CanMove = False
                        
                        ' Set the new direction if they weren't facing that direction
                        If d <> DIR_LEFT Then
                            Call SendPlayerDir
                        End If
                        Exit Function
                    End If
                End If
            Next i
        Else
            ' Check if they can warp to a new map
            If Map.Left > 0 Then
                Call SendPlayerRequestNewMap
                GettingMap = True
            End If
            CanMove = False
            Exit Function
        End If
    End If
        
    If DirRight Then
        Call SetPlayerDir(MyIndex, DIR_RIGHT)
        
        ' Check to see if they are trying to go out of bounds
        If GetPlayerX(MyIndex) < MAX_MAPX Then
            ' Check to see if the map tile is blocked or not
            If Map.Tile(GetPlayerX(MyIndex) + 1, GetPlayerY(MyIndex)).Type = TILE_TYPE_BLOCKED Or Map.Tile(GetPlayerX(MyIndex) + 1, GetPlayerY(MyIndex)).Type = TILE_TYPE_SIGN Then
                CanMove = False
                
                ' Set the new direction if they weren't facing that direction
                If d <> DIR_RIGHT Then
                    Call SendPlayerDir
                End If
                Exit Function
            End If
                        
            If Map.Tile(GetPlayerX(MyIndex) + 1, GetPlayerY(MyIndex)).Type = TILE_TYPE_CBLOCK Then
                If Map.Tile(GetPlayerX(MyIndex) + 1, GetPlayerY(MyIndex)).Data1 = Player(MyIndex).Class Then Exit Function
                If Map.Tile(GetPlayerX(MyIndex) + 1, GetPlayerY(MyIndex)).Data2 = Player(MyIndex).Class Then Exit Function
                If Map.Tile(GetPlayerX(MyIndex) + 1, GetPlayerY(MyIndex)).Data3 = Player(MyIndex).Class Then Exit Function
                CanMove = False
                
                ' Set the new direction if they weren't facing that direction
                If d <> DIR_RIGHT Then
                    Call SendPlayerDir
                End If
                Exit Function
            End If
                                        
            ' Check to see if the key door is open or not
            If Map.Tile(GetPlayerX(MyIndex) + 1, GetPlayerY(MyIndex)).Type = TILE_TYPE_KEY Or Map.Tile(GetPlayerX(MyIndex) + 1, GetPlayerY(MyIndex)).Type = TILE_TYPE_DOOR Then
                ' This actually checks if its open or not
                If TempTile(GetPlayerX(MyIndex) + 1, GetPlayerY(MyIndex)).DoorOpen = NO Then
                    CanMove = False
                    
                    ' Set the new direction if they weren't facing that direction
                    If d <> DIR_RIGHT Then
                        Call SendPlayerDir
                    End If
                    Exit Function
                End If
            End If
                        
            ' Check to see if a player is already on that tile
            For i = 1 To MAX_PLAYERS
                If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                    If (GetPlayerX(i) = GetPlayerX(MyIndex) + 1) And (GetPlayerY(i) = GetPlayerY(MyIndex)) Then
                        CanMove = False
                        
                        ' Set the new direction if they weren't facing that direction
                        If d <> DIR_RIGHT Then
                            Call SendPlayerDir
                        End If
                        Exit Function
                    End If
                End If
            Next i
        
            ' Check to see if a npc is already on that tile
            For i = 1 To MAX_MAP_NPCS
                If MapNpc(i).Num > 0 Then
                    If (MapNpc(i).X = GetPlayerX(MyIndex) + 1) And (MapNpc(i).Y = GetPlayerY(MyIndex)) Then
                        CanMove = False
                        
                        ' Set the new direction if they weren't facing that direction
                        If d <> DIR_RIGHT Then
                            Call SendPlayerDir
                        End If
                        Exit Function
                    End If
                End If
            Next i
        Else
            ' Check if they can warp to a new map
            If Map.Right > 0 Then
                Call SendPlayerRequestNewMap
                GettingMap = True
            End If
            CanMove = False
            Exit Function
        End If
    End If
End Function

Sub CheckMovement()
Dim Num As Long
    If GettingMap = False Then
        If IsTryingToMove Then
            If CanMove Then
                ' Check if player has the shift key down for running
                If GetPlayerSP(MyIndex) > 0 Then
                    If ShiftDown Then
                        Player(MyIndex).Moving = MOVING_RUNNING
                        If Player(MyIndex).Moving = MOVING_RUNNING Then
                        Num = RandomNumber(0, 1)
                            Call SetPlayerSP(MyIndex, GetPlayerSP(MyIndex) - Num)
                        End If
                    Else
                        Player(MyIndex).Moving = MOVING_WALKING
                    End If
                Else
                    Player(MyIndex).Moving = MOVING_WALKING
                End If
            
                Select Case GetPlayerDir(MyIndex)
                    Case DIR_UP
                        Call SendPlayerMove
                        Player(MyIndex).YOffset = PIC_Y
                        Call SetPlayerY(MyIndex, GetPlayerY(MyIndex) - 1)
                
                    Case DIR_DOWN
                        Call SendPlayerMove
                        Player(MyIndex).YOffset = PIC_Y * -1
                        Call SetPlayerY(MyIndex, GetPlayerY(MyIndex) + 1)
                
                    Case DIR_LEFT
                        Call SendPlayerMove
                        Player(MyIndex).XOffset = PIC_X
                        Call SetPlayerX(MyIndex, GetPlayerX(MyIndex) - 1)
                
                    Case DIR_RIGHT
                        Call SendPlayerMove
                        Player(MyIndex).XOffset = PIC_X * -1
                        Call SetPlayerX(MyIndex, GetPlayerX(MyIndex) + 1)
                End Select
            
                ' Gotta check :)
                If Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex)).Type = TILE_TYPE_WARP Then
                    GettingMap = True
                End If
            End If
        End If
    End If
End Sub

Function FindPlayer(ByVal Name As String) As Long
Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            ' Make sure we dont try to check a name thats to small
            If Len(GetPlayerName(i)) >= Len(Trim$(Name)) Then
                If UCase(Mid$(GetPlayerName(i), 1, Len(Trim$(Name)))) = UCase(Trim$(Name)) Then
                    FindPlayer = i
                    Exit Function
                End If
            End If
        End If
    Next i
    
    FindPlayer = 0
End Function

Public Sub EditorInit()
    SaveMap = Map
    InEditor = True
    frmEditMap.Show vbModeless, frmEndieko
    'frmEndieko.picMapEditor.Visible = True
    'With frmEditMap.picBackSelect
        '.Width = 14 * PIC_X
        '.Height = 16384
        '.Picture = LoadPicture(App.Path + "\Graphics\tiles.bmp")
    'End With
    frmEditMap.picBackSelect.Picture = LoadPicture(App.Path + "\Graphics\tiles.bmp")
    frmEditMap.MouseSelected.Picture = LoadPicture(App.Path + "\Graphics\tiles.bmp")
End Sub

Public Sub EditorMouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim x1, y1 As Long
Dim x2 As Long, y2 As Long

    If InEditor Then
        x1 = Int(X / PIC_X)
        y1 = Int(Y / PIC_Y)
        If (Button = 1) And (x1 >= 0) And (x1 <= MAX_MAPX) And (y1 >= 0) And (y1 <= MAX_MAPY) Then
            If frmEditMap.shpSelected.Height <= 32 And frmEditMap.shpSelected.Width <= 32 Then
                If frmEditMap.optLayers.Value = True Then
                    With Map.Tile(x1, y1)
                        If frmEditMap.optGround.Value = True Then .Ground = EditorTileY * 14 + EditorTileX
                            If frmEditMap.optMask.Value = True Then .Mask = EditorTileY * 14 + EditorTileX
                            If frmEditMap.optAnim.Value = True Then .Anim = EditorTileY * 14 + EditorTileX
                            If frmEditMap.optMask2.Value = True Then .Mask2 = EditorTileY * 14 + EditorTileX
                            If frmEditMap.optM2Anim.Value = True Then .M2Anim = EditorTileY * 14 + EditorTileX
                            If frmEditMap.optFringe.Value = True Then .Fringe = EditorTileY * 14 + EditorTileX
                            If frmEditMap.optFAnim.Value = True Then .FAnim = EditorTileY * 14 + EditorTileX
                            If frmEditMap.optFringe2.Value = True Then .Fringe2 = EditorTileY * 14 + EditorTileX
                            If frmEditMap.optF2Anim.Value = True Then .F2Anim = EditorTileY * 14 + EditorTileX
                    End With
                Else
                    With Map.Tile(x1, y1)
                        If frmEditMap.optBlocked.Value = True Then .Type = TILE_TYPE_BLOCKED
                        If frmEditMap.optWarp.Value = True Then
                            .Type = TILE_TYPE_WARP
                            .Data1 = EditorWarpMap
                            .Data2 = EditorWarpX
                            .Data3 = EditorWarpY
                            .String1 = ""
                            .String2 = ""
                            .String3 = ""
                        End If
    
                        If frmEditMap.optHeal.Value = True Then
                            .Type = TILE_TYPE_HEAL
                            .Data1 = 0
                            .Data2 = 0
                            .Data3 = 0
                            .String1 = ""
                            .String2 = ""
                            .String3 = ""
                        End If
                        If frmEditMap.optKill.Value = True Then
                            .Type = TILE_TYPE_KILL
                            .Data1 = 0
                            .Data2 = 0
                            .Data3 = 0
                            .String1 = ""
                            .String2 = ""
                            .String3 = ""
                        End If
                        If frmEditMap.optBank.Value = True Then
                            .Type = TILE_TYPE_BANK
                            .Data1 = 0
                            .Data2 = 0
                            .Data3 = 0
                            .String1 = ""
                            .String2 = ""
                            .String3 = ""
                        End If
                        If frmEditMap.optItem.Value = True Then
                            .Type = TILE_TYPE_ITEM
                            .Data1 = ItemEditorNum
                            .Data2 = ItemEditorValue
                            .Data3 = 0
                            .String1 = ""
                            .String2 = ""
                            .String3 = ""
                        End If
                        If frmEditMap.optNpcAvoid.Value = True Then
                            .Type = TILE_TYPE_NPCAVOID
                            .Data1 = 0
                            .Data2 = 0
                            .Data3 = 0
                            .String1 = ""
                            .String2 = ""
                            .String3 = ""
                        End If
                        If frmEditMap.optKey.Value = True Then
                            .Type = TILE_TYPE_KEY
                            .Data1 = KeyEditorNum
                            .Data2 = KeyEditorTake
                            .Data3 = 0
                            .String1 = ""
                            .String2 = ""
                            .String3 = ""
                        End If
                        If frmEditMap.optKeyOpen.Value = True Then
                            .Type = TILE_TYPE_KEYOPEN
                            .Data1 = KeyOpenEditorX
                            .Data2 = KeyOpenEditorY
                            .Data3 = 0
                            .String1 = KeyOpenEditorMsg
                            .String2 = ""
                            .String3 = ""
                        End If
                        If frmEditMap.optShop.Value = True Then
                            .Type = TILE_TYPE_SHOP
                            .Data1 = EditorShopNum
                            .Data2 = 0
                            .Data3 = 0
                            .String1 = ""
                            .String2 = ""
                            .String3 = ""
                        End If
                        If frmEditMap.optCBlock.Value = True Then
                            .Type = TILE_TYPE_CBLOCK
                            .Data1 = EditorItemNum1
                            .Data2 = EditorItemNum2
                            .Data3 = EditorItemNum3
                            .String1 = ""
                            .String2 = ""
                            .String3 = ""
                        End If
                        If frmEditMap.optArena.Value = True Then
                            .Type = TILE_TYPE_ARENA
                            .Data1 = Arena1
                            .Data2 = Arena2
                            .Data3 = Arena3
                            .String1 = ""
                            .String2 = ""
                            .String3 = ""
                        End If
                        If frmEditMap.optSound.Value = True Then
                            .Type = TILE_TYPE_SOUND
                            .Data1 = 0
                            .Data2 = 0
                            .Data3 = 0
                            .String1 = SoundFileName
                            .String2 = ""
                            .String3 = ""
                        End If
                        If frmEditMap.optSprite.Value = True Then
                            .Type = TILE_TYPE_SPRITE_CHANGE
                            .Data1 = SpritePic
                            .Data2 = SpriteItem
                            .Data3 = SpritePrice
                            .String1 = ""
                            .String2 = ""
                            .String3 = ""
                        End If
                        If frmEditMap.optSign.Value = True Then
                            .Type = TILE_TYPE_SIGN
                            .Data1 = 0
                            .Data2 = 0
                            .Data3 = 0
                            .String1 = SignLine1
                            .String2 = SignLine2
                            .String3 = SignLine3
                        End If
                        If frmEditMap.optDoor.Value = True Then
                            .Type = TILE_TYPE_DOOR
                            .Data1 = 0
                            .Data2 = 0
                            .Data3 = 0
                            .String1 = ""
                            .String2 = ""
                            .String3 = ""
                        End If
                        If frmEditMap.optNotice.Value = True Then
                            .Type = TILE_TYPE_NOTICE
                            .Data1 = 0
                            .Data2 = 0
                            .Data3 = 0
                            .String1 = NoticeTitle
                            .String2 = NoticeText
                            .String3 = NoticeSound
                        End If
                        If frmEditMap.optChest.Value = True Then
                            .Type = TILE_TYPE_CHEST
                            .Data1 = 0
                            .Data2 = 0
                            .Data3 = 0
                            .String1 = ""
                            .String2 = ""
                            .String3 = ""
                        End If
                        If frmEditMap.optClassChange.Value = True Then
                            .Type = TILE_TYPE_CLASS_CHANGE
                            .Data1 = ClassChange
                            .Data2 = ClassChangeReq
                            .Data3 = 0
                            .String1 = ""
                            .String2 = ""
                            .String3 = ""
                        End If
                        If frmEditMap.optScripted.Value = True Then
                            .Type = TILE_TYPE_SCRIPTED
                            .Data1 = ScriptNum
                            .Data2 = 0
                            .Data3 = 0
                            .String1 = ""
                            .String2 = ""
                            .String3 = ""
                        End If
                    End With
                End If
            Else
                 For y2 = 0 To Int(frmEditMap.shpSelected.Height / PIC_Y) - 1
                     For x2 = 0 To Int(frmEditMap.shpSelected.Width / PIC_X) - 1
                           If x1 + x2 <= MAX_MAPX Then
                               If y1 + y2 <= MAX_MAPY Then
                                   If frmEditMap.optLayers.Value = True Then
                                         With Map.Tile(x1 + x2, y1 + y2)
                                             If frmEditMap.optGround.Value = True Then .Ground = (EditorTileY + y2) * 14 + (EditorTileX + x2)
                                             If frmEditMap.optMask.Value = True Then .Mask = (EditorTileY + y2) * 14 + (EditorTileX + x2)
                                             If frmEditMap.optAnim.Value = True Then .Anim = (EditorTileY + y2) * 14 + (EditorTileX + x2)
                                             If frmEditMap.optMask2.Value = True Then .Mask2 = (EditorTileY + y2) * 14 + (EditorTileX + x2)
                                             If frmEditMap.optM2Anim.Value = True Then .M2Anim = (EditorTileY + y2) * 14 + (EditorTileX + x2)
                                             If frmEditMap.optFringe.Value = True Then .Fringe = (EditorTileY + y2) * 14 + (EditorTileX + x2)
                                             If frmEditMap.optFAnim.Value = True Then .FAnim = (EditorTileY + y2) * 14 + (EditorTileX + x2)
                                             If frmEditMap.optFringe2.Value = True Then .Fringe2 = (EditorTileY + y2) * 14 + (EditorTileX + x2)
                                             If frmEditMap.optF2Anim.Value = True Then .F2Anim = (EditorTileY + y2) * 14 + (EditorTileX + x2)
                                         End With
                                   End If
                               End If
                           End If
                     Next x2
                 Next y2
             End If
        End If
        
        If (Button = 2) And (x1 >= 0) And (x1 <= MAX_MAPX) And (y1 >= 0) And (y1 <= MAX_MAPY) Then
            If frmEditMap.optLayers.Value = True Then
                With Map.Tile(x1, y1)
                    If frmEditMap.optGround.Value = True Then .Ground = 0
                    If frmEditMap.optMask.Value = True Then .Mask = 0
                    If frmEditMap.optAnim.Value = True Then .Anim = 0
                    If frmEditMap.optMask2.Value = True Then .Mask2 = 0
                    If frmEditMap.optM2Anim.Value = True Then .M2Anim = 0
                    If frmEditMap.optFringe.Value = True Then .Fringe = 0
                    If frmEditMap.optFAnim.Value = True Then .FAnim = 0
                    If frmEditMap.optFringe2.Value = True Then .Fringe2 = 0
                    If frmEditMap.optF2Anim.Value = True Then .F2Anim = 0
                End With
            Else
                With Map.Tile(x1, y1)
                    .Type = 0
                    .Data1 = 0
                    .Data2 = 0
                    .Data3 = 0
                    .String1 = ""
                    .String2 = ""
                    .String3 = ""
                End With
            End If
        End If
    End If
End Sub

Public Sub EditorChooseTile(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        EditorTileX = Int(X / PIC_X)
        EditorTileY = Int(Y / PIC_Y)
    End If
    frmEditMap.shpSelected.Top = Int(EditorTileY * PIC_Y)
    frmEditMap.shpSelected.Left = Int(EditorTileX * PIC_Y)
    'Call BitBlt(frmEditMap.picSelect.hDC, 0, 0, PIC_X, PIC_Y, frmEditMap.picBackSelect.hDC, EditorTileX * PIC_X, EditorTileY * PIC_Y, SRCCOPY)
End Sub

Public Sub EditorTileScroll()
    frmEditMap.picBackSelect.Top = (frmEditMap.scrlPicture.Value * PIC_Y) * -1
End Sub

Public Sub EditorSend()
    Call SendMap
    Call EditorCancel
End Sub

Public Sub EditorCancel()
    Map = SaveMap
    InEditor = False
    frmEditMap.Visible = False
    frmEndieko.Show
    'frmEndieko.picMapEditor.Visible = False
End Sub

Public Sub EditorClearLayer()
Dim YesNo As Long, X As Long, Y As Long

    ' Ground layer
    If frmEditMap.optGround.Value = True Then
    YesNo = MsgBox("Are you sure you wish to clear the ground layer?", vbYesNo, GAME_NAME)
        If YesNo = vbYes Then
            For Y = 0 To MAX_MAPY
                For X = 0 To MAX_MAPX
                    Map.Tile(X, Y).Ground = 0
                Next X
            Next Y
        End If
    End If

    ' Mask layer
    If frmEditMap.optMask.Value = True Then
    YesNo = MsgBox("Are you sure you wish to clear the mask layer?", vbYesNo, GAME_NAME)
        If YesNo = vbYes Then
            For Y = 0 To MAX_MAPY
                For X = 0 To MAX_MAPX
                    Map.Tile(X, Y).Mask = 0
                Next X
            Next Y
        End If
    End If
    
    ' Mask Animation layer
    If frmEditMap.optAnim.Value = True Then
    YesNo = MsgBox("Are you sure you wish to clear the animation layer?", vbYesNo, GAME_NAME)
        If YesNo = vbYes Then
            For Y = 0 To MAX_MAPY
                For X = 0 To MAX_MAPX
                    Map.Tile(X, Y).Anim = 0
                Next X
            Next Y
        End If
    End If
    
    ' Mask 2 layer
    If frmEditMap.optMask2.Value = True Then
    YesNo = MsgBox("Are you sure you wish to clear the mask 2 layer?", vbYesNo, GAME_NAME)
        If YesNo = vbYes Then
            For Y = 0 To MAX_MAPY
                For X = 0 To MAX_MAPX
                    Map.Tile(X, Y).Mask2 = 0
                Next X
            Next Y
        End If
    End If
    
    ' Mask 2 Animation layer
    If frmEditMap.optM2Anim.Value = True Then
    YesNo = MsgBox("Are you sure you wish to clear the mask 2 animation layer?", vbYesNo, GAME_NAME)
        If YesNo = vbYes Then
            For Y = 0 To MAX_MAPY
                For X = 0 To MAX_MAPX
                    Map.Tile(X, Y).M2Anim = 0
                Next X
            Next Y
        End If
    End If
    
    ' Fringe layer
    If frmEditMap.optFringe.Value = True Then
    YesNo = MsgBox("Are you sure you wish to clear the fringe layer?", vbYesNo, GAME_NAME)
        If YesNo = vbYes Then
            For Y = 0 To MAX_MAPY
                For X = 0 To MAX_MAPX
                    Map.Tile(X, Y).Fringe = 0
                Next X
            Next Y
        End If
    End If
    
    ' Fringe Animation layer
    If frmEditMap.optFAnim.Value = True Then
    YesNo = MsgBox("Are you sure you wish to clear the fringe animation layer?", vbYesNo, GAME_NAME)
        If YesNo = vbYes Then
            For Y = 0 To MAX_MAPY
                For X = 0 To MAX_MAPX
                    Map.Tile(X, Y).FAnim = 0
                Next X
            Next Y
        End If
    End If
    
    ' Fringe 2 layer
    If frmEditMap.optFringe2.Value = True Then
    YesNo = MsgBox("Are you sure you wish to clear the fringe 2 layer?", vbYesNo, GAME_NAME)
        If YesNo = vbYes Then
            For Y = 0 To MAX_MAPY
                For X = 0 To MAX_MAPX
                    Map.Tile(X, Y).Fringe2 = 0
                Next X
            Next Y
        End If
    End If
    
    ' Fringe 2 Animation layer
    If frmEditMap.optF2Anim.Value = True Then
    YesNo = MsgBox("Are you sure you wish to clear the fringe 2 animation layer?", vbYesNo, GAME_NAME)
        If YesNo = vbYes Then
            For Y = 0 To MAX_MAPY
                For X = 0 To MAX_MAPX
                    Map.Tile(X, Y).F2Anim = 0
                Next X
            Next Y
        End If
    End If
End Sub

Public Sub EditorClearAttribs()
Dim YesNo As Long, X As Long, Y As Long

    YesNo = MsgBox("Are you sure you wish to clear the attributes on this map?", vbYesNo, GAME_NAME)
    
    If YesNo = vbYes Then
        For Y = 0 To MAX_MAPY
            For X = 0 To MAX_MAPX
                Map.Tile(X, Y).Type = 0
            Next X
        Next Y
    End If
End Sub

Public Sub EmoticonEditorInit()
    frmEditEmoticon.scrlEmoticon.Max = MAX_EMOTICONS
    frmEditEmoticon.scrlEmoticon.Value = Emoticons(EditorIndex - 1).Pic
    frmEditEmoticon.txtCommand.Text = Trim$(Emoticons(EditorIndex - 1).Command)
    frmEditEmoticon.picEmoticons.Picture = LoadPicture(App.Path & "\Graphics\emoticons.bmp")
    frmEditEmoticon.Show vbModal
End Sub

Public Sub EmoticonEditorOk()
    Emoticons(EditorIndex - 1).Pic = frmEditEmoticon.scrlEmoticon.Value
    If frmEditEmoticon.txtCommand.Text <> "/" Then
        Emoticons(EditorIndex - 1).Command = frmEditEmoticon.txtCommand.Text
    Else
        Emoticons(EditorIndex - 1).Command = ""
    End If
    
    Call SendSaveEmoticon(EditorIndex - 1)
    Call EmoticonEditorCancel
End Sub

Public Sub EmoticonEditorCancel()
    InEmoticonEditor = False
    Unload frmEditEmoticon
End Sub

Public Sub ItemEditorInit()
On Error Resume Next
Dim i As Long

    EditorItemY = Int(Item(EditorIndex).Pic / 6)
    EditorItemX = (Item(EditorIndex).Pic - Int(Item(EditorIndex).Pic / 6) * 6)
    
    frmEditItem.scrlClassReq.Max = Max_Classes

    frmEditItem.picItems.Picture = LoadPicture(App.Path & "\Graphics\items.bmp")
    
    frmEditItem.txtName.Text = Trim$(Item(EditorIndex).Name)
    frmEditItem.txtDesc.Text = Trim$(Item(EditorIndex).desc)
    frmEditItem.cmbType.ListIndex = Item(EditorIndex).Type
    
    If (frmEditItem.cmbType.ListIndex >= ITEM_TYPE_WEAPON) And (frmEditItem.cmbType.ListIndex <= ITEM_TYPE_BOOTS) Then
        frmEditItem.fraEquipment.Visible = True
        frmEditItem.txtDurability.Text = Item(EditorIndex).Data1
        frmEditItem.txtDamage.Text = Item(EditorIndex).Data2
        frmEditItem.txtStrReq.Text = Item(EditorIndex).StrReq
        frmEditItem.txtDefReq.Text = Item(EditorIndex).DefReq
        frmEditItem.txtSpeedReq.Text = Item(EditorIndex).SpeedReq
        frmEditItem.txtMagiReq.Text = Item(EditorIndex).MagiReq
        frmEditItem.scrlClassReq.Value = Item(EditorIndex).ClassReq
        frmEditItem.scrlAccessReq.Value = Item(EditorIndex).AccessReq
        frmEditItem.txtAddHP.Text = Item(EditorIndex).AddHP
        frmEditItem.txtAddMP.Text = Item(EditorIndex).AddMP
        frmEditItem.txtAddSP.Text = Item(EditorIndex).AddSP
        frmEditItem.txtAddStr.Text = Item(EditorIndex).AddStr
        frmEditItem.txtAddDef.Text = Item(EditorIndex).AddDef
        frmEditItem.txtAddMagi.Text = Item(EditorIndex).AddMagi
        frmEditItem.txtAddSpeed.Text = Item(EditorIndex).AddSpeed
        frmEditItem.txtAddExp.Text = Item(EditorIndex).AddEXP
        frmEditItem.chkFix.Value = Item(EditorIndex).CannotBeRepaired
        frmEditItem.chkDrop.Value = Item(EditorIndex).DropOnDeath
        
        If Item(EditorIndex).Data3 > 0 Then
             frmEditItem.chkBow.Value = Checked
        Else
             frmEditItem.chkBow.Value = Unchecked
        End If
       
        frmEditItem.cmbBow.Clear
        If frmEditItem.chkBow.Value = Checked Then
             For i = 1 To 100
                 frmEditItem.cmbBow.AddItem i & ": " & Arrows(i).Name
             Next i
             frmEditItem.cmbBow.ListIndex = Item(EditorIndex).Data3 - 1
             frmEditItem.picBow.Top = (Arrows(Item(EditorIndex).Data3).Pic * 32) * -1
             frmEditItem.cmbBow.Enabled = True
        Else
             frmEditItem.cmbBow.AddItem "None"
             frmEditItem.cmbBow.ListIndex = 0
             frmEditItem.cmbBow.Enabled = False
        End If
    Else
        frmEditItem.fraEquipment.Visible = False
    End If
    
    If (frmEditItem.cmbType.ListIndex >= ITEM_TYPE_POTIONADDHP) And (frmEditItem.cmbType.ListIndex <= ITEM_TYPE_POTIONSUBSP) Then
        frmEditItem.fraVitals.Visible = True
        frmEditItem.scrlVitalMod.Value = Item(EditorIndex).Data1
    Else
        frmEditItem.fraVitals.Visible = False
    End If
    
    If (frmEditItem.cmbType.ListIndex = ITEM_TYPE_SPELL) Then
        frmEditItem.fraSpell.Visible = True
        frmEditItem.scrlSpell.Value = Item(EditorIndex).Data1
    Else
        frmEditItem.fraSpell.Visible = False
    End If

    ' Set the Form
    If (frmEditItem.cmbType.ListIndex >= ITEM_TYPE_WEAPON) And (frmEditItem.cmbType.ListIndex <= ITEM_TYPE_BOOTS) Then
            If frmEditItem.cmbType.ListIndex = ITEM_TYPE_WEAPON Then
                 frmEditItem.Label3.Caption = "Damage :"
            Else
                 frmEditItem.Label3.Caption = "Defence :"
            End If
            frmEditItem.fraEquipment.Visible = True
            frmEditItem.fraAttributes.Visible = True
            frmEditItem.SSTab1.Width = 403
            frmEditItem.Width = 6345
        Else
            frmEditItem.fraEquipment.Visible = False
            frmEditItem.fraAttributes.Visible = False
        End If
    
        If (frmEditItem.cmbType.ListIndex >= ITEM_TYPE_POTIONADDHP) And (frmEditItem.cmbType.ListIndex <= ITEM_TYPE_POTIONSUBSP) Then
            frmEditItem.fraVitals.Visible = True
            frmEditItem.fraAttributes.Visible = False
            frmEditItem.SSTab1.Width = 235
            frmEditItem.Width = 3825
        Else
            frmEditItem.fraVitals.Visible = False
        End If
    
        If (frmEditItem.cmbType.ListIndex = ITEM_TYPE_SPELL) Then
            frmEditItem.fraSpell.Visible = True
            frmEditItem.fraAttributes.Visible = False
            frmEditItem.SSTab1.Width = 235
            frmEditItem.Width = 3825
        Else
            frmEditItem.fraSpell.Visible = False
        End If
        
        If (frmEditItem.cmbType.ListIndex = ITEM_TYPE_CURRENCY) Then
            frmEditItem.fraAttributes.Visible = False
            frmEditItem.SSTab1.Width = 235
            frmEditItem.Width = 3825
        End If
    
    frmEditItem.Show vbModal
End Sub

Public Sub ItemEditorOk()
    Item(EditorIndex).Name = frmEditItem.txtName.Text
    Item(EditorIndex).desc = frmEditItem.txtDesc.Text
    Item(EditorIndex).Pic = EditorItemY * 6 + EditorItemX
    Item(EditorIndex).Type = frmEditItem.cmbType.ListIndex

    If (frmEditItem.cmbType.ListIndex >= ITEM_TYPE_WEAPON) And (frmEditItem.cmbType.ListIndex <= ITEM_TYPE_BOOTS) Then
        Item(EditorIndex).Data1 = frmEditItem.txtDurability.Text
        Item(EditorIndex).Data2 = frmEditItem.txtDamage.Text
        Item(EditorIndex).CannotBeRepaired = frmEditItem.chkFix.Value
        Item(EditorIndex).DropOnDeath = frmEditItem.chkDrop.Value
        If frmEditItem.chkBow.Value = Checked Then
             Item(EditorIndex).Data3 = frmEditItem.cmbBow.ListIndex + 1
        Else
             Item(EditorIndex).Data3 = 0
        End If
        Item(EditorIndex).StrReq = frmEditItem.txtStrReq.Text
        Item(EditorIndex).DefReq = frmEditItem.txtDefReq.Text
        Item(EditorIndex).SpeedReq = frmEditItem.txtSpeedReq.Text
        Item(EditorIndex).MagiReq = frmEditItem.txtMagiReq.Text
        
        Item(EditorIndex).ClassReq = frmEditItem.scrlClassReq.Value
        Item(EditorIndex).AccessReq = frmEditItem.scrlAccessReq.Value
        
        Item(EditorIndex).AddHP = frmEditItem.txtAddHP.Text
        Item(EditorIndex).AddMP = frmEditItem.txtAddMP.Text
        Item(EditorIndex).AddSP = frmEditItem.txtAddSP.Text
        Item(EditorIndex).AddStr = frmEditItem.txtAddStr.Text
        Item(EditorIndex).AddDef = frmEditItem.txtAddDef.Text
        Item(EditorIndex).AddMagi = frmEditItem.txtAddMagi.Text
        Item(EditorIndex).AddSpeed = frmEditItem.txtAddSpeed.Text
        Item(EditorIndex).AddEXP = frmEditItem.txtAddExp.Text
    End If
    
    If (frmEditItem.cmbType.ListIndex >= ITEM_TYPE_POTIONADDHP) And (frmEditItem.cmbType.ListIndex <= ITEM_TYPE_POTIONSUBSP) Then
        Item(EditorIndex).Data1 = frmEditItem.scrlVitalMod.Value
        Item(EditorIndex).Data2 = 0
        Item(EditorIndex).Data3 = 0
        Item(EditorIndex).StrReq = 0
        Item(EditorIndex).DefReq = 0
        Item(EditorIndex).SpeedReq = 0
        Item(EditorIndex).MagiReq = 0
        Item(EditorIndex).ClassReq = -1
        Item(EditorIndex).AccessReq = 0
        
        Item(EditorIndex).AddHP = 0
        Item(EditorIndex).AddMP = 0
        Item(EditorIndex).AddSP = 0
        Item(EditorIndex).AddStr = 0
        Item(EditorIndex).AddDef = 0
        Item(EditorIndex).AddMagi = 0
        Item(EditorIndex).AddSpeed = 0
        Item(EditorIndex).AddEXP = 0
    End If
    
    If (frmEditItem.cmbType.ListIndex = ITEM_TYPE_SPELL) Then
        Item(EditorIndex).Data1 = frmEditItem.scrlSpell.Value
        Item(EditorIndex).Data2 = 0
        Item(EditorIndex).Data3 = 0
        Item(EditorIndex).StrReq = 0
        Item(EditorIndex).DefReq = 0
        Item(EditorIndex).SpeedReq = 0
        Item(EditorIndex).MagiReq = 0
        Item(EditorIndex).ClassReq = -1
        Item(EditorIndex).AccessReq = 0
        
        Item(EditorIndex).AddHP = 0
        Item(EditorIndex).AddMP = 0
        Item(EditorIndex).AddSP = 0
        Item(EditorIndex).AddStr = 0
        Item(EditorIndex).AddDef = 0
        Item(EditorIndex).AddMagi = 0
        Item(EditorIndex).AddSpeed = 0
        Item(EditorIndex).AddEXP = 0
    End If
    
    Call SendSaveItem(EditorIndex)
    InItemsEditor = False
    Unload frmEditItem
End Sub

Public Sub ItemEditorCancel()
    InItemsEditor = False
    Unload frmEditItem
End Sub

Public Sub NpcEditorInit()
On Error Resume Next
    
    frmEditNpc.picSprites.Picture = LoadPicture(App.Path & "\Graphics\sprites.bmp")
    
    frmEditNpc.txtName.Text = Trim$(Npc(EditorIndex).Name)
    frmEditNpc.txtAttackSay.Text = Trim$(Npc(EditorIndex).AttackSay)
    frmEditNpc.scrlSprite.Value = Npc(EditorIndex).Sprite
    frmEditNpc.txtSpawnSecs.Text = STR$(Npc(EditorIndex).SpawnSecs)
    frmEditNpc.cmbBehavior.ListIndex = Npc(EditorIndex).Behavior
    frmEditNpc.txtRange.Text = Npc(EditorIndex).Range
    frmEditNpc.txtStrength.Text = Npc(EditorIndex).STR
    frmEditNpc.txtDef.Text = Npc(EditorIndex).DEF
    frmEditNpc.txtSPEED.Text = Npc(EditorIndex).Speed
    frmEditNpc.txtMagi.Text = Npc(EditorIndex).MAGI
    frmEditNpc.BigNpc.Value = Npc(EditorIndex).Big
    frmEditNpc.txtStartingHP.Text = Npc(EditorIndex).MaxHp
    frmEditNpc.txtExp.Text = Npc(EditorIndex).EXP
    frmEditNpc.txtKarma.Text = Npc(EditorIndex).Alignment
    frmEditNpc.txtChance.Text = STR$(Npc(EditorIndex).ItemNPC(1).Chance)
    frmEditNpc.scrlNum.Value = Npc(EditorIndex).ItemNPC(1).ItemNum
    frmEditNpc.scrlValue.Value = Npc(EditorIndex).ItemNPC(1).ItemValue
    
    frmEditNpc.Show vbModal
End Sub

Public Sub NpcEditorOk()
    Npc(EditorIndex).Name = frmEditNpc.txtName.Text
    Npc(EditorIndex).AttackSay = frmEditNpc.txtAttackSay.Text
    Npc(EditorIndex).Sprite = frmEditNpc.scrlSprite.Value
    Npc(EditorIndex).SpawnSecs = Val(frmEditNpc.txtSpawnSecs.Text)
    Npc(EditorIndex).Behavior = frmEditNpc.cmbBehavior.ListIndex
    Npc(EditorIndex).Range = Val(frmEditNpc.txtRange.Text)
    Npc(EditorIndex).STR = Val(frmEditNpc.txtStrength.Text)
    Npc(EditorIndex).DEF = Val(frmEditNpc.txtDef.Text)
    Npc(EditorIndex).Speed = Val(frmEditNpc.txtSPEED.Text)
    Npc(EditorIndex).MAGI = Val(frmEditNpc.txtMagi.Text)
    Npc(EditorIndex).Big = frmEditNpc.BigNpc.Value
    Npc(EditorIndex).MaxHp = Val(frmEditNpc.txtStartingHP.Text)
    Npc(EditorIndex).EXP = Val(frmEditNpc.txtExp.Text)
    Npc(EditorIndex).Alignment = Val(frmEditNpc.txtKarma.Text)
    
    Call SendSaveNpc(EditorIndex)
    InNpcEditor = False
    Unload frmEditNpc
End Sub

Public Sub NpcEditorCancel()
    InNpcEditor = False
    Unload frmEditNpc
End Sub

Public Sub NpcEditorBltSprite()
    If frmEditNpc.BigNpc.Value = Checked Then
        Call BitBlt(frmEditNpc.picSprite.hDC, 0, 0, 64, 64, frmEditNpc.picSprites.hDC, 3 * 64, frmEditNpc.scrlSprite.Value * 64, SRCCOPY)
    Else
        Call BitBlt(frmEditNpc.picSprite.hDC, 0, 0, PIC_X, PIC_Y, frmEditNpc.picSprites.hDC, 3 * PIC_X, frmEditNpc.scrlSprite.Value * PIC_Y, SRCCOPY)
    End If
End Sub

Public Sub ShopEditorInit()
On Error Resume Next

Dim i As Long

    frmEditShop.txtName.Text = Trim$(Shop(EditorIndex).Name)
    frmEditShop.txtJoinSay.Text = Trim$(Shop(EditorIndex).JoinSay)
    frmEditShop.txtLeaveSay.Text = Trim$(Shop(EditorIndex).LeaveSay)
    frmEditShop.chkFixesItems.Value = Shop(EditorIndex).FixesItems
    
    frmEditShop.cmbItemGive.Clear
    frmEditShop.cmbItemGive.AddItem "None"
    frmEditShop.cmbItemGet.Clear
    frmEditShop.cmbItemGet.AddItem "None"
    For i = 1 To MAX_ITEMS
        frmEditShop.cmbItemGive.AddItem i & ": " & Trim$(Item(i).Name)
        frmEditShop.cmbItemGet.AddItem i & ": " & Trim$(Item(i).Name)
    Next i
    frmEditShop.cmbItemGive.ListIndex = 0
    frmEditShop.cmbItemGet.ListIndex = 0
    
    Call UpdateShopTrade
    
    frmEditShop.Show vbModal
End Sub

Public Sub UpdateShopTrade()
Dim i As Long, GetItem As Long, GetValue As Long, GiveItem As Long, GiveValue As Long
    
    frmEditShop.lstTradeItem.Clear
    For i = 1 To MAX_TRADES
        GetItem = Shop(EditorIndex).TradeItem(i).GetItem
        GetValue = Shop(EditorIndex).TradeItem(i).GetValue
        GiveItem = Shop(EditorIndex).TradeItem(i).GiveItem
        GiveValue = Shop(EditorIndex).TradeItem(i).GiveValue
        
        If GetItem > 0 And GiveItem > 0 Then
            frmEditShop.lstTradeItem.AddItem i & ": " & GiveValue & " " & Trim$(Item(GiveItem).Name) & " for " & GetValue & " " & Trim$(Item(GetItem).Name)
        Else
            frmEditShop.lstTradeItem.AddItem "Empty Trade Slot"
        End If
    Next i
    frmEditShop.lstTradeItem.ListIndex = 0
End Sub

Public Sub ShopEditorOk()
    Shop(EditorIndex).Name = frmEditShop.txtName.Text
    Shop(EditorIndex).JoinSay = frmEditShop.txtJoinSay.Text
    Shop(EditorIndex).LeaveSay = frmEditShop.txtLeaveSay.Text
    Shop(EditorIndex).FixesItems = frmEditShop.chkFixesItems.Value
    
    Call SendSaveShop(EditorIndex)
    InShopEditor = False
    Unload frmEditShop
End Sub

Public Sub ShopEditorCancel()
    InShopEditor = False
    Unload frmEditShop
End Sub

Public Sub SpellEditorInit()
On Error Resume Next
Dim i As Long

    frmEditSpell.cmbClassReq.AddItem "All Classes"
    For i = 0 To Max_Classes
        frmEditSpell.cmbClassReq.AddItem Trim$(Class(i).Name)
    Next i
    
    frmEditSpell.txtName.Text = Trim$(Spell(EditorIndex).Name)
    frmEditSpell.picSpells.Picture = LoadPicture(App.Path & "\graphics\spellicons.bmp")
    frmEditSpell.scrlSpellPic.Value = Spell(EditorIndex).Pic
    frmEditSpell.cmbClassReq.ListIndex = Spell(EditorIndex).ClassReq
    frmEditSpell.scrlLevelReq.Value = Spell(EditorIndex).LevelReq
        
    frmEditSpell.cmbType.ListIndex = Spell(EditorIndex).Type
    frmEditSpell.scrlVitalMod.Value = Spell(EditorIndex).Data1
    
    frmEditSpell.scrlCost.Value = Spell(EditorIndex).MPCost
    frmEditSpell.scrlSound.Value = Spell(EditorIndex).Sound
        
    frmEditSpell.Show vbModal
End Sub

Public Sub SpellEditorOk()
    Spell(EditorIndex).Name = frmEditSpell.txtName.Text
    Spell(EditorIndex).Pic = frmEditSpell.scrlSpellPic.Value
    Spell(EditorIndex).ClassReq = frmEditSpell.cmbClassReq.ListIndex
    Spell(EditorIndex).LevelReq = frmEditSpell.scrlLevelReq.Value
    Spell(EditorIndex).Type = frmEditSpell.cmbType.ListIndex
    Spell(EditorIndex).Data1 = frmEditSpell.scrlVitalMod.Value
    Spell(EditorIndex).Data3 = 0
    Spell(EditorIndex).MPCost = frmEditSpell.scrlCost.Value
    Spell(EditorIndex).Sound = frmEditSpell.scrlSound.Value
    
    Call SendSaveSpell(EditorIndex)
    InSpellEditor = False
    Unload frmEditSpell
End Sub

Public Sub SpellEditorCancel()
    InSpellEditor = False
    Unload frmEditSpell
End Sub

Public Sub UpdateInventory()
Dim i As Long

    frmEndieko.lstInv.Clear
    
    ' Show the inventory
    For i = 1 To MAX_INV
        If GetPlayerInvItemNum(MyIndex, i) > 0 And GetPlayerInvItemNum(MyIndex, i) <= MAX_ITEMS Then
            If Item(GetPlayerInvItemNum(MyIndex, i)).Type = ITEM_TYPE_CURRENCY Then
                frmEndieko.lstInv.AddItem i & ": " & Trim$(Item(GetPlayerInvItemNum(MyIndex, i)).Name) & " (" & GetPlayerInvItemValue(MyIndex, i) & ")"
            Else
                ' Check if this item is being worn
                If GetPlayerWeaponSlot(MyIndex) = i Or GetPlayerArmorSlot(MyIndex) = i Or GetPlayerHelmetSlot(MyIndex) = i Or GetPlayerShieldSlot(MyIndex) = i Then
                    frmEndieko.lstInv.AddItem i & ": " & Trim$(Item(GetPlayerInvItemNum(MyIndex, i)).Name) & " (worn)"
                Else
                    frmEndieko.lstInv.AddItem i & ": " & Trim$(Item(GetPlayerInvItemNum(MyIndex, i)).Name)
                End If
            End If
        Else
            frmEndieko.lstInv.AddItem "<free inventory slot>"
        End If
    Next i
    
    frmEndieko.lstInv.ListIndex = 0
End Sub

Public Sub UpdateTradeInventory()
Dim i As Long

    frmPlayerTrade.PlayerInv1.Clear
    
For i = 1 To MAX_INV
    If GetPlayerInvItemNum(MyIndex, i) > 0 And GetPlayerInvItemNum(MyIndex, i) <= MAX_ITEMS Then
        If Item(GetPlayerInvItemNum(MyIndex, i)).Type = ITEM_TYPE_CURRENCY Then
            frmPlayerTrade.PlayerInv1.AddItem i & ": " & Trim$(Item(GetPlayerInvItemNum(MyIndex, i)).Name) & " (" & GetPlayerInvItemValue(MyIndex, i) & ")"
        Else
            If GetPlayerWeaponSlot(MyIndex) = i Or GetPlayerArmorSlot(MyIndex) = i Or GetPlayerHelmetSlot(MyIndex) = i Or GetPlayerShieldSlot(MyIndex) = i Then
                frmPlayerTrade.PlayerInv1.AddItem i & ": " & Trim$(Item(GetPlayerInvItemNum(MyIndex, i)).Name) & " (worn)"
            Else
                frmPlayerTrade.PlayerInv1.AddItem i & ": " & Trim$(Item(GetPlayerInvItemNum(MyIndex, i)).Name)
            End If
        End If
    Else
        frmPlayerTrade.PlayerInv1.AddItem "<Nothing>"
    End If
Next i
    
    frmPlayerTrade.PlayerInv1.ListIndex = 0
End Sub

Sub PlayerSearch(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim x1 As Long, y1 As Long

    x1 = Int(X / PIC_X)
    y1 = Int(Y / PIC_Y)
    
    If (x1 >= 0) And (x1 <= MAX_MAPX) And (y1 >= 0) And (y1 <= MAX_MAPY) Then
        Call SendData("search" & SEP_CHAR & x1 & SEP_CHAR & y1 & SEP_CHAR & END_CHAR)
    End If
    MouseDownX = x1
    MouseDownY = y1
End Sub
Sub BltTile2(ByVal X As Long, ByVal Y As Long, ByVal Tile As Long)
    rec.Top = Int(Tile / 14) * PIC_Y
    rec.Bottom = rec.Top + PIC_Y
    rec.Left = (Tile - Int(Tile / 14) * 14) * PIC_X
    rec.Right = rec.Left + PIC_X
    Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) + sx - NewXOffset, Y - (NewPlayerY * PIC_Y) + sx - NewYOffset, DD_TileSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End Sub

Sub BltPlayerBars(ByVal Index As Long)
Dim X As Long, Y As Long

X = (GetPlayerX(Index) * PIC_X + sx + Player(Index).XOffset) - (NewPlayerX * PIC_X) - NewXOffset
Y = (GetPlayerY(Index) * PIC_Y + sx + Player(Index).YOffset) - (NewPlayerY * PIC_Y) - NewYOffset

If Player(Index).HP = 0 Then Exit Sub
    'draws the back bars
    Call DD_BackBuffer.SetFillColor(RGB(255, 0, 0))
    Call DD_BackBuffer.DrawBox(X, Y + 32, X + 32, Y + 36)
    
    'draws HP
    Call DD_BackBuffer.SetFillColor(RGB(0, 255, 0))
    Call DD_BackBuffer.DrawBox(X, Y + 32, X + ((Player(Index).HP / 100) / (Player(Index).MaxHp / 100) * 32), Y + 36)
End Sub

Public Sub UpdateVisInv()
Dim Index As Long
Dim d As Long
    
    For Index = 1 To MAX_INV
        If GetPlayerShieldSlot(MyIndex) <> Index Then frmEndieko.ShieldImage.Picture = LoadPicture()
        If GetPlayerWeaponSlot(MyIndex) <> Index Then frmEndieko.WeaponImage.Picture = LoadPicture()
        If GetPlayerHelmetSlot(MyIndex) <> Index Then frmEndieko.HelmetImage.Picture = LoadPicture()
        If GetPlayerArmorSlot(MyIndex) <> Index Then frmEndieko.ArmorImage.Picture = LoadPicture()
        If GetPlayerLegSlot(MyIndex) <> Index Then frmEndieko.LegImage.Picture = LoadPicture()
        If GetPlayerBootSlot(MyIndex) <> Index Then frmEndieko.BootImage.Picture = LoadPicture()
    Next Index
    
    For Index = 1 To MAX_INV
        If GetPlayerShieldSlot(MyIndex) = Index Then Call BitBlt(frmEndieko.ShieldImage.hDC, 0, 0, PIC_X, PIC_Y, frmEndieko.picItems.hDC, (Item(GetPlayerInvItemNum(MyIndex, Index)).Pic - Int(Item(GetPlayerInvItemNum(MyIndex, Index)).Pic / 6) * 6) * PIC_X, Int(Item(GetPlayerInvItemNum(MyIndex, Index)).Pic / 6) * PIC_Y, SRCCOPY)
        If GetPlayerWeaponSlot(MyIndex) = Index Then Call BitBlt(frmEndieko.WeaponImage.hDC, 0, 0, PIC_X, PIC_Y, frmEndieko.picItems.hDC, (Item(GetPlayerInvItemNum(MyIndex, Index)).Pic - Int(Item(GetPlayerInvItemNum(MyIndex, Index)).Pic / 6) * 6) * PIC_X, Int(Item(GetPlayerInvItemNum(MyIndex, Index)).Pic / 6) * PIC_Y, SRCCOPY)
        If GetPlayerHelmetSlot(MyIndex) = Index Then Call BitBlt(frmEndieko.HelmetImage.hDC, 0, 0, PIC_X, PIC_Y, frmEndieko.picItems.hDC, (Item(GetPlayerInvItemNum(MyIndex, Index)).Pic - Int(Item(GetPlayerInvItemNum(MyIndex, Index)).Pic / 6) * 6) * PIC_X, Int(Item(GetPlayerInvItemNum(MyIndex, Index)).Pic / 6) * PIC_Y, SRCCOPY)
        If GetPlayerArmorSlot(MyIndex) = Index Then Call BitBlt(frmEndieko.ArmorImage.hDC, 0, 0, PIC_X, PIC_Y, frmEndieko.picItems.hDC, (Item(GetPlayerInvItemNum(MyIndex, Index)).Pic - Int(Item(GetPlayerInvItemNum(MyIndex, Index)).Pic / 6) * 6) * PIC_X, Int(Item(GetPlayerInvItemNum(MyIndex, Index)).Pic / 6) * PIC_Y, SRCCOPY)
        If GetPlayerLegSlot(MyIndex) = Index Then Call BitBlt(frmEndieko.LegImage.hDC, 0, 0, PIC_X, PIC_Y, frmEndieko.picItems.hDC, (Item(GetPlayerInvItemNum(MyIndex, Index)).Pic - Int(Item(GetPlayerInvItemNum(MyIndex, Index)).Pic / 6) * 6) * PIC_X, Int(Item(GetPlayerInvItemNum(MyIndex, Index)).Pic / 6) * PIC_Y, SRCCOPY)
        If GetPlayerBootSlot(MyIndex) = Index Then Call BitBlt(frmEndieko.BootImage.hDC, 0, 0, PIC_X, PIC_Y, frmEndieko.picItems.hDC, (Item(GetPlayerInvItemNum(MyIndex, Index)).Pic - Int(Item(GetPlayerInvItemNum(MyIndex, Index)).Pic / 6) * 6) * PIC_X, Int(Item(GetPlayerInvItemNum(MyIndex, Index)).Pic / 6) * PIC_Y, SRCCOPY)
    Next Index
        
    frmEndieko.SelectedItem.Top = frmEndieko.picInv(frmEndieko.lstInv.ListIndex).Top - 1
    frmEndieko.SelectedItem.Left = frmEndieko.picInv(frmEndieko.lstInv.ListIndex).Left - 1
    
    frmEndieko.EquipS(0).Visible = False
    frmEndieko.EquipS(1).Visible = False
    frmEndieko.EquipS(2).Visible = False
    frmEndieko.EquipS(3).Visible = False
    frmEndieko.EquipS(4).Visible = False
    frmEndieko.EquipS(5).Visible = False

    For d = 0 To MAX_INV - 1
        If Player(MyIndex).Inv(d + 1).Num > 0 Then
            If Item(GetPlayerInvItemNum(MyIndex, d + 1)).Type <> ITEM_TYPE_CURRENCY Then
                If GetPlayerWeaponSlot(MyIndex) = d + 1 Then
                    frmEndieko.EquipS(0).Visible = True
                    frmEndieko.EquipS(0).Top = frmEndieko.picInv(d).Top - 2
                    frmEndieko.EquipS(0).Left = frmEndieko.picInv(d).Left - 2
                ElseIf GetPlayerArmorSlot(MyIndex) = d + 1 Then
                    frmEndieko.EquipS(1).Visible = True
                    frmEndieko.EquipS(1).Top = frmEndieko.picInv(d).Top - 2
                    frmEndieko.EquipS(1).Left = frmEndieko.picInv(d).Left - 2
                ElseIf GetPlayerHelmetSlot(MyIndex) = d + 1 Then
                    frmEndieko.EquipS(2).Visible = True
                    frmEndieko.EquipS(2).Top = frmEndieko.picInv(d).Top - 2
                    frmEndieko.EquipS(2).Left = frmEndieko.picInv(d).Left - 2
                ElseIf GetPlayerShieldSlot(MyIndex) = d + 1 Then
                    frmEndieko.EquipS(3).Visible = True
                    frmEndieko.EquipS(3).Top = frmEndieko.picInv(d).Top - 2
                    frmEndieko.EquipS(3).Left = frmEndieko.picInv(d).Left - 2
                ElseIf GetPlayerBootSlot(MyIndex) = d + 1 Then
                    frmEndieko.EquipS(4).Visible = True
                    frmEndieko.EquipS(4).Top = frmEndieko.picInv(d).Top - 2
                    frmEndieko.EquipS(4).Left = frmEndieko.picInv(d).Left - 2
                ElseIf GetPlayerLegSlot(MyIndex) = d + 1 Then
                    frmEndieko.EquipS(5).Visible = True
                    frmEndieko.EquipS(5).Top = frmEndieko.picInv(d).Top - 2
                    frmEndieko.EquipS(5).Left = frmEndieko.picInv(d).Left - 2
                End If
            End If
        End If
    Next d
End Sub

Sub BltSpriteChange(ByVal X As Long, ByVal Y As Long)
Dim x2 As Long, y2 As Long
    
    If Map.Tile(X, Y).Type = TILE_TYPE_SPRITE_CHANGE Then

        ' Only used if ever want to switch to blt rather then bltfast
        With rec_pos
            .Top = Y * PIC_Y
            .Bottom = .Top + PIC_Y
            .Left = X * PIC_X
            .Right = .Left + PIC_X
        End With
                                        
        rec.Top = Map.Tile(X, Y).Data1 * PIC_Y + 16
        rec.Bottom = rec.Top + PIC_Y - 16
        rec.Left = 96
        rec.Right = rec.Left + PIC_X
        
        x2 = X * PIC_X + sx
        y2 = Y * PIC_Y + sx
                                           
        'Call DD_BackBuffer.Blt(rec_pos, DD_SpriteSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
        Call DD_BackBuffer.BltFast(x2 - (NewPlayerX * PIC_X) - NewXOffset, y2 - (NewPlayerY * PIC_Y) - NewYOffset, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    End If
End Sub

Sub BltSpriteChange2(ByVal X As Long, ByVal Y As Long)
Dim x2 As Long, y2 As Long
    
    If Map.Tile(X, Y).Type = TILE_TYPE_SPRITE_CHANGE Then

        ' Only used if ever want to switch to blt rather then bltfast
        With rec_pos
            .Top = Y * PIC_Y
            .Bottom = .Top + PIC_Y
            .Left = X * PIC_X
            .Right = .Left + PIC_X
        End With
                                        
        rec.Top = Map.Tile(X, Y).Data1 * PIC_Y
        rec.Bottom = rec.Top + PIC_Y - 16
        rec.Left = 96
        rec.Right = rec.Left + PIC_X
        
        x2 = X * PIC_X + sx
        y2 = Y * PIC_Y + (sx / 2) '- 16
               
        If y2 < 0 Then
            Exit Sub
        End If
                            
        'Call DD_BackBuffer.Blt(rec_pos, DD_SpriteSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
        Call DD_BackBuffer.BltFast(x2 - (NewPlayerX * PIC_X) - NewXOffset, y2 - (NewPlayerY * PIC_Y) - NewYOffset, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    End If
End Sub

Sub BltEmoticons(ByVal Index As Long)
Dim x2 As Long, y2 As Long
   
    If Player(Index).Emoticon < 0 Then Exit Sub
    
    If Player(Index).EmoticonT + 5000 > GetTickCount Then
        rec.Top = Player(Index).Emoticon * PIC_Y
        rec.Bottom = rec.Top + PIC_Y
        rec.Left = 0
        rec.Right = PIC_X
        
        If Index = MyIndex Then
            x2 = NewX + sx + 16
            y2 = NewY + sx - 32
            
            If y2 < 0 Then
                Exit Sub
            End If
            
            Call DD_BackBuffer.BltFast(x2, y2, DD_EmoticonSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        Else
            x2 = GetPlayerX(Index) * PIC_X + sx + Player(Index).XOffset + 16
            y2 = GetPlayerY(Index) * PIC_Y + sx + Player(Index).YOffset - 32
            
            If y2 < 0 Then
                Exit Sub
            End If
            
            Call DD_BackBuffer.BltFast(x2 - (NewPlayerX * PIC_X) - NewXOffset, y2 - (NewPlayerY * PIC_Y) - NewYOffset, DD_EmoticonSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    End If
End Sub

Public Sub ArrowEditorInit()
Dim i As Long
    frmEditArrows.scrlArrow.Max = MAX_ARROWS
    If Arrows(EditorIndex).Pic = 0 Then Arrows(EditorIndex).Pic = 1
    frmEditArrows.scrlArrow.Value = Arrows(EditorIndex).Pic
    frmEditArrows.txtName.Text = Arrows(EditorIndex).Name
    If Arrows(EditorIndex).Range = 0 Then Arrows(EditorIndex).Range = 1
    frmEditArrows.scrlRange.Value = Arrows(EditorIndex).Range
    frmEditArrows.picArrows.Picture = LoadPicture(App.Path & "\graphics\arrows.bmp")
    If Arrows(EditorIndex).HasAmmo = 1 Then
        frmEditArrows.chkHasAmmo.Value = 1
        frmEditArrows.fraAmmo.Enabled = True
    End If
    frmEditArrows.cmbAmmo.AddItem "None", 0
    For i = 1 To MAX_ITEMS
        frmEditArrows.cmbAmmo.AddItem i & ": " & Item(i).Name
    Next i
    frmEditArrows.cmbAmmo.ListIndex = Arrows(EditorIndex).Ammunition
    frmEditArrows.Show vbModal
End Sub

Public Sub ArrowEditorOk()
    Arrows(EditorIndex).Pic = frmEditArrows.scrlArrow.Value
    Arrows(EditorIndex).Range = frmEditArrows.scrlRange.Value
    Arrows(EditorIndex).Name = frmEditArrows.txtName.Text
    Arrows(EditorIndex).HasAmmo = frmEditArrows.chkHasAmmo.Value
    If frmEditArrows.chkHasAmmo.Value = 1 Then
        Arrows(EditorIndex).Ammunition = frmEditArrows.cmbAmmo.ListIndex
    Else
        Arrows(EditorIndex).Ammunition = 0
    End If
    Call SendSaveArrow(EditorIndex)
    Call ArrowEditorCancel
End Sub

Public Sub ArrowEditorCancel()
    InArrowEditor = False
    Unload frmEditArrows
End Sub

Public Sub UpdateBank()
Dim i As Long
'Dim strToolTip As String
    ' Show the Bank
    For i = 0 To MAX_INV - 1
        If i <= MAX_BANK - 1 Then
            If GetPlayerBankItemNum(MyIndex, i + 1) > 0 And GetPlayerBankItemNum(MyIndex, i + 1) <= MAX_ITEMS Then
                Call BitBlt(frmBank.picBank(i).hDC, 0, 0, PIC_X, PIC_Y, frmEndieko.picItems.hDC, (Item(GetPlayerBankItemNum(MyIndex, i + 1)).Pic - Int(Item(GetPlayerBankItemNum(MyIndex, i + 1)).Pic / 6) * 6) * PIC_X, Int(Item(GetPlayerBankItemNum(MyIndex, i + 1)).Pic / 6) * PIC_Y, SRCCOPY)
                If Item(GetPlayerBankItemNum(MyIndex, i + 1)).Type = ITEM_TYPE_CURRENCY Then
                    frmBank.picBank(i).ToolTipText = i & ": " & Trim$(Item(GetPlayerBankItemNum(MyIndex, i + 1)).Name) & " (" & GetPlayerBankItemValue(MyIndex, i + 1) & ")"
                Else
                    ' Check if this item is being worn
                    If GetPlayerWeaponSlot(MyIndex) = i Or GetPlayerArmorSlot(MyIndex) = i Or GetPlayerHelmetSlot(MyIndex) = i Or GetPlayerShieldSlot(MyIndex) = i Then
                        frmBank.picBank(i).ToolTipText = i & ": " & Trim$(Item(GetPlayerBankItemNum(MyIndex, i + 1)).Name) & " (worn)"
                    Else
                        frmBank.picBank(i).ToolTipText = i & ": " & Trim$(Item(GetPlayerBankItemNum(MyIndex, i + 1)).Name)
                    End If
                End If
            Else
                frmBank.picBank(i).Picture = LoadPicture
                frmBank.picBank(i).ToolTipText = ""
            End If
        End If
        If GetPlayerInvItemNum(MyIndex, i + 1) > 0 And GetPlayerInvItemNum(MyIndex, i + 1) <= MAX_ITEMS Then
            Call BitBlt(frmBank.picInv(i).hDC, 0, 0, PIC_X, PIC_Y, frmEndieko.picItems.hDC, (Item(GetPlayerInvItemNum(MyIndex, i + 1)).Pic - Int(Item(GetPlayerInvItemNum(MyIndex, i + 1)).Pic / 6) * 6) * PIC_X, Int(Item(GetPlayerInvItemNum(MyIndex, i + 1)).Pic / 6) * PIC_Y, SRCCOPY)
            If Item(GetPlayerInvItemNum(MyIndex, i + 1)).Type = ITEM_TYPE_CURRENCY Then
                frmBank.picInv(i).ToolTipText = i & ": " & Trim$(Item(GetPlayerInvItemNum(MyIndex, i + 1)).Name) & " (" & GetPlayerInvItemValue(MyIndex, i + 1) & ")"
            Else
                ' Check if this item is being worn
                If GetPlayerWeaponSlot(MyIndex) = i Or GetPlayerArmorSlot(MyIndex) = i Or GetPlayerHelmetSlot(MyIndex) = i Or GetPlayerShieldSlot(MyIndex) = i Then
                    frmBank.picInv(i).ToolTipText = i & ": " & Trim$(Item(GetPlayerInvItemNum(MyIndex, i + 1)).Name) & " (worn)"
                Else
                    frmBank.picInv(i).ToolTipText = i & ": " & Trim$(Item(GetPlayerInvItemNum(MyIndex, i + 1)).Name)
                End If
            End If
        Else
            frmBank.picInv(i).Picture = LoadPicture
            frmBank.picInv(i).ToolTipText = ""
        End If
    Next i
    BankSelect = 0
    Inventory = 0
End Sub

Public Sub UpdateVisSpell()
Dim Index As Long
Dim d As Long
    frmEndieko.SelectedSpell.Top = frmEndieko.picSpell(frmEndieko.lstSpells.ListIndex).Top - 2
    frmEndieko.SelectedSpell.Left = frmEndieko.picSpell(frmEndieko.lstSpells.ListIndex).Left - 2
End Sub

Sub ItemSelected(ByVal Index As Long, ByVal Selected As Long)
Dim index2 As Long
index2 = Trade(Selected).ItemS(Index).ItemGetNum

    frmTrade.shpSelect.Top = frmTrade.picItem(Index - 1).Top - 1
    frmTrade.shpSelect.Left = frmTrade.picItem(Index - 1).Left - 1

    If index2 <= 0 Then
        Call ClearItemSelected
        Exit Sub
    End If

    frmTrade.descName.Caption = "Name: " & Trim(Item(index2).Name)
    frmTrade.descQuantity.Caption = "Quantity: " & Trade(Selected).ItemS(Index).ItemGetVal
    
    frmTrade.descStr.Caption = "Strength Req: " & Item(index2).StrReq
    frmTrade.descDef.Caption = "Defence Req: " & Item(index2).DefReq
    frmTrade.descDef.Caption = "Speed Req: " & Item(index2).SpeedReq
    
    frmTrade.descAStr.Caption = "Strength: " & Item(index2).AddStr
    frmTrade.descADef.Caption = "Defence: " & Item(index2).AddDef
    frmTrade.descAMagi.Caption = "Magic: " & Item(index2).AddMagi
    frmTrade.descASpeed.Caption = "Speed: " & Item(index2).AddSpeed
    
    frmTrade.descHp.Caption = "Hp: " & Item(index2).AddHP
    frmTrade.descMp.Caption = "Mp: " & Item(index2).AddMP
    frmTrade.descSp.Caption = "Sp: " & Item(index2).AddSP

    frmTrade.descAExp.Caption = "Exp: " & Item(index2).AddEXP
    frmTrade.desc.Caption = Trim(Item(index2).desc)
    
    frmTrade.lblTradeFor.Caption = Trim(Item(Trade(Selected).ItemS(Index).ItemGiveNum).Name)
    frmTrade.lblQuantity.Caption = Trade(Selected).ItemS(Index).ItemGiveVal
End Sub

Sub ClearItemSelected()
    frmTrade.lblTradeFor.Caption = ""
    frmTrade.lblQuantity.Caption = ""
    
    frmTrade.descName.Caption = "Name: " & ""
    frmTrade.descQuantity.Caption = "Quantity: " & ""
    
    frmTrade.descStr.Caption = "Strength Req: " & 0
    frmTrade.descDef.Caption = "Defence Req: " & 0
    frmTrade.descDef.Caption = "Speed Req: " & 0
    
    frmTrade.descAStr.Caption = "Strength: " & 0
    frmTrade.descADef.Caption = "Defence: " & 0
    frmTrade.descAMagi.Caption = "Magic: " & 0
    frmTrade.descASpeed.Caption = "Speed: " & 0
    
    frmTrade.descHp.Caption = "Hp: " & 0
    frmTrade.descMp.Caption = "Mp: " & 0
    frmTrade.descSp.Caption = "Sp: " & 0

    frmTrade.descAExp.Caption = "Exp: " & 0
    frmTrade.desc.Caption = ""
End Sub

Public Sub EffectEditorInit()
On Error Resume Next
Dim i As Long
    
    frmEditEffect.txtName.Text = Trim$(Effect(EditorIndex).Name)
    frmEditEffect.cmbTime.ListIndex = Effect(EditorIndex).Time
    frmEditEffect.cmbEffect.ListIndex = Effect(EditorIndex).Effect
    Select Case frmEditEffect.cmbEffect.ListIndex
        Case 0
            frmEditEffect.fraDrain.Visible = True
            frmEditEffect.fraFortify.Visible = False
            frmEditEffect.cmbDrain.ListIndex = Effect(EditorIndex).Data1
            frmEditEffect.scrlDrainAmount.Value = Effect(EditorIndex).Data2
        Case 1
            frmEditEffect.fraDrain.Visible = False
            frmEditEffect.fraFortify.Visible = True
            frmEditEffect.cmbFortify.ListIndex = Effect(EditorIndex).Data1
            frmEditEffect.scrlFortifyAmount.Value = Effect(EditorIndex).Data2
        Case 2
            frmEditEffect.fraDrain.Visible = False
            frmEditEffect.fraFortify.Visible = False
    End Select
    frmEditEffect.Show vbModal
End Sub

Public Sub EffectEditorOk()
    Effect(EditorIndex).Name = frmEditEffect.txtName.Text
    Effect(EditorIndex).Effect = frmEditEffect.cmbEffect.ListIndex
    Effect(EditorIndex).Time = frmEditEffect.cmbTime.ListIndex
    
    Select Case frmEditEffect.cmbEffect.ListIndex
        Case EFFECT_TYPE_DRAIN
            Effect(EditorIndex).Data1 = frmEditEffect.cmbDrain.ListIndex
            Effect(EditorIndex).Data2 = frmEditEffect.scrlDrainAmount.Value
            Effect(EditorIndex).Data3 = 0
        Case EFFECT_TYPE_FORTIFY
            Effect(EditorIndex).Data1 = frmEditEffect.cmbFortify.ListIndex
            Effect(EditorIndex).Data2 = frmEditEffect.scrlFortifyAmount.Value
            Effect(EditorIndex).Data3 = 0
        Case EFFECT_TYPE_FREEZE
            Effect(EditorIndex).Data1 = 0
            Effect(EditorIndex).Data2 = 0
            Effect(EditorIndex).Data3 = 0
    End Select
    
    Call SendSaveEffect(EditorIndex)
    InEffectEditor = False
    Unload frmEditEffect
End Sub

Public Sub EffectEditorCancel()
    InEffectEditor = False
    Unload frmEditEffect
End Sub
