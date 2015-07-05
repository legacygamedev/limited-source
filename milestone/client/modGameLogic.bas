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
Public SaveMapItem() As MapItemRec
Public SaveMapNpc(1 To MAX_MAP_NPCS) As MapNpcRec

' Used for index based editors
Public InItemsEditor As Boolean
Public InNpcEditor As Boolean
Public InShopEditor As Boolean
Public InSpellEditor As Boolean
Public InEmoticonEditor As Boolean
Public EditorIndex As Long

' Game fps
Public GameFPS As Long

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
    If LCase(Dir(App.Path & "\Maps", vbDirectory)) <> "maps" Then
        Call MkDir(App.Path & "\Maps")
    End If
    If UCase(Dir(App.Path & "\GFX", vbDirectory)) <> "GFX" Then
        Call MkDir(App.Path & "\GFX")
    End If
    If UCase(Dir(App.Path & "\GUI", vbDirectory)) <> "GUI" Then
        Call MkDir(App.Path & "\GUI")
    End If
    If UCase(Dir(App.Path & "\Music", vbDirectory)) <> "MUSIC" Then
        Call MkDir(App.Path & "\Music")
    End If
    If UCase(Dir(App.Path & "\SFX", vbDirectory)) <> "SFX" Then
        Call MkDir(App.Path & "\SFX")
    End If
    If UCase(Dir(App.Path & "\Flashs", vbDirectory)) <> "FLASHS" Then
        Call MkDir(App.Path & "\Flashs")
    End If
    
    Call SetStatus("Loading GUI...")
    DoEvents
    
    If FileExist("GUI\Loading.gif") Then frmSendGetData.Picture = LoadPicture(App.Path & "\GUI\Loading.gif")
    If FileExist("GUI\CharacterSelect.gif") Then frmChars.Picture = LoadPicture(App.Path & "\GUI\CharacterSelect.gif")
    If FileExist("GUI\Credits.gif") Then frmCredits.Picture = LoadPicture(App.Path & "\GUI\Credits.gif")
    If FileExist("GUI\DeleteAccount.gif") Then frmDeleteAccount.Picture = LoadPicture(App.Path & "\GUI\DeleteAccount.gif")
    If FileExist("GUI\FixItems.gif") Then frmFixItem.Picture = LoadPicture(App.Path & "\GUI\FixItems.gif")
    If FileExist("GUI\IPConfig.gif") Then frmIpconfig.Picture = LoadPicture(App.Path & "\GUI\IPConfig.gif")
    If FileExist("GUI\Notes.gif") Then frmKeepNotes.Picture = LoadPicture(App.Path & "\GUI\Notes.gif")
    If FileExist("GUI\Login.gif") Then frmLogin.Picture = LoadPicture(App.Path & "\GUI\Login.gif")
    If FileExist("GUI\MainMenu.gif") Then frmMainMenu.Picture = LoadPicture(App.Path & "\GUI\MainMenu.gif")
    If FileExist("GUI\NewAccount.gif") Then frmNewAccount.Picture = LoadPicture(App.Path & "\GUI\NewAccount.gif")
    If FileExist("GUI\NewCharacter.gif") Then frmNewChar.Picture = LoadPicture(App.Path & "\GUI\NewCharacter.gif")
    If FileExist("GUI\Shop.gif") Then frmTrade.Picture = LoadPicture(App.Path & "\GUI\Shop.gif")
    If FileExist("GUI\Training.gif") Then frmTraining.Picture = LoadPicture(App.Path & "\GUI\Training.gif")
    If FileExist("GUI\CreateGuild.gif") Then frmGuild.Picture = LoadPicture(App.Path & "\GUI\CreateGuild.gif")
    If FileExist("GUI\Game.gif") Then frmMirage.Picture = LoadPicture(App.Path & "\GUI\Game.gif")
    If FileExist("GUI\Trade.gif") Then frmPlayerTrade.Picture = LoadPicture(App.Path & "\GUI\Trade.gif")
    If FileExist("GUI\Player Chat.gif") Then frmPlayerChat.Picture = LoadPicture(App.Path & "\GUI\Player Chat.gif")
    If FileExist("GUI\FlashTheatre.gif") Then frmFlash.Picture = LoadPicture(App.Path & "\GUI\FlashTheatre.gif")
    
    If FileExist("GUI\Loading.jpg") Then frmSendGetData.Picture = LoadPicture(App.Path & "\GUI\Loading.jpg")
    If FileExist("GUI\CharacterSelect.jpg") Then frmChars.Picture = LoadPicture(App.Path & "\GUI\CharacterSelect.jpg")
    If FileExist("GUI\Credits.jpg") Then frmCredits.Picture = LoadPicture(App.Path & "\GUI\Credits.jpg")
    If FileExist("GUI\DeleteAccount.jpg") Then frmDeleteAccount.Picture = LoadPicture(App.Path & "\GUI\DeleteAccount.jpg")
    If FileExist("GUI\FixItems.jpg") Then frmFixItem.Picture = LoadPicture(App.Path & "\GUI\FixItems.jpg")
    If FileExist("GUI\IPConfig.jpg") Then frmIpconfig.Picture = LoadPicture(App.Path & "\GUI\IPConfig.jpg")
    If FileExist("GUI\Notes.jpg") Then frmKeepNotes.Picture = LoadPicture(App.Path & "\GUI\Notes.jpg")
    If FileExist("GUI\Login.jpg") Then frmLogin.Picture = LoadPicture(App.Path & "\GUI\Login.jpg")
    If FileExist("GUI\MainMenu.jpg") Then frmMainMenu.Picture = LoadPicture(App.Path & "\GUI\MainMenu.jpg")
    If FileExist("GUI\NewAccount.jpg") Then frmNewAccount.Picture = LoadPicture(App.Path & "\GUI\NewAccount.jpg")
    If FileExist("GUI\NewCharacter.jpg") Then frmNewChar.Picture = LoadPicture(App.Path & "\GUI\NewCharacter.jpg")
    If FileExist("GUI\Shop.jpg") Then frmTrade.Picture = LoadPicture(App.Path & "\GUI\Shop.jpg")
    If FileExist("GUI\Training.jpg") Then frmTraining.Picture = LoadPicture(App.Path & "\GUI\Training.jpg")
    If FileExist("GUI\CreateGuild.jpg") Then frmGuild.Picture = LoadPicture(App.Path & "\GUI\CreateGuild.jpg")
    If FileExist("GUI\Game.jpg") Then frmMirage.Picture = LoadPicture(App.Path & "\GUI\Game.jpg")
    If FileExist("GUI\Trade.jpg") Then frmPlayerTrade.Picture = LoadPicture(App.Path & "\GUI\Trade.jpg")
    If FileExist("GUI\Player Chat.jpg") Then frmPlayerChat.Picture = LoadPicture(App.Path & "\GUI\Player Chat.jpg")
    If FileExist("GUI\FlashTheatre.jpg") Then frmFlash.Picture = LoadPicture(App.Path & "\GUI\FlashTheatre.jpg")
    
    If FileExist("GUI\Loading.png") Then frmSendGetData.Picture = LoadPicture(App.Path & "\GUI\Loading.png")
    If FileExist("GUI\CharacterSelect.png") Then frmChars.Picture = LoadPicture(App.Path & "\GUI\CharacterSelect.png")
    If FileExist("GUI\Credits.png") Then frmCredits.Picture = LoadPicture(App.Path & "\GUI\Credits.png")
    If FileExist("GUI\DeleteAccount.png") Then frmDeleteAccount.Picture = LoadPicture(App.Path & "\GUI\DeleteAccount.png")
    If FileExist("GUI\FixItems.png") Then frmFixItem.Picture = LoadPicture(App.Path & "\GUI\FixItems.png")
    If FileExist("GUI\IPConfig.png") Then frmIpconfig.Picture = LoadPicture(App.Path & "\GUI\IPConfig.png")
    If FileExist("GUI\Notes.png") Then frmKeepNotes.Picture = LoadPicture(App.Path & "\GUI\Notes.png")
    If FileExist("GUI\Login.png") Then frmLogin.Picture = LoadPicture(App.Path & "\GUI\Login.png")
    If FileExist("GUI\MainMenu.png") Then frmMainMenu.Picture = LoadPicture(App.Path & "\GUI\MainMenu.png")
    If FileExist("GUI\NewAccount.png") Then frmNewAccount.Picture = LoadPicture(App.Path & "\GUI\NewAccount.png")
    If FileExist("GUI\NewCharacter.png") Then frmNewChar.Picture = LoadPicture(App.Path & "\GUI\NewCharacter.png")
    If FileExist("GUI\Shop.png") Then frmTrade.Picture = LoadPicture(App.Path & "\GUI\Shop.png")
    If FileExist("GUI\Training.png") Then frmTraining.Picture = LoadPicture(App.Path & "\GUI\Training.png")
    If FileExist("GUI\CreateGuild.png") Then frmGuild.Picture = LoadPicture(App.Path & "\GUI\CreateGuild.png")
    If FileExist("GUI\Game.png") Then frmMirage.Picture = LoadPicture(App.Path & "\GUI\Game.png")
    If FileExist("GUI\Trade.png") Then frmPlayerTrade.Picture = LoadPicture(App.Path & "\GUI\Trade.png")
    If FileExist("GUI\Player Chat.png") Then frmPlayerChat.Picture = LoadPicture(App.Path & "\GUI\Player Chat.png")
    If FileExist("GUI\FlashTheatre.png") Then frmFlash.Picture = LoadPicture(App.Path & "\GUI\FlashTheatre.png")
    
    ' Make sure we set that we aren't in the game
    InGame = False
    GettingMap = True
    InEditor = False
    InItemsEditor = False
    InNpcEditor = False
    InShopEditor = False
    InEmoticonEditor = False
    
    frmMirage.picItems.Picture = LoadPicture(App.Path & "\GFX\items.bmp")
    frmSpriteChange.picSprites.Picture = LoadPicture(App.Path & "\GFX\sprites.bmp")
    
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
                    Call SendAddChar(frmNewChar.txtName, 0, frmNewChar.cmbClass.ListIndex, frmChars.lstChars.ListIndex + 1)
                Else
                    Call SendAddChar(frmNewChar.txtName, 1, frmNewChar.cmbClass.ListIndex, frmChars.lstChars.ListIndex + 1)
                End If
            End If
        
        Case MENU_STATE_DELCHAR
            frmChars.Visible = False
            If ConnectToServer = True Then
                Call SetStatus("Connected, sending character deletion request...")
                Call SendDelChar(frmChars.lstChars.ListIndex + 1)
            End If
            
        Case MENU_STATE_USECHAR
            frmChars.Visible = False
            If ConnectToServer = True Then
                Call SetStatus("Connected, sending char data...")
                Call SendUseChar(frmChars.lstChars.ListIndex + 1)
            End If
    End Select

    If Not IsConnected And Connucted = True Then
        frmMainMenu.Visible = True
        frmSendGetData.Visible = False
        Call MsgBox("Sorry, the server seems to be down.  Please try to reconnect in a few minutes or visit " & WEBSITE, vbOKOnly, GAME_NAME)
    End If
End Sub
Sub GameInit()
    frmMirage.Visible = True
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
    frmMirage.picScreen.SetFocus
    
    ' Set font
    Call SetFont("Fixedsys", 18)
                
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
        
        If GetPlayerPOINTS(MyIndex) > 0 Then
            frmMirage.cmdSTR.Visible = True
            frmMirage.cmdDEF.Visible = True
            frmMirage.cmdSpeed.Visible = True
            frmMirage.cmdEnergy.Visible = True
        Else
            frmMirage.cmdSTR.Visible = False
            frmMirage.cmdDEF.Visible = False
            frmMirage.cmdSpeed.Visible = False
            frmMirage.cmdEnergy.Visible = False
        End If
        
    If GettingMap = False Then
        ' Visual Inventory
        Dim Q As Long
        Dim Qq As Long
        Dim IT As Long
               
        If GetTickCount > IT + 500 And frmMirage.picInv3.Visible = True Then
            For Q = 0 To MAX_INV - 1
                Qq = Player(MyIndex).Inv(Q + 1).Num
               
                If frmMirage.picInv(Q).Picture <> LoadPicture() Then
                    frmMirage.picInv(Q).Picture = LoadPicture()
                Else
                    If Qq = 0 Then
                        frmMirage.picInv(Q).Picture = LoadPicture()
                    Else
                        Call BitBlt(frmMirage.picInv(Q).hDC, 0, 0, PIC_X, PIC_Y, frmMirage.picItems.hDC, (Item(Qq).Pic - Int(Item(Qq).Pic / 6) * 6) * PIC_X, Int(Item(Qq).Pic / 6) * PIC_Y, SRCCOPY)
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
        ' Blit out players hp/mp bars
        'Call BltPlayerBar
        If ReadINI("CONFIG", "NpcBar", App.Path & "\config.ini") = 1 Then
            ' Blit out NPC hp bars
            Dim GSD2 As Long
            If GetTickCount > GSD2 + 1000 Then
                Call SendData("getnpchp" & SEP_CHAR & END_CHAR)
                GSD2 = GetTickCount
            End If
            
            For i = 1 To MAX_MAP_NPCS
                Call BltNpcBars(i)
            Next i
        End If
                       
        ' Blit players bar
        If ReadINI("CONFIG", "PlayerBar", App.Path & "\config.ini") = 1 Then
           Dim GSD As Long
           If GetTickCount > GSD + 500 Then
               Call SendData("playershpmp" & SEP_CHAR & END_CHAR)
               GSD = GetTickCount
           End If
    
           For i = 1 To MAX_PLAYERS
               If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                   Call BltPlayerBars(i)
               End If
           Next i
        End If
        
        ' Blit out the sprite change attribute
        For Y = 0 To MAX_MAPY
            For X = 0 To MAX_MAPX
                Call BltSpriteChange(X, Y)
            Next X
        Next Y

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
      
    If ScreenMode = 0 Then
        ' Blit out the npcs
        For i = 1 To MAX_MAP_NPCS
            If Map.Tile(MapNpc(i).X, MapNpc(i).Y).Fringe < 1 Then
                If Map.Tile(MapNpc(i).X, MapNpc(i).Y).FAnim < 1 Then
                    If Map.Tile(MapNpc(i).X, MapNpc(i).Y).Fringe2 < 1 Then
                        If Map.Tile(MapNpc(i).X, MapNpc(i).Y).F2Anim < 1 Then
                            Call BltNpcTop(i)
                        End If
                    End If
                End If
            End If
        Next i
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
        If ReadINI("CONFIG", "NPCDamage", App.Path & "\config.ini") = 1 Then
            If ReadINI("CONFIG", "PlayerName", App.Path & "\config.ini") = 0 Then
                If GetTickCount < NPCDmgTime + 2000 Then
                    Call DrawText(TexthDC, (Int(Len(NPCDmgDamage)) / 2) * 3 + NewX + sx, NewY - 22 - ii + sx, NPCDmgDamage, QBColor(BrightRed))
                End If
            Else
                If GetPlayerGuild(MyIndex) <> "" Then
                    If GetTickCount < NPCDmgTime + 2000 Then
                        Call DrawText(TexthDC, (Int(Len(NPCDmgDamage)) / 2) * 3 + NewX + sx, NewY - 42 - ii + sx, NPCDmgDamage, QBColor(BrightRed))
                    End If
                Else
                    If GetTickCount < NPCDmgTime + 2000 Then
                        Call DrawText(TexthDC, (Int(Len(NPCDmgDamage)) / 2) * 3 + NewX + sx, NewY - 22 - ii + sx, NPCDmgDamage, QBColor(BrightRed))
                    End If
                End If
            End If
            ii = ii + 1
        End If
        
        If ReadINI("CONFIG", "PlayerDamage", App.Path & "\config.ini") = 1 Then
            If NPCWho > 0 Then
                If MapNpc(NPCWho).Num > 0 Then
                    If ReadINI("CONFIG", "NPCName", App.Path & "\config.ini") = 0 Then
                        If Npc(MapNpc(NPCWho).Num).Big = 0 Then
                            If GetTickCount < DmgTime + 2000 Then
                                Call DrawText(TexthDC, (MapNpc(NPCWho).X - NewPlayerX) * PIC_X + sx + (Int(Len(DmgDamage)) / 2) * 3 + MapNpc(NPCWho).XOffset - NewXOffset, (MapNpc(NPCWho).Y - NewPlayerY) * PIC_Y + sx - 20 + MapNpc(NPCWho).YOffset - NewYOffset - iii, DmgDamage, QBColor(White))
                            End If
                        Else
                            If GetTickCount < DmgTime + 2000 Then
                                Call DrawText(TexthDC, (MapNpc(NPCWho).X - NewPlayerX) * PIC_X + sx + (Int(Len(DmgDamage)) / 2) * 3 + MapNpc(NPCWho).XOffset - NewXOffset, (MapNpc(NPCWho).Y - NewPlayerY) * PIC_Y + sx - 47 + MapNpc(NPCWho).YOffset - NewYOffset - iii, DmgDamage, QBColor(White))
                            End If
                        End If
                    Else
                        If Npc(MapNpc(NPCWho).Num).Big = 0 Then
                            If GetTickCount < DmgTime + 2000 Then
                                Call DrawText(TexthDC, (MapNpc(NPCWho).X - NewPlayerX) * PIC_X + sx + (Int(Len(DmgDamage)) / 2) * 3 + MapNpc(NPCWho).XOffset - NewXOffset, (MapNpc(NPCWho).Y - NewPlayerY) * PIC_Y + sx - 30 + MapNpc(NPCWho).YOffset - NewYOffset - iii, DmgDamage, QBColor(White))
                            End If
                        Else
                            If GetTickCount < DmgTime + 2000 Then
                                Call DrawText(TexthDC, (MapNpc(NPCWho).X - NewPlayerX) * PIC_X + sx + (Int(Len(DmgDamage)) / 2) * 3 + MapNpc(NPCWho).XOffset - NewXOffset, (MapNpc(NPCWho).Y - NewPlayerY) * PIC_Y + sx - 57 + MapNpc(NPCWho).YOffset - NewYOffset - iii, DmgDamage, QBColor(White))
                            End If
                        End If
                    End If
                    iii = iii + 1
                End If
            End If
        End If
        
        If ReadINI("CONFIG", "PlayerName", App.Path & "\config.ini") = 1 Then
            For i = 1 To MAX_PLAYERS
                If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                    Call BltPlayerGuildName(i)
                    Call BltPlayerName(i)
                End If
            Next i
        End If
 
        ' speech bubble stuffs
        If ReadINI("CONFIG", "SpeechBubbles", App.Path & "\config.ini") = 1 Then
            For i = 1 To MAX_PLAYERS
                If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                    If Bubble(i).Text <> "" Then
                        Call BltPlayerText(i)
                    End If
    
                    If GetTickCount() > Bubble(i).Created + DISPLAY_BUBBLE_TIME Then
                        Bubble(i).Text = ""
                    End If
                End If
            Next i
        End If

        'Draw NPC Names
        If ReadINI("CONFIG", "NPCName", App.Path & "\config.ini") = 1 Then
            For i = LBound(MapNpc) To UBound(MapNpc)
                If MapNpc(i).Num > 0 Then
                    Call BltMapNPCName(i)
                End If
            Next i
        End If
                
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
                    End With
                Next X
            Next Y
        End If

        ' Blit the text they are putting in
        frmMirage.txtMyTextBox.Text = MyText
        
        If Len(MyText) > 4 Then
            frmMirage.txtMyTextBox.SelStart = Len(frmMirage.txtMyTextBox.Text) + 1
        End If
                
        ' Draw map name
        If Map.Moral = MAP_MORAL_NONE Then
            Call DrawText(TexthDC, Int((20.5) * PIC_X / 2) - (Int(Len(Trim(Map.Name)) / 2) * 8) + sx, 2 + sx, Trim(Map.Name), QBColor(BrightRed))
        ElseIf Map.Moral = MAP_MORAL_SAFE Then
            Call DrawText(TexthDC, Int((20.5) * PIC_X / 2) - (Int(Len(Trim(Map.Name)) / 2) * 8) + sx, 2 + sx, Trim(Map.Name), QBColor(White))
        ElseIf Map.Moral = MAP_MORAL_NO_PENALTY Then
            Call DrawText(TexthDC, Int((20.5) * PIC_X / 2) - (Int(Len(Trim(Map.Name)) / 2) * 8) + sx, 2 + sx, Trim(Map.Name), QBColor(Black))
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
        Call DX.GetWindowRect(frmMirage.picScreen.hWnd, rec_pos)
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
                
        ' Lock fps
        Do While GetTickCount < Tick + 35
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
    
    frmMirage.Visible = False
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
        rec.Top = Npc(MapNpc(MapNpcNum).Num).Sprite * PIC_Y
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
                Color = QBColor(Brown)
            Case 1
                Color = QBColor(DarkGrey)
            Case 2
                Color = QBColor(Cyan)
            Case 3
                Color = QBColor(Blue)
            Case 4
                Color = QBColor(Pink)
        End Select
    Else
        Color = QBColor(BrightRed)
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

Sub BltPlayerGuildName(ByVal Index As Long)
Dim TextX As Long
Dim TextY As Long
Dim Color As Long

    ' Check access level
    If GetPlayerPK(Index) = NO Then
        Select Case GetPlayerGuildAccess(Index)
            Case 0
                If GetPlayerSTR(Index) > 0 Then
                    Color = QBColor(Red)
                Else
                    Color = QBColor(Red)
                End If
            Case 1
                Color = QBColor(BrightCyan)
            Case 2
                Color = QBColor(Yellow)
            Case 3
                Color = QBColor(BrightGreen)
            Case 4
                Color = QBColor(Yellow)
        End Select
    Else
        Color = QBColor(BrightRed)
    End If

If Index = MyIndex Then
    TextX = NewX + sx + Int(PIC_X / 2) - ((Len(GetPlayerGuild(MyIndex)) / 2) * 8)
    TextY = NewY + sx - Int(PIC_Y / 4) - 20
    
    Call DrawText(TexthDC, TextX, TextY, GetPlayerGuild(MyIndex), Color)
Else
    ' Draw name
    TextX = GetPlayerX(Index) * PIC_X + sx + Player(Index).XOffset + Int(PIC_X / 2) - ((Len(GetPlayerGuild(Index)) / 2) * 8)
    TextY = GetPlayerY(Index) * PIC_Y + sx + Player(Index).YOffset - Int(PIC_Y / 2) - 12
    Call DrawText(TexthDC, TextX - (NewPlayerX * PIC_X) - NewXOffset, TextY - (NewPlayerY * PIC_Y) - NewYOffset, GetPlayerGuild(Index), Color)
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
                If Trim(Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).String1) <> "" Then
                    Call AddText(Trim(Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).String1), Grey)
                End If
                If Trim(Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).String2) <> "" Then
                    Call AddText(Trim(Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).String2), Grey)
                End If
                If Trim(Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).String3) <> "" Then
                    Call AddText(Trim(Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).String3), Grey)
                End If
            Exit Sub
            End If
        End If
        ' Broadcast message
        If Mid(MyText, 1, 1) = "'" Then
            ChatText = Mid(MyText, 2, Len(MyText) - 1)
            If Len(Trim(ChatText)) > 0 Then
                Call BroadcastMsg(ChatText)
            End If
            MyText = ""
            Exit Sub
        End If
        
        ' Emote message
        If Mid(MyText, 1, 1) = "-" Then
            ChatText = Mid(MyText, 2, Len(MyText) - 1)
            If Len(Trim(ChatText)) > 0 Then
                Call EmoteMsg(ChatText)
            End If
            MyText = ""
            Exit Sub
        End If
        
        ' Player message
        If Mid(MyText, 1, 1) = "!" Then
            ChatText = Mid(MyText, 2, Len(MyText) - 1)
            Name = ""
                    
            ' Get the desired player from the user text
            For i = 1 To Len(ChatText)
                If Mid(ChatText, i, 1) <> " " Then
                    Name = Name & Mid(ChatText, i, 1)
                Else
                    Exit For
                End If
            Next i
                    
            ' Make sure they are actually sending something
            If Len(ChatText) - i > 0 Then
                ChatText = Mid(ChatText, i + 1, Len(ChatText) - i)
                    
                ' Send the message to the player
                Call PlayerMsg(ChatText, Name)
            Else
                Call AddText("Usage: !playername msghere", AlertColor)
            End If
            MyText = ""
            Exit Sub
        End If
            
        ' // Commands //
        ' Help
        If LCase(Mid(MyText, 1, 5)) = "/help" Then
            Call AddText("Social Commands:", HelpColor)
            Call AddText("'msghere = Broadcast Message", HelpColor)
            Call AddText("-msghere = Emote Message", HelpColor)
            Call AddText("!namehere msghere = Player Message", HelpColor)
            Call AddText("Available Commands: /help, /info, /who, /fps, /inv, /stats, /train, /party, /join, /leave, /refresh", HelpColor)
            MyText = ""
            Exit Sub
        End If
        
        ' Call upon admins
        If LCase(Mid(MyText, 1, 11)) = "/calladmins" Then
            Call SendData("calladmins" & SEP_CHAR & END_CHAR)
            MyText = ""
            Exit Sub
        End If
        
        ' Verification User
        If LCase(Mid(MyText, 1, 5)) = "/info" Then
            ChatText = Mid(MyText, 6, Len(MyText) - 5)
            Call SendData("playerinforequest" & SEP_CHAR & ChatText & SEP_CHAR & END_CHAR)
            MyText = ""
            Exit Sub
        End If
        
        ' Whos Online
        If LCase(Mid(MyText, 1, 4)) = "/who" Then
            Call SendWhosOnline
            MyText = ""
            Exit Sub
        End If
                        
        ' Checking fps
        If LCase(Mid(MyText, 1, 4)) = "/fps" Then
            Call AddText("FPS: " & GameFPS, Pink)
            MyText = ""
            Exit Sub
        End If
                
        ' Show inventory
        If LCase(Mid(MyText, 1, 4)) = "/inv" Then
            Call UpdateInventory
            frmMirage.picInv3.Visible = True
            MyText = ""
            Exit Sub
        End If
        
        ' Request stats
        If LCase(Mid(MyText, 1, 6)) = "/stats" Then
            Call SendData("getstats" & SEP_CHAR & END_CHAR)
            MyText = ""
            Exit Sub
        End If
    
        ' Show training
        If LCase(Mid(MyText, 1, 6)) = "/train" Then
            frmTraining.Show vbModal
            MyText = ""
            Exit Sub
        End If
     
        ' Refresh Player
        If LCase(Mid(MyText, 1, 8)) = "/refresh" Then
            Call SendData("refresh" & SEP_CHAR & END_CHAR)
            MyText = ""
            Exit Sub
        End If
        
        ' Decline Chat
        If LCase(Mid(MyText, 1, 12)) = "/chatdecline" Then
            Call SendData("dchat" & SEP_CHAR & END_CHAR)
            MyText = ""
            Exit Sub
        End If
        
        ' Accept Chat
        If LCase(Mid(MyText, 1, 5)) = "/chat" Then
            Call SendData("achat" & SEP_CHAR & END_CHAR)
            MyText = ""
            Exit Sub
        End If
        
        If LCase(Mid(MyText, 1, 6)) = "/trade" Then
            ' Make sure they are actually sending something
            If Len(MyText) > 7 Then
                ChatText = Mid(MyText, 8, Len(MyText) - 7)
                Call SendTradeRequest(ChatText)
            Else
                Call AddText("Usage: /trade playernamehere", AlertColor)
            End If
            MyText = ""
            Exit Sub
        End If
        
        ' Accept Trade
        If LCase(Mid(MyText, 1, 7)) = "/accept" Then
            Call SendAcceptTrade
            MyText = ""
            Exit Sub
        End If
        
        ' Decline Trade
        If LCase(Mid(MyText, 1, 8)) = "/decline" Then
            Call SendDeclineTrade
            MyText = ""
            Exit Sub
        End If
        
        ' Party request
        If LCase(Mid(MyText, 1, 6)) = "/party" Then
            ' Make sure they are actually sending something
            If Len(MyText) > 7 Then
                ChatText = Mid(MyText, 8, Len(MyText) - 7)
                Call SendPartyRequest(ChatText)
            Else
                Call AddText("Usage: /party playernamehere", AlertColor)
            End If
            MyText = ""
            Exit Sub
        End If
        
        ' Join party
        If LCase(Mid(MyText, 1, 5)) = "/join" Then
            Call SendJoinParty
            MyText = ""
            Exit Sub
        End If
        
        ' Leave party
        If LCase(Mid(MyText, 1, 6)) = "/leave" Then
            Call SendLeaveParty
            MyText = ""
            Exit Sub
        End If
        
        ' // Moniter Admin Commands //
        If GetPlayerAccess(MyIndex) > 0 Then
            ' Admin Help
            If LCase(Mid(MyText, 1, 6)) = "/admin" Then
                Call AddText("Social Commands:", HelpColor)
                Call AddText("""msghere = Global Admin Message", HelpColor)
                Call AddText("=msghere = Private Admin Message", HelpColor)
                Call AddText("Available Commands: /admin, /loc, /mapeditor, /warpmeto, /warptome, /warpto, /setsprite, /mapreport, /kick, /ban, /edititem, /respawn, /editnpc, /motd, /editshop, /ban, /editspell", HelpColor)
                MyText = ""
                Exit Sub
            End If
            
            ' Kicking a player
            If LCase(Mid(MyText, 1, 5)) = "/kick" Then
                If Len(MyText) > 6 Then
                    MyText = Mid(MyText, 7, Len(MyText) - 6)
                    Call SendKick(MyText)
                End If
                MyText = ""
                Exit Sub
            End If
        
            ' Global Message
            If Mid(MyText, 1, 1) = """" Then
                ChatText = Mid(MyText, 2, Len(MyText) - 1)
                If Len(Trim(ChatText)) > 0 Then
                    Call GlobalMsg(ChatText)
                End If
                MyText = ""
                Exit Sub
            End If
        
            ' Admin Message
            If Mid(MyText, 1, 1) = "=" Then
                ChatText = Mid(MyText, 2, Len(MyText) - 1)
                If Len(Trim(ChatText)) > 0 Then
                    Call AdminMsg(ChatText)
                End If
                MyText = ""
                Exit Sub
            End If
        End If
        
        ' // Mapper Admin Commands //
        If GetPlayerAccess(MyIndex) >= ADMIN_MAPPER Then
            ' Location
            If LCase(Mid(MyText, 1, 4)) = "/loc" Then
                Call SendRequestLocation
                MyText = ""
                Exit Sub
            End If
            
            ' Map Editor
            If LCase(Mid(MyText, 1, 10)) = "/mapeditor" Then
                Call SendRequestEditMap
                MyText = ""
                Exit Sub
            End If
            
            ' Map report
            If LCase(Mid(MyText, 1, 10)) = "/mapreport" Then
                Call SendData("mapreport" & SEP_CHAR & END_CHAR)
                MyText = ""
                Exit Sub
            End If
            
            ' Warping to a player
            If LCase(Mid(MyText, 1, 9)) = "/warpmeto" Then
                If Len(MyText) > 10 Then
                    MyText = Mid(MyText, 10, Len(MyText) - 9)
                    Call WarpMeTo(MyText)
                End If
                MyText = ""
                Exit Sub
            End If
                        
            ' Warping a player to you
            If LCase(Mid(MyText, 1, 9)) = "/warptome" Then
                If Len(MyText) > 10 Then
                    MyText = Mid(MyText, 10, Len(MyText) - 9)
                    Call WarpToMe(MyText)
                End If
                MyText = ""
                Exit Sub
            End If
                        
            ' Warping to a map
            If LCase(Mid(MyText, 1, 7)) = "/warpto" Then
                If Len(MyText) > 8 Then
                    MyText = Mid(MyText, 8, Len(MyText) - 7)
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
            If LCase(Mid(MyText, 1, 10)) = "/setsprite" Then
                If Len(MyText) > 11 Then
                    ' Get sprite #
                    MyText = Mid(MyText, 12, Len(MyText) - 11)
                
                    Call SendSetSprite(Val(MyText))
                End If
                MyText = ""
                Exit Sub
            End If
            
            ' Setting player sprite
            If LCase(Mid(MyText, 1, 16)) = "/setplayersprite" Then
                If Len(MyText) > 19 Then
                    i = Val(Mid(MyText, 17, 1))
                
                    MyText = Mid(MyText, 18, Len(MyText) - 17)
                    Call SendSetPlayerSprite(i, Val(MyText))
                End If
                MyText = ""
                Exit Sub
            End If
            
            ' Map report
            'If LCase(Mid(MyText, 1, 10)) = "/mapreport" Then
                'Call SendData("mapreport" & SEP_CHAR & END_CHAR)
                'MyText = ""
                'Exit Sub
            'End If
        
            ' Respawn request
            If Mid(MyText, 1, 8) = "/respawn" Then
                Call SendMapRespawn
                MyText = ""
                Exit Sub
            End If
        
            ' MOTD change
            If Mid(MyText, 1, 5) = "/motd" Then
                If Len(MyText) > 6 Then
                    MyText = Mid(MyText, 7, Len(MyText) - 6)
                    If Trim(MyText) <> "" Then
                        Call SendMOTDChange(MyText)
                    End If
                End If
                MyText = ""
                Exit Sub
            End If
            
            ' Check the ban list
            If Mid(MyText, 1, 3) = "/banlist" Then
                Call SendBanList
                MyText = ""
                Exit Sub
            End If
            
            ' Banning a player
            If LCase(Mid(MyText, 1, 4)) = "/ban" Then
                If Len(MyText) > 5 Then
                    MyText = Mid(MyText, 6, Len(MyText) - 5)
                    Call SendBan(MyText)
                    MyText = ""
                End If
                Exit Sub
            End If
        End If
            
        ' // Developer Admin Commands //
        If GetPlayerAccess(MyIndex) >= ADMIN_DEVELOPER Then
            ' Editing item request
            If Mid(MyText, 1, 9) = "/edititem" Then
                Call SendRequestEditItem
                MyText = ""
                Exit Sub
            End If
            
            ' Editing emoticon request
            If Mid(MyText, 1, 13) = "/editemoticon" Then
                Call SendRequestEditEmoticon
                MyText = ""
                Exit Sub
            End If
            
            ' Editing npc request
            If Mid(MyText, 1, 8) = "/editnpc" Then
                Call SendRequestEditNpc
                MyText = ""
                Exit Sub
            End If
            
            ' Editing shop request
            If Mid(MyText, 1, 9) = "/editshop" Then
                Call SendRequestEditShop
                MyText = ""
                Exit Sub
            End If
        
            ' Editing spell request
            If LCase(Trim(MyText)) = "/editspell" Then
            'If Mid(MyText, 1, 10) = "/editspell" Then
                Call SendRequestEditSpell
                MyText = ""
                Exit Sub
            End If
        End If
        
        ' // Creator Admin Commands //
        If GetPlayerAccess(MyIndex) >= ADMIN_CREATOR Then
            ' Giving another player access
            If LCase(Mid(MyText, 1, 10)) = "/setaccess" Then
                ' Get access #
                i = Val(Mid(MyText, 12, 1))
                
                MyText = Mid(MyText, 14, Len(MyText) - 13)
                
                Call SendSetAccess(MyText, i)
                MyText = ""
                Exit Sub
            End If
            
            ' Ban destroy
            If LCase(Mid(MyText, 1, 15)) = "/destroybanlist" Then
                Call SendBanDestroy
                MyText = ""
                Exit Sub
            End If
        End If
        
        ' Tell them its not a valid command
        If Left$(Trim(MyText), 1) = "/" Then
            For i = 0 To MAX_EMOTICONS
                If Trim(Emoticons(i).Command) = Trim(MyText) And Trim(Emoticons(i).Command) <> "/" Then
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
        If Len(Trim(MyText)) > 0 Then
            Call SayMsg(MyText)
        End If
        MyText = ""
        Exit Sub
    End If
    
    ' Handle when the user presses the backspace key
    If (KeyAscii = vbKeyBack) Then
        If Len(MyText) > 0 Then
            MyText = Mid(MyText, 1, Len(MyText) - 1)
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
    If GetTickCount > Player(MyIndex).MapGetTimer + 250 And Trim(MyText) = "" Then
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
    If GettingMap = False Then
        If IsTryingToMove Then
            If CanMove Then
                ' Check if player has the shift key down for running
                If ShiftDown Then
                    Player(MyIndex).Moving = MOVING_RUNNING
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
            If Len(GetPlayerName(i)) >= Len(Trim(Name)) Then
                If UCase(Mid(GetPlayerName(i), 1, Len(Trim(Name)))) = UCase(Trim(Name)) Then
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
    frmMapEditor.Show vbModeless, frmMirage
    'frmMirage.picMapEditor.Visible = True
    'With frmMapEditor.picBackSelect
        '.Width = 14 * PIC_X
        '.Height = 16384
        '.Picture = LoadPicture(App.Path + "\GFX\tiles.bmp")
    'End With
    frmMapEditor.picBackSelect.Picture = LoadPicture(App.Path + "\GFX\tiles.bmp")
    frmMapEditor.MouseSelected.Picture = LoadPicture(App.Path + "\GFX\tiles.bmp")
End Sub

Public Sub EditorMouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim x1, y1 As Long

    If InEditor Then
        x1 = Int(X / PIC_X)
        y1 = Int(Y / PIC_Y)
        If (Button = 1) And (x1 >= 0) And (x1 <= MAX_MAPX) And (y1 >= 0) And (y1 <= MAX_MAPY) Then
            If frmMapEditor.optLayers.Value = True Then
                With Map.Tile(x1, y1)
                    If frmMapEditor.optGround.Value = True Then .Ground = EditorTileY * 14 + EditorTileX
                        If frmMapEditor.optMask.Value = True Then .Mask = EditorTileY * 14 + EditorTileX
                        If frmMapEditor.optAnim.Value = True Then .Anim = EditorTileY * 14 + EditorTileX
                        If frmMapEditor.optMask2.Value = True Then .Mask2 = EditorTileY * 14 + EditorTileX
                        If frmMapEditor.optM2Anim.Value = True Then .M2Anim = EditorTileY * 14 + EditorTileX
                        If frmMapEditor.optFringe.Value = True Then .Fringe = EditorTileY * 14 + EditorTileX
                        If frmMapEditor.optFAnim.Value = True Then .FAnim = EditorTileY * 14 + EditorTileX
                        If frmMapEditor.optFringe2.Value = True Then .Fringe2 = EditorTileY * 14 + EditorTileX
                        If frmMapEditor.optF2Anim.Value = True Then .F2Anim = EditorTileY * 14 + EditorTileX
                End With
            Else
                With Map.Tile(x1, y1)
                    If frmMapEditor.optBlocked.Value = True Then .Type = TILE_TYPE_BLOCKED
                    If frmMapEditor.optWarp.Value = True Then
                        .Type = TILE_TYPE_WARP
                        .Data1 = EditorWarpMap
                        .Data2 = EditorWarpX
                        .Data3 = EditorWarpY
                        .String1 = ""
                        .String2 = ""
                        .String3 = ""
                    End If

                    If frmMapEditor.optHeal.Value = True Then
                        .Type = TILE_TYPE_HEAL
                        .Data1 = 0
                        .Data2 = 0
                        .Data3 = 0
                        .String1 = ""
                        .String2 = ""
                        .String3 = ""
                    End If

                    If frmMapEditor.optKill.Value = True Then
                        .Type = TILE_TYPE_KILL
                        .Data1 = 0
                        .Data2 = 0
                        .Data3 = 0
                        .String1 = ""
                        .String2 = ""
                        .String3 = ""
                    End If

                    If frmMapEditor.optItem.Value = True Then
                        .Type = TILE_TYPE_ITEM
                        .Data1 = ItemEditorNum
                        .Data2 = ItemEditorValue
                        .Data3 = 0
                        .String1 = ""
                        .String2 = ""
                        .String3 = ""
                    End If
                    If frmMapEditor.optNpcAvoid.Value = True Then
                        .Type = TILE_TYPE_NPCAVOID
                        .Data1 = 0
                        .Data2 = 0
                        .Data3 = 0
                        .String1 = ""
                        .String2 = ""
                        .String3 = ""
                    End If
                    If frmMapEditor.optKey.Value = True Then
                        .Type = TILE_TYPE_KEY
                        .Data1 = KeyEditorNum
                        .Data2 = KeyEditorTake
                        .Data3 = 0
                        .String1 = ""
                        .String2 = ""
                        .String3 = ""
                    End If
                    If frmMapEditor.optKeyOpen.Value = True Then
                        .Type = TILE_TYPE_KEYOPEN
                        .Data1 = KeyOpenEditorX
                        .Data2 = KeyOpenEditorY
                        .Data3 = 0
                        .String1 = KeyOpenEditorMsg
                        .String2 = ""
                        .String3 = ""
                    End If
                    If frmMapEditor.optShop.Value = True Then
                        .Type = TILE_TYPE_SHOP
                        .Data1 = EditorShopNum
                        .Data2 = 0
                        .Data3 = 0
                        .String1 = ""
                        .String2 = ""
                        .String3 = ""
                    End If
                    If frmMapEditor.optCBlock.Value = True Then
                        .Type = TILE_TYPE_CBLOCK
                        .Data1 = EditorItemNum1
                        .Data2 = EditorItemNum2
                        .Data3 = EditorItemNum3
                        .String1 = ""
                        .String2 = ""
                        .String3 = ""
                    End If
                    If frmMapEditor.optArena.Value = True Then
                        .Type = TILE_TYPE_ARENA
                        .Data1 = Arena1
                        .Data2 = Arena2
                        .Data3 = Arena3
                        .String1 = ""
                        .String2 = ""
                        .String3 = ""
                    End If
                    If frmMapEditor.optSound.Value = True Then
                        .Type = TILE_TYPE_SOUND
                        .Data1 = 0
                        .Data2 = 0
                        .Data3 = 0
                        .String1 = SoundFileName
                        .String2 = ""
                        .String3 = ""
                    End If
                    If frmMapEditor.optSprite.Value = True Then
                        .Type = TILE_TYPE_SPRITE_CHANGE
                        .Data1 = SpritePic
                        .Data2 = SpriteItem
                        .Data3 = SpritePrice
                        .String1 = ""
                        .String2 = ""
                        .String3 = ""
                    End If
                    If frmMapEditor.optSign.Value = True Then
                        .Type = TILE_TYPE_SIGN
                        .Data1 = 0
                        .Data2 = 0
                        .Data3 = 0
                        .String1 = SignLine1
                        .String2 = SignLine2
                        .String3 = SignLine3
                    End If
                    If frmMapEditor.optDoor.Value = True Then
                        .Type = TILE_TYPE_DOOR
                        .Data1 = 0
                        .Data2 = 0
                        .Data3 = 0
                        .String1 = ""
                        .String2 = ""
                        .String3 = ""
                    End If
                    If frmMapEditor.optNotice.Value = True Then
                        .Type = TILE_TYPE_NOTICE
                        .Data1 = 0
                        .Data2 = 0
                        .Data3 = 0
                        .String1 = NoticeTitle
                        .String2 = NoticeText
                        .String3 = NoticeSound
                    End If
                    If frmMapEditor.optChest.Value = True Then
                        .Type = TILE_TYPE_CHEST
                        .Data1 = 0
                        .Data2 = 0
                        .Data3 = 0
                        .String1 = ""
                        .String2 = ""
                        .String3 = ""
                    End If
                    If frmMapEditor.optClassChange.Value = True Then
                        .Type = TILE_TYPE_CLASS_CHANGE
                        .Data1 = ClassChange
                        .Data2 = ClassChangeReq
                        .Data3 = 0
                        .String1 = ""
                        .String2 = ""
                        .String3 = ""
                    End If
                    If frmMapEditor.optScripted.Value = True Then
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
        End If
        
        If (Button = 2) And (x1 >= 0) And (x1 <= MAX_MAPX) And (y1 >= 0) And (y1 <= MAX_MAPY) Then
            If frmMapEditor.optLayers.Value = True Then
                With Map.Tile(x1, y1)
                    If frmMapEditor.optGround.Value = True Then .Ground = 0
                    If frmMapEditor.optMask.Value = True Then .Mask = 0
                    If frmMapEditor.optAnim.Value = True Then .Anim = 0
                    If frmMapEditor.optMask2.Value = True Then .Mask2 = 0
                    If frmMapEditor.optM2Anim.Value = True Then .M2Anim = 0
                    If frmMapEditor.optFringe.Value = True Then .Fringe = 0
                    If frmMapEditor.optFAnim.Value = True Then .FAnim = 0
                    If frmMapEditor.optFringe2.Value = True Then .Fringe2 = 0
                    If frmMapEditor.optF2Anim.Value = True Then .F2Anim = 0
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
    frmMapEditor.MouseSelected.Top = -(EditorTileY * PIC_Y)
    frmMapEditor.MouseSelected.Left = -(EditorTileX * PIC_Y)
    'Call BitBlt(frmMapEditor.picSelect.hDC, 0, 0, PIC_X, PIC_Y, frmMapEditor.picBackSelect.hDC, EditorTileX * PIC_X, EditorTileY * PIC_Y, SRCCOPY)
End Sub

Public Sub EditorTileScroll()
    frmMapEditor.picBackSelect.Top = (frmMapEditor.scrlPicture.Value * PIC_Y) * -1
End Sub

Public Sub EditorSend()
    Call SendMap
    Call EditorCancel
End Sub

Public Sub EditorCancel()
    Map = SaveMap
    InEditor = False
    frmMapEditor.Visible = False
    frmMirage.Show
    'frmMirage.picMapEditor.Visible = False
End Sub

Public Sub EditorClearLayer()
Dim YesNo As Long, X As Long, Y As Long

    ' Ground layer
    If frmMapEditor.optGround.Value = True Then
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
    If frmMapEditor.optMask.Value = True Then
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
    If frmMapEditor.optAnim.Value = True Then
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
    If frmMapEditor.optMask2.Value = True Then
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
    If frmMapEditor.optM2Anim.Value = True Then
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
    If frmMapEditor.optFringe.Value = True Then
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
    If frmMapEditor.optFAnim.Value = True Then
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
    If frmMapEditor.optFringe2.Value = True Then
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
    If frmMapEditor.optF2Anim.Value = True Then
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
    frmEmoticonEditor.scrlEmoticon.Max = MAX_EMOTICONS
    frmEmoticonEditor.scrlEmoticon.Value = Emoticons(EditorIndex - 1).Pic
    frmEmoticonEditor.txtCommand.Text = Trim(Emoticons(EditorIndex - 1).Command)
    frmEmoticonEditor.picEmoticons.Picture = LoadPicture(App.Path & "\GFX\emoticons.bmp")
    frmEmoticonEditor.Show vbModal
End Sub

Public Sub EmoticonEditorOk()
    Emoticons(EditorIndex - 1).Pic = frmEmoticonEditor.scrlEmoticon.Value
    If frmEmoticonEditor.txtCommand.Text <> "/" Then
        Emoticons(EditorIndex - 1).Command = frmEmoticonEditor.txtCommand.Text
    Else
        Emoticons(EditorIndex - 1).Command = ""
    End If
    
    Call SendSaveEmoticon(EditorIndex - 1)
    Call EmoticonEditorCancel
End Sub

Public Sub EmoticonEditorCancel()
    InEmoticonEditor = False
    Unload frmEmoticonEditor
End Sub

Public Sub ItemEditorInit()
On Error Resume Next

    EditorItemY = Int(Item(EditorIndex).Pic / 6)
    EditorItemX = Int((Item(EditorIndex).Pic / 6) - EditorItemY)
    
    frmItemEditor.scrlClassReq.Max = Max_Classes

    frmItemEditor.picItems.Picture = LoadPicture(App.Path & "\GFX\items.bmp")
    
    frmItemEditor.txtName.Text = Trim(Item(EditorIndex).Name)
    frmItemEditor.txtDesc.Text = Trim(Item(EditorIndex).desc)
    frmItemEditor.cmbType.ListIndex = Item(EditorIndex).Type
    
    If (frmItemEditor.cmbType.ListIndex >= ITEM_TYPE_WEAPON) And (frmItemEditor.cmbType.ListIndex <= ITEM_TYPE_SHIELD) Then
        frmItemEditor.fraEquipment.Visible = True
        frmItemEditor.scrlDurability.Value = Item(EditorIndex).Data1
        frmItemEditor.scrlStrength.Value = Item(EditorIndex).Data2
        frmItemEditor.scrlStrReq.Value = Item(EditorIndex).StrReq
        frmItemEditor.scrlDefReq.Value = Item(EditorIndex).DefReq
        frmItemEditor.scrlSpeedReq.Value = Item(EditorIndex).SpeedReq
        frmItemEditor.scrlClassReq.Value = Item(EditorIndex).ClassReq
        frmItemEditor.scrlAccessReq.Value = Item(EditorIndex).AccessReq
        frmItemEditor.scrlAddHP.Value = Item(EditorIndex).AddHP
        frmItemEditor.scrlAddMP.Value = Item(EditorIndex).AddMP
        frmItemEditor.scrlAddSP.Value = Item(EditorIndex).AddSP
        frmItemEditor.scrlAddStr.Value = Item(EditorIndex).AddStr
        frmItemEditor.scrlAddDef.Value = Item(EditorIndex).AddDef
        frmItemEditor.scrlAddMagi.Value = Item(EditorIndex).AddMagi
        frmItemEditor.scrlAddSpeed.Value = Item(EditorIndex).AddSpeed
        frmItemEditor.scrlAddEXP.Value = Item(EditorIndex).AddEXP
    Else
        frmItemEditor.fraEquipment.Visible = False
    End If
    
    If (frmItemEditor.cmbType.ListIndex >= ITEM_TYPE_POTIONADDHP) And (frmItemEditor.cmbType.ListIndex <= ITEM_TYPE_POTIONSUBSP) Then
        frmItemEditor.fraVitals.Visible = True
        frmItemEditor.scrlVitalMod.Value = Item(EditorIndex).Data1
    Else
        frmItemEditor.fraVitals.Visible = False
    End If
    
    If (frmItemEditor.cmbType.ListIndex = ITEM_TYPE_SPELL) Then
        frmItemEditor.fraSpell.Visible = True
        frmItemEditor.scrlSpell.Value = Item(EditorIndex).Data1
    Else
        frmItemEditor.fraSpell.Visible = False
    End If
    
    frmItemEditor.Show vbModal
End Sub

Public Sub ItemEditorOk()
    Item(EditorIndex).Name = frmItemEditor.txtName.Text
    Item(EditorIndex).desc = frmItemEditor.txtDesc.Text
    Item(EditorIndex).Pic = EditorItemY * 6 + EditorItemX
    Item(EditorIndex).Type = frmItemEditor.cmbType.ListIndex

    If (frmItemEditor.cmbType.ListIndex >= ITEM_TYPE_WEAPON) And (frmItemEditor.cmbType.ListIndex <= ITEM_TYPE_SHIELD) Then
        Item(EditorIndex).Data1 = frmItemEditor.scrlDurability.Value
        Item(EditorIndex).Data2 = frmItemEditor.scrlStrength.Value
        Item(EditorIndex).Data3 = 0
        Item(EditorIndex).StrReq = frmItemEditor.scrlStrReq.Value
        Item(EditorIndex).DefReq = frmItemEditor.scrlDefReq.Value
        Item(EditorIndex).SpeedReq = frmItemEditor.scrlSpeedReq.Value
        
        Item(EditorIndex).ClassReq = frmItemEditor.scrlClassReq.Value
        Item(EditorIndex).AccessReq = frmItemEditor.scrlAccessReq.Value
        
        Item(EditorIndex).AddHP = frmItemEditor.scrlAddHP.Value
        Item(EditorIndex).AddMP = frmItemEditor.scrlAddMP.Value
        Item(EditorIndex).AddSP = frmItemEditor.scrlAddSP.Value
        Item(EditorIndex).AddStr = frmItemEditor.scrlAddStr.Value
        Item(EditorIndex).AddDef = frmItemEditor.scrlAddDef.Value
        Item(EditorIndex).AddMagi = frmItemEditor.scrlAddMagi.Value
        Item(EditorIndex).AddSpeed = frmItemEditor.scrlAddSpeed.Value
        Item(EditorIndex).AddEXP = frmItemEditor.scrlAddEXP.Value
    End If
    
    If (frmItemEditor.cmbType.ListIndex >= ITEM_TYPE_POTIONADDHP) And (frmItemEditor.cmbType.ListIndex <= ITEM_TYPE_POTIONSUBSP) Then
        Item(EditorIndex).Data1 = frmItemEditor.scrlVitalMod.Value
        Item(EditorIndex).Data2 = 0
        Item(EditorIndex).Data3 = 0
        Item(EditorIndex).StrReq = 0
        Item(EditorIndex).DefReq = 0
        Item(EditorIndex).SpeedReq = 0
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
    
    If (frmItemEditor.cmbType.ListIndex = ITEM_TYPE_SPELL) Then
        Item(EditorIndex).Data1 = frmItemEditor.scrlSpell.Value
        Item(EditorIndex).Data2 = 0
        Item(EditorIndex).Data3 = 0
        Item(EditorIndex).StrReq = 0
        Item(EditorIndex).DefReq = 0
        Item(EditorIndex).SpeedReq = 0
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
    Unload frmItemEditor
End Sub

Public Sub ItemEditorCancel()
    InItemsEditor = False
    Unload frmItemEditor
End Sub

Public Sub NpcEditorInit()
On Error Resume Next
    
    frmNpcEditor.picSprites.Picture = LoadPicture(App.Path & "\GFX\sprites.bmp")
    
    frmNpcEditor.txtName.Text = Trim(Npc(EditorIndex).Name)
    frmNpcEditor.txtAttackSay.Text = Trim(Npc(EditorIndex).AttackSay)
    frmNpcEditor.scrlSprite.Value = Npc(EditorIndex).Sprite
    frmNpcEditor.txtSpawnSecs.Text = STR(Npc(EditorIndex).SpawnSecs)
    frmNpcEditor.cmbBehavior.ListIndex = Npc(EditorIndex).Behavior
    frmNpcEditor.scrlRange.Value = Npc(EditorIndex).Range
    frmNpcEditor.scrlSTR.Value = Npc(EditorIndex).STR
    frmNpcEditor.scrlDEF.Value = Npc(EditorIndex).DEF
    frmNpcEditor.scrlSPEED.Value = Npc(EditorIndex).SPEED
    frmNpcEditor.scrlMAGI.Value = Npc(EditorIndex).MAGI
    frmNpcEditor.BigNpc.Value = Npc(EditorIndex).Big
    frmNpcEditor.StartHP.Value = Npc(EditorIndex).MaxHp
    frmNpcEditor.ExpGive.Value = Npc(EditorIndex).EXP
    frmNpcEditor.txtChance.Text = STR(Npc(EditorIndex).ItemNPC(1).Chance)
    frmNpcEditor.scrlNum.Value = Npc(EditorIndex).ItemNPC(1).ItemNum
    frmNpcEditor.scrlValue.Value = Npc(EditorIndex).ItemNPC(1).ItemValue
    
    frmNpcEditor.Show vbModal
End Sub

Public Sub NpcEditorOk()
    Npc(EditorIndex).Name = frmNpcEditor.txtName.Text
    Npc(EditorIndex).AttackSay = frmNpcEditor.txtAttackSay.Text
    Npc(EditorIndex).Sprite = frmNpcEditor.scrlSprite.Value
    Npc(EditorIndex).SpawnSecs = Val(frmNpcEditor.txtSpawnSecs.Text)
    Npc(EditorIndex).Behavior = frmNpcEditor.cmbBehavior.ListIndex
    Npc(EditorIndex).Range = frmNpcEditor.scrlRange.Value
    Npc(EditorIndex).STR = frmNpcEditor.scrlSTR.Value
    Npc(EditorIndex).DEF = frmNpcEditor.scrlDEF.Value
    Npc(EditorIndex).SPEED = frmNpcEditor.scrlSPEED.Value
    Npc(EditorIndex).MAGI = frmNpcEditor.scrlMAGI.Value
    Npc(EditorIndex).Big = frmNpcEditor.BigNpc.Value
    Npc(EditorIndex).MaxHp = frmNpcEditor.StartHP.Value
    Npc(EditorIndex).EXP = frmNpcEditor.ExpGive.Value
    
    Call SendSaveNpc(EditorIndex)
    InNpcEditor = False
    Unload frmNpcEditor
End Sub

Public Sub NpcEditorCancel()
    InNpcEditor = False
    Unload frmNpcEditor
End Sub

Public Sub NpcEditorBltSprite()
    If frmNpcEditor.BigNpc.Value = Checked Then
        Call BitBlt(frmNpcEditor.picSprite.hDC, 0, 0, 64, 64, frmNpcEditor.picSprites.hDC, 3 * 64, frmNpcEditor.scrlSprite.Value * 64, SRCCOPY)
    Else
        Call BitBlt(frmNpcEditor.picSprite.hDC, 0, 0, PIC_X, PIC_Y, frmNpcEditor.picSprites.hDC, 3 * PIC_X, frmNpcEditor.scrlSprite.Value * PIC_Y, SRCCOPY)
    End If
End Sub

Public Sub ShopEditorInit()
On Error Resume Next

Dim i As Long

    frmShopEditor.txtName.Text = Trim(Shop(EditorIndex).Name)
    frmShopEditor.txtJoinSay.Text = Trim(Shop(EditorIndex).JoinSay)
    frmShopEditor.txtLeaveSay.Text = Trim(Shop(EditorIndex).LeaveSay)
    frmShopEditor.chkFixesItems.Value = Shop(EditorIndex).FixesItems
    
    frmShopEditor.cmbItemGive.Clear
    frmShopEditor.cmbItemGive.AddItem "None"
    frmShopEditor.cmbItemGet.Clear
    frmShopEditor.cmbItemGet.AddItem "None"
    For i = 1 To MAX_ITEMS
        frmShopEditor.cmbItemGive.AddItem i & ": " & Trim(Item(i).Name)
        frmShopEditor.cmbItemGet.AddItem i & ": " & Trim(Item(i).Name)
    Next i
    frmShopEditor.cmbItemGive.ListIndex = 0
    frmShopEditor.cmbItemGet.ListIndex = 0
    
    Call UpdateShopTrade
    
    frmShopEditor.Show vbModal
End Sub

Public Sub UpdateShopTrade()
Dim i As Long, GetItem As Long, GetValue As Long, GiveItem As Long, GiveValue As Long
    
    frmShopEditor.lstTradeItem.Clear
    For i = 1 To MAX_TRADES
        GetItem = Shop(EditorIndex).TradeItem(i).GetItem
        GetValue = Shop(EditorIndex).TradeItem(i).GetValue
        GiveItem = Shop(EditorIndex).TradeItem(i).GiveItem
        GiveValue = Shop(EditorIndex).TradeItem(i).GiveValue
        
        If GetItem > 0 And GiveItem > 0 Then
            frmShopEditor.lstTradeItem.AddItem i & ": " & GiveValue & " " & Trim(Item(GiveItem).Name) & " for " & GetValue & " " & Trim(Item(GetItem).Name)
        Else
            frmShopEditor.lstTradeItem.AddItem "Empty Trade Slot"
        End If
    Next i
    frmShopEditor.lstTradeItem.ListIndex = 0
End Sub

Public Sub ShopEditorOk()
    Shop(EditorIndex).Name = frmShopEditor.txtName.Text
    Shop(EditorIndex).JoinSay = frmShopEditor.txtJoinSay.Text
    Shop(EditorIndex).LeaveSay = frmShopEditor.txtLeaveSay.Text
    Shop(EditorIndex).FixesItems = frmShopEditor.chkFixesItems.Value
    
    Call SendSaveShop(EditorIndex)
    InShopEditor = False
    Unload frmShopEditor
End Sub

Public Sub ShopEditorCancel()
    InShopEditor = False
    Unload frmShopEditor
End Sub

Public Sub SpellEditorInit()
On Error Resume Next

Dim i As Long

    frmSpellEditor.cmbClassReq.AddItem "All Classes"
    For i = 0 To Max_Classes
        frmSpellEditor.cmbClassReq.AddItem Trim(Class(i).Name)
    Next i
    
    frmSpellEditor.txtName.Text = Trim(Spell(EditorIndex).Name)
    frmSpellEditor.cmbClassReq.ListIndex = Spell(EditorIndex).ClassReq
    frmSpellEditor.scrlLevelReq.Value = Spell(EditorIndex).LevelReq
        
    frmSpellEditor.cmbType.ListIndex = Spell(EditorIndex).Type
    frmSpellEditor.scrlVitalMod.Value = Spell(EditorIndex).Data1
    
    frmSpellEditor.scrlCost.Value = Spell(EditorIndex).MPCost
    frmSpellEditor.scrlSound.Value = Spell(EditorIndex).Sound
        
    frmSpellEditor.Show vbModal
End Sub

Public Sub SpellEditorOk()
    Spell(EditorIndex).Name = frmSpellEditor.txtName.Text
    Spell(EditorIndex).ClassReq = frmSpellEditor.cmbClassReq.ListIndex
    Spell(EditorIndex).LevelReq = frmSpellEditor.scrlLevelReq.Value
    Spell(EditorIndex).Type = frmSpellEditor.cmbType.ListIndex
    Spell(EditorIndex).Data1 = frmSpellEditor.scrlVitalMod.Value
    Spell(EditorIndex).Data3 = 0
    Spell(EditorIndex).MPCost = frmSpellEditor.scrlCost.Value
    Spell(EditorIndex).Sound = frmSpellEditor.scrlSound.Value
    
    Call SendSaveSpell(EditorIndex)
    InSpellEditor = False
    Unload frmSpellEditor
End Sub

Public Sub SpellEditorCancel()
    InSpellEditor = False
    Unload frmSpellEditor
End Sub

Public Sub UpdateInventory()
Dim i As Long

    frmMirage.lstInv.Clear
    
    ' Show the inventory
    For i = 1 To MAX_INV
        If GetPlayerInvItemNum(MyIndex, i) > 0 And GetPlayerInvItemNum(MyIndex, i) <= MAX_ITEMS Then
            If Item(GetPlayerInvItemNum(MyIndex, i)).Type = ITEM_TYPE_CURRENCY Then
                frmMirage.lstInv.AddItem i & ": " & Trim(Item(GetPlayerInvItemNum(MyIndex, i)).Name) & " (" & GetPlayerInvItemValue(MyIndex, i) & ")"
            Else
                ' Check if this item is being worn
                If GetPlayerWeaponSlot(MyIndex) = i Or GetPlayerArmorSlot(MyIndex) = i Or GetPlayerHelmetSlot(MyIndex) = i Or GetPlayerShieldSlot(MyIndex) = i Then
                    frmMirage.lstInv.AddItem i & ": " & Trim(Item(GetPlayerInvItemNum(MyIndex, i)).Name) & " (worn)"
                Else
                    frmMirage.lstInv.AddItem i & ": " & Trim(Item(GetPlayerInvItemNum(MyIndex, i)).Name)
                End If
            End If
        Else
            frmMirage.lstInv.AddItem "<free inventory slot>"
        End If
    Next i
    
    frmMirage.lstInv.ListIndex = 0
End Sub

Public Sub UpdateTradeInventory()
Dim i As Long

    frmPlayerTrade.PlayerInv1.Clear
    
For i = 1 To MAX_INV
    If GetPlayerInvItemNum(MyIndex, i) > 0 And GetPlayerInvItemNum(MyIndex, i) <= MAX_ITEMS Then
        If Item(GetPlayerInvItemNum(MyIndex, i)).Type = ITEM_TYPE_CURRENCY Then
            frmPlayerTrade.PlayerInv1.AddItem i & ": " & Trim(Item(GetPlayerInvItemNum(MyIndex, i)).Name) & " (" & GetPlayerInvItemValue(MyIndex, i) & ")"
        Else
            If GetPlayerWeaponSlot(MyIndex) = i Or GetPlayerArmorSlot(MyIndex) = i Or GetPlayerHelmetSlot(MyIndex) = i Or GetPlayerShieldSlot(MyIndex) = i Then
                frmPlayerTrade.PlayerInv1.AddItem i & ": " & Trim(Item(GetPlayerInvItemNum(MyIndex, i)).Name) & " (worn)"
            Else
                frmPlayerTrade.PlayerInv1.AddItem i & ": " & Trim(Item(GetPlayerInvItemNum(MyIndex, i)).Name)
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

Sub BltPlayerText(ByVal Index As Long)
Dim TextX As Long
Dim TextY As Long
Dim intLoop As Integer
Dim intLoop2 As Integer

Dim bytLineCount As Byte
Dim bytLineLength As Byte
Dim strLine(0 To MAX_LINES - 1) As String
Dim strWords() As String
    strWords() = Split(Bubble(Index).Text, " ")
    
    If Len(Bubble(Index).Text) < MAX_LINE_LENGTH Then
        DISPLAY_BUBBLE_WIDTH = 2 + ((Len(Bubble(Index).Text) * 9) \ PIC_X)
        
        If DISPLAY_BUBBLE_WIDTH > MAX_BUBBLE_WIDTH Then
            DISPLAY_BUBBLE_WIDTH = MAX_BUBBLE_WIDTH
        End If
    Else
        DISPLAY_BUBBLE_WIDTH = 6
    End If
    
    TextX = GetPlayerX(Index) * PIC_X + Player(Index).XOffset + Int(PIC_X) - ((DISPLAY_BUBBLE_WIDTH * 32) / 2) - 6
    TextY = GetPlayerY(Index) * PIC_Y + Player(Index).YOffset - Int(PIC_Y) + 85
    
    Call DD_BackBuffer.ReleaseDC(TexthDC)
    
    ' Draw the fancy box with tiles.
    Call BltTile2(TextX - 10, TextY - 10, 2)
    Call BltTile2(TextX + (DISPLAY_BUBBLE_WIDTH * PIC_X) - PIC_X - 10, TextY - 10, 3)
    
    For intLoop = 1 To (DISPLAY_BUBBLE_WIDTH - 2)
        Call BltTile2(TextX - 10 + (intLoop * PIC_X), TextY - 10, 17)
    Next intLoop
    
    TexthDC = DD_BackBuffer.GetDC
    
    ' Loop through all the words.
    For intLoop = 0 To UBound(strWords)
        ' Increment the line length.
        bytLineLength = bytLineLength + Len(strWords(intLoop)) + 1
            
        ' If we have room on the current line.
        If bytLineLength < MAX_LINE_LENGTH Then
            ' Add the text to the current line.
            strLine(bytLineCount) = strLine(bytLineCount) & strWords(intLoop) & " "
        Else
            bytLineCount = bytLineCount + 1
            
            If bytLineCount = MAX_LINES Then
                bytLineCount = bytLineCount - 1
                Exit For
            End If
            
            strLine(bytLineCount) = Trim(strWords(intLoop)) & " "
            bytLineLength = 0
        End If
    Next intLoop
    
    Call DD_BackBuffer.ReleaseDC(TexthDC)
    
    If bytLineCount > 0 Then
        For intLoop = 6 To (bytLineCount - 2) * PIC_Y + 6
            Call BltTile2(TextX - 10, TextY - 10 + intLoop, 19)
            Call BltTile2(TextX - 10 + (DISPLAY_BUBBLE_WIDTH * PIC_X) - PIC_X, TextY - 10 + intLoop, 18)
            
            For intLoop2 = 1 To DISPLAY_BUBBLE_WIDTH - 2
                Call BltTile2(TextX - 10 + (intLoop2 * PIC_X), TextY + intLoop - 10, 6)
            Next intLoop2
        Next intLoop
    End If

    Call BltTile2(TextX - 10, TextY + (bytLineCount * 16) - 4, 4)
    Call BltTile2(TextX + (DISPLAY_BUBBLE_WIDTH * PIC_X) - PIC_X - 10, TextY + (bytLineCount * 16) - 4, 5)
    
    For intLoop = 1 To (DISPLAY_BUBBLE_WIDTH - 2)
        Call BltTile2(TextX - 10 + (intLoop * PIC_X), TextY + (bytLineCount * 16) - 4, 16)
    Next intLoop
    
    TexthDC = DD_BackBuffer.GetDC
    
    For intLoop = 0 To (MAX_LINES - 1)
        If strLine(intLoop) <> "" Then
            Call DrawText(TexthDC, TextX - (NewPlayerX * PIC_X) + sx - NewXOffset + (((DISPLAY_BUBBLE_WIDTH) * PIC_X) / 2) - ((Len(strLine(intLoop)) * 8) \ 2) - 7, TextY - (NewPlayerY * PIC_Y) + sx - NewYOffset, strLine(intLoop), QBColor(DarkGrey))
            TextY = TextY + 16
        End If
    Next intLoop
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
Sub BltNpcBars(ByVal Index As Long)
Dim X As Long, Y As Long

If MapNpc(Index).HP = 0 Then Exit Sub
If MapNpc(Index).Num < 1 Then Exit Sub

    If Npc(MapNpc(Index).Num).Big = 1 Then
        X = (MapNpc(Index).X * PIC_X + sx - 9 + MapNpc(Index).XOffset) - (NewPlayerX * PIC_X) - NewXOffset
        Y = (MapNpc(Index).Y * PIC_Y + sx + MapNpc(Index).YOffset) - (NewPlayerY * PIC_Y) - NewYOffset
        
        Call DD_BackBuffer.SetFillColor(RGB(255, 0, 0))
        Call DD_BackBuffer.DrawBox(X, Y + 32, X + 50, Y + 36)
        Call DD_BackBuffer.SetFillColor(RGB(0, 255, 0))
        Call DD_BackBuffer.DrawBox(X, Y + 32, X + ((MapNpc(Index).HP / 100) / (MapNpc(Index).MaxHp / 100) * 50), Y + 36)
    Else
        X = (MapNpc(Index).X * PIC_X + sx + MapNpc(Index).XOffset) - (NewPlayerX * PIC_X) - NewXOffset
        Y = (MapNpc(Index).Y * PIC_Y + sx + MapNpc(Index).YOffset) - (NewPlayerY * PIC_Y) - NewYOffset
        
        Call DD_BackBuffer.SetFillColor(RGB(255, 0, 0))
        Call DD_BackBuffer.DrawBox(X, Y + 32, X + 32, Y + 36)
        Call DD_BackBuffer.SetFillColor(RGB(0, 255, 0))
        Call DD_BackBuffer.DrawBox(X, Y + 32, X + ((MapNpc(Index).HP / 100) / (MapNpc(Index).MaxHp / 100) * 32), Y + 36)
    End If
End Sub

Public Sub UpdateVisInv()
Dim Index As Long
Dim d As Long

    For Index = 1 To MAX_INV
        If GetPlayerShieldSlot(MyIndex) <> Index Then frmMirage.ShieldImage.Picture = LoadPicture()
        If GetPlayerWeaponSlot(MyIndex) <> Index Then frmMirage.WeaponImage.Picture = LoadPicture()
        If GetPlayerHelmetSlot(MyIndex) <> Index Then frmMirage.HelmetImage.Picture = LoadPicture()
        If GetPlayerArmorSlot(MyIndex) <> Index Then frmMirage.ArmorImage.Picture = LoadPicture()
    Next Index
    
    For Index = 1 To MAX_INV
        If GetPlayerShieldSlot(MyIndex) = Index Then Call BitBlt(frmMirage.ShieldImage.hDC, 0, 0, PIC_X, PIC_Y, frmMirage.picItems.hDC, (Item(GetPlayerInvItemNum(MyIndex, Index)).Pic - Int(Item(GetPlayerInvItemNum(MyIndex, Index)).Pic / 6) * 6) * PIC_X, Int(Item(GetPlayerInvItemNum(MyIndex, Index)).Pic / 6) * PIC_Y, SRCCOPY)
        If GetPlayerWeaponSlot(MyIndex) = Index Then Call BitBlt(frmMirage.WeaponImage.hDC, 0, 0, PIC_X, PIC_Y, frmMirage.picItems.hDC, (Item(GetPlayerInvItemNum(MyIndex, Index)).Pic - Int(Item(GetPlayerInvItemNum(MyIndex, Index)).Pic / 6) * 6) * PIC_X, Int(Item(GetPlayerInvItemNum(MyIndex, Index)).Pic / 6) * PIC_Y, SRCCOPY)
        If GetPlayerHelmetSlot(MyIndex) = Index Then Call BitBlt(frmMirage.HelmetImage.hDC, 0, 0, PIC_X, PIC_Y, frmMirage.picItems.hDC, (Item(GetPlayerInvItemNum(MyIndex, Index)).Pic - Int(Item(GetPlayerInvItemNum(MyIndex, Index)).Pic / 6) * 6) * PIC_X, Int(Item(GetPlayerInvItemNum(MyIndex, Index)).Pic / 6) * PIC_Y, SRCCOPY)
        If GetPlayerArmorSlot(MyIndex) = Index Then Call BitBlt(frmMirage.ArmorImage.hDC, 0, 0, PIC_X, PIC_Y, frmMirage.picItems.hDC, (Item(GetPlayerInvItemNum(MyIndex, Index)).Pic - Int(Item(GetPlayerInvItemNum(MyIndex, Index)).Pic / 6) * 6) * PIC_X, Int(Item(GetPlayerInvItemNum(MyIndex, Index)).Pic / 6) * PIC_Y, SRCCOPY)
    Next Index
    
    frmMirage.SelectedItem.Top = frmMirage.picInv(frmMirage.lstInv.ListIndex).Top - 1
    frmMirage.SelectedItem.Left = frmMirage.picInv(frmMirage.lstInv.ListIndex).Left - 1
    
    frmMirage.EquipS(0).Visible = False
    frmMirage.EquipS(1).Visible = False
    frmMirage.EquipS(2).Visible = False
    frmMirage.EquipS(3).Visible = False

    For d = 0 To MAX_INV - 1
        If Player(MyIndex).Inv(d + 1).Num > 0 Then
            If Item(GetPlayerInvItemNum(MyIndex, d + 1)).Type <> ITEM_TYPE_CURRENCY Then
                'frmMirage.descName.Caption = Trim(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Name) & " (" & GetPlayerInvItemValue(MyIndex, d + 1) & ")"
            'Else
                If GetPlayerWeaponSlot(MyIndex) = d + 1 Then
                    'frmMirage.picInv(d).ToolTipText = Trim(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Name) & " (worn)"
                    frmMirage.EquipS(0).Visible = True
                    frmMirage.EquipS(0).Top = frmMirage.picInv(d).Top - 2
                    frmMirage.EquipS(0).Left = frmMirage.picInv(d).Left - 2
                ElseIf GetPlayerArmorSlot(MyIndex) = d + 1 Then
                    'frmMirage.picInv(d).ToolTipText = Trim(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Name) & " (worn)"
                    frmMirage.EquipS(1).Visible = True
                    frmMirage.EquipS(1).Top = frmMirage.picInv(d).Top - 2
                    frmMirage.EquipS(1).Left = frmMirage.picInv(d).Left - 2
                ElseIf GetPlayerHelmetSlot(MyIndex) = d + 1 Then
                    'frmMirage.picInv(d).ToolTipText = Trim(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Name) & " (worn)"
                    frmMirage.EquipS(2).Visible = True
                    frmMirage.EquipS(2).Top = frmMirage.picInv(d).Top - 2
                    frmMirage.EquipS(2).Left = frmMirage.picInv(d).Left - 2
                ElseIf GetPlayerShieldSlot(MyIndex) = d + 1 Then
                    'frmMirage.picInv(d).ToolTipText = Trim(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Name) & " (worn)"
                    frmMirage.EquipS(3).Visible = True
                    frmMirage.EquipS(3).Top = frmMirage.picInv(d).Top - 2
                    frmMirage.EquipS(3).Left = frmMirage.picInv(d).Left - 2
                Else
                    'frmMirage.picInv(d).ToolTipText = Trim(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Name)
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
