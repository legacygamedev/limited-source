Attribute VB_Name = "modGameLogic"
'   Copyright (c) 2006 Joshua Bendig
'   This file is part of Asgard.
'
'    Asgard is free software; you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation; either version 2 of the License, or
'    (at your option) any later version.
'
'    Asgard is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License
'    along with Asgard; if not, write to the Free Software
'    Foundation, Inc., 51 Franklin St, Fifth Floor, Boston, MA  02110-1301  USA

Option Explicit

Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
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
Public EditorSet As Byte

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
Public SaveMapItem() As MapItemRec
Public SaveMapNpc(1 To MAX_MAP_NPCS) As MapNpcRec

' Used for index based editors
Public InItemsEditor As Boolean
Public InNpcEditor As Boolean
Public InShopEditor As Boolean
Public InSpellEditor As Boolean
Public InEmoticonEditor As Boolean
Public InArrowEditor As Boolean
Public EditorIndex As Long

' Game fps
Public GameFPS As Long

' Used for atmosphere
Public GameWeather As Long
Public GameTime As Long
Public RainIntensity As Long

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

Public NPCSpawnNum As Long
Public NPCSpawnAmount As Long
Public NPCSpawnRange As Long

Public Conf_PlayerBar As Boolean
Public Conf_NPCbar As Boolean
Public Conf_NPCdamage As Boolean
Public Conf_PlayerName As Boolean
Public Conf_PlayerDamage As Boolean
Public Conf_NPCname As Boolean
Public Conf_SpeechBubbles As Boolean

Public Conf_MapGridFlag As Boolean

                    
Sub Main()
Dim i As Long
Dim Ending As String
    
    ScreenMode = 0

    frmSendGetData.Visible = True
    Call SetStatus("Initializing...")
    DoEvents
    
    ' Initialize FMOD for the BGMs
    Call FSOUND_Init(44100, 32, 0)
    
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
    
    Dim filename As String
    filename = App.Path & "\config.ini"
    If FileExist("config.ini") Then
        frmMirage.chkbubblebar.value = ReadINI("CONFIG", "SpeechBubbles", filename)
        frmMirage.chknpcbar.value = ReadINI("CONFIG", "NpcBar", App.Path & "\config.ini")
        frmMirage.chknpcname.value = ReadINI("CONFIG", "NPCName", filename)
        frmMirage.chkplayerbar.value = ReadINI("CONFIG", "PlayerBar", App.Path & "\config.ini")
        frmMirage.chkplayername.value = ReadINI("CONFIG", "PlayerName", filename)
        frmMirage.chkplayerdamage.value = ReadINI("CONFIG", "NPCDamage", filename)
        frmMirage.chknpcdamage.value = ReadINI("CONFIG", "PlayerDamage", filename)
        frmMirage.chkmusic.value = ReadINI("CONFIG", "Music", filename)
        frmMirage.chksound.value = ReadINI("CONFIG", "Sound", filename)
        frmMirage.chkAutoScroll.value = ReadINI("CONFIG", "AutoScroll", filename)
        Conf_NPCbar = ReadINI("CONFIG", "NpcBar", App.Path & "\config.ini")
        Conf_PlayerBar = ReadINI("CONFIG", "PlayerBar", App.Path & "\config.ini")
        Conf_MapGridFlag = True
        Conf_NPCdamage = ReadINI("CONFIG", "NPCDamage", App.Path & "\config.ini")
        Conf_PlayerName = ReadINI("CONFIG", "PlayerName", App.Path & "\config.ini")
        Conf_PlayerDamage = ReadINI("CONFIG", "PlayerDamage", App.Path & "\config.ini")
        Conf_NPCname = ReadINI("CONFIG", "NPCName", App.Path & "\config.ini")
        Conf_SpeechBubbles = ReadINI("CONFIG", "SpeechBubbles", App.Path & "\config.ini")


        If ReadINI("CONFIG", "MapGrid", filename) = 0 Then
            frmMapEditor.mnuMapGrid.Checked = False
        Else
            frmMapEditor.mnuMapGrid.Checked = True
        End If
    Else
        WriteINI "UPDATER", "FileName", "Konfuze.exe", App.Path & "\config.ini"
        WriteINI "UPDATER", "WebSite", "", App.Path & "\config.ini"
        WriteINI "IPCONFIG", "IP", "127.0.0.1", App.Path & "\config.ini"
        WriteINI "IPCONFIG", "PORT", 4000, App.Path & "\config.ini"
        WriteINI "CONFIG", "Account", "", App.Path & "\config.ini"
        WriteINI "CONFIG", "Password", "", App.Path & "\config.ini"
        WriteINI "CONFIG", "WebSite", "", App.Path & "\config.ini"
        WriteINI "CONFIG", "SpeechBubbles", 1, App.Path & "\config.ini"
        WriteINI "CONFIG", "NpcBar", 1, App.Path & "\config.ini"
        WriteINI "CONFIG", "NPCName", 1, App.Path & "\config.ini"
        WriteINI "CONFIG", "NPCDamage", 1, App.Path & "\config.ini"
        WriteINI "CONFIG", "PlayerBar", 1, App.Path & "\config.ini"
        WriteINI "CONFIG", "PlayerName", 1, App.Path & "\config.ini"
        WriteINI "CONFIG", "PlayerDamage", 1, App.Path & "\config.ini"
        WriteINI "CONFIG", "MapGrid", 1, App.Path & "\config.ini"
        WriteINI "CONFIG", "Music", 1, App.Path & "\config.ini"
        WriteINI "CONFIG", "Sound", 1, App.Path & "\config.ini"
        WriteINI "CONFIG", "AutoScroll", 1, App.Path & "\config.ini"
    End If
    
    If FileExist("GUI\Colors.txt") = False Then
        WriteINI "CHATBOX", "R", 152, App.Path & "\GUI\Colors.txt"
        WriteINI "CHATBOX", "G", 146, App.Path & "\GUI\Colors.txt"
        WriteINI "CHATBOX", "B", 120, App.Path & "\GUI\Colors.txt"
        
        WriteINI "CHATTEXTBOX", "R", 152, App.Path & "\GUI\Colors.txt"
        WriteINI "CHATTEXTBOX", "G", 146, App.Path & "\GUI\Colors.txt"
        WriteINI "CHATTEXTBOX", "B", 120, App.Path & "\GUI\Colors.txt"
        
        WriteINI "BACKGROUND", "R", 152, App.Path & "\GUI\Colors.txt"
        WriteINI "BACKGROUND", "G", 146, App.Path & "\GUI\Colors.txt"
        WriteINI "BACKGROUND", "B", 120, App.Path & "\GUI\Colors.txt"
        
        WriteINI "SPELLLIST", "R", 152, App.Path & "\GUI\Colors.txt"
        WriteINI "SPELLLIST", "G", 146, App.Path & "\GUI\Colors.txt"
        WriteINI "SPELLLIST", "B", 120, App.Path & "\GUI\Colors.txt"

        WriteINI "WHOLIST", "R", 152, App.Path & "\GUI\Colors.txt"
        WriteINI "WHOLIST", "G", 146, App.Path & "\GUI\Colors.txt"
        WriteINI "WHOLIST", "B", 120, App.Path & "\GUI\Colors.txt"
        
        WriteINI "NEWCHAR", "R", 152, App.Path & "\GUI\Colors.txt"
        WriteINI "NEWCHAR", "G", 146, App.Path & "\GUI\Colors.txt"
        WriteINI "NEWCHAR", "B", 120, App.Path & "\GUI\Colors.txt"
    End If
    
    Dim R1 As Long, G1 As Long, B1 As Long
    R1 = Val(ReadINI("CHATBOX", "R", App.Path & "\GUI\Colors.txt"))
    G1 = Val(ReadINI("CHATBOX", "G", App.Path & "\GUI\Colors.txt"))
    B1 = Val(ReadINI("CHATBOX", "B", App.Path & "\GUI\Colors.txt"))
    frmMirage.txtChat.BackColor = RGB(R1, G1, B1)
    
    R1 = Val(ReadINI("CHATTEXTBOX", "R", App.Path & "\GUI\Colors.txt"))
    G1 = Val(ReadINI("CHATTEXTBOX", "G", App.Path & "\GUI\Colors.txt"))
    B1 = Val(ReadINI("CHATTEXTBOX", "B", App.Path & "\GUI\Colors.txt"))
        
    R1 = Val(ReadINI("BACKGROUND", "R", App.Path & "\GUI\Colors.txt"))
    G1 = Val(ReadINI("BACKGROUND", "G", App.Path & "\GUI\Colors.txt"))
    B1 = Val(ReadINI("BACKGROUND", "B", App.Path & "\GUI\Colors.txt"))
    
    R1 = Val(ReadINI("SPELLLIST", "R", App.Path & "\GUI\Colors.txt"))
    G1 = Val(ReadINI("SPELLLIST", "G", App.Path & "\GUI\Colors.txt"))
    B1 = Val(ReadINI("SPELLLIST", "B", App.Path & "\GUI\Colors.txt"))
    frmMirage.lstSpells.BackColor = RGB(R1, G1, B1)
    
    R1 = Val(ReadINI("WHOLIST", "R", App.Path & "\GUI\Colors.txt"))
    G1 = Val(ReadINI("WHOLIST", "G", App.Path & "\GUI\Colors.txt"))
    B1 = Val(ReadINI("WHOLIST", "B", App.Path & "\GUI\Colors.txt"))
    
    R1 = Val(ReadINI("NEWCHAR", "R", App.Path & "\GUI\Colors.txt"))
    G1 = Val(ReadINI("NEWCHAR", "G", App.Path & "\GUI\Colors.txt"))
    B1 = Val(ReadINI("NEWCHAR", "B", App.Path & "\GUI\Colors.txt"))
    frmNewChar.optMale.BackColor = RGB(R1, G1, B1)
    frmNewChar.optFemale.BackColor = RGB(R1, G1, B1)
    
    Call SetStatus("Checking status...")
    DoEvents
    
    ' Make sure we set that we aren't in the game
    InGame = False
    GettingMap = True
    InEditor = False
    InItemsEditor = False
    InNpcEditor = False
    InShopEditor = False
    InEmoticonEditor = False
    InArrowEditor = False
    
    frmMirage.picItems.Picture = LoadPicture(App.Path & "\GFX\items.bmp")
    frmSpriteChange.picSprites.Picture = LoadPicture(App.Path & "\GFX\sprites.bmp")
    
    Call SetStatus("Initializing TCP Settings...")
    DoEvents
    
    Call TcpInit
    PlayMidi (ReadINI("CONFIG", "TitleBGM", App.Path & "\config.ini"))
    frmLogin.Visible = True
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
                Call SetStatus("Please wait...")
                Call SendNewAccount(frmNewAccount.txtName.Text, frmNewAccount.txtPassword.Text)
                Call frmMirage.Socket.Close
                frmLogin.Visible = True
                frmSendGetData.Visible = False
                Exit Sub
            End If
            
        Case MENU_STATE_DELACCOUNT
            frmDeleteAccount.Visible = False
            If ConnectToServer = True Then
                Call SetStatus("Please wait...")
                Call SendDelAccount(frmDeleteAccount.txtName.Text, frmDeleteAccount.txtPassword.Text)
            End If
        
        Case MENU_STATE_LOGIN
            frmLogin.Visible = False
            If ConnectToServer = True Then
                Call SetStatus("Please wait...")
                Call SendLogin(frmLogin.txtName.Text, frmLogin.txtPassword.Text)
            End If
        
        Case MENU_STATE_NEWCHAR
            frmChars.Visible = False
            If ConnectToServer = True Then
                Call SetStatus("Please wait...")
                Call SendGetClasses
            End If
            
        Case MENU_STATE_ADDCHAR
            frmNewChar.Visible = False
            If ConnectToServer = True Then
                Call SetStatus("Please wait...")
                If frmNewChar.optMale.value = True Then
                    Call SendAddChar(frmNewChar.txtName, 0, frmNewChar.cmbClass.ListIndex, frmChars.lstChars.ListIndex + 1)
                Else
                    Call SendAddChar(frmNewChar.txtName, 1, frmNewChar.cmbClass.ListIndex, frmChars.lstChars.ListIndex + 1)
                End If
            End If
        
        Case MENU_STATE_DELCHAR
            frmChars.Visible = False
            If ConnectToServer = True Then
                Call SetStatus("Please wait...")
                Call SendDelChar(frmChars.lstChars.ListIndex + 1)
            End If
            
        Case MENU_STATE_USECHAR
            frmChars.Visible = False
            If ConnectToServer = True Then
                Call StopMidi
                Call SetStatus("Please wait...")
                Call SendUseChar(frmChars.lstChars.ListIndex + 1)
            End If
    End Select

    If Not IsConnected And Connucted = True Then
        frmLogin.Visible = True
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
Dim x As Long
Dim y As Long
Dim i As Long
Dim rec_back As RECT
    
    ' Set the focus
    'frmMirage.picScreen.SetFocus
    
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
                
    If GettingMap = False Then
        If GetPlayerPOINTS(MyIndex) > 0 Then
            frmMirage.AddStr.Visible = True
            frmMirage.AddDef.Visible = True
            frmMirage.AddSpeed.Visible = True
            frmMirage.AddMagi.Visible = True
        Else
            frmMirage.AddStr.Visible = False
            frmMirage.AddDef.Visible = False
            frmMirage.AddSpeed.Visible = False
            frmMirage.AddMagi.Visible = False
        End If
        ' Visual Inventory
        Dim Q As Long
        Dim Qq As Long
        Dim IT As Long
               
        If GetTickCount > IT + 500 And frmMirage.picInv3.Visible = True Then
            For Q = 0 To MAX_INV - 1
                Qq = Player(MyIndex).Inv(Q + 1).num
               
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
                        
        NewX = 13
        NewY = 8 'oldval 8
                          
        NewPlayerY = Player(MyIndex).y - NewY
        NewPlayerX = Player(MyIndex).x - NewX
        
        NewX = NewX * PIC_X
        NewY = NewY * PIC_Y
        
        NewXOffset = Player(MyIndex).XOffset
        NewYOffset = Player(MyIndex).YOffset
        
        'NewX / NewY = Player Onscreen Position
        'NewPlayerX / NewPlayerY = Player Offset (all players / NPCs)
        'NewXOffset / NewYOffset = Map Offset
        
        'If Player(MyIndex).x <= 13 Then
        '    NewX = (Player(MyIndex).x * PIC_Y) + Player(MyIndex).XOffset
        '    NewPlayerX = 0
        '    NewXOffset = 0
        'End If
        
        'If Player(MyIndex).x >= MAX_MAPX - 13 Then
        '    NewX = ((Player(MyIndex).x) * PIC_X) + Player(MyIndex).XOffset
        '    NewPlayerX = 0
        '    NewXOffset = MAX_MAPX - 13
        'End If
        
        
        If Player(MyIndex).y * 32 + Player(MyIndex).YOffset - (7.5 * 32) < 16 Then
            NewY = Player(MyIndex).y * PIC_Y + Player(MyIndex).YOffset
            NewYOffset = 0
            NewPlayerY = 0
        ElseIf Player(MyIndex).y * 32 + Player(MyIndex).YOffset + (7.5 * 32) > (MAX_MAPY + 0.5) * 32 Then
            NewY = (Player(MyIndex).y - 15) * PIC_Y + Player(MyIndex).YOffset
            NewYOffset = 0
            NewPlayerY = MAX_MAPY - 15
        End If
        
        If (Player(MyIndex).x * 32 + Player(MyIndex).XOffset - (13 * 32)) < 1 Then
            NewX = Player(MyIndex).x * PIC_X + Player(MyIndex).XOffset
            NewXOffset = 0
            NewPlayerX = 0
        ElseIf (Player(MyIndex).x * 32) + Player(MyIndex).XOffset + (13 * 32) > (MAX_MAPX) * 32 Then
            NewX = (Player(MyIndex).x - 4) * PIC_X + Player(MyIndex).XOffset
            NewXOffset = 0
            NewPlayerX = MAX_MAPX - 26
        End If
       
        sx = 32
        ' Blit out tiles layers ground/anim1/anim2
        For y = 0 To MAX_MAPY
            For x = 0 To MAX_MAPX
                If IsVisible(x, y) Then
                    Call BltTile(x, y)
                End If
            Next x
        Next y
       
    If ScreenMode = 0 Then
        ' Blit out the items
        For i = 1 To MAX_MAP_ITEMS
            If MapItem(i).num > 0 Then
                Call BltItem(i)
            End If
        Next i
        
        If frmMirage.chknpcbar.value = 1 Then
            ' Blit out NPC hp bars
            For i = 1 To MAX_MAP_NPCS
                Call BltNpcBars(i)
            Next i
            
            For y = 0 To MAX_MAPY
                For x = 0 To MAX_MAPX
                    If Map(GetPlayerMap(MyIndex)).Tile(x, y).Type = TILE_TYPE_NPC_SPAWN Then
                        For i = 1 To MAX_ATTRIBUTE_NPCS
                            Call BltAttributeNpcBars(i, x, y)
                        Next i
                    End If
                Next x
            Next y
        End If
              
        If frmMirage.chkplayerbar.value = 1 Then
            ' Blit players bar
            'For i = 1 To MAX_PLAYERS
                'If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                    Call BltPlayerBar 's(i)
                'End If
            'Next i
        End If
        
        ' Blit out the sprite change attribute
        For y = 0 To MAX_MAPY
            For x = 0 To MAX_MAPX
                If IsVisible(x, y) Then
                    Call BltSpriteChange(x, y)
                End If
            Next x
        Next y
        
        ' Blit out arrows
        For i = 1 To MAX_PLAYERS
            If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                Call BltArrow(i)
            End If
        Next i

        ' Blit out players
        For i = 1 To MAX_PLAYERS
            If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                Call BltPlayer(i)
            End If
        Next i
        
        For y = 0 To MAX_MAPY
            For x = 0 To MAX_MAPX
                If Map(GetPlayerMap(MyIndex)).Tile(x, y).Type = TILE_TYPE_NPC_SPAWN Then
                    For i = 1 To MAX_ATTRIBUTE_NPCS
                        Call BltAttributeNpc(i, x, y)
                    Next i
                End If
            Next x
        Next y
        
        ' Blit out the npcs
        For i = 1 To MAX_MAP_NPCS
            Call BltNpc(i)
        Next i
                
        For i = 1 To MAX_PLAYERS
            If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                Call BltSpell(i)
            End If
        Next i
        
        ' Blit out the sprite change attribute
        For y = 0 To MAX_MAPY
            For x = 0 To MAX_MAPX
                If IsVisible(x, y) Then
                    Call BltSpriteChange2(x, y)
                End If
            Next x
        Next y
        
        ' Blit out the npcs
        For i = 1 To MAX_MAP_NPCS
            Call BltNpcTop(i)
        Next i
        
        For y = 0 To MAX_MAPY
            For x = 0 To MAX_MAPX
                If Map(GetPlayerMap(MyIndex)).Tile(x, y).Type = TILE_TYPE_NPC_SPAWN Then
                    For i = 1 To MAX_ATTRIBUTE_NPCS
                        Call BltAttributeNpcTop(i, x, y)
                    Next i
                End If
            Next x
        Next y
    End If
                
    ' Blit out tile layer fringe
    For y = 0 To MAX_MAPY
        For x = 0 To MAX_MAPX
            If IsVisible(x, y) Then
                Call BltFringeTile(x, y)
            End If
        Next x
    Next y
      
    If ScreenMode = 0 Then
        ' Blit out the npcs
        For i = 1 To MAX_MAP_NPCS
            If Map(GetPlayerMap(MyIndex)).Tile(MapNpc(i).x, MapNpc(i).y).Fringe < 1 Then
                If Map(GetPlayerMap(MyIndex)).Tile(MapNpc(i).x, MapNpc(i).y).FAnim < 1 Then
                    If Map(GetPlayerMap(MyIndex)).Tile(MapNpc(i).x, MapNpc(i).y).Fringe2 < 1 Then
                        If Map(GetPlayerMap(MyIndex)).Tile(MapNpc(i).x, MapNpc(i).y).F2Anim < 1 Then
                            Call BltNpcTop(i)
                        End If
                    End If
                End If
            End If
        Next i
        
        For y = 0 To MAX_MAPY
            For x = 0 To MAX_MAPX
                If Map(GetPlayerMap(MyIndex)).Tile(x, y).Type = TILE_TYPE_NPC_SPAWN Then
                    For i = 1 To MAX_ATTRIBUTE_NPCS
                        If Map(GetPlayerMap(MyIndex)).Tile(MapAttributeNpc(i, x, y).x, MapAttributeNpc(i, x, y).y).Fringe < 1 Then
                            If Map(GetPlayerMap(MyIndex)).Tile(MapAttributeNpc(i, x, y).x, MapAttributeNpc(i, x, y).y).FAnim < 1 Then
                                If Map(GetPlayerMap(MyIndex)).Tile(MapAttributeNpc(i, x, y).x, MapAttributeNpc(i, x, y).y).Fringe2 < 1 Then
                                    If Map(GetPlayerMap(MyIndex)).Tile(MapAttributeNpc(i, x, y).x, MapAttributeNpc(i, x, y).y).F2Anim < 1 Then
                                        Call BltAttributeNpcTop(i, x, y)
                                    End If
                                End If
                            End If
                        End If
                    Next i
                End If
            Next x
        Next y
    End If
    
    For i = 1 To MAX_PLAYERS
            If IsPlaying(i) = True Then
                If Player(i).LevelUpT + 3000 > GetTickCount Then
                    rec.Top = Int(32 / TilesInSheets) * PIC_Y
                    rec.Bottom = rec.Top + PIC_Y
                    rec.Left = (32 - Int(32 / TilesInSheets) * TilesInSheets) * PIC_X
                    rec.Right = rec.Left + 96
                    
                    If i = MyIndex Then
                        x = NewX + sx
                        y = NewY + sx
                        Call DD_BackBuffer.BltFast(x - 32, y - 10 - Player(i).LevelUp, DD_TileSurf(6), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    Else
                        x = GetPlayerX(i) * PIC_X + sx + Player(i).XOffset
                        y = GetPlayerY(i) * PIC_Y + sx + Player(i).YOffset
                        Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) - 32 - NewXOffset, y - (NewPlayerY * PIC_Y) - 10 - Player(i).LevelUp - NewYOffset, DD_TileSurf(6), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If
                    If Player(i).LevelUp >= 3 Then
                        Player(i).LevelUp = Player(i).LevelUp - 1
                    ElseIf Player(i).LevelUp >= 1 Then
                        Player(i).LevelUp = Player(i).LevelUp + 1
                    End If
                Else
                    Player(i).LevelUpT = 0
                End If
            End If
        Next i
        If GettingMap = False Then
            If GameTime = TIME_NIGHT And Map(GetPlayerMap(MyIndex)).Indoors = 0 And InEditor = False Then
                Call Night
            End If
            If frmMapEditor.mnuDayNight.Checked = True And InEditor = True Then
                Call Night
            End If
            If Map(GetPlayerMap(MyIndex)).Indoors = 0 Then Call BltWeather
        End If

        If InEditor = True And Conf_MapGridFlag = 1 Then
            For y = 0 To MAX_MAPY
                For x = 0 To MAX_MAPX
                    Call BltTile2(x * 32, y * 32, 0)
                Next x
            Next y
        End If
    End If
    
        ' Lock the backbuffer so we can draw text and names
        TexthDC = DD_BackBuffer.GetDC
    If GettingMap = False Then
        If ScreenMode = 0 Then
            If frmMirage.chknpcdamage.value = 1 Then
                If frmMirage.chkplayername.value = 0 Then
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
            
            If frmMirage.chkplayerdamage.value = 1 Then
                If NPCWho > 0 Then
                    If MapNpc(NPCWho).num > 0 Then
                        If frmMirage.chknpcdamage.value = 0 Then
                            If Npc(MapNpc(NPCWho).num).Big = 0 Then
                                If GetTickCount < DmgTime + 2000 Then
                                    Call DrawText(TexthDC, (MapNpc(NPCWho).x - NewPlayerX) * PIC_X + sx + (Int(Len(DmgDamage)) / 2) * 3 + MapNpc(NPCWho).XOffset - NewXOffset, (MapNpc(NPCWho).y - NewPlayerY) * PIC_Y + sx - 20 + MapNpc(NPCWho).YOffset - NewYOffset - iii, DmgDamage, QBColor(White))
                                End If
                            Else
                                If GetTickCount < DmgTime + 2000 Then
                                    Call DrawText(TexthDC, (MapNpc(NPCWho).x - NewPlayerX) * PIC_X + sx + (Int(Len(DmgDamage)) / 2) * 3 + MapNpc(NPCWho).XOffset - NewXOffset, (MapNpc(NPCWho).y - NewPlayerY) * PIC_Y + sx - 47 + MapNpc(NPCWho).YOffset - NewYOffset - iii, DmgDamage, QBColor(White))
                                End If
                            End If
                        Else
                            If Npc(MapNpc(NPCWho).num).Big = 0 Then
                                If GetTickCount < DmgTime + 2000 Then
                                    Call DrawText(TexthDC, (MapNpc(NPCWho).x - NewPlayerX) * PIC_X + sx + (Int(Len(DmgDamage)) / 2) * 3 + MapNpc(NPCWho).XOffset - NewXOffset, (MapNpc(NPCWho).y - NewPlayerY) * PIC_Y + sx - 30 + MapNpc(NPCWho).YOffset - NewYOffset - iii, DmgDamage, QBColor(White))
                                End If
                            Else
                                If GetTickCount < DmgTime + 2000 Then
                                    Call DrawText(TexthDC, (MapNpc(NPCWho).x - NewPlayerX) * PIC_X + sx + (Int(Len(DmgDamage)) / 2) * 3 + MapNpc(NPCWho).XOffset - NewXOffset, (MapNpc(NPCWho).y - NewPlayerY) * PIC_Y + sx - 57 + MapNpc(NPCWho).YOffset - NewYOffset - iii, DmgDamage, QBColor(White))
                                End If
                            End If
                        End If
                        iii = iii + 1
                    End If
                End If
            End If
            
            If frmMirage.chkplayername.value = 1 Then
                For i = 1 To MAX_PLAYERS
                    If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                        Call BltPlayerGuildName(i)
                        Call BltPlayerName(i)
                    End If
                Next i
            End If
     
            ' speech bubble stuffs
            If Conf_SpeechBubbles = 1 Then
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
            If frmMirage.chknpcdamage.value = 1 Then
                For i = LBound(MapNpc) To UBound(MapNpc)
                    If MapNpc(i).num > 0 Then
                        Call BltMapNPCName(i)
                    End If
                Next i
                
                For y = 0 To MAX_MAPY
                    For x = 0 To MAX_MAPX
                        If Map(GetPlayerMap(MyIndex)).Tile(x, y).Type = TILE_TYPE_NPC_SPAWN Then
                            For i = 1 To MAX_ATTRIBUTE_NPCS
                                Call BltAttributeNPCName(i, x, y)
                            Next i
                        End If
                    Next x
                Next y
            End If
            
            ' Blit out attribs if in editor
            If InEditor Then
                For y = 0 To MAX_MAPY
                    For x = 0 To MAX_MAPX
                        With Map(GetPlayerMap(MyIndex)).Tile(x, y)
                            If .Type = TILE_TYPE_BLOCKED Then Call DrawText(TexthDC, x * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "B", QBColor(BrightRed))
                            If .Type = TILE_TYPE_WARP Then Call DrawText(TexthDC, x * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "W", QBColor(BrightBlue))
                            If .Type = TILE_TYPE_ITEM Then Call DrawText(TexthDC, x * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "I", QBColor(White))
                            If .Type = TILE_TYPE_NPCAVOID Then Call DrawText(TexthDC, x * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "N", QBColor(White))
                            If .Type = TILE_TYPE_KEY Then Call DrawText(TexthDC, x * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "K", QBColor(White))
                            If .Type = TILE_TYPE_KEYOPEN Then Call DrawText(TexthDC, x * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "O", QBColor(White))
                            If .Type = TILE_TYPE_HEAL Then Call DrawText(TexthDC, x * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "H", QBColor(BrightGreen))
                            If .Type = TILE_TYPE_KILL Then Call DrawText(TexthDC, x * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "K", QBColor(BrightRed))
                            If .Type = TILE_TYPE_SHOP Then Call DrawText(TexthDC, x * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "S", QBColor(Yellow))
                            If .Type = TILE_TYPE_CBLOCK Then Call DrawText(TexthDC, x * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "CB", QBColor(Black))
                            If .Type = TILE_TYPE_ARENA Then Call DrawText(TexthDC, x * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "A", QBColor(BrightGreen))
                            If .Type = TILE_TYPE_SOUND Then Call DrawText(TexthDC, x * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "PS", QBColor(Yellow))
                            If .Type = TILE_TYPE_SPRITE_CHANGE Then Call DrawText(TexthDC, x * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "SC", QBColor(Grey))
                            If .Type = TILE_TYPE_SIGN Then Call DrawText(TexthDC, x * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "SI", QBColor(Yellow))
                            If .Type = TILE_TYPE_DOOR Then Call DrawText(TexthDC, x * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "D", QBColor(Black))
                            If .Type = TILE_TYPE_NOTICE Then Call DrawText(TexthDC, x * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "N", QBColor(BrightGreen))
                            If .Type = TILE_TYPE_CHEST Then Call DrawText(TexthDC, x * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "C", QBColor(Brown))
                            If .Type = TILE_TYPE_CLASS_CHANGE Then Call DrawText(TexthDC, x * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "CG", QBColor(White))
                            If .Type = TILE_TYPE_SCRIPTED Then Call DrawText(TexthDC, x * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "SC", QBColor(Yellow))
                            If .Type = TILE_TYPE_NPC_SPAWN Then Call DrawText(TexthDC, x * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "NPC", QBColor(BrightGreen))
                            If .Light > 0 Then Call DrawText(TexthDC, x * PIC_X + sx + 18 - (NewPlayerX * PIC_X) - NewXOffset, y * PIC_Y + sx + 14 - (NewPlayerY * PIC_Y) - NewYOffset, "L", QBColor(Yellow))
                        End With
                    Next x
                Next y
            End If
            
            ' Blit the text they are putting in
            'MyText = frmMirage.txtMyTextBox.Text
            'frmMirage.txtMyTextBox.Text = MyText
            
            'If Len(MyText) > 4 Then
                'frmMirage.txtMyTextBox.SelStart = Len(frmMirage.txtMyTextBox.Text) + 1
            'End If
                    
            ' Draw map name (disabled since I don't want a map name ;p)
            'If Map(GetPlayerMap(MyIndex)).Moral = MAP_MORAL_NONE Then
            '    Call DrawText(TexthDC, Int((20.5) * PIC_X / 2) - (Int(Len(Trim(Map(GetPlayerMap(MyIndex)).name)) / 2) * 8) + sx, 2 + sx, Trim(Map(GetPlayerMap(MyIndex)).name), QBColor(BrightRed))
            'ElseIf Map(GetPlayerMap(MyIndex)).Moral = MAP_MORAL_SAFE Then
            '    Call DrawText(TexthDC, Int((20.5) * PIC_X / 2) - (Int(Len(Trim(Map(GetPlayerMap(MyIndex)).name)) / 2) * 8) + sx, 2 + sx, Trim(Map(GetPlayerMap(MyIndex)).name), QBColor(White))
            'ElseIf Map(GetPlayerMap(MyIndex)).Moral = MAP_MORAL_NO_PENALTY Then
            '    Call DrawText(TexthDC, Int((20.5) * PIC_X / 2) - (Int(Len(Trim(Map(GetPlayerMap(MyIndex)).name)) / 2) * 8) + sx, 2 + sx, Trim(Map(GetPlayerMap(MyIndex)).name), QBColor(Black))
            'End If
            
            For i = 1 To MAX_BLT_LINE
                If BattlePMsg(i).index > 0 Then
                    If BattlePMsg(i).Time + 7000 > GetTickCount Then
                        Call DrawText(TexthDC, 1 + sx, BattlePMsg(i).y + frmMirage.picScreen.Height - 15 + sx, Trim(BattlePMsg(i).Msg), QBColor(BattlePMsg(i).Color))
                    Else
                        BattlePMsg(i).Done = 0
                    End If
                End If
                
                If BattleMMsg(i).index > 0 Then
                    If BattleMMsg(i).Time + 7000 > GetTickCount Then
                        Call DrawText(TexthDC, (frmMirage.picScreen.Width - (Len(BattleMMsg(i).Msg) * 8)) + sx, BattleMMsg(i).y + frmMirage.picScreen.Height - 15 + sx, Trim(BattleMMsg(i).Msg), QBColor(BattleMMsg(i).Color))
                    Else
                        BattleMMsg(i).Done = 0
                    End If
                End If
            Next i
        End If
    End If

        ' Check if we are getting a map, and if we are tell them so
        ' Probably gonna insert loading screens here
        If GettingMap = True Then
            Call DrawText(TexthDC, 36, 36, "Receiving map...", QBColor(BrightCyan))
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
        Call DX.GetWindowRect(frmMirage.picScreen.hwnd, rec_pos)
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
            If Map(GetPlayerMap(MyIndex)).Npc(i) > 0 Then
                Call ProcessNpcMovement(i)
            End If
        Next i
        
        ' Process npc movements (actually move them)
        For x = 0 To MAX_MAPX
            For y = 0 To MAX_MAPY
                For i = 1 To MAX_ATTRIBUTE_NPCS
                    If Map(GetPlayerMap(MyIndex)).Tile(x, y).Type = TILE_TYPE_NPC_SPAWN Then
                        If MapAttributeNpc(i, x, y).num > 0 Then
                            Call ProcessAttributeNpcMovement(i, x, y)
                        End If
                    End If
                Next i
            Next y
        Next x
  
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
    FSOUND_Close
    End
End Sub

Sub BltTile(ByVal x As Long, ByVal y As Long)
Dim Ground As Long
Dim Anim1 As Long
Dim Anim2 As Long
Dim Mask2 As Long
Dim M2Anim As Long
Dim GroundTileSet As Byte
Dim MaskTileSet As Byte
Dim AnimTileSet As Byte
Dim Mask2TileSet As Byte
Dim M2AnimTileSet As Byte

    Ground = Map(GetPlayerMap(MyIndex)).Tile(x, y).Ground
    Anim1 = Map(GetPlayerMap(MyIndex)).Tile(x, y).Mask
    Anim2 = Map(GetPlayerMap(MyIndex)).Tile(x, y).Anim
    Mask2 = Map(GetPlayerMap(MyIndex)).Tile(x, y).Mask2
    M2Anim = Map(GetPlayerMap(MyIndex)).Tile(x, y).M2Anim
    
    GroundTileSet = Map(GetPlayerMap(MyIndex)).Tile(x, y).GroundSet
    MaskTileSet = Map(GetPlayerMap(MyIndex)).Tile(x, y).MaskSet
    AnimTileSet = Map(GetPlayerMap(MyIndex)).Tile(x, y).AnimSet
    Mask2TileSet = Map(GetPlayerMap(MyIndex)).Tile(x, y).Mask2Set
    M2AnimTileSet = Map(GetPlayerMap(MyIndex)).Tile(x, y).M2AnimSet
        
    ' Only used if ever want to switch to blt rather then bltfast
    With rec_pos
        .Top = (y - NewPlayerY) * PIC_Y + sx - NewYOffset
        .Bottom = .Top + PIC_Y
        .Left = (x - NewPlayerX) * PIC_X + sx - NewXOffset
        .Right = .Left + PIC_X
    End With
    
    If TileFile(GroundTileSet) = 0 Then Exit Sub
    rec.Top = Int(Ground / TilesInSheets) * PIC_Y
    rec.Bottom = rec.Top + PIC_Y
    rec.Left = (Ground - Int(Ground / TilesInSheets) * TilesInSheets) * PIC_X
    rec.Right = rec.Left + PIC_X
    Call DD_BackBuffer.Blt(rec_pos, DD_TileSurf(GroundTileSet), rec, DDBLT_WAIT)
    'Call DD_BackBuffer.BltFast((X - NewPlayerX) * PIC_X - NewXOffset, (Y - NewPlayerY) * PIC_Y - NewYOffset, DD_TileSurf, rec, DDBLTFAST_WAIT)
    
    If (MapAnim = 0) Or (Anim2 <= 0) Then
        ' Is there an animation tile to plot?
        If Anim1 > 0 And TempTile(x, y).DoorOpen = NO Then
            If TileFile(MaskTileSet) = 0 Then Exit Sub
            rec.Top = Int(Anim1 / TilesInSheets) * PIC_Y
            rec.Bottom = rec.Top + PIC_Y
            rec.Left = (Anim1 - Int(Anim1 / TilesInSheets) * TilesInSheets) * PIC_X
            rec.Right = rec.Left + PIC_X
            Call DD_BackBuffer.Blt(rec_pos, DD_TileSurf(MaskTileSet), rec, DDBLT_WAIT Or DDBLT_KEYSRC)
            'Call DD_BackBuffer.BltFast((X - NewPlayerX) * PIC_X - NewXOffset, (Y - NewPlayerY) * PIC_Y - NewYOffset, DD_TileSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    Else
        ' Is there a second animation tile to plot?
        If Anim2 > 0 Then
            If TileFile(AnimTileSet) = 0 Then Exit Sub
            rec.Top = Int(Anim2 / TilesInSheets) * PIC_Y
            rec.Bottom = rec.Top + PIC_Y
            rec.Left = (Anim2 - Int(Anim2 / TilesInSheets) * TilesInSheets) * PIC_X
            rec.Right = rec.Left + PIC_X
            Call DD_BackBuffer.Blt(rec_pos, DD_TileSurf(AnimTileSet), rec, DDBLT_WAIT Or DDBLT_KEYSRC)
            'Call DD_BackBuffer.BltFast((X - NewPlayerX) * PIC_X - NewXOffset, (Y - NewPlayerY) * PIC_Y - NewYOffset, DD_TileSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    End If
    
    If (MapAnim = 0) Or (M2Anim <= 0) Then
        ' Is there an animation tile to plot?
        If Mask2 > 0 Then
            If TileFile(Mask2TileSet) = 0 Then Exit Sub
            rec.Top = Int(Mask2 / TilesInSheets) * PIC_Y
            rec.Bottom = rec.Top + PIC_Y
            rec.Left = (Mask2 - Int(Mask2 / TilesInSheets) * TilesInSheets) * PIC_X
            rec.Right = rec.Left + PIC_X
            Call DD_BackBuffer.Blt(rec_pos, DD_TileSurf(Mask2TileSet), rec, DDBLT_WAIT Or DDBLT_KEYSRC)
            'Call DD_BackBuffer.BltFast((X - NewPlayerX) * PIC_X - NewXOffset, (Y - NewPlayerY) * PIC_Y - NewYOffset, DD_TileSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    Else
        ' Is there a second animation tile to plot?
        If M2Anim > 0 Then
            If TileFile(M2AnimTileSet) = 0 Then Exit Sub
            rec.Top = Int(M2Anim / TilesInSheets) * PIC_Y
            rec.Bottom = rec.Top + PIC_Y
            rec.Left = (M2Anim - Int(M2Anim / TilesInSheets) * TilesInSheets) * PIC_X
            rec.Right = rec.Left + PIC_X
            Call DD_BackBuffer.Blt(rec_pos, DD_TileSurf(M2AnimTileSet), rec, DDBLT_WAIT Or DDBLT_KEYSRC)
            'Call DD_BackBuffer.BltFast((X - NewPlayerX) * PIC_X - NewXOffset, (Y - NewPlayerY) * PIC_Y - NewYOffset, DD_TileSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    End If
End Sub

Sub BltItem(ByVal ItemNum As Long)
    ' Only used if ever want to switch to blt rather then bltfast
    With rec_pos
        .Top = MapItem(ItemNum).y * PIC_Y
        .Bottom = .Top + PIC_Y
        .Left = MapItem(ItemNum).x * PIC_X
        .Right = .Left + PIC_X
    End With
    
    rec.Top = Int(Item(MapItem(ItemNum).num).Pic / 6) * PIC_Y
    rec.Bottom = rec.Top + PIC_Y
    rec.Left = (Item(MapItem(ItemNum).num).Pic - Int(Item(MapItem(ItemNum).num).Pic / 6) * 6) * PIC_X
    rec.Right = rec.Left + PIC_X
    
    'Call DD_BackBuffer.Blt(rec_pos, DD_ItemSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
    Call DD_BackBuffer.BltFast((MapItem(ItemNum).x - NewPlayerX) * PIC_X + sx - NewXOffset, (MapItem(ItemNum).y - NewPlayerY) * PIC_Y + sx - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End Sub

Sub BltFringeTile(ByVal x As Long, ByVal y As Long)
Dim Fringe As Long
Dim FAnim As Long
Dim Fringe2 As Long
Dim F2Anim As Long
Dim FringeTileSet As Byte
Dim FAnimTileSet As Byte
Dim Fringe2TileSet As Byte
Dim F2AnimTileSet As Byte

    ' Only used if ever want to switch to blt rather then bltfast
    With rec_pos
        .Top = y * PIC_Y
        .Bottom = .Top + PIC_Y
        .Left = x * PIC_X
        .Right = .Left + PIC_X
    End With
    Fringe = Map(GetPlayerMap(MyIndex)).Tile(x, y).Fringe
    FAnim = Map(GetPlayerMap(MyIndex)).Tile(x, y).FAnim
    Fringe2 = Map(GetPlayerMap(MyIndex)).Tile(x, y).Fringe2
    F2Anim = Map(GetPlayerMap(MyIndex)).Tile(x, y).F2Anim
    
    FringeTileSet = Map(GetPlayerMap(MyIndex)).Tile(x, y).FringeSet
    FAnimTileSet = Map(GetPlayerMap(MyIndex)).Tile(x, y).FAnimSet
    Fringe2TileSet = Map(GetPlayerMap(MyIndex)).Tile(x, y).Fringe2Set
    F2AnimTileSet = Map(GetPlayerMap(MyIndex)).Tile(x, y).F2AnimSet
        
    If (MapAnim = 0) Or (FAnim <= 0) Then
        ' Is there an animation tile to plot?
        
        If Fringe > 0 Then
        If TileFile(FringeTileSet) = 0 Then Exit Sub
        rec.Top = Int(Fringe / TilesInSheets) * PIC_Y
        rec.Bottom = rec.Top + PIC_Y
        rec.Left = (Fringe - Int(Fringe / TilesInSheets) * TilesInSheets) * PIC_X
        rec.Right = rec.Left + PIC_X
        'Call DD_BackBuffer.Blt(rec_pos, DD_TileSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
        Call DD_BackBuffer.BltFast((x - NewPlayerX) * PIC_X + sx - NewXOffset, (y - NewPlayerY) * PIC_Y + sx - NewYOffset, DD_TileSurf(FringeTileSet), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    
    Else
    
        If FAnim > 0 Then
        If TileFile(FAnimTileSet) = 0 Then Exit Sub
        rec.Top = Int(FAnim / TilesInSheets) * PIC_Y
        rec.Bottom = rec.Top + PIC_Y
        rec.Left = (FAnim - Int(FAnim / TilesInSheets) * TilesInSheets) * PIC_X
        rec.Right = rec.Left + PIC_X
        'Call DD_BackBuffer.Blt(rec_pos, DD_TileSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
        Call DD_BackBuffer.BltFast((x - NewPlayerX) * PIC_X + sx - NewXOffset, (y - NewPlayerY) * PIC_Y + sx - NewYOffset, DD_TileSurf(FAnimTileSet), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
        
    End If

    If (MapAnim = 0) Or (F2Anim <= 0) Then
        ' Is there an animation tile to plot?
        
        If Fringe2 > 0 Then
        If TileFile(Fringe2TileSet) = 0 Then Exit Sub
        rec.Top = Int(Fringe2 / TilesInSheets) * PIC_Y
        rec.Bottom = rec.Top + PIC_Y
        rec.Left = (Fringe2 - Int(Fringe2 / TilesInSheets) * TilesInSheets) * PIC_X
        rec.Right = rec.Left + PIC_X
        'Call DD_BackBuffer.Blt(rec_pos, DD_TileSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
        Call DD_BackBuffer.BltFast((x - NewPlayerX) * PIC_X + sx - NewXOffset, (y - NewPlayerY) * PIC_Y + sx - NewYOffset, DD_TileSurf(Fringe2TileSet), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    Else
        If F2Anim > 0 Then
        If TileFile(F2AnimTileSet) = 0 Then Exit Sub
        rec.Top = Int(F2Anim / TilesInSheets) * PIC_Y
        rec.Bottom = rec.Top + PIC_Y
        rec.Left = (F2Anim - Int(F2Anim / TilesInSheets) * TilesInSheets) * PIC_X
        rec.Right = rec.Left + PIC_X
        'Call DD_BackBuffer.Blt(rec_pos, DD_TileSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
        Call DD_BackBuffer.BltFast((x - NewPlayerX) * PIC_X + sx - NewXOffset, (y - NewPlayerY) * PIC_Y + sx - NewYOffset, DD_TileSurf(F2AnimTileSet), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    End If
End Sub

Sub BltPlayer(ByVal index As Long)
Dim Anim As Byte
Dim x As Long, y As Long
Dim AttackSpeed As Long

    If GetPlayerWeaponSlot(index) > 0 Then
        AttackSpeed = Item(GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index))).AttackSpeed
    Else
        AttackSpeed = 1000
    End If

    ' Only used if ever want to switch to blt rather then bltfast
    With rec_pos
        .Top = GetPlayerY(index) * PIC_Y + Player(index).YOffset
        .Bottom = .Top + PIC_Y
        .Left = GetPlayerX(index) * PIC_X + Player(index).XOffset
        .Right = .Left + PIC_X
    End With
    
    ' Check for animation
    Anim = 0
    If Player(index).Attacking = 0 Then
        Select Case GetPlayerDir(index)
            Case DIR_UP
                If (Player(index).YOffset < PIC_Y / 2) Then Anim = 1
            Case DIR_DOWN
                If (Player(index).YOffset < PIC_Y / 2 * -1) Then Anim = 1
            Case DIR_LEFT
                If (Player(index).XOffset < PIC_Y / 2) Then Anim = 1
            Case DIR_RIGHT
                If (Player(index).XOffset < PIC_Y / 2 * -1) Then Anim = 1
        End Select
    Else
        If Player(index).AttackTimer + Int(AttackSpeed / 2) > GetTickCount Then
            Anim = 2
        End If
    End If
    
    ' Check to see if we want to stop making him attack
    If Player(index).AttackTimer + AttackSpeed < GetTickCount Then
        Player(index).Attacking = 0
        Player(index).AttackTimer = 0
    End If
    
    rec.Top = GetPlayerSprite(index) * PIC_Y
    rec.Bottom = rec.Top + PIC_Y
    rec.Left = (GetPlayerDir(index) * 3 + Anim) * PIC_X
    rec.Right = rec.Left + PIC_X

If index = MyIndex Then
    x = NewX + sx
    y = NewY + sx
        
    Call DD_BackBuffer.BltFast(x, y, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
Else
    x = GetPlayerX(index) * PIC_X + sx + Player(index).XOffset
    y = GetPlayerY(index) * PIC_Y + sx + Player(index).YOffset '- 4
    
    If y < 0 Then
        y = 0
        rec.Top = rec.Top + (y * -1)
    End If
    
    Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End If
End Sub

Sub BltMapNPCName(ByVal index As Long)
Dim TextX As Long
Dim TextY As Long

If Npc(MapNpc(index).num).Big = 0 Then
    With Npc(MapNpc(index).num)
    'Draw name
        TextX = MapNpc(index).x * PIC_X + sx + MapNpc(index).XOffset + CLng(PIC_X / 2) - ((Len(Trim$(.name)) / 2) * 8)
        TextY = MapNpc(index).y * PIC_Y + sx + MapNpc(index).YOffset - CLng(PIC_Y / 2) - 4
        DrawPlayerNameText TexthDC, TextX - (NewPlayerX * PIC_X) - NewXOffset, TextY - (NewPlayerY * PIC_Y) - NewYOffset, Trim$(.name), vbWhite
    End With
Else
    With Npc(MapNpc(index).num)
    'Draw name
        TextX = MapNpc(index).x * PIC_X + sx + MapNpc(index).XOffset + CLng(PIC_X / 2) - ((Len(Trim$(.name)) / 2) * 8)
        TextY = MapNpc(index).y * PIC_Y + sx + MapNpc(index).YOffset - CLng(PIC_Y / 2) - 32
        DrawPlayerNameText TexthDC, TextX - (NewPlayerX * PIC_X) - NewXOffset, TextY - (NewPlayerY * PIC_Y) - NewYOffset, Trim$(.name), vbWhite
    End With
End If
End Sub

Sub BltNpc(ByVal MapNpcNum As Long)
Dim Anim As Byte
Dim x As Long, y As Long

    ' Make sure that theres an npc there, and if not exit the sub
    If MapNpc(MapNpcNum).num <= 0 Then
        Exit Sub
    End If
    
    ' Only used if ever want to switch to blt rather then bltfast
    With rec_pos
        .Top = MapNpc(MapNpcNum).y * PIC_Y + MapNpc(MapNpcNum).YOffset
        .Bottom = .Top + PIC_Y
        .Left = MapNpc(MapNpcNum).x * PIC_X + MapNpc(MapNpcNum).XOffset
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
        
    If Npc(MapNpc(MapNpcNum).num).Big = 0 Then
        rec.Top = Npc(MapNpc(MapNpcNum).num).Sprite * PIC_Y
        rec.Bottom = rec.Top + PIC_Y
        rec.Left = (MapNpc(MapNpcNum).Dir * 3 + Anim) * PIC_X
        rec.Right = rec.Left + PIC_X
        
        x = MapNpc(MapNpcNum).x * PIC_X + sx + MapNpc(MapNpcNum).XOffset
        y = MapNpc(MapNpcNum).y * PIC_Y + sx + MapNpc(MapNpcNum).YOffset
        
        ' Check if its out of bounds because of the offset
        If y < 0 Then
            y = 0
            rec.Top = rec.Top + (y * -1)
        End If
            
        'Call DD_BackBuffer.Blt(rec_pos, DD_SpriteSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
        Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    Else
        rec.Top = Npc(MapNpc(MapNpcNum).num).Sprite * 64 + 32
        rec.Bottom = rec.Top + 32
        rec.Left = (MapNpc(MapNpcNum).Dir * 3 + Anim) * 64
        rec.Right = rec.Left + 64
    
        x = MapNpc(MapNpcNum).x * 32 + sx - 16 + MapNpc(MapNpcNum).XOffset
        y = MapNpc(MapNpcNum).y * 32 + sx + MapNpc(MapNpcNum).YOffset
   
        If y < 0 Then
            rec.Top = Npc(MapNpc(MapNpcNum).num).Sprite * 64 + 32
            rec.Bottom = rec.Top + 32
            y = MapNpc(MapNpcNum).YOffset + sx
        End If
        
        If x < 0 Then
            rec.Left = (MapNpc(MapNpcNum).Dir * 3 + Anim) * 64 + 16
            rec.Right = rec.Left + 48
            x = MapNpc(MapNpcNum).XOffset + sx
        End If
        
        If x > MAX_MAPX * 32 Then
            rec.Left = (MapNpc(MapNpcNum).Dir * 3 + Anim) * 64
            rec.Right = rec.Left + 48
            x = MAX_MAPX * 32 + sx - 16 + MapNpc(MapNpcNum).XOffset
        End If

        Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_BigSpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    End If
End Sub

Sub BltNpcTop(ByVal MapNpcNum As Long)
Dim Anim As Byte
Dim x As Long, y As Long

    ' Make sure that theres an npc there, and if not exit the sub
    If MapNpc(MapNpcNum).num <= 0 Then
        Exit Sub
    End If
    
    If Npc(MapNpc(MapNpcNum).num).Big = 0 Then Exit Sub
    
    ' Only used if ever want to switch to blt rather then bltfast
    With rec_pos
        .Top = MapNpc(MapNpcNum).y * PIC_Y + MapNpc(MapNpcNum).YOffset
        .Bottom = .Top + PIC_Y
        .Left = MapNpc(MapNpcNum).x * PIC_X + MapNpc(MapNpcNum).XOffset
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
    
    rec.Top = Npc(MapNpc(MapNpcNum).num).Sprite * PIC_Y
        
     rec.Top = Npc(MapNpc(MapNpcNum).num).Sprite * 64
     rec.Bottom = rec.Top + 32
     rec.Left = (MapNpc(MapNpcNum).Dir * 3 + Anim) * 64
     rec.Right = rec.Left + 64
 
     x = MapNpc(MapNpcNum).x * 32 + sx - 16 + MapNpc(MapNpcNum).XOffset
     y = MapNpc(MapNpcNum).y * 32 + sx - 32 + MapNpc(MapNpcNum).YOffset

     If y < 0 Then
         rec.Top = Npc(MapNpc(MapNpcNum).num).Sprite * 64 + 32
         rec.Bottom = rec.Top
         y = MapNpc(MapNpcNum).YOffset + sx
     End If
     
     If x < 0 Then
         rec.Left = (MapNpc(MapNpcNum).Dir * 3 + Anim) * 64 + 16
         rec.Right = rec.Left + 48
         x = MapNpc(MapNpcNum).XOffset + sx
     End If
     
     If x > MAX_MAPX * 32 Then
         rec.Left = (MapNpc(MapNpcNum).Dir * 3 + Anim) * 64
         rec.Right = rec.Left + 48
         x = MAX_MAPX * 32 + sx - 16 + MapNpc(MapNpcNum).XOffset
     End If

     Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_BigSpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End Sub

Sub BltPlayerName(ByVal index As Long)
Dim TextX As Long
Dim TextY As Long
Dim Color As Long
    
    ' Check access level
    If GetPlayerPK(index) = NO Then
        Select Case GetPlayerAccess(index)
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
        
    
If index = MyIndex Then
    TextX = NewX + sx + Int(PIC_X / 2) - ((Len(GetPlayerName(MyIndex)) / 2) * 8)
    TextY = NewY + sx - Int(PIC_Y / 2)
    
    Call DrawText(TexthDC, TextX, TextY, GetPlayerName(MyIndex), Color)
Else
    ' Draw name
    TextX = GetPlayerX(index) * PIC_X + sx + Player(index).XOffset + Int(PIC_X / 2) - ((Len(GetPlayerName(index)) / 2) * 8)
    TextY = GetPlayerY(index) * PIC_Y + sx + Player(index).YOffset - Int(PIC_Y / 2)
    Call DrawText(TexthDC, TextX - (NewPlayerX * PIC_X) - NewXOffset, TextY - (NewPlayerY * PIC_Y) - NewYOffset, GetPlayerName(index), Color)
End If
End Sub

Sub BltPlayerGuildName(ByVal index As Long)
Dim TextX As Long
Dim TextY As Long
Dim Color As Long

    ' Check access level
    If GetPlayerPK(index) = NO Then
        Select Case GetPlayerGuildAccess(index)
            Case 0
                If GetPlayerSTR(index) > 0 Then
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

If index = MyIndex Then
    TextX = NewX + sx + Int(PIC_X / 2) - ((Len(GetPlayerGuild(MyIndex)) / 2) * 8)
    TextY = NewY + sx - Int(PIC_Y / 4) - 20
    
    Call DrawText(TexthDC, TextX, TextY, GetPlayerGuild(MyIndex), Color)
Else
    ' Draw name
    TextX = GetPlayerX(index) * PIC_X + sx + Player(index).XOffset + Int(PIC_X / 2) - ((Len(GetPlayerGuild(index)) / 2) * 8)
    TextY = GetPlayerY(index) * PIC_Y + sx + Player(index).YOffset - Int(PIC_Y / 2) - 12
    Call DrawText(TexthDC, TextX - (NewPlayerX * PIC_X) - NewXOffset, TextY - (NewPlayerY * PIC_Y) - NewYOffset, GetPlayerGuild(index), Color)
End If
End Sub

Sub ProcessMovement(ByVal index As Long)
    ' Check if player is walking, and if so process moving them over
If Player(index).Moving = MOVING_WALKING Then
        If Player(index).Access > 0 Then
            Select Case GetPlayerDir(index)
                Case DIR_UP
                    Player(index).YOffset = Player(index).YOffset - GM_WALK_SPEED
                Case DIR_DOWN
                    Player(index).YOffset = Player(index).YOffset + GM_WALK_SPEED
                Case DIR_LEFT
                    Player(index).XOffset = Player(index).XOffset - GM_WALK_SPEED
                Case DIR_RIGHT
                    Player(index).XOffset = Player(index).XOffset + GM_WALK_SPEED
            End Select
        Else
            Select Case GetPlayerDir(index)
                Case DIR_UP
                    Player(index).YOffset = Player(index).YOffset - WALK_SPEED
                Case DIR_DOWN
                    Player(index).YOffset = Player(index).YOffset + WALK_SPEED
                Case DIR_LEFT
                    Player(index).XOffset = Player(index).XOffset - WALK_SPEED
                Case DIR_RIGHT
                    Player(index).XOffset = Player(index).XOffset + WALK_SPEED
            End Select
        End If
        
        ' Check if completed walking over to the next tile
        If (Player(index).XOffset = 0) And (Player(index).YOffset = 0) Then
            Player(index).Moving = 0
        End If
    End If

    ' Check if player is running, and if so process moving them over
If Player(index).Moving = MOVING_RUNNING Then
            If Player(index).Access > 0 Then
            Select Case GetPlayerDir(index)
                Case DIR_UP
                    Player(index).YOffset = Player(index).YOffset - GM_RUN_SPEED
                Case DIR_DOWN
                    Player(index).YOffset = Player(index).YOffset + GM_RUN_SPEED
                Case DIR_LEFT
                    Player(index).XOffset = Player(index).XOffset - GM_RUN_SPEED
                Case DIR_RIGHT
                    Player(index).XOffset = Player(index).XOffset + GM_RUN_SPEED
            End Select
        Else
            Select Case GetPlayerDir(index)
                Case DIR_UP
                    Player(index).YOffset = Player(index).YOffset - RUN_SPEED
                Case DIR_DOWN
                    Player(index).YOffset = Player(index).YOffset + RUN_SPEED
                Case DIR_LEFT
                    Player(index).XOffset = Player(index).XOffset - RUN_SPEED
                Case DIR_RIGHT
                    Player(index).XOffset = Player(index).XOffset + RUN_SPEED
            End Select
        End If
        
        ' Check if completed walking over to the next tile
        If (Player(index).XOffset = 0) And (Player(index).YOffset = 0) Then
            Player(index).Moving = 0
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
Dim name As String
Dim i As Long
Dim n As Long

MyText = frmMirage.txtMyTextBox.Text

    ' Handle when the player presses the return key
    If (KeyAscii = vbKeyReturn) Then
    frmMirage.txtMyTextBox.Text = ""
        If Player(MyIndex).y - 1 > -1 Then
            If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).Type = TILE_TYPE_SIGN And Player(MyIndex).Dir = DIR_UP Then
                Call AddText("The Sign Reads:", Black)
                If Trim(Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).String1) <> "" Then
                    Call AddText(Trim(Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).String1), Grey)
                End If
                If Trim(Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).String2) <> "" Then
                    Call AddText(Trim(Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).String2), Grey)
                End If
                If Trim(Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).String3) <> "" Then
                    Call AddText(Trim(Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).String3), Grey)
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
            name = ""
                    
            ' Get the desired player from the user text
            For i = 1 To Len(ChatText)
                If Mid(ChatText, i, 1) <> " " Then
                    name = name & Mid(ChatText, i, 1)
                Else
                    Exit For
                End If
            Next i
                    
            ' Make sure they are actually sending something
            If Len(ChatText) - i > 0 Then
                ChatText = Mid(ChatText, i + 1, Len(ChatText) - i)
                    
                ' Send the message to the player
                Call PlayerMsg(ChatText, name)
            Else
                Call AddText("Usage: !playername msghere", AlertColor)
            End If
            MyText = ""
            Exit Sub
        End If
            
        ' // Commands //
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
            ' day night command
            If LCase(Mid(MyText, 1, 9)) = "/daynight" Then
                If GameTime = TIME_DAY Then
                    GameTime = TIME_NIGHT
                Else
                    GameTime = TIME_DAY
                End If
                Call SendGameTime
                MyText = ""
                Exit Sub
            End If
            
            ' weather command
            If LCase(Mid(MyText, 1, 8)) = "/weather" Then
                If Len(MyText) > 8 Then
                    MyText = Mid(MyText, 9, Len(MyText) - 8)
                    If IsNumeric(MyText) = True Then
                        Call SendData("weather" & SEP_CHAR & Val(MyText) & SEP_CHAR & END_CHAR)
                    Else
                        If Trim(LCase(MyText)) = "none" Then i = 0
                        If Trim(LCase(MyText)) = "rain" Then i = 1
                        If Trim(LCase(MyText)) = "snow" Then i = 2
                        If Trim(LCase(MyText)) = "thunder" Then i = 3
                        Call SendData("weather" & SEP_CHAR & i & SEP_CHAR & END_CHAR)
                    End If
                End If
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
            
            ' Editing arrow request
            If Mid(MyText, 1, 13) = "/editarrow" Then
                Call SendRequestEditArrow
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
    
    'frmMirage.txtMyTextBox.SetFocus
    
    ' Handle when the user presses the backspace key
    If (KeyAscii = vbKeyBack) Then
        If Len(MyText) > 0 Then
            'MyText = Mid(MyText, 1, Len(MyText) - 1)
        End If
    End If
    
    ' And if neither, then add the character to the user's text buffer
    If (KeyAscii <> vbKeyReturn) And (KeyAscii <> vbKeyBack) Then
        ' Make sure they just use standard keys, no gay shitty macro keys
        If KeyAscii >= 32 And KeyAscii <= 126 Then
            'frmMirage.txtMyTextBox.Text = frmMirage.txtMyTextBox.Text & Chr(KeyAscii)
            'MyText = MyText & Chr(KeyAscii)
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
Dim AttackSpeed As Long
    If GetPlayerWeaponSlot(MyIndex) > 0 Then
        AttackSpeed = Item(GetPlayerInvItemNum(MyIndex, GetPlayerWeaponSlot(MyIndex))).AttackSpeed
    Else
        AttackSpeed = 1000
    End If
    
    If ControlDown = True And Player(MyIndex).AttackTimer + AttackSpeed < GetTickCount And Player(MyIndex).Attacking = 0 Then
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
Dim x As Long, y As Long

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
            If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).Type = TILE_TYPE_BLOCKED Or Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).Type = TILE_TYPE_SIGN Then
                CanMove = False
                
                ' Set the new direction if they weren't facing that direction
                If d <> DIR_UP Then
                    Call SendPlayerDir
                End If
                Exit Function
            End If
            
            If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).Type = TILE_TYPE_CBLOCK Then
                If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).Data1 = Player(MyIndex).Class Then Exit Function
                If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).Data2 = Player(MyIndex).Class Then Exit Function
                If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).Data3 = Player(MyIndex).Class Then Exit Function
                CanMove = False
                
                ' Set the new direction if they weren't facing that direction
                If d <> DIR_DOWN Then
                    Call SendPlayerDir
                End If
            End If
                                                    
            ' Check to see if the key door is open or not
            If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).Type = TILE_TYPE_KEY Or Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).Type = TILE_TYPE_DOOR Then
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
                If MapNpc(i).num > 0 Then
                    If (MapNpc(i).x = GetPlayerX(MyIndex)) And (MapNpc(i).y = GetPlayerY(MyIndex) - 1) Then
                        CanMove = False
                        
                        ' Set the new direction if they weren't facing that direction
                        If d <> DIR_UP Then
                            Call SendPlayerDir
                        End If
                        Exit Function
                    End If
                End If
            Next i
            
            
            If CanAttributeNPCMove(DIR_UP) = False Then
                CanMove = False
                        
                ' Set the new direction if they weren't facing that direction
                If d <> DIR_UP Then
                    Call SendPlayerDir
                End If
                Exit Function
            End If
        Else
            ' Check if they can warp to a new map
            If Map(GetPlayerMap(MyIndex)).Up > 0 Then
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
            If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) + 1).Type = TILE_TYPE_BLOCKED Or Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) + 1).Type = TILE_TYPE_SIGN Then
                CanMove = False
                
                ' Set the new direction if they weren't facing that direction
                If d <> DIR_DOWN Then
                    Call SendPlayerDir
                End If
                Exit Function
            End If
                        
            If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) + 1).Type = TILE_TYPE_CBLOCK Then
                If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) + 1).Data1 = Player(MyIndex).Class Then Exit Function
                If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) + 1).Data2 = Player(MyIndex).Class Then Exit Function
                If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) + 1).Data3 = Player(MyIndex).Class Then Exit Function
                CanMove = False
                
                ' Set the new direction if they weren't facing that direction
                If d <> DIR_DOWN Then
                    Call SendPlayerDir
                End If
                Exit Function
            End If
                                        
            ' Check to see if the key door is open or not
            If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) + 1).Type = TILE_TYPE_KEY Or Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) + 1).Type = TILE_TYPE_DOOR Then
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
                If MapNpc(i).num > 0 Then
                    If (MapNpc(i).x = GetPlayerX(MyIndex)) And (MapNpc(i).y = GetPlayerY(MyIndex) + 1) Then
                        CanMove = False
                        
                        ' Set the new direction if they weren't facing that direction
                        If d <> DIR_DOWN Then
                            Call SendPlayerDir
                        End If
                        Exit Function
                    End If
                End If
            Next i
            
            If CanAttributeNPCMove(DIR_DOWN) = False Then
                CanMove = False
                        
                ' Set the new direction if they weren't facing that direction
                If d <> DIR_DOWN Then
                    Call SendPlayerDir
                End If
                Exit Function
            End If
        Else
            ' Check if they can warp to a new map
            If Map(GetPlayerMap(MyIndex)).Down > 0 Then
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
            If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) - 1, GetPlayerY(MyIndex)).Type = TILE_TYPE_BLOCKED Or Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) - 1, GetPlayerY(MyIndex)).Type = TILE_TYPE_SIGN Then
                CanMove = False
                
                ' Set the new direction if they weren't facing that direction
                If d <> DIR_LEFT Then
                    Call SendPlayerDir
                End If
                Exit Function
            End If
                        
            If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) - 1, GetPlayerY(MyIndex)).Type = TILE_TYPE_CBLOCK Then
                If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) - 1, GetPlayerY(MyIndex)).Data1 = Player(MyIndex).Class Then Exit Function
                If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) - 1, GetPlayerY(MyIndex)).Data2 = Player(MyIndex).Class Then Exit Function
                If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) - 1, GetPlayerY(MyIndex)).Data3 = Player(MyIndex).Class Then Exit Function
                CanMove = False
                
                ' Set the new direction if they weren't facing that direction
                If d <> DIR_LEFT Then
                    Call SendPlayerDir
                End If
                Exit Function
            End If
                                        
            ' Check to see if the key door is open or not
            If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) - 1, GetPlayerY(MyIndex)).Type = TILE_TYPE_KEY Or Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) - 1, GetPlayerY(MyIndex)).Type = TILE_TYPE_DOOR Then
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
                If MapNpc(i).num > 0 Then
                    If (MapNpc(i).x = GetPlayerX(MyIndex) - 1) And (MapNpc(i).y = GetPlayerY(MyIndex)) Then
                        CanMove = False
                        
                        ' Set the new direction if they weren't facing that direction
                        If d <> DIR_LEFT Then
                            Call SendPlayerDir
                        End If
                        Exit Function
                    End If
                End If
            Next i
            
            If CanAttributeNPCMove(DIR_LEFT) = False Then
                CanMove = False
                        
                ' Set the new direction if they weren't facing that direction
                If d <> DIR_LEFT Then
                    Call SendPlayerDir
                End If
                Exit Function
            End If
        Else
            ' Check if they can warp to a new map
            If Map(GetPlayerMap(MyIndex)).Left > 0 Then
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
            If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + 1, GetPlayerY(MyIndex)).Type = TILE_TYPE_BLOCKED Or Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + 1, GetPlayerY(MyIndex)).Type = TILE_TYPE_SIGN Then
                CanMove = False
                
                ' Set the new direction if they weren't facing that direction
                If d <> DIR_RIGHT Then
                    Call SendPlayerDir
                End If
                Exit Function
            End If
                        
            If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + 1, GetPlayerY(MyIndex)).Type = TILE_TYPE_CBLOCK Then
                If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + 1, GetPlayerY(MyIndex)).Data1 = Player(MyIndex).Class Then Exit Function
                If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + 1, GetPlayerY(MyIndex)).Data2 = Player(MyIndex).Class Then Exit Function
                If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + 1, GetPlayerY(MyIndex)).Data3 = Player(MyIndex).Class Then Exit Function
                CanMove = False
                
                ' Set the new direction if they weren't facing that direction
                If d <> DIR_RIGHT Then
                    Call SendPlayerDir
                End If
                Exit Function
            End If
                                        
            ' Check to see if the key door is open or not
            If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + 1, GetPlayerY(MyIndex)).Type = TILE_TYPE_KEY Or Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + 1, GetPlayerY(MyIndex)).Type = TILE_TYPE_DOOR Then
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
                If MapNpc(i).num > 0 Then
                    If (MapNpc(i).x = GetPlayerX(MyIndex) + 1) And (MapNpc(i).y = GetPlayerY(MyIndex)) Then
                        CanMove = False
                        
                        ' Set the new direction if they weren't facing that direction
                        If d <> DIR_RIGHT Then
                            Call SendPlayerDir
                        End If
                        Exit Function
                    End If
                End If
            Next i
            
            If CanAttributeNPCMove(DIR_RIGHT) = False Then
                CanMove = False
                        
                ' Set the new direction if they weren't facing that direction
                If d <> DIR_RIGHT Then
                    Call SendPlayerDir
                End If
                Exit Function
            End If
        Else
            ' Check if they can warp to a new map
            If Map(GetPlayerMap(MyIndex)).Right > 0 Then
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
                If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex)).Type = TILE_TYPE_WARP Then
                    GettingMap = True
                End If
            End If
        End If
    End If
End Sub

Function FindPlayer(ByVal name As String) As Long
Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            ' Make sure we dont try to check a name thats to small
            If Len(GetPlayerName(i)) >= Len(Trim(name)) Then
                If UCase(Mid(GetPlayerName(i), 1, Len(Trim(name)))) = UCase(Trim(name)) Then
                    FindPlayer = i
                    Exit Function
                End If
            End If
        End If
    Next i
    
    FindPlayer = 0
End Function

Public Sub EditorInit()
    InEditor = True
    frmAttributes.Show vbModeless, frmMirage
    frmMapEditor.Show vbModeless, frmMirage
    EditorSet = 0

    frmMapEditor.picBackSelect.Picture = LoadPicture(App.Path + "\GFX\tiles0.bmp")
    frmMapEditor.scrlPicture.max = Int((frmMapEditor.picBackSelect.Height - frmMapEditor.picBack.Height) / PIC_Y)
    frmMapEditor.picBack.Width = frmMapEditor.picBackSelect.Width
End Sub

Public Sub EditorMouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim x1, y1 As Long
Dim x2 As Long, y2 As Long, PicX As Long

    If InEditor Then
        x1 = Int(x / PIC_X)
        y1 = Int(y / PIC_Y)
        
        If frmMapEditor.MousePointer = 2 Then
            If frmMapEditor.mnuType(1).Checked = True Then
                With Map(GetPlayerMap(MyIndex)).Tile(x1, y1)
                    If frmAttributes.optGround.value = True Then
                        PicX = .Ground
                        EditorSet = .GroundSet
                    End If
                    If frmAttributes.optMask.value = True Then
                        PicX = .Mask
                        EditorSet = .MaskSet
                    End If
                    If frmAttributes.optAnim.value = True Then
                        PicX = .Anim
                        EditorSet = .AnimSet
                    End If
                    If frmAttributes.optMask2.value = True Then
                        PicX = .Mask2
                        EditorSet = .Mask2Set
                    End If
                    If frmAttributes.optM2Anim.value = True Then
                        PicX = .M2Anim
                        EditorSet = .M2AnimSet
                    End If
                    If frmAttributes.optFringe.value = True Then
                        PicX = .Fringe
                        EditorSet = .FringeSet
                    End If
                    If frmAttributes.optFAnim.value = True Then
                        PicX = .FAnim
                        EditorSet = .FAnimSet
                    End If
                    If frmAttributes.optFringe2.value = True Then
                        PicX = .Fringe2
                        EditorSet = .Fringe2Set
                    End If
                    If frmAttributes.optF2Anim.value = True Then
                        PicX = .F2Anim
                        EditorSet = .F2AnimSet
                    End If
                    
                    EditorTileY = Int(PicX / TilesInSheets)
                    EditorTileX = (PicX - Int(PicX / TilesInSheets) * TilesInSheets)
                    frmMapEditor.shpSelected.Top = Int(EditorTileY * PIC_Y)
                    frmMapEditor.shpSelected.Left = Int(EditorTileX * PIC_Y)
                    frmMapEditor.shpSelected.Height = PIC_Y
                    frmMapEditor.shpSelected.Width = PIC_X
                End With
            ElseIf frmMapEditor.mnuType(3).Checked = True Then
                EditorTileY = Int(Map(GetPlayerMap(MyIndex)).Tile(x1, y1).Light / TilesInSheets)
                EditorTileX = (Map(GetPlayerMap(MyIndex)).Tile(x1, y1).Light - Int(Map(GetPlayerMap(MyIndex)).Tile(x1, y1).Light / TilesInSheets) * TilesInSheets)
                frmMapEditor.shpSelected.Top = Int(EditorTileY * PIC_Y)
                frmMapEditor.shpSelected.Left = Int(EditorTileX * PIC_Y)
                frmMapEditor.shpSelected.Height = PIC_Y
                frmMapEditor.shpSelected.Width = PIC_X
            ElseIf frmMapEditor.mnuType(2).Checked = True Then
                With Map(GetPlayerMap(MyIndex)).Tile(x1, y1)
                    If .Type = TILE_TYPE_BLOCKED Then frmAttributes.optBlocked.value = True
                    If .Type = TILE_TYPE_WARP Then
                        EditorWarpMap = .Data1
                        EditorWarpX = .Data2
                        EditorWarpY = .Data3
                        frmAttributes.optWarp.value = True
                    End If
                    If .Type = TILE_TYPE_HEAL Then frmAttributes.optHeal.value = True
                    If .Type = TILE_TYPE_KILL Then frmAttributes.optKill.value = True
                    If .Type = TILE_TYPE_ITEM Then
                        ItemEditorNum = .Data1
                        ItemEditorValue = .Data2
                        frmAttributes.optItem.value = True
                    End If
                    If .Type = TILE_TYPE_NPCAVOID Then frmAttributes.optNpcAvoid.value = True
                    If .Type = TILE_TYPE_KEY Then
                        KeyEditorNum = .Data1
                        KeyEditorTake = .Data2
                        frmAttributes.optKey.value = True
                    End If
                    If .Type = TILE_TYPE_KEYOPEN Then
                        KeyOpenEditorX = .Data1
                        KeyOpenEditorY = .Data2
                        KeyOpenEditorMsg = .String1
                        frmAttributes.optKeyOpen.value = True
                    End If
                    If .Type = TILE_TYPE_SHOP Then
                        EditorShopNum = .Data1
                        frmAttributes.optShop.value = True
                    End If
                    If .Type = TILE_TYPE_CBLOCK Then
                        EditorItemNum1 = .Data1
                        EditorItemNum2 = .Data2
                        EditorItemNum3 = .Data3
                        frmAttributes.optCBlock.value = True
                    End If
                    If .Type = TILE_TYPE_ARENA Then
                        Arena1 = .Data1
                        Arena2 = .Data2
                        Arena3 = .Data3
                        frmAttributes.optArena.value = True
                    End If
                    If .Type = TILE_TYPE_SOUND Then
                        SoundFileName = .String1
                        frmAttributes.optSound.value = True
                    End If
                    If .Type = TILE_TYPE_SPRITE_CHANGE Then
                        SpritePic = .Data1
                        SpriteItem = .Data2
                        SpritePrice = .Data3
                        frmAttributes.optSprite.value = True
                    End If
                    If .Type = TILE_TYPE_SIGN Then
                        SignLine1 = .String1
                        SignLine2 = .String2
                        SignLine3 = .String3
                        frmAttributes.optSign.value = True
                    End If
                    If .Type = TILE_TYPE_DOOR Then frmAttributes.optDoor.value = True
                    If .Type = TILE_TYPE_NOTICE Then
                        NoticeTitle = .String1
                        NoticeText = .String2
                        NoticeSound = .String3
                        frmAttributes.optNotice.value = True
                    End If
                    If .Type = TILE_TYPE_CHEST Then frmAttributes.optChest.value = True
                    If .Type = TILE_TYPE_CLASS_CHANGE Then
                        ClassChange = .Data1
                        ClassChangeReq = .Data2
                        frmAttributes.optClassChange.value = True
                    End If
                    If .Type = TILE_TYPE_SCRIPTED Then
                        ScriptNum = .Data1
                        frmAttributes.optScripted.value = True
                    End If
                    If .Type = TILE_TYPE_NPC_SPAWN Then
                        NPCSpawnNum = .Data1
                        NPCSpawnAmount = .Data2
                        NPCSpawnRange = .Data3
                        frmAttributes.optNPC.value = True
                    End If
                End With
            End If
            frmMapEditor.MousePointer = 1
            frmMirage.MousePointer = 1
        Else
            If (Button = 1) And (x1 >= 0) And (x1 <= MAX_MAPX) And (y1 >= 0) And (y1 <= MAX_MAPY) Then
                If frmMapEditor.shpSelected.Height <= PIC_Y And frmMapEditor.shpSelected.Width <= PIC_X Then
                    If frmMapEditor.mnuType(1).Checked = True Then
                        With Map(GetPlayerMap(MyIndex)).Tile(x1, y1)
                            If frmAttributes.optGround.value = True Then
                                .Ground = EditorTileY * TilesInSheets + EditorTileX
                                .GroundSet = EditorSet
                            End If
                            If frmAttributes.optMask.value = True Then
                                .Mask = EditorTileY * TilesInSheets + EditorTileX
                                .MaskSet = EditorSet
                            End If
                            If frmAttributes.optAnim.value = True Then
                                .Anim = EditorTileY * TilesInSheets + EditorTileX
                                .AnimSet = EditorSet
                            End If
                            If frmAttributes.optMask2.value = True Then
                                .Mask2 = EditorTileY * TilesInSheets + EditorTileX
                                .Mask2Set = EditorSet
                            End If
                            If frmAttributes.optM2Anim.value = True Then
                                .M2Anim = EditorTileY * TilesInSheets + EditorTileX
                                .M2AnimSet = EditorSet
                            End If
                            If frmAttributes.optFringe.value = True Then
                                .Fringe = EditorTileY * TilesInSheets + EditorTileX
                                .FringeSet = EditorSet
                            End If
                            If frmAttributes.optFAnim.value = True Then
                                .FAnim = EditorTileY * TilesInSheets + EditorTileX
                                .FAnimSet = EditorSet
                            End If
                            If frmAttributes.optFringe2.value = True Then
                                .Fringe2 = EditorTileY * TilesInSheets + EditorTileX
                                .Fringe2Set = EditorSet
                            End If
                            If frmAttributes.optF2Anim.value = True Then
                                .F2Anim = EditorTileY * TilesInSheets + EditorTileX
                                .F2AnimSet = EditorSet
                            End If
                        End With
                    ElseIf frmMapEditor.mnuType(3).Checked = True Then
                        Map(GetPlayerMap(MyIndex)).Tile(x1, y1).Light = EditorTileY * TilesInSheets + EditorTileX
                    ElseIf frmMapEditor.mnuType(2).Checked = True Then
                        With Map(GetPlayerMap(MyIndex)).Tile(x1, y1)
                            If frmAttributes.optBlocked.value = True Then .Type = TILE_TYPE_BLOCKED
                            If frmAttributes.optWarp.value = True Then
                                .Type = TILE_TYPE_WARP
                                .Data1 = EditorWarpMap
                                .Data2 = EditorWarpX
                                .Data3 = EditorWarpY
                                .String1 = ""
                                .String2 = ""
                                .String3 = ""
                            End If
        
                            If frmAttributes.optHeal.value = True Then
                                .Type = TILE_TYPE_HEAL
                                .Data1 = 0
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = ""
                                .String2 = ""
                                .String3 = ""
                            End If
        
                            If frmAttributes.optKill.value = True Then
                                .Type = TILE_TYPE_KILL
                                .Data1 = 0
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = ""
                                .String2 = ""
                                .String3 = ""
                            End If
        
                            If frmAttributes.optItem.value = True Then
                                .Type = TILE_TYPE_ITEM
                                .Data1 = ItemEditorNum
                                .Data2 = ItemEditorValue
                                .Data3 = 0
                                .String1 = ""
                                .String2 = ""
                                .String3 = ""
                            End If
                            If frmAttributes.optNpcAvoid.value = True Then
                                .Type = TILE_TYPE_NPCAVOID
                                .Data1 = 0
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = ""
                                .String2 = ""
                                .String3 = ""
                            End If
                            If frmAttributes.optKey.value = True Then
                                .Type = TILE_TYPE_KEY
                                .Data1 = KeyEditorNum
                                .Data2 = KeyEditorTake
                                .Data3 = 0
                                .String1 = ""
                                .String2 = ""
                                .String3 = ""
                            End If
                            If frmAttributes.optKeyOpen.value = True Then
                                .Type = TILE_TYPE_KEYOPEN
                                .Data1 = KeyOpenEditorX
                                .Data2 = KeyOpenEditorY
                                .Data3 = 0
                                .String1 = KeyOpenEditorMsg
                                .String2 = ""
                                .String3 = ""
                            End If
                            If frmAttributes.optShop.value = True Then
                                .Type = TILE_TYPE_SHOP
                                .Data1 = EditorShopNum
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = ""
                                .String2 = ""
                                .String3 = ""
                            End If
                            If frmAttributes.optCBlock.value = True Then
                                .Type = TILE_TYPE_CBLOCK
                                .Data1 = EditorItemNum1
                                .Data2 = EditorItemNum2
                                .Data3 = EditorItemNum3
                                .String1 = ""
                                .String2 = ""
                                .String3 = ""
                            End If
                            If frmAttributes.optArena.value = True Then
                                .Type = TILE_TYPE_ARENA
                                .Data1 = Arena1
                                .Data2 = Arena2
                                .Data3 = Arena3
                                .String1 = ""
                                .String2 = ""
                                .String3 = ""
                            End If
                            If frmAttributes.optSound.value = True Then
                                .Type = TILE_TYPE_SOUND
                                .Data1 = 0
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = SoundFileName
                                .String2 = ""
                                .String3 = ""
                            End If
                            If frmAttributes.optSprite.value = True Then
                                .Type = TILE_TYPE_SPRITE_CHANGE
                                .Data1 = SpritePic
                                .Data2 = SpriteItem
                                .Data3 = SpritePrice
                                .String1 = ""
                                .String2 = ""
                                .String3 = ""
                            End If
                            If frmAttributes.optSign.value = True Then
                                .Type = TILE_TYPE_SIGN
                                .Data1 = 0
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = SignLine1
                                .String2 = SignLine2
                                .String3 = SignLine3
                            End If
                            If frmAttributes.optDoor.value = True Then
                                .Type = TILE_TYPE_DOOR
                                .Data1 = 0
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = ""
                                .String2 = ""
                                .String3 = ""
                            End If
                            If frmAttributes.optNotice.value = True Then
                                .Type = TILE_TYPE_NOTICE
                                .Data1 = 0
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = NoticeTitle
                                .String2 = NoticeText
                                .String3 = NoticeSound
                            End If
                            If frmAttributes.optChest.value = True Then
                                .Type = TILE_TYPE_CHEST
                                .Data1 = 0
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = ""
                                .String2 = ""
                                .String3 = ""
                            End If
                            If frmAttributes.optClassChange.value = True Then
                                .Type = TILE_TYPE_CLASS_CHANGE
                                .Data1 = ClassChange
                                .Data2 = ClassChangeReq
                                .Data3 = 0
                                .String1 = ""
                                .String2 = ""
                                .String3 = ""
                            End If
                            If frmAttributes.optScripted.value = True Then
                                .Type = TILE_TYPE_SCRIPTED
                                .Data1 = ScriptNum
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = ""
                                .String2 = ""
                                .String3 = ""
                            End If
                            If frmAttributes.optNPC.value = True Then
                                .Type = TILE_TYPE_NPC_SPAWN
                                .Data1 = NPCSpawnNum
                                .Data2 = NPCSpawnAmount
                                .Data3 = NPCSpawnRange
                                .String1 = ""
                                .String2 = ""
                                .String3 = ""
                            End If
                        End With
                    End If
                Else
                    For y2 = 0 To Int(frmMapEditor.shpSelected.Height / PIC_Y) - 1
                        For x2 = 0 To Int(frmMapEditor.shpSelected.Width / PIC_X) - 1
                            If x1 + x2 <= MAX_MAPX Then
                                If y1 + y2 <= MAX_MAPY Then
                                    If frmMapEditor.mnuType(1).Checked = True Then
                                        With Map(GetPlayerMap(MyIndex)).Tile(x1 + x2, y1 + y2)
                                            If frmAttributes.optGround.value = True Then
                                                .Ground = (EditorTileY + y2) * TilesInSheets + (EditorTileX + x2)
                                                .GroundSet = EditorSet
                                            End If
                                            If frmAttributes.optMask.value = True Then
                                                .Mask = (EditorTileY + y2) * TilesInSheets + (EditorTileX + x2)
                                                .MaskSet = EditorSet
                                            End If
                                            If frmAttributes.optAnim.value = True Then
                                                .Anim = (EditorTileY + y2) * TilesInSheets + (EditorTileX + x2)
                                                .AnimSet = EditorSet
                                            End If
                                            If frmAttributes.optMask2.value = True Then
                                                .Mask2 = (EditorTileY + y2) * TilesInSheets + (EditorTileX + x2)
                                                .Mask2Set = EditorSet
                                            End If
                                            If frmAttributes.optM2Anim.value = True Then
                                                .M2Anim = (EditorTileY + y2) * TilesInSheets + (EditorTileX + x2)
                                                .M2AnimSet = EditorSet
                                            End If
                                            If frmAttributes.optFringe.value = True Then
                                                .Fringe = (EditorTileY + y2) * TilesInSheets + (EditorTileX + x2)
                                                .FringeSet = EditorSet
                                            End If
                                            If frmAttributes.optFAnim.value = True Then
                                                .FAnim = (EditorTileY + y2) * TilesInSheets + (EditorTileX + x2)
                                                .FAnimSet = EditorSet
                                            End If
                                            If frmAttributes.optFringe2.value = True Then
                                                .Fringe2 = (EditorTileY + y2) * TilesInSheets + (EditorTileX + x2)
                                                .Fringe2Set = EditorSet
                                            End If
                                            If frmAttributes.optF2Anim.value = True Then
                                                .F2Anim = (EditorTileY + y2) * TilesInSheets + (EditorTileX + x2)
                                                .F2AnimSet = EditorSet
                                            End If
                                        End With
                                    ElseIf frmMapEditor.mnuType(3).Checked = True Then
                                        Map(GetPlayerMap(MyIndex)).Tile(x1 + x2, y1 + y2).Light = (EditorTileY + y2) * TilesInSheets + (EditorTileX + x2)
                                    End If
                                End If
                            End If
                        Next x2
                    Next y2
                End If
            End If
            
            If (Button = 2) And (x1 >= 0) And (x1 <= MAX_MAPX) And (y1 >= 0) And (y1 <= MAX_MAPY) Then
                If frmMapEditor.mnuType(1).Checked = True Then
                    With Map(GetPlayerMap(MyIndex)).Tile(x1, y1)
                        If frmAttributes.optGround.value = True Then .Ground = 0
                        If frmAttributes.optMask.value = True Then .Mask = 0
                        If frmAttributes.optAnim.value = True Then .Anim = 0
                        If frmAttributes.optMask2.value = True Then .Mask2 = 0
                        If frmAttributes.optM2Anim.value = True Then .M2Anim = 0
                        If frmAttributes.optFringe.value = True Then .Fringe = 0
                        If frmAttributes.optFAnim.value = True Then .FAnim = 0
                        If frmAttributes.optFringe2.value = True Then .Fringe2 = 0
                        If frmAttributes.optF2Anim.value = True Then .F2Anim = 0
                    End With
                ElseIf frmMapEditor.mnuType(3).Checked = True Then
                    Map(GetPlayerMap(MyIndex)).Tile(x1, y1).Light = 0
                ElseIf frmMapEditor.mnuType(2).Checked = True Then
                    With Map(GetPlayerMap(MyIndex)).Tile(x1, y1)
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
    End If
End Sub

Public Sub EditorChooseTile(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        EditorTileX = Int(x / PIC_X)
        EditorTileY = Int(y / PIC_Y)
    End If
    frmMapEditor.shpSelected.Top = Int(EditorTileY * PIC_Y)
    frmMapEditor.shpSelected.Left = Int(EditorTileX * PIC_Y)
    'Call BitBlt(frmMapEditor.picSelect.hDC, 0, 0, PIC_X, PIC_Y, frmMapEditor.picBackSelect.hDC, EditorTileX * PIC_X, EditorTileY * PIC_Y, SRCCOPY)
End Sub

Public Sub EditorTileScroll()
    frmMapEditor.picBackSelect.Top = (frmMapEditor.scrlPicture.value * PIC_Y) * -1
End Sub

Public Sub EditorSend()
    Call SendMap
    Call EditorCancel
End Sub

Public Sub EditorCancel()
    InEditor = False
    frmMapEditor.Visible = False
    frmAttributes.Visible = False
    frmMirage.Show
    frmMapEditor.MousePointer = 1
    frmMirage.MousePointer = 1
    'frmMirage.picMapEditor.Visible = False
End Sub

Public Sub EditorClearLayer()
Dim YesNo As Long, x As Long, y As Long

    ' Ground layer
    If frmAttributes.optGround.value = True Then
    YesNo = MsgBox("Are you sure you wish to clear the ground layer?", vbYesNo, GAME_NAME)
        If YesNo = vbYes Then
            For y = 0 To MAX_MAPY
                For x = 0 To MAX_MAPX
                    Map(GetPlayerMap(MyIndex)).Tile(x, y).Ground = 0
                    Map(GetPlayerMap(MyIndex)).Tile(x, y).GroundSet = 0
                Next x
            Next y
        End If
    End If

    ' Mask layer
    If frmAttributes.optMask.value = True Then
    YesNo = MsgBox("Are you sure you wish to clear the mask layer?", vbYesNo, GAME_NAME)
        If YesNo = vbYes Then
            For y = 0 To MAX_MAPY
                For x = 0 To MAX_MAPX
                    Map(GetPlayerMap(MyIndex)).Tile(x, y).Mask = 0
                    Map(GetPlayerMap(MyIndex)).Tile(x, y).MaskSet = 0
                Next x
            Next y
        End If
    End If
    
    ' Mask Animation layer
    If frmAttributes.optAnim.value = True Then
    YesNo = MsgBox("Are you sure you wish to clear the animation layer?", vbYesNo, GAME_NAME)
        If YesNo = vbYes Then
            For y = 0 To MAX_MAPY
                For x = 0 To MAX_MAPX
                    Map(GetPlayerMap(MyIndex)).Tile(x, y).Anim = 0
                    Map(GetPlayerMap(MyIndex)).Tile(x, y).AnimSet = 0
                Next x
            Next y
        End If
    End If
    
    ' Mask 2 layer
    If frmAttributes.optMask2.value = True Then
    YesNo = MsgBox("Are you sure you wish to clear the mask 2 layer?", vbYesNo, GAME_NAME)
        If YesNo = vbYes Then
            For y = 0 To MAX_MAPY
                For x = 0 To MAX_MAPX
                    Map(GetPlayerMap(MyIndex)).Tile(x, y).Mask2 = 0
                    Map(GetPlayerMap(MyIndex)).Tile(x, y).Mask2Set = 0
                Next x
            Next y
        End If
    End If
    
    ' Mask 2 Animation layer
    If frmAttributes.optM2Anim.value = True Then
    YesNo = MsgBox("Are you sure you wish to clear the mask 2 animation layer?", vbYesNo, GAME_NAME)
        If YesNo = vbYes Then
            For y = 0 To MAX_MAPY
                For x = 0 To MAX_MAPX
                    Map(GetPlayerMap(MyIndex)).Tile(x, y).M2Anim = 0
                    Map(GetPlayerMap(MyIndex)).Tile(x, y).M2AnimSet = 0
                Next x
            Next y
        End If
    End If
    
    ' Fringe layer
    If frmAttributes.optFringe.value = True Then
    YesNo = MsgBox("Are you sure you wish to clear the fringe layer?", vbYesNo, GAME_NAME)
        If YesNo = vbYes Then
            For y = 0 To MAX_MAPY
                For x = 0 To MAX_MAPX
                    Map(GetPlayerMap(MyIndex)).Tile(x, y).Fringe = 0
                    Map(GetPlayerMap(MyIndex)).Tile(x, y).FringeSet = 0
                Next x
            Next y
        End If
    End If
    
    ' Fringe Animation layer
    If frmAttributes.optFAnim.value = True Then
    YesNo = MsgBox("Are you sure you wish to clear the fringe animation layer?", vbYesNo, GAME_NAME)
        If YesNo = vbYes Then
            For y = 0 To MAX_MAPY
                For x = 0 To MAX_MAPX
                    Map(GetPlayerMap(MyIndex)).Tile(x, y).FAnim = 0
                    Map(GetPlayerMap(MyIndex)).Tile(x, y).FAnimSet = 0
                Next x
            Next y
        End If
    End If
    
    ' Fringe 2 layer
    If frmAttributes.optFringe2.value = True Then
    YesNo = MsgBox("Are you sure you wish to clear the fringe 2 layer?", vbYesNo, GAME_NAME)
        If YesNo = vbYes Then
            For y = 0 To MAX_MAPY
                For x = 0 To MAX_MAPX
                    Map(GetPlayerMap(MyIndex)).Tile(x, y).Fringe2 = 0
                    Map(GetPlayerMap(MyIndex)).Tile(x, y).Fringe2Set = 0
                Next x
            Next y
        End If
    End If
    
    ' Fringe 2 Animation layer
    If frmAttributes.optF2Anim.value = True Then
    YesNo = MsgBox("Are you sure you wish to clear the fringe 2 animation layer?", vbYesNo, GAME_NAME)
        If YesNo = vbYes Then
            For y = 0 To MAX_MAPY
                For x = 0 To MAX_MAPX
                    Map(GetPlayerMap(MyIndex)).Tile(x, y).F2Anim = 0
                    Map(GetPlayerMap(MyIndex)).Tile(x, y).F2AnimSet = 0
                Next x
            Next y
        End If
    End If
End Sub

Public Sub EditorClearAttribs()
Dim YesNo As Long, x As Long, y As Long

    YesNo = MsgBox("Are you sure you wish to clear the attributes on this map?", vbYesNo, GAME_NAME)
    
    If YesNo = vbYes Then
        For y = 0 To MAX_MAPY
            For x = 0 To MAX_MAPX
                Map(GetPlayerMap(MyIndex)).Tile(x, y).Type = 0
            Next x
        Next y
    End If
End Sub

Public Sub EmoticonEditorInit()
    frmEmoticonEditor.scrlEmoticon.max = MAX_EMOTICONS
    frmEmoticonEditor.scrlEmoticon.value = Emoticons(EditorIndex - 1).Pic
    frmEmoticonEditor.txtCommand.Text = Trim(Emoticons(EditorIndex - 1).Command)
    frmEmoticonEditor.picEmoticons.Picture = LoadPicture(App.Path & "\GFX\emoticons.bmp")
    frmEmoticonEditor.Show vbModal
End Sub

Public Sub EmoticonEditorOk()
    Emoticons(EditorIndex - 1).Pic = frmEmoticonEditor.scrlEmoticon.value
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

Public Sub ArrowEditorInit()
    frmEditArrows.scrlArrow.max = MAX_ARROWS
    If Arrows(EditorIndex).Pic = 0 Then Arrows(EditorIndex).Pic = 1
    frmEditArrows.scrlArrow.value = Arrows(EditorIndex).Pic
    frmEditArrows.txtName.Text = Arrows(EditorIndex).name
    If Arrows(EditorIndex).Range = 0 Then Arrows(EditorIndex).Range = 1
    frmEditArrows.scrlRange.value = Arrows(EditorIndex).Range
    frmEditArrows.picArrows.Picture = LoadPicture(App.Path & "\GFX\arrows.bmp")
    frmEditArrows.Show vbModal
End Sub

Public Sub ArrowEditorOk()
    Arrows(EditorIndex).Pic = frmEditArrows.scrlArrow.value
    Arrows(EditorIndex).Range = frmEditArrows.scrlRange.value
    Arrows(EditorIndex).name = frmEditArrows.txtName.Text
    Call SendSaveArrow(EditorIndex)
    Call ArrowEditorCancel
End Sub

Public Sub ArrowEditorCancel()
    InArrowEditor = False
    Unload frmEditArrows
End Sub

Public Sub ItemEditorInit()
Dim i As Long
    EditorItemY = Int(Item(EditorIndex).Pic / 6)
    EditorItemX = (Item(EditorIndex).Pic - Int(Item(EditorIndex).Pic / 6) * 6)
    
    frmItemEditor.scrlClassReq.max = Max_Classes

    frmItemEditor.picItems.Picture = LoadPicture(App.Path & "\GFX\items.bmp")
    
    frmItemEditor.txtName.Text = Trim(Item(EditorIndex).name)
    frmItemEditor.txtDesc.Text = Trim(Item(EditorIndex).desc)
    frmItemEditor.cmbType.ListIndex = Item(EditorIndex).Type
    
    If (frmItemEditor.cmbType.ListIndex >= ITEM_TYPE_WEAPON) And (frmItemEditor.cmbType.ListIndex <= ITEM_TYPE_SHIELD) Then
        frmItemEditor.fraEquipment.Visible = True
        frmItemEditor.fraAttributes.Visible = True
        frmItemEditor.fraBow.Visible = True
        
        frmItemEditor.scrlDurability.value = Item(EditorIndex).Data1
        frmItemEditor.scrlStrength.value = Item(EditorIndex).Data2
        frmItemEditor.scrlStrReq.value = Item(EditorIndex).StrReq
        frmItemEditor.scrlDefReq.value = Item(EditorIndex).DefReq
        frmItemEditor.scrlSpeedReq.value = Item(EditorIndex).SpeedReq
        frmItemEditor.scrlClassReq.value = Item(EditorIndex).ClassReq
        frmItemEditor.scrlAccessReq.value = Item(EditorIndex).AccessReq
        frmItemEditor.scrlAddHP.value = Item(EditorIndex).AddHP
        frmItemEditor.scrlAddMP.value = Item(EditorIndex).AddMP
        frmItemEditor.scrlAddSP.value = Item(EditorIndex).AddSP
        frmItemEditor.scrlAddStr.value = Item(EditorIndex).AddStr
        frmItemEditor.scrlAddDef.value = Item(EditorIndex).AddDef
        frmItemEditor.scrlAddMagi.value = Item(EditorIndex).AddMagi
        frmItemEditor.scrlAddSpeed.value = Item(EditorIndex).AddSpeed
        frmItemEditor.scrlAddEXP.value = Item(EditorIndex).AddEXP
        frmItemEditor.scrlAttackSpeed.value = Item(EditorIndex).AttackSpeed
        If Item(EditorIndex).Data3 > 0 Then
            frmItemEditor.chkBow.value = Checked
        Else
            frmItemEditor.chkBow.value = Unchecked
        End If
        
        frmItemEditor.cmbBow.Clear
        If frmItemEditor.chkBow.value = Checked Then
            For i = 1 To 100
                frmItemEditor.cmbBow.AddItem i & ": " & Arrows(i).name
            Next i
            frmItemEditor.cmbBow.ListIndex = Item(EditorIndex).Data3 - 1
            frmItemEditor.picBow.Top = (Arrows(Item(EditorIndex).Data3).Pic * 32) * -1
            frmItemEditor.cmbBow.Enabled = True
        Else
            frmItemEditor.cmbBow.AddItem "None"
            frmItemEditor.cmbBow.ListIndex = 0
            frmItemEditor.cmbBow.Enabled = False
        End If
    Else
        frmItemEditor.fraEquipment.Visible = False
    End If
    
    If (frmItemEditor.cmbType.ListIndex >= ITEM_TYPE_POTIONADDHP) And (frmItemEditor.cmbType.ListIndex <= ITEM_TYPE_POTIONSUBSP) Then
        frmItemEditor.fraVitals.Visible = True
        frmItemEditor.scrlVitalMod.value = Item(EditorIndex).Data1
    Else
        frmItemEditor.fraVitals.Visible = False
    End If
    
    If (frmItemEditor.cmbType.ListIndex = ITEM_TYPE_SPELL) Then
        frmItemEditor.fraSpell.Visible = True
        frmItemEditor.scrlSpell.value = Item(EditorIndex).Data1
    Else
        frmItemEditor.fraSpell.Visible = False
    End If

    frmItemEditor.Show vbModal
End Sub

Public Sub ItemEditorOk()
    Item(EditorIndex).name = frmItemEditor.txtName.Text
    Item(EditorIndex).desc = frmItemEditor.txtDesc.Text
    Item(EditorIndex).Pic = EditorItemY * 6 + EditorItemX
    Item(EditorIndex).Type = frmItemEditor.cmbType.ListIndex

    If (frmItemEditor.cmbType.ListIndex >= ITEM_TYPE_WEAPON) And (frmItemEditor.cmbType.ListIndex <= ITEM_TYPE_SHIELD) Then
        Item(EditorIndex).Data1 = frmItemEditor.scrlDurability.value
        Item(EditorIndex).Data2 = frmItemEditor.scrlStrength.value
        If frmItemEditor.chkBow.value = Checked Then
            Item(EditorIndex).Data3 = frmItemEditor.cmbBow.ListIndex + 1
        Else
            Item(EditorIndex).Data3 = 0
        End If
        Item(EditorIndex).StrReq = frmItemEditor.scrlStrReq.value
        Item(EditorIndex).DefReq = frmItemEditor.scrlDefReq.value
        Item(EditorIndex).SpeedReq = frmItemEditor.scrlSpeedReq.value
        
        Item(EditorIndex).ClassReq = frmItemEditor.scrlClassReq.value
        Item(EditorIndex).AccessReq = frmItemEditor.scrlAccessReq.value
        
        Item(EditorIndex).AddHP = frmItemEditor.scrlAddHP.value
        Item(EditorIndex).AddMP = frmItemEditor.scrlAddMP.value
        Item(EditorIndex).AddSP = frmItemEditor.scrlAddSP.value
        Item(EditorIndex).AddStr = frmItemEditor.scrlAddStr.value
        Item(EditorIndex).AddDef = frmItemEditor.scrlAddDef.value
        Item(EditorIndex).AddMagi = frmItemEditor.scrlAddMagi.value
        Item(EditorIndex).AddSpeed = frmItemEditor.scrlAddSpeed.value
        Item(EditorIndex).AddEXP = frmItemEditor.scrlAddEXP.value
        Item(EditorIndex).AttackSpeed = frmItemEditor.scrlAttackSpeed.value
    End If
    
    If (frmItemEditor.cmbType.ListIndex >= ITEM_TYPE_POTIONADDHP) And (frmItemEditor.cmbType.ListIndex <= ITEM_TYPE_POTIONSUBSP) Then
        Item(EditorIndex).Data1 = frmItemEditor.scrlVitalMod.value
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
        Item(EditorIndex).AttackSpeed = 0
    End If
    
    If (frmItemEditor.cmbType.ListIndex = ITEM_TYPE_SPELL) Then
        Item(EditorIndex).Data1 = frmItemEditor.scrlSpell.value
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
        Item(EditorIndex).AttackSpeed = 0
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
    
    frmNpcEditor.picSprites.Picture = LoadPicture(App.Path & "\GFX\sprites.bmp")
    
    frmNpcEditor.txtName.Text = Trim(Npc(EditorIndex).name)
    frmNpcEditor.txtAttackSay.Text = Trim(Npc(EditorIndex).AttackSay)
    frmNpcEditor.scrlSprite.value = Npc(EditorIndex).Sprite
    frmNpcEditor.txtSpawnSecs.Text = STR(Npc(EditorIndex).SpawnSecs)
    frmNpcEditor.cmbBehavior.ListIndex = Npc(EditorIndex).Behavior
    frmNpcEditor.scrlRange.value = Npc(EditorIndex).Range
    frmNpcEditor.scrlSTR.value = Npc(EditorIndex).STR
    frmNpcEditor.scrlDEF.value = Npc(EditorIndex).DEF
    frmNpcEditor.scrlSPEED.value = Npc(EditorIndex).speed
    frmNpcEditor.scrlMAGI.value = Npc(EditorIndex).MAGI
    frmNpcEditor.BigNpc.value = Npc(EditorIndex).Big
    frmNpcEditor.StartHP.value = Npc(EditorIndex).MaxHp
    frmNpcEditor.ExpGive.value = Npc(EditorIndex).EXP
    frmNpcEditor.txtChance.Text = STR(Npc(EditorIndex).ItemNPC(1).Chance)
    frmNpcEditor.scrlNum.value = Npc(EditorIndex).ItemNPC(1).ItemNum
    frmNpcEditor.scrlValue.value = Npc(EditorIndex).ItemNPC(1).ItemValue
    If Npc(EditorIndex).SpawnTime = 0 Then
        frmNpcEditor.chkDay.value = Checked
        frmNpcEditor.chkNight.value = Checked
    ElseIf Npc(EditorIndex).SpawnTime = 1 Then
        frmNpcEditor.chkDay.value = Checked
        frmNpcEditor.chkNight.value = Unchecked
    ElseIf Npc(EditorIndex).SpawnTime = 2 Then
        frmNpcEditor.chkDay.value = Unchecked
        frmNpcEditor.chkNight.value = Checked
    End If
    
    frmNpcEditor.Show vbModal
End Sub

Public Sub NpcEditorOk()
    Npc(EditorIndex).name = frmNpcEditor.txtName.Text
    Npc(EditorIndex).AttackSay = frmNpcEditor.txtAttackSay.Text
    Npc(EditorIndex).Sprite = frmNpcEditor.scrlSprite.value
    Npc(EditorIndex).SpawnSecs = Val(frmNpcEditor.txtSpawnSecs.Text)
    Npc(EditorIndex).Behavior = frmNpcEditor.cmbBehavior.ListIndex
    Npc(EditorIndex).Range = frmNpcEditor.scrlRange.value
    Npc(EditorIndex).STR = frmNpcEditor.scrlSTR.value
    Npc(EditorIndex).DEF = frmNpcEditor.scrlDEF.value
    Npc(EditorIndex).speed = frmNpcEditor.scrlSPEED.value
    Npc(EditorIndex).MAGI = frmNpcEditor.scrlMAGI.value
    Npc(EditorIndex).Big = frmNpcEditor.BigNpc.value
    Npc(EditorIndex).MaxHp = frmNpcEditor.StartHP.value
    Npc(EditorIndex).EXP = frmNpcEditor.ExpGive.value
    
    If frmNpcEditor.chkDay.value = Checked And frmNpcEditor.chkNight.value = Checked Then
        Npc(EditorIndex).SpawnTime = 0
    ElseIf frmNpcEditor.chkDay.value = Checked And frmNpcEditor.chkNight.value = Unchecked Then
        Npc(EditorIndex).SpawnTime = 1
    ElseIf frmNpcEditor.chkDay.value = Unchecked And frmNpcEditor.chkNight.value = Checked Then
        Npc(EditorIndex).SpawnTime = 2
    End If
    
    Call SendSaveNpc(EditorIndex)
    InNpcEditor = False
    Unload frmNpcEditor
End Sub

Public Sub NpcEditorCancel()
    InNpcEditor = False
    Unload frmNpcEditor
End Sub

Public Sub NpcEditorBltSprite()
    If frmNpcEditor.BigNpc.value = Checked Then
        Call BitBlt(frmNpcEditor.picSprite.hDC, 0, 0, 64, 64, frmNpcEditor.picSprites.hDC, 3 * 64, frmNpcEditor.scrlSprite.value * 64, SRCCOPY)
    Else
        Call BitBlt(frmNpcEditor.picSprite.hDC, 0, 0, PIC_X, PIC_Y, frmNpcEditor.picSprites.hDC, 3 * PIC_X, frmNpcEditor.scrlSprite.value * PIC_Y, SRCCOPY)
    End If
End Sub

Public Sub ShopEditorInit()
Dim i As Long

    frmShopEditor.txtName.Text = Trim(Shop(EditorIndex).name)
    frmShopEditor.txtJoinSay.Text = Trim(Shop(EditorIndex).JoinSay)
    frmShopEditor.txtLeaveSay.Text = Trim(Shop(EditorIndex).LeaveSay)
    frmShopEditor.chkFixesItems.value = Shop(EditorIndex).FixesItems
    
    frmShopEditor.cmbItemGive.Clear
    frmShopEditor.cmbItemGive.AddItem "None"
    frmShopEditor.cmbItemGet.Clear
    frmShopEditor.cmbItemGet.AddItem "None"
    For i = 1 To MAX_ITEMS
        frmShopEditor.cmbItemGive.AddItem i & ": " & Trim(Item(i).name)
        frmShopEditor.cmbItemGet.AddItem i & ": " & Trim(Item(i).name)
    Next i
    frmShopEditor.cmbItemGive.ListIndex = 0
    frmShopEditor.cmbItemGet.ListIndex = 0
    
    Call UpdateShopTrade
    
    frmShopEditor.Show vbModal
End Sub

Public Sub UpdateShopTrade()
Dim i As Long, GetItem As Long, GetValue As Long, GiveItem As Long, GiveValue As Long, C As Long
    
    For i = 0 To 5
        frmShopEditor.lstTradeItem(i).Clear
    Next i
    
    For C = 1 To 6
        For i = 1 To MAX_TRADES
            GetItem = Shop(EditorIndex).TradeItem(C).value(i).GetItem
            GetValue = Shop(EditorIndex).TradeItem(C).value(i).GetValue
            GiveItem = Shop(EditorIndex).TradeItem(C).value(i).GiveItem
            GiveValue = Shop(EditorIndex).TradeItem(C).value(i).GiveValue

            If GetItem > 0 And GiveItem > 0 Then
                frmShopEditor.lstTradeItem(C - 1).AddItem i & ": " & GiveValue & " " & Trim(Item(GiveItem).name) & " for " & GetValue & " " & Trim(Item(GetItem).name)
            Else
                frmShopEditor.lstTradeItem(C - 1).AddItem "Empty Trade Slot"
            End If
        Next i
    Next C
    
    For i = 0 To 5
        frmShopEditor.lstTradeItem(i).ListIndex = 0
    Next i
End Sub

Public Sub ShopEditorOk()
    Shop(EditorIndex).name = frmShopEditor.txtName.Text
    Shop(EditorIndex).JoinSay = frmShopEditor.txtJoinSay.Text
    Shop(EditorIndex).LeaveSay = frmShopEditor.txtLeaveSay.Text
    Shop(EditorIndex).FixesItems = frmShopEditor.chkFixesItems.value
    
    Call SendSaveShop(EditorIndex)
    InShopEditor = False
    Unload frmShopEditor
End Sub

Public Sub ShopEditorCancel()
    InShopEditor = False
    Unload frmShopEditor
End Sub

Public Sub SpellEditorInit()
Dim i As Long

    frmSpellEditor.cmbClassReq.AddItem "All Classes"
    For i = 0 To Max_Classes
        frmSpellEditor.cmbClassReq.AddItem Trim(Class(i).name)
    Next i
    
    frmSpellEditor.txtName.Text = Trim(Spell(EditorIndex).name)
    frmSpellEditor.cmbClassReq.ListIndex = Spell(EditorIndex).ClassReq
    frmSpellEditor.scrlLevelReq.value = Spell(EditorIndex).LevelReq
        
    frmSpellEditor.cmbType.ListIndex = Spell(EditorIndex).Type
    frmSpellEditor.scrlVitalMod.value = Spell(EditorIndex).Data1
    
    frmSpellEditor.scrlCost.value = Spell(EditorIndex).MPCost
    frmSpellEditor.scrlSound.value = Spell(EditorIndex).Sound
    
    If Spell(EditorIndex).Range = 0 Then Spell(EditorIndex).Range = 1
    frmSpellEditor.scrlRange.value = Spell(EditorIndex).Range
    
    frmSpellEditor.scrlSpellAnim.value = Spell(EditorIndex).SpellAnim
    frmSpellEditor.scrlSpellTime.value = Spell(EditorIndex).SpellTime
    frmSpellEditor.scrlSpellDone.value = Spell(EditorIndex).SpellDone
    
    frmSpellEditor.chkArea.value = Spell(EditorIndex).AE
        
    frmSpellEditor.Show vbModal
End Sub

Public Sub SpellEditorOk()
    Spell(EditorIndex).name = frmSpellEditor.txtName.Text
    Spell(EditorIndex).ClassReq = frmSpellEditor.cmbClassReq.ListIndex
    Spell(EditorIndex).LevelReq = frmSpellEditor.scrlLevelReq.value
    Spell(EditorIndex).Type = frmSpellEditor.cmbType.ListIndex
    Spell(EditorIndex).Data1 = frmSpellEditor.scrlVitalMod.value
    Spell(EditorIndex).Data3 = 0
    Spell(EditorIndex).MPCost = frmSpellEditor.scrlCost.value
    Spell(EditorIndex).Sound = frmSpellEditor.scrlSound.value
    Spell(EditorIndex).Range = frmSpellEditor.scrlRange.value
    
    Spell(EditorIndex).SpellAnim = frmSpellEditor.scrlSpellAnim.value
    Spell(EditorIndex).SpellTime = frmSpellEditor.scrlSpellTime.value
    Spell(EditorIndex).SpellDone = frmSpellEditor.scrlSpellDone.value
    
    Spell(EditorIndex).AE = frmSpellEditor.chkArea.value
    
    Call SendSaveSpell(EditorIndex)
    InSpellEditor = False
    Unload frmSpellEditor
End Sub

Public Sub SpellEditorCancel()
    InSpellEditor = False
    Unload frmSpellEditor
End Sub

Public Sub UpdateTradeInventory()
Dim i As Long

    frmPlayerTrade.PlayerInv1.Clear
    
For i = 1 To MAX_INV
    If GetPlayerInvItemNum(MyIndex, i) > 0 And GetPlayerInvItemNum(MyIndex, i) <= MAX_ITEMS Then
        If Item(GetPlayerInvItemNum(MyIndex, i)).Type = ITEM_TYPE_CURRENCY Then
            frmPlayerTrade.PlayerInv1.AddItem i & ": " & Trim(Item(GetPlayerInvItemNum(MyIndex, i)).name) & " (" & GetPlayerInvItemValue(MyIndex, i) & ")"
        Else
            If GetPlayerWeaponSlot(MyIndex) = i Or GetPlayerArmorSlot(MyIndex) = i Or GetPlayerHelmetSlot(MyIndex) = i Or GetPlayerShieldSlot(MyIndex) = i Then
                frmPlayerTrade.PlayerInv1.AddItem i & ": " & Trim(Item(GetPlayerInvItemNum(MyIndex, i)).name) & " (worn)"
            Else
                frmPlayerTrade.PlayerInv1.AddItem i & ": " & Trim(Item(GetPlayerInvItemNum(MyIndex, i)).name)
            End If
        End If
    Else
        frmPlayerTrade.PlayerInv1.AddItem "<Nothing>"
    End If
Next i
    
    frmPlayerTrade.PlayerInv1.ListIndex = 0
End Sub

Sub PlayerSearch(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim x1 As Long, y1 As Long

    x1 = Int(x / PIC_X)
    y1 = Int(y / PIC_Y)
    
    If (x1 >= 0) And (x1 <= MAX_MAPX) And (y1 >= 0) And (y1 <= MAX_MAPY) Then
        Call SendData("search" & SEP_CHAR & x1 & SEP_CHAR & y1 & SEP_CHAR & END_CHAR)
    End If
    MouseDownX = x1
    MouseDownY = y1
End Sub
Sub BltTile2(ByVal x As Long, ByVal y As Long, ByVal Tile As Long)
If TileFile(6) = 0 Then Exit Sub

    rec.Top = Int(Tile / TilesInSheets) * PIC_Y
    rec.Bottom = rec.Top + PIC_Y
    rec.Left = (Tile - Int(Tile / TilesInSheets) * TilesInSheets) * PIC_X
    rec.Right = rec.Left + PIC_X
    Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) + sx - NewXOffset, y - (NewPlayerY * PIC_Y) + sx - NewYOffset, DD_TileSurf(6), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End Sub

Sub BltPlayerText(ByVal index As Long)
Dim TextX As Long
Dim TextY As Long
Dim intLoop As Integer
Dim intLoop2 As Integer

Dim bytLineCount As Byte
Dim bytLineLength As Byte
Dim strLine(0 To MAX_LINES - 1) As String
Dim strWords() As String
    strWords() = Split(Bubble(index).Text, " ")
    
    If Len(Bubble(index).Text) < MAX_LINE_LENGTH Then
        DISPLAY_BUBBLE_WIDTH = 2 + ((Len(Bubble(index).Text) * 9) \ PIC_X)
        
        If DISPLAY_BUBBLE_WIDTH > MAX_BUBBLE_WIDTH Then
            DISPLAY_BUBBLE_WIDTH = MAX_BUBBLE_WIDTH
        End If
    Else
        DISPLAY_BUBBLE_WIDTH = 6
    End If
    
    TextX = GetPlayerX(index) * PIC_X + Player(index).XOffset + Int(PIC_X) - ((DISPLAY_BUBBLE_WIDTH * 32) / 2) - 6
    TextY = GetPlayerY(index) * PIC_Y + Player(index).YOffset - Int(PIC_Y) + 85
    
    Call DD_BackBuffer.ReleaseDC(TexthDC)
    
    ' Draw the fancy box with tiles.
    Call BltTile2(TextX - 10, TextY - 10, 1)
    Call BltTile2(TextX + (DISPLAY_BUBBLE_WIDTH * PIC_X) - PIC_X - 10, TextY - 10, 2)
    
    For intLoop = 1 To (DISPLAY_BUBBLE_WIDTH - 2)
        Call BltTile2(TextX - 10 + (intLoop * PIC_X), TextY - 10, 16)
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
            Call BltTile2(TextX - 10 + (DISPLAY_BUBBLE_WIDTH * PIC_X) - PIC_X, TextY - 10 + intLoop, 17)
            
            For intLoop2 = 1 To DISPLAY_BUBBLE_WIDTH - 2
                Call BltTile2(TextX - 10 + (intLoop2 * PIC_X), TextY + intLoop - 10, 5)
            Next intLoop2
        Next intLoop
    End If

    Call BltTile2(TextX - 10, TextY + (bytLineCount * 16) - 4, 3)
    Call BltTile2(TextX + (DISPLAY_BUBBLE_WIDTH * PIC_X) - PIC_X - 10, TextY + (bytLineCount * 16) - 4, 4)
    
    For intLoop = 1 To (DISPLAY_BUBBLE_WIDTH - 2)
        Call BltTile2(TextX - 10 + (intLoop * PIC_X), TextY + (bytLineCount * 16) - 4, 15)
    Next intLoop
    
    TexthDC = DD_BackBuffer.GetDC
    
    For intLoop = 0 To (MAX_LINES - 1)
        If strLine(intLoop) <> "" Then
            Call DrawText(TexthDC, TextX - (NewPlayerX * PIC_X) + sx - NewXOffset + (((DISPLAY_BUBBLE_WIDTH) * PIC_X) / 2) - ((Len(strLine(intLoop)) * 8) \ 2) - 7, TextY - (NewPlayerY * PIC_Y) + sx - NewYOffset, strLine(intLoop), QBColor(DarkGrey))
            TextY = TextY + 16
        End If
    Next intLoop
End Sub
Sub BltPlayerBar() 's(ByVal Index As Long)
Dim x As Long, y As Long, index As Long

index = MyIndex

x = (GetPlayerX(index) * PIC_X + sx + Player(index).XOffset) - (NewPlayerX * PIC_X) - NewXOffset
y = (GetPlayerY(index) * PIC_Y + sx + Player(index).YOffset) - (NewPlayerY * PIC_Y) - NewYOffset

If Player(index).HP = 0 Then Exit Sub
    'draws the back bars
    Call DD_BackBuffer.SetFillColor(RGB(255, 0, 0))
    Call DD_BackBuffer.DrawBox(x, y + 2, x + 32, y - 2)
    
    'draws HP
    Call DD_BackBuffer.SetFillColor(RGB(0, 255, 0))
    Call DD_BackBuffer.DrawBox(x, y + 2, x + ((Player(index).HP / 100) / (Player(index).MaxHp / 100) * 32), y - 2)
End Sub
Sub BltNpcBars(ByVal index As Long)
Dim x As Long, y As Long

If MapNpc(index).HP = 0 Then Exit Sub
If MapNpc(index).num < 1 Then Exit Sub

    If Npc(MapNpc(index).num).Big = 1 Then
        x = (MapNpc(index).x * PIC_X + sx - 9 + MapNpc(index).XOffset) - (NewPlayerX * PIC_X) - NewXOffset
        y = (MapNpc(index).y * PIC_Y + sx + MapNpc(index).YOffset) - (NewPlayerY * PIC_Y) - NewYOffset
        
        Call DD_BackBuffer.SetFillColor(RGB(255, 0, 0))
        Call DD_BackBuffer.DrawBox(x, y + 32, x + 50, y + 36)
        Call DD_BackBuffer.SetFillColor(RGB(0, 255, 0))
        Call DD_BackBuffer.DrawBox(x, y + 32, x + ((MapNpc(index).HP / 100) / (MapNpc(index).MaxHp / 100) * 50), y + 36)
    Else
        x = (MapNpc(index).x * PIC_X + sx + MapNpc(index).XOffset) - (NewPlayerX * PIC_X) - NewXOffset
        y = (MapNpc(index).y * PIC_Y + sx + MapNpc(index).YOffset) - (NewPlayerY * PIC_Y) - NewYOffset
        
        Call DD_BackBuffer.SetFillColor(RGB(255, 0, 0))
        Call DD_BackBuffer.DrawBox(x, y + 32, x + 32, y + 36)
        Call DD_BackBuffer.SetFillColor(RGB(0, 255, 0))
        Call DD_BackBuffer.DrawBox(x, y + 32, x + ((MapNpc(index).HP / 100) / (MapNpc(index).MaxHp / 100) * 32), y + 36)
    End If
End Sub

Public Sub UpdateVisInv()
Dim index As Long
Dim d As Long

    For index = 1 To MAX_INV
        If GetPlayerShieldSlot(MyIndex) <> index Then frmMirage.ShieldImage.Picture = LoadPicture()
        If GetPlayerWeaponSlot(MyIndex) <> index Then frmMirage.WeaponImage.Picture = LoadPicture()
        If GetPlayerHelmetSlot(MyIndex) <> index Then frmMirage.HelmetImage.Picture = LoadPicture()
        If GetPlayerArmorSlot(MyIndex) <> index Then frmMirage.ArmorImage.Picture = LoadPicture()
    Next index
    
    For index = 1 To MAX_INV
        If GetPlayerShieldSlot(MyIndex) = index Then Call BitBlt(frmMirage.ShieldImage.hDC, 0, 0, PIC_X, PIC_Y, frmMirage.picItems.hDC, (Item(GetPlayerInvItemNum(MyIndex, index)).Pic - Int(Item(GetPlayerInvItemNum(MyIndex, index)).Pic / 6) * 6) * PIC_X, Int(Item(GetPlayerInvItemNum(MyIndex, index)).Pic / 6) * PIC_Y, SRCCOPY)
        If GetPlayerWeaponSlot(MyIndex) = index Then Call BitBlt(frmMirage.WeaponImage.hDC, 0, 0, PIC_X, PIC_Y, frmMirage.picItems.hDC, (Item(GetPlayerInvItemNum(MyIndex, index)).Pic - Int(Item(GetPlayerInvItemNum(MyIndex, index)).Pic / 6) * 6) * PIC_X, Int(Item(GetPlayerInvItemNum(MyIndex, index)).Pic / 6) * PIC_Y, SRCCOPY)
        If GetPlayerHelmetSlot(MyIndex) = index Then Call BitBlt(frmMirage.HelmetImage.hDC, 0, 0, PIC_X, PIC_Y, frmMirage.picItems.hDC, (Item(GetPlayerInvItemNum(MyIndex, index)).Pic - Int(Item(GetPlayerInvItemNum(MyIndex, index)).Pic / 6) * 6) * PIC_X, Int(Item(GetPlayerInvItemNum(MyIndex, index)).Pic / 6) * PIC_Y, SRCCOPY)
        If GetPlayerArmorSlot(MyIndex) = index Then Call BitBlt(frmMirage.ArmorImage.hDC, 0, 0, PIC_X, PIC_Y, frmMirage.picItems.hDC, (Item(GetPlayerInvItemNum(MyIndex, index)).Pic - Int(Item(GetPlayerInvItemNum(MyIndex, index)).Pic / 6) * 6) * PIC_X, Int(Item(GetPlayerInvItemNum(MyIndex, index)).Pic / 6) * PIC_Y, SRCCOPY)
    Next index
        
    frmMirage.EquipS(0).Visible = False
    frmMirage.EquipS(1).Visible = False
    frmMirage.EquipS(2).Visible = False
    frmMirage.EquipS(3).Visible = False

    For d = 0 To MAX_INV - 1
        If Player(MyIndex).Inv(d + 1).num > 0 Then
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

Sub BltSpriteChange(ByVal x As Long, ByVal y As Long)
Dim x2 As Long, y2 As Long
    
    If Map(GetPlayerMap(MyIndex)).Tile(x, y).Type = TILE_TYPE_SPRITE_CHANGE Then

        ' Only used if ever want to switch to blt rather then bltfast
        With rec_pos
            .Top = y * PIC_Y
            .Bottom = .Top + PIC_Y
            .Left = x * PIC_X
            .Right = .Left + PIC_X
        End With
                                        
        rec.Top = Map(GetPlayerMap(MyIndex)).Tile(x, y).Data1 * PIC_Y + 16
        rec.Bottom = rec.Top + PIC_Y - 16
        rec.Left = 96
        rec.Right = rec.Left + PIC_X
        
        x2 = x * PIC_X + sx
        y2 = y * PIC_Y + sx
                                           
        'Call DD_BackBuffer.Blt(rec_pos, DD_SpriteSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
        Call DD_BackBuffer.BltFast(x2 - (NewPlayerX * PIC_X) - NewXOffset, y2 - (NewPlayerY * PIC_Y) - NewYOffset, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    End If
End Sub

Sub BltSpriteChange2(ByVal x As Long, ByVal y As Long)
Dim x2 As Long, y2 As Long
    
    If Map(GetPlayerMap(MyIndex)).Tile(x, y).Type = TILE_TYPE_SPRITE_CHANGE Then

        ' Only used if ever want to switch to blt rather then bltfast
        With rec_pos
            .Top = y * PIC_Y
            .Bottom = .Top + PIC_Y
            .Left = x * PIC_X
            .Right = .Left + PIC_X
        End With
                                        
        rec.Top = Map(GetPlayerMap(MyIndex)).Tile(x, y).Data1 * PIC_Y
        rec.Bottom = rec.Top + PIC_Y - 16
        rec.Left = 96
        rec.Right = rec.Left + PIC_X
        
        x2 = x * PIC_X + sx
        y2 = y * PIC_Y + (sx / 2) '- 16
               
        If y2 < 0 Then
            Exit Sub
        End If
                            
        'Call DD_BackBuffer.Blt(rec_pos, DD_SpriteSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
        Call DD_BackBuffer.BltFast(x2 - (NewPlayerX * PIC_X) - NewXOffset, y2 - (NewPlayerY * PIC_Y) - NewYOffset, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    End If
End Sub

Sub SendGameTime()
Dim Packet As String

Packet = "GmTime" & SEP_CHAR & GameTime & SEP_CHAR & END_CHAR
Call SendData(Packet)
End Sub

Sub ItemSelected(ByVal index As Long, ByVal Selected As Long)
Dim index2 As Long
index2 = Trade(Selected).Items(index).ItemGetNum

    frmTrade.shpSelect.Top = frmTrade.picItem(index - 1).Top - 1
    frmTrade.shpSelect.Left = frmTrade.picItem(index - 1).Left - 1

    If index2 <= 0 Then
        Call clearItemSelected
        Exit Sub
    End If

    frmTrade.descName.Caption = Trim(Item(index2).name)
    frmTrade.descQuantity.Caption = Trade(Selected).Items(index).ItemGetVal
    
    frmTrade.descStr.Caption = Item(index2).StrReq
    frmTrade.descDef.Caption = Item(index2).DefReq
    frmTrade.descSpeed.Caption = Item(index2).SpeedReq
    
    frmTrade.descAStr.Caption = Item(index2).AddStr
    frmTrade.descADef.Caption = Item(index2).AddDef
    frmTrade.descAMagi.Caption = Item(index2).AddMagi
    frmTrade.descASpeed.Caption = Item(index2).AddSpeed
    
    frmTrade.descHp.Caption = Item(index2).AddHP
    frmTrade.descMp.Caption = Item(index2).AddMP
    frmTrade.descSp.Caption = Item(index2).AddSP

    frmTrade.descAExp.Caption = Item(index2).AddEXP
    frmTrade.desc.Caption = Trim(Item(index2).desc)
    
    frmTrade.lblTradeFor.Caption = Trim(Item(Trade(Selected).Items(index).ItemGiveNum).name)
    frmTrade.lblQuantity.Caption = Trade(Selected).Items(index).ItemGiveVal
End Sub

Sub clearItemSelected()
    frmTrade.lblTradeFor.Caption = ""
    frmTrade.lblQuantity.Caption = ""
    
    frmTrade.descName.Caption = ""
    frmTrade.descQuantity.Caption = ""
    
    frmTrade.descStr.Caption = 0
    frmTrade.descDef.Caption = 0
    frmTrade.descSpeed.Caption = 0
    
    frmTrade.descAStr.Caption = 0
    frmTrade.descADef.Caption = 0
    frmTrade.descAMagi.Caption = 0
    frmTrade.descASpeed.Caption = 0
    
    frmTrade.descHp.Caption = 0
    frmTrade.descMp.Caption = 0
    frmTrade.descSp.Caption = 0

    frmTrade.descAExp.Caption = 0
    frmTrade.desc.Caption = ""
End Sub

Function IsVisible(x, y)
    Dim DistLeft As Integer
    Dim DistRight As Integer
    Dim DistUp As Integer
    Dim DistDown As Integer
    
    DistLeft = Player(MyIndex).x
    DistRight = MAX_MAPX - Player(MyIndex).x
    DistUp = Player(MyIndex).y
    DistDown = MAX_MAPY - Player(MyIndex).y
    
    If DistLeft < 12 Then DistRight = DistRight + (26 - DistLeft)
    If DistRight < 12 Then DistLeft = DistLeft + (26 - DistRight)
    If DistUp < 7 Then DistDown = DistDown + (15 - DistUp)
    If DistDown < 7 Then DistUp = DistUp + (15 - DistDown)
    
    If (x >= (Player(MyIndex).x - DistLeft)) And (x <= (Player(MyIndex).x + DistRight)) And (y >= (Player(MyIndex).y - DistUp)) And (y <= (Player(MyIndex).y + DistDown)) Then
        IsVisible = True
    Else
        IsVisible = False
    End If
End Function
