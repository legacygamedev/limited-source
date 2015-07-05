Attribute VB_Name = "modGameLogic"

' Copyright (c) 2006 Chaos Engine Source. All rights reserved.
' This code is licensed under the Chaos Engine General License.

Option Explicit

Private Declare Function GetVersion Lib "kernel32" () As Long
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crey As Byte, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Const SRCAND = &H8800C6
Public Const SRCCOPY = &HCC0020
Public Const SRCPAINT = &HEE0086

' Translucecy stuff
Public Const GWL_EXSTYLE = (-20)
Public Const WS_EX_LAYERED = &H80000
Public Const WS_EX_TRANSPARENT = &H20&
Public Const LWA_ALPHA = &H2&

Public Const VK_UP = &H26
Public Const VK_DOWN = &H28
Public Const VK_LEFT = &H25
Public Const VK_RIGHT = &H27
Public Const VK_SHIFT = &H10
Public Const VK_RETURN = &HD
Public Const VK_CONTROL = &H11

' Full Screen or Windowed
Public mclsStyle As clsWindowed

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
Public Const GM_WALK_SPEED = 4

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

' Alignment
Public AlignmentBarTime As Long

Public CorpseIndex As Integer

' Used to check if in editor or not and variables for use in editor
Public InEditor As Boolean
Public EditorTileX As Long
Public EditorTileY As Long
Public EditorWarpMap As Long
Public EditorWarpX As Long
Public EditorWarpY As Long
Public EditorSet As Byte

Public EditorSpellX As Long
Public EditorSpellY As Long

' Used for map item editor
Public ItemEditorNum As Long
Public ItemEditorValue As Long

' Used for map furniture editor
Public FurnitureNum As Long

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
Public InSpeechEditor As Boolean
Public InItemsEditor As Boolean
Public InNpcEditor As Boolean
Public InShopEditor As Boolean
Public InSpellEditor As Boolean
Public InElementEditor As Boolean
Public InEmoticonEditor As Boolean
Public InArrowEditor As Boolean
Public InSpawnEditor As Boolean
Public EditorIndex As Long

' Used to know what npc we are choosing the spawn for
Public SpawnLocator As Long

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

Public MiniMap As Boolean

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

Public HouseItem As Long
Public HousePrice As Long

Public SelectorWidth As Long
Public SelectorHeight As Long

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

Public SpeechEditorCurrentNumber As Long
Public SpeechConvo1 As Long
Public SpeechConvo2 As Long
Public SpeechConvo3 As Long

Public ShopNum As Long

Public GoDebug As Long

Public MouseX As Long
Public MouseY As Long
Public XToGo As Long
Public YToGo As Long
                    
Sub Main()
Dim i As Long
Dim Ending As String
Dim sDc As Long
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
    If UCase(Dir(App.Path & "\Music", vbDirectory)) <> "MUSIC" Then
        Call MkDir(App.Path & "\Music")
    End If
    If UCase(Dir(App.Path & "\SFX", vbDirectory)) <> "SFX" Then
        Call MkDir(App.Path & "\SFX")
    End If
    If UCase(Dir(App.Path & "\Flashs", vbDirectory)) <> "FLASHS" Then
        Call MkDir(App.Path & "\Flashs")
    End If
    
    'Load the DX Surfaces here..
    Call InitDirectX(True)
    
    Dim filename As String
    filename = App.Path & "\config.ini"
    If FileExist("config.ini") Then
        frmMirage.chkbubblebar.Value = Val(GetVar(filename, "CONFIG", "SpeechBubbles"))
        frmMirage.chkEmoSound.Value = Val(GetVar(filename, "CONFIG", "EmoticonSound"))
        frmMirage.chknpcname.Value = Val(GetVar(filename, "CONFIG", "NPCName"))
        frmMirage.chkplayername.Value = Val(GetVar(filename, "CONFIG", "PlayerName"))
        frmMirage.chkplayerdamage.Value = Val(GetVar(filename, "CONFIG", "NPCDamage"))
        frmMirage.chknpcdamage.Value = Val(GetVar(filename, "CONFIG", "PlayerDamage"))
        frmMirage.chkmusic.Value = Val(GetVar(filename, "CONFIG", "Music"))
        frmMirage.chkSound.Value = Val(GetVar(filename, "CONFIG", "Sound"))
        frmMirage.chkAutoScroll.Value = Val(GetVar(filename, "CONFIG", "AutoScroll"))

        If Val(GetVar(filename, "CONFIG", "MapGrid")) = 0 Then
            frmMapEditor.chkGrid.Value = 0
        Else
            frmMapEditor.chkGrid.Value = 1
        End If
    Else
        Call PutVar(App.Path & "\config.ini", "UPDATER", "FileName", "Chaos Client.exe")
        Call PutVar(App.Path & "\config.ini", "UPDATER", "WebSite", "")
        Call PutVar(App.Path & "\config.ini", "IPCONFIG", "IP", "127.0.0.1")
        Call PutVar(App.Path & "\config.ini", "IPCONFIG", "PORT", 4000)
        Call PutVar(App.Path & "\config.ini", "CONFIG", "Account", "")
        Call PutVar(App.Path & "\config.ini", "CONFIG", "Password", "")
        Call PutVar(App.Path & "\config.ini", "CONFIG", "WebSite", "")
        Call PutVar(App.Path & "\config.ini", "CONFIG", "SpeechBubbles", 1)
        Call PutVar(App.Path & "\config.ini", "CONFIG", "EmoticonSound", 1)
        Call PutVar(App.Path & "\config.ini", "CONFIG", "NPCName", 1)
        Call PutVar(App.Path & "\config.ini", "CONFIG", "NPCDamage", 1)
        Call PutVar(App.Path & "\config.ini", "CONFIG", "PlayerName", 1)
        Call PutVar(App.Path & "\config.ini", "CONFIG", "PlayerDamage", 1)
        Call PutVar(App.Path & "\config.ini", "CONFIG", "MapGrid", 1)
        Call PutVar(App.Path & "\config.ini", "CONFIG", "Music", 1)
        Call PutVar(App.Path & "\config.ini", "CONFIG", "Sound", 1)
        Call PutVar(App.Path & "\config.ini", "CONFIG", "AutoScroll", 1)
    End If
    
    If FileExist("News.ini") = False Then
        WriteINI "DATA", "News", "News: -The Chaos Engine has been released", App.Path & "\News.ini"
    End If
    
    If FileExist("GUI\Colors.txt") = False Then
        Call PutVar(App.Path & "\Colors.txt", "CHATBOX", "R", 152)
        Call PutVar(App.Path & "\Colors.txt", "CHATBOX", "G", 146)
        Call PutVar(App.Path & "\Colors.txt", "CHATBOX", "B", 120)
        
        Call PutVar(App.Path & "\Colors.txt", "CHATTEXTBOX", "R", 152)
        Call PutVar(App.Path & "\Colors.txt", "CHATTEXTBOX", "G", 146)
        Call PutVar(App.Path & "\Colors.txt", "CHATTEXTBOX", "B", 120)
        
        Call PutVar(App.Path & "\Colors.txt", "BACKGROUND", "R", 152)
        Call PutVar(App.Path & "\Colors.txt", "BACKGROUND", "G", 146)
        Call PutVar(App.Path & "\Colors.txt", "BACKGROUND", "B", 120)
        
        Call PutVar(App.Path & "\Colors.txt", "SPELLLIST", "R", 152)
        Call PutVar(App.Path & "\Colors.txt", "SPELLLIST", "G", 146)
        Call PutVar(App.Path & "\Colors.txt", "SPELLLIST", "B", 120)

        Call PutVar(App.Path & "\Colors.txt", "WHOLIST", "R", 152)
        Call PutVar(App.Path & "\Colors.txt", "WHOLIST", "G", 146)
        Call PutVar(App.Path & "\Colors.txt", "WHOLIST", "B", 120)
        
        Call PutVar(App.Path & "\Colors.txt", "NEWCHAR", "R", 152)
        Call PutVar(App.Path & "\Colors.txt", "NEWCHAR", "G", 146)
        Call PutVar(App.Path & "\Colors.txt", "NEWCHAR", "B", 120)
    End If
    
    Dim R1 As Long, G1 As Long, B1 As Long
    R1 = Val(GetVar(App.Path & "\Colors.txt", "CHATBOX", "R"))
    G1 = Val(GetVar(App.Path & "\Colors.txt", "CHATBOX", "G"))
    B1 = Val(GetVar(App.Path & "\Colors.txt", "CHATBOX", "B"))
    frmMirage.txtChat.BackColor = RGB(R1, G1, B1)
    
    R1 = Val(GetVar(App.Path & "\Colors.txt", "CHATTEXTBOX", "R"))
    G1 = Val(GetVar(App.Path & "\Colors.txt", "CHATTEXTBOX", "G"))
    B1 = Val(GetVar(App.Path & "\Colors.txt", "CHATTEXTBOX", "B"))
    frmMirage.txtMyTextBox.BackColor = RGB(R1, G1, B1)
    
    R1 = Val(GetVar(App.Path & "\Colors.txt", "BACKGROUND", "R"))
    G1 = Val(GetVar(App.Path & "\Colors.txt", "BACKGROUND", "G"))
    B1 = Val(GetVar(App.Path & "\Colors.txt", "BACKGROUND", "B"))
    frmMirage.Picture9.BackColor = RGB(R1, G1, B1)
    frmMirage.Picture8.BackColor = RGB(R1, G1, B1)
    frmMirage.picInv3.BackColor = RGB(R1, G1, B1)
    frmMirage.itmDesc.BackColor = RGB(R1, G1, B1)
    frmMirage.picWhosOnline.BackColor = RGB(R1, G1, B1)
    frmMirage.picGuildAdmin.BackColor = RGB(R1, G1, B1)
    frmMirage.picGuild.BackColor = RGB(R1, G1, B1)
    frmMirage.picEquip.BackColor = RGB(R1, G1, B1)
    frmMirage.picPlayerSpells.BackColor = RGB(R1, G1, B1)
    frmMirage.picOptions.BackColor = RGB(R1, G1, B1)
    
    frmMirage.chkbubblebar.BackColor = RGB(R1, G1, B1)
    frmMirage.chkEmoSound.BackColor = RGB(R1, G1, B1)
    frmMirage.chknpcname.BackColor = RGB(R1, G1, B1)
    frmMirage.chkplayername.BackColor = RGB(R1, G1, B1)
    frmMirage.chkplayerdamage.BackColor = RGB(R1, G1, B1)
    frmMirage.chknpcdamage.BackColor = RGB(R1, G1, B1)
    frmMirage.chkmusic.BackColor = RGB(R1, G1, B1)
    frmMirage.chkSound.BackColor = RGB(R1, G1, B1)
    frmMirage.chkAutoScroll.BackColor = RGB(R1, G1, B1)
    
    R1 = Val(GetVar(App.Path & "\Colors.txt", "SPELLLIST", "R"))
    G1 = Val(GetVar(App.Path & "\Colors.txt", "SPELLLIST", "G"))
    B1 = Val(GetVar(App.Path & "\Colors.txt", "SPELLLIST", "B"))
    'frmMirage.lstSpells.BackColor = RGB(R1, G1, B1)
    
    R1 = Val(GetVar(App.Path & "\Colors.txt", "WHOLIST", "R"))
    G1 = Val(GetVar(App.Path & "\Colors.txt", "WHOLIST", "G"))
    B1 = Val(GetVar(App.Path & "\Colors.txt", "WHOLIST", "B"))
    frmMirage.lstOnline.BackColor = RGB(R1, G1, B1)
    
    R1 = Val(GetVar(App.Path & "\Colors.txt", "FRIENDLIST", "R"))
    G1 = Val(GetVar(App.Path & "\Colors.txt", "FRIENDLIST", "G"))
    B1 = Val(GetVar(App.Path & "\Colors.txt", "FRIENDLIST", "B"))
    frmMirage.lstFriend.BackColor = RGB(R1, G1, B1)

    R1 = Val(GetVar(App.Path & "\Colors.txt", "NEWCHAR", "R"))
    G1 = Val(GetVar(App.Path & "\Colors.txt", "NEWCHAR", "G"))
    B1 = Val(GetVar(App.Path & "\Colors.txt", "NEWCHAR", "B"))
    frmNewChar.optMale.BackColor = RGB(R1, G1, B1)
    frmNewChar.optFemale.BackColor = RGB(R1, G1, B1)
    
    Call SetStatus("Checking status...")
    DoEvents
    
    ' Make sure we set that we aren't in the game
    InGame = False
    GettingMap = True
    InEditor = False
    InSpeechEditor = False
    InItemsEditor = False
    InNpcEditor = False
    InShopEditor = False
    InElementEditor = False
    InEmoticonEditor = False
    InArrowEditor = False
    InSpawnEditor = False
    
    sDc = DD_ItemSurf.GetDC
    With frmMirage.picItems
        .Width = DDSD_Item.lWidth
        .Height = DDSD_Item.lHeight
        .Cls
        Call BitBlt(.hDC, 0, 0, .Width, .Height, sDc, 0, 0, SRCCOPY)
    End With
    Call DD_ItemSurf.ReleaseDC(sDc)
    
    'frmMirage.picItems.Picture = LoadPicture(App.Path & "\GFX\items.bmp")
    
    sDc = DD_SpriteSurf.GetDC
    With frmSpriteChange.picSprite
        .Width = DDSD_Sprite.lWidth
        .Height = DDSD_Sprite.lHeight
        .Cls
        Call BitBlt(.hDC, 0, 0, .Width, .Height, sDc, 0, 0, SRCCOPY)
    End With
    Call DD_SpriteSurf.ReleaseDC(sDc)
    
    sDc = DD_Icon.GetDC
    With frmMirage.picSpellIcons
        .Width = DDSD_Icon.lWidth
        .Height = DDSD_Icon.lHeight
        .Cls
        Call BitBlt(.hDC, 0, 0, .Width, .Height, sDc, 0, 0, SRCCOPY)
    End With
    Call DD_Icon.ReleaseDC(sDc)
    
    sDc = DD_ItemSurf.GetDC
    With frmMirage.picItems
        .Width = DDSD_Item.lWidth
        .Height = DDSD_Item.lHeight
        .Cls
        Call BitBlt(.hDC, 0, 0, .Width, .Height, sDc, 0, 0, SRCCOPY)
    End With
    Call DD_ItemSurf.ReleaseDC(sDc)
    
    sDc = DD_SpriteSurf.GetDC
    With frmChars.picSpriteloader
        .Width = DDSD_Sprite.lWidth
        .Height = DDSD_Sprite.lHeight
        .Cls
        Call BitBlt(.hDC, 0, 0, .Width, .Height, sDc, 0, 0, SRCCOPY)
    End With
    Call DD_SpriteSurf.ReleaseDC(sDc)
    
    Call SetStatus("Initializing TCP Settings...")
    DoEvents
    
    Call SetStatus("Complete.")
    Call TcpInit
    PlayMidi (GetVar(App.Path & "\config.ini", "CONFIG", "TitleBGM"))
    frmMainMenu.Visible = True
    frmSendGetData.Visible = False
    MiniMap = False
End Sub

Sub SetStatus(ByVal Caption As String)
Dim s As String
  
    s = vbNewLine & Caption
    frmSendGetData.txtStatus.SelText = s
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
            If ConnectToServer = True Then
                Call SetStatus("Connected, getting available classes...")
                Call SendGetClasses
            End If
            
        Case MENU_STATE_ADDCHAR
            frmNewChar.Visible = False
            If ConnectToServer = True Then
                Call SetStatus("Connected, sending character addition data...")
                If frmNewChar.optMale.Value = True Then
                    Call SendAddChar(frmNewChar.txtName, 0, frmNewChar.cmbClass.ListIndex + 1, frmChars.lstChars.ListIndex + 1)
                Else
                    Call SendAddChar(frmNewChar.txtName, 1, frmNewChar.cmbClass.ListIndex + 1, frmChars.lstChars.ListIndex + 1)
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
Dim x As Long
Dim y As Long
Dim i As Long
Dim rec_back As RECT
Dim sDc As Long
Dim FlashCntr As Long
Dim FlashSwitch As Byte
    
    ' Set the focus
    frmMirage.picScreen.SetFocus
    
    ' Set font
    Call SetFont("Fixedsys", 18, 0, 0, 0, 0)
    ' Fixedsys's size can't be changed and bold doesn't seem to work
                
    ' Used for calculating fps
    TickFPS = GetTickCount
    FPS = 0
    FlashCntr = GetTickCount
    FlashSwitch = 0
    
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
                Qq = Player(MyIndex).Inv(Q + 1).Num
               
                If frmMirage.picInv(Q).Picture <> LoadPicture() Then
                    frmMirage.picInv(Q).Picture = LoadPicture()
                Else
                    If Qq = 0 Then
                        frmMirage.picInv(Q).Picture = LoadPicture()
                    Else
                     sDc = DD_ItemSurf.GetDC
                        Call BitBlt(frmMirage.picInv(Q).hDC, 0, 0, PIC_X, PIC_Y, frmMirage.picItems.hDC, (Item(Qq).pic - Int(Item(Qq).pic / 6) * 6) * PIC_X, Int(Item(Qq).pic / 6) * PIC_Y, SRCCOPY)
                     DD_ItemSurf.ReleaseDC (sDc)
                    End If
                End If
            Next Q
        End If
        
        ' Icons
        Dim M As Long
        Dim Mm As Long
        Dim TI As Long
               
        If GetTickCount > TI + 500 And frmMirage.picPlayerSpells.Visible = True Then
            For M = 0 To MAX_PLAYER_SPELLS - 1
                Mm = Player(MyIndex).Spell(M + 1)
               
                If frmMirage.picSpell(M).Picture <> LoadPicture() Then
                    frmMirage.picSpell(M).Picture = LoadPicture()
                Else
                    If Mm = 0 Then
                        frmMirage.picSpell(M).Picture = LoadPicture()
                    Else
                        Call BitBlt(frmMirage.picSpell(M).hDC, 0, 0, PIC_X, PIC_Y, frmMirage.picSpellIcons.hDC, (Spell(Mm).pic - Int(Spell(Mm).pic / 6) * 6) * PIC_X, Int(Spell(Mm).pic / 6) * PIC_Y, SRCCOPY)
                    End If
                End If
            Next M
        End If
                        
        NewX = roundUp(SCREEN_X / 2)
        NewY = roundUp(SCREEN_Y / 2)
       
        NewPlayerY = Player(MyIndex).y - NewY
        NewPlayerX = Player(MyIndex).x - NewX
       
        NewX = NewX * PIC_X
        NewY = NewY * PIC_Y
       
        NewXOffset = Player(MyIndex).XOffset
        NewYOffset = Player(MyIndex).YOffset

        If Player(MyIndex).y <= roundUp(SCREEN_Y / 2) Then
            NewY = Player(MyIndex).y * PIC_Y + Player(MyIndex).YOffset
            NewYOffset = 0
            NewPlayerY = 0
            If Player(MyIndex).y = roundUp(SCREEN_Y / 2) And Player(MyIndex).Dir = DIR_UP Then
                NewPlayerY = Player(MyIndex).y - roundUp(SCREEN_Y / 2)
                NewY = roundUp(SCREEN_Y / 2) * PIC_Y
                NewYOffset = Player(MyIndex).YOffset
            End If
        ElseIf Player(MyIndex).y >= MAX_MAPY - roundDown(SCREEN_Y / 2) Then
            NewY = (Player(MyIndex).y - (MAX_MAPY - SCREEN_Y)) * PIC_Y + Player(MyIndex).YOffset
            NewYOffset = 0
            NewPlayerY = MAX_MAPY - SCREEN_Y
            If Player(MyIndex).y = MAX_MAPY - roundDown(SCREEN_Y / 2) And Player(MyIndex).Dir = DIR_DOWN Then
                NewPlayerY = Player(MyIndex).y - roundUp(SCREEN_Y / 2)
                NewY = roundUp(SCREEN_Y / 2) * PIC_Y
                NewYOffset = Player(MyIndex).YOffset
            End If
        End If
       
        If Player(MyIndex).x <= roundUp(SCREEN_X / 2) Then
            NewX = Player(MyIndex).x * PIC_Y + Player(MyIndex).XOffset
            NewXOffset = 0
            NewPlayerX = 0
            If Player(MyIndex).x = roundUp(SCREEN_X / 2) And Player(MyIndex).Dir = DIR_LEFT Then
                NewPlayerX = Player(MyIndex).x - roundUp(SCREEN_X / 2)
                NewX = roundUp(SCREEN_X / 2) * PIC_X
                NewXOffset = Player(MyIndex).XOffset
            End If
        ElseIf Player(MyIndex).x >= MAX_MAPX - roundDown(SCREEN_X / 2) Then
            NewX = (Player(MyIndex).x - (MAX_MAPX - SCREEN_X)) * PIC_Y + Player(MyIndex).XOffset
            NewXOffset = 0
            NewPlayerX = MAX_MAPX - SCREEN_X
            If Player(MyIndex).x = MAX_MAPX - roundDown(SCREEN_X / 2) And Player(MyIndex).Dir = DIR_RIGHT Then
                NewPlayerX = Player(MyIndex).x - roundUp(SCREEN_X / 2)
                NewX = roundUp(SCREEN_X / 2) * PIC_X
                NewXOffset = Player(MyIndex).XOffset
            End If
        End If
       
        sx = 32
        If MAX_MAPX = SCREEN_X And MAX_MAPY = SCREEN_Y Then
            NewX = Player(MyIndex).x * PIC_X + Player(MyIndex).XOffset
            NewXOffset = 0
            NewPlayerX = 0
            NewY = Player(MyIndex).y * PIC_Y + Player(MyIndex).YOffset
            NewYOffset = 0
            NewPlayerY = 0
            sx = 0
        End If
        
        ' Blit out tiles layers ground/anim1/anim2
        For y = 0 To MAX_MAPY
            For x = 0 To MAX_MAPX
                Call BltTile(x, y)
            Next x
        Next y
       
    If ScreenMode = 0 Then
        ' Blit out the items
        For i = 1 To MAX_MAP_ITEMS
            If MapItem(i).Num > 0 Then
                Call BltItem(i)
            End If
        Next i
        
        ' Blit out NPC hp bars
        For i = 1 To MAX_MAP_NPCS
            If GetTickCount < MapNpc(i).LastAttack + 5000 Then
                Call BltNpcBars(i)
            End If
        Next i
              
        ' Blit players bar
        For i = 1 To MAX_PLAYERS
            If IsPlaying(i) Then
                If GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                    If GetTickCount < Player(MyIndex).LastAttack + 5000 Then
                        Call BltPlayerBars(i)
                    End If
                End If
                If Player(i).Pet.Map = GetPlayerMap(MyIndex) And Player(i).Pet.Alive = YES Then
                    If GetTickCount < Player(MyIndex).Pet.LastAttack + 5000 Then
                        Call BltPetBars(i)
                    End If
                End If
            End If
        Next i
        
        ' Blit out the sprite change attribute
        For y = 0 To MAX_MAPY
            For x = 0 To MAX_MAPX
                Call BltSpriteChange(x, y)
            Next x
        Next y
        
        ' Blit out the furniture attribute
        For y = 0 To MAX_MAPY
            For x = 0 To MAX_MAPX
                Call BltFurniture(x, y)
            Next x
        Next y

        
        ' Blit out arrows
        For i = 1 To MAX_PLAYERS
            If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                Call BltArrow(i)
            End If
        Next i
        
        For i = 1 To MAX_PLAYERS
            If IsPlaying(i) Then
                If Player(i).CorpseMap = GetPlayerMap(MyIndex) Then
                    Call BltPlayerCorpse(i)
                End If
            End If
        Next i
        
        ' Blit out the npcs
        For i = 1 To MAX_MAP_NPCS
            Call BltNpc(i)
        Next i
        
        ' Blit out players
        For i = 1 To MAX_PLAYERS
            If IsPlaying(i) Then
                If GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                    If Player(i).Pet.Alive = YES Then
                        Call BltPet(i)
                    End If
                    Call BltPlayer(i)
                End If
            End If
        Next i
        
        If SIZE_Y > PIC_Y Then
            For i = 1 To MAX_PLAYERS
                If IsPlaying(i) Then
                    If GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                    If Player(i).Pet.Alive = YES Then
                            Call BltPetTop(i)
                        End If
                        Call BltPlayerTop(i)
                    End If
                End If
            Next i
        End If
        
        ' Blit the spells
        For i = 1 To MAX_PLAYERS
            If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                Call BltSpell(i)
            End If
        Next i
        
        ' Blit out the sprite change attribute
        For y = 0 To MAX_MAPY
            For x = 0 To MAX_MAPX
                Call BltSpriteChange2(x, y)
            Next x
        Next y
        
        ' Blit out the npc's top
        For i = 1 To MAX_MAP_NPCS
            Call BltNpcTop(i)
        Next i
    End If
                
    ' Blit out tile layer fringe
    For y = 0 To MAX_MAPY
        For x = 0 To MAX_MAPX
            Call BltFringeTile(x, y)
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
        If frmMapEditor.chkDayNight.Value = 1 And InEditor = True Then
            Call Night
        End If
        If Map(GetPlayerMap(MyIndex)).Indoors = 0 Then Call BltWeather
    End If

    If InEditor = True And Val(GetVar(App.Path & "\config.ini", "CONFIG", "MapGrid")) = 1 Then
        For y = 0 To MAX_MAPY
            For x = 0 To MAX_MAPX
                Call BltTile2(x * 32, y * 32, 0, 6)
            Next x
        Next y
    End If
End If

    If InEditor = True And SelectorWidth <> 0 And SelectorHeight <> 0 And frmMapEditor.fraLayers.Visible = True And GetTickCount Mod 1000 < 700 Then
        For y = 0 To SelectorHeight - 1
            For x = 0 To SelectorWidth - 1
                Call BltTile2(MouseX + (x * PIC_X), MouseY + (y * PIC_Y), ((EditorTileY + y) * TilesInSheets) + (EditorTileX + x), EditorSet)
            Next x
        Next y
    End If
    
    ' Lock the backbuffer so we can draw text and names
    TexthDC = DD_BackBuffer.GetDC
    If GettingMap = False Then
        If ScreenMode = 0 Then
            If Val(GetVar(App.Path & "\config.ini", "CONFIG", "NPCDamage")) = 1 Then
                If Val(GetVar(App.Path & "\config.ini", "CONFIG", "PlayerName")) = 0 Then
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
            
            If Val(GetVar(App.Path & "\config.ini", "CONFIG", "PlayerDamage")) = 1 Then
                If NPCWho > 0 Then
                    If MapNpc(NPCWho).Num > 0 Then
                        If Val(GetVar(App.Path & "\config.ini", "CONFIG", "NPCName")) = 0 Then
                            If Npc(MapNpc(NPCWho).Num).Big = 0 Then
                                If GetTickCount < DmgTime + 2000 Then
                                    Call DrawText(TexthDC, (MapNpc(NPCWho).x - NewPlayerX) * PIC_X + sx + (Int(Len(DmgDamage)) / 2) * 3 + MapNpc(NPCWho).XOffset - NewXOffset, (MapNpc(NPCWho).y - NewPlayerY) * PIC_Y + sx - 20 + MapNpc(NPCWho).YOffset - NewYOffset - iii, DmgDamage, QBColor(White))
                                End If
                            Else
                                If GetTickCount < DmgTime + 2000 Then
                                    Call DrawText(TexthDC, (MapNpc(NPCWho).x - NewPlayerX) * PIC_X + sx + (Int(Len(DmgDamage)) / 2) * 3 + MapNpc(NPCWho).XOffset - NewXOffset, (MapNpc(NPCWho).y - NewPlayerY) * PIC_Y + sx - 47 + MapNpc(NPCWho).YOffset - NewYOffset - iii, DmgDamage, QBColor(White))
                                End If
                            End If
                        Else
                            If Npc(MapNpc(NPCWho).Num).Big = 0 Then
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
            
            'Draw NPC Names
            If Val(GetVar(App.Path & "\config.ini", "CONFIG", "NPCName")) = 1 Then
                For i = LBound(MapNpc) To UBound(MapNpc)
                    If MapNpc(i).Num > 0 Then
                        Call BltMapNPCName(i)
                    End If
                Next i
            End If
            
            ' Draw Player Names
            If Val(GetVar(App.Path & "\config.ini", "CONFIG", "PlayerName")) = 1 Then
                For i = 1 To MAX_PLAYERS
                    If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                        Call BltPlayerGuildName(i)
                        If GetTickCount > FlashCntr + 250 Then
                            If FlashSwitch = 1 Then
                                FlashSwitch = 0
                            Else
                                FlashSwitch = 1
                            End If
                            FlashCntr = GetTickCount
                        End If
                        Call BltPlayerName(i, FlashSwitch)
                        If Player(i).Pet.Alive = YES And Player(i).Pet.Map = GetPlayerMap(MyIndex) Then
                            Call BltPetName(i)
                        End If
                        ' XCORPSEX
                        If Player(i).CorpseMap = GetPlayerMap(MyIndex) Then
                       Call BltPlayerCorpseName(i)
                        End If
                        ' XCORPSEX
                    End If
                Next i
            End If
     
            ' speech bubble stuffs
            If Val(GetVar(App.Path & "\config.ini", "CONFIG", "SpeechBubbles")) = 1 Then
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
                            If .Type = TILE_TYPE_BANK Then Call DrawText(TexthDC, x * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "BANK", QBColor(BrightRed))
                            If .Type = TILE_TYPE_HOUSE_BUY Then Call DrawText(TexthDC, x * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "PHB", QBColor(Yellow))
                            If .Type = TILE_TYPE_HOUSE Then Call DrawText(TexthDC, x * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "PH", QBColor(White))
                            If .Type = TILE_TYPE_FURNITURE Then Call DrawText(TexthDC, x * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "F", QBColor(BrightRed))
                            If .Light > 0 Then Call DrawText(TexthDC, x * PIC_X + sx + 18 - (NewPlayerX * PIC_X) - NewXOffset, y * PIC_Y + sx + 14 - (NewPlayerY * PIC_Y) - NewYOffset, "L", QBColor(Yellow))
                        End With
                        
                        If InSpawnEditor Then
                            For i = 1 To MAX_MAP_NPCS
                                If TempNpcSpawn(i).Used = YES Then
                                    If x = TempNpcSpawn(i).x And y = TempNpcSpawn(i).y Then
                                        Call DrawText(TexthDC, x * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, i, QBColor(White))
                                    End If
                                End If
                            Next i
                        End If
                    Next x
                Next y
            End If
            
            ' Blit the text they are putting in
            'MyText = frmMirage.txtMyTextBox.Text
            'frmMirage.txtMyTextBox.Text = MyText
            
            'If Len(MyText) > 4 Then
                'frmMirage.txtMyTextBox.SelStart = Len(frmMirage.txtMyTextBox.Text) + 1
            'End If
                    
            ' Draw map name
            If Map(GetPlayerMap(MyIndex)).Moral = MAP_MORAL_NONE Then
                Call DrawText(TexthDC, Int((20.5) * PIC_X / 2) - (Int(Len(Trim(Map(GetPlayerMap(MyIndex)).name)) / 2) * 8) + sx, 2 + sx, Trim(Map(GetPlayerMap(MyIndex)).name), QBColor(BrightRed))
            ElseIf Map(GetPlayerMap(MyIndex)).Moral = MAP_MORAL_SAFE Then
                Call DrawText(TexthDC, Int((20.5) * PIC_X / 2) - (Int(Len(Trim(Map(GetPlayerMap(MyIndex)).name)) / 2) * 8) + sx, 2 + sx, Trim(Map(GetPlayerMap(MyIndex)).name), QBColor(White))
            ElseIf Map(GetPlayerMap(MyIndex)).Moral = MAP_MORAL_NO_PENALTY Then
                Call DrawText(TexthDC, Int((20.5) * PIC_X / 2) - (Int(Len(Trim(Map(GetPlayerMap(MyIndex)).name)) / 2) * 8) + sx, 2 + sx, Trim(Map(GetPlayerMap(MyIndex)).name), QBColor(Black))
                ElseIf Map(GetPlayerMap(MyIndex)).Moral = MAP_MORAL_HOUSE Then
                Call DrawText(TexthDC, Int((20.5) * PIC_X / 2) - (Int(Len(Trim(Map(GetPlayerMap(MyIndex)).name)) / 2) * 8) + sx, 2 + sx, Trim(Map(GetPlayerMap(MyIndex)).name), QBColor(Yellow))
            End If
            
            ' Battle messages
            For i = 1 To MAX_BLT_LINE
                If BattlePMsg(i).Index > 0 Then
                    If BattlePMsg(i).Time + 7000 > GetTickCount Then
                        Call DrawText(TexthDC, 1 + sx, BattlePMsg(i).y + frmMirage.picScreen.Height - 15 + sx, Trim(BattlePMsg(i).Msg), QBColor(BattlePMsg(i).Color))
                    Else
                        BattlePMsg(i).Done = 0
                    End If
                End If
                
                If BattleMMsg(i).Index > 0 Then
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
        
        ' Blit out MiniMap
        If MiniMap = True Then
            Call BltMiniMap
        End If
        
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
        
        If XToGo <> -1 Or YToGo <> -1 Then
            Dim XDif As Long
            Dim YDif As Long
            
            XDif = Abs(GetPlayerX(MyIndex) - XToGo)
            YDif = Abs(GetPlayerY(MyIndex) - YToGo)
            
            If XToGo = GetPlayerX(MyIndex) Or XToGo = -1 Then
                XToGo = -1
                XDif = 0
            Else
                XDif = Abs(GetPlayerX(MyIndex) - XToGo)
            End If
            
            If YToGo = GetPlayerY(MyIndex) Or YToGo = -1 Then
                YToGo = -1
                YDif = 0
            Else
                YDif = Abs(GetPlayerY(MyIndex) - YToGo)
            End If
            
            Debug.Print (XDif & " " & YDif)
            
            If XDif > YDif Then
                If GetPlayerX(MyIndex) - XToGo > 0 Then
                    DirLeft = True
                Else
                    DirRight = True
                End If
            End If
            
            If YDif > XDif Then
                If GetPlayerY(MyIndex) - YToGo > 0 Then
                    DirUp = True
                Else
                    DirDown = True
                End If
            End If
            
            If XDif = YDif And XDif <> 0 And YDif <> 0 Then
                ' I'll be nice and give you the non-directional movement code
                'If Int(Rnd * 2) = 0 Then
                If GetPlayerX(MyIndex) - XToGo > 0 Then
                    DirLeft = True
                Else
                    DirRight = True
                End If
                ' Else
                If GetPlayerY(MyIndex) - YToGo > 0 Then
                    DirUp = True
                Else
                    DirDown = True
                End If
                'End If
            End If
        End If
        
        ' Check if player is trying to move
        Call CheckMovement
        
        ' Check to see if player is trying to attack
        Call CheckAttack
        
        ' Process player and pet movements (actually move them)
        For i = 1 To MAX_PLAYERS
            If IsPlaying(i) Then
                Call ProcessMovement(i)
                If Player(i).Pet.Alive = YES Then
                    Call ProcessPetMovement(i)
                End If
            End If
        Next i
        
        ' Process npc movements (actually move them)
        For i = 1 To MAX_MAP_NPCS
            
            If Map(GetPlayerMap(MyIndex)).Npc(i) > 0 Then
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

Sub BltTile(ByVal x As Long, ByVal y As Long)
Dim Ground As Long
Dim Mask As Long
Dim Anim As Long
Dim Mask2 As Long
Dim M2Anim As Long
Dim GroundTileSet As Long
Dim MaskTileSet As Long
Dim AnimTileSet As Long
Dim Mask2TileSet As Long
Dim M2AnimTileSet As Long

    Ground = Map(GetPlayerMap(MyIndex)).Tile(x, y).Ground
    Mask = Map(GetPlayerMap(MyIndex)).Tile(x, y).Mask
    Anim = Map(GetPlayerMap(MyIndex)).Tile(x, y).Anim
    Mask2 = Map(GetPlayerMap(MyIndex)).Tile(x, y).Mask2
    M2Anim = Map(GetPlayerMap(MyIndex)).Tile(x, y).M2Anim
    
    If TempTile(x, y).Ground <> 0 Then Ground = TempTile(x, y).Ground
    If TempTile(x, y).Mask <> 0 Then Mask = TempTile(x, y).Mask
    If TempTile(x, y).Anim <> 0 Then Anim = TempTile(x, y).Anim
    If TempTile(x, y).Mask2 <> 0 Then Mask2 = TempTile(x, y).Mask2
    If TempTile(x, y).M2Anim <> 0 Then M2Anim = TempTile(x, y).M2Anim
    
    GroundTileSet = Map(GetPlayerMap(MyIndex)).Tile(x, y).GroundSet
    MaskTileSet = Map(GetPlayerMap(MyIndex)).Tile(x, y).MaskSet
    AnimTileSet = Map(GetPlayerMap(MyIndex)).Tile(x, y).AnimSet
    Mask2TileSet = Map(GetPlayerMap(MyIndex)).Tile(x, y).Mask2Set
    M2AnimTileSet = Map(GetPlayerMap(MyIndex)).Tile(x, y).M2AnimSet
    
    ' If (GroundTileSet >= 0 And TileFile(GroundTileSet) = 0) Or (MaskTileSet >= 0 And TileFile(MaskTileSet) = 0) Or (AnimTileSet >= 0 And TileFile(AnimTileSet) = 0) Or (Mask2TileSet >= 0 And TileFile(Mask2TileSet) = 0) Or (M2AnimTileSet >= 0 And TileFile(M2AnimTileSet) = 0) Then Exit Sub
    
    ' Only used if ever want to switch to blt rather then bltfast
    With rec_pos
        .Top = (y - NewPlayerY) * PIC_Y + sx - NewYOffset
        .Bottom = .Top + PIC_Y
        .Left = (x - NewPlayerX) * PIC_X + sx - NewXOffset
        .Right = .Left + PIC_X
    End With
    
    If GroundTileSet < 0 Then
        GroundTileSet = 0
        Ground = 0
    End If
    
    rec.Top = Int(Ground / TilesInSheets) * PIC_Y
    rec.Bottom = rec.Top + PIC_Y
    rec.Left = (Ground - Int(Ground / TilesInSheets) * TilesInSheets) * PIC_X
    rec.Right = rec.Left + PIC_X
    Call DD_BackBuffer.Blt(rec_pos, DD_TileSurf(GroundTileSet), rec, DDBLT_WAIT)
    'Call DD_BackBuffer.BltFast((X - NewPlayerX) * PIC_X - NewXOffset, (Y - NewPlayerY) * PIC_Y - NewYOffset, DD_TileSurf, rec, DDBLTFAST_WAIT)
    
    ' Is there an animation tile to plot?
    If (MapAnim = 0) Or (Anim <= 0) Then
        If Mask > 0 And MaskTileSet >= 0 And TempTile(x, y).DoorOpen = NO Then
            rec.Top = Int(Mask / TilesInSheets) * PIC_Y
            rec.Bottom = rec.Top + PIC_Y
            rec.Left = (Mask - Int(Mask / TilesInSheets) * TilesInSheets) * PIC_X
            rec.Right = rec.Left + PIC_X
            Call DD_BackBuffer.Blt(rec_pos, DD_TileSurf(MaskTileSet), rec, DDBLT_WAIT Or DDBLT_KEYSRC)
            'Call DD_BackBuffer.BltFast((X - NewPlayerX) * PIC_X - NewXOffset, (Y - NewPlayerY) * PIC_Y - NewYOffset, DD_TileSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    Else
        If Anim > 0 And AnimTileSet >= 0 Then
            rec.Top = Int(Anim / TilesInSheets) * PIC_Y
            rec.Bottom = rec.Top + PIC_Y
            rec.Left = (Anim - Int(Anim / TilesInSheets) * TilesInSheets) * PIC_X
            rec.Right = rec.Left + PIC_X
            Call DD_BackBuffer.Blt(rec_pos, DD_TileSurf(AnimTileSet), rec, DDBLT_WAIT Or DDBLT_KEYSRC)
            'Call DD_BackBuffer.BltFast((X - NewPlayerX) * PIC_X - NewXOffset, (Y - NewPlayerY) * PIC_Y - NewYOffset, DD_TileSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    End If
    
    ' Is there a second animation tile to plot?
    If (MapAnim = 0) Or (M2Anim <= 0) Then
        If Mask2 > 0 And Mask2TileSet >= 0 Then
            rec.Top = Int(Mask2 / TilesInSheets) * PIC_Y
            rec.Bottom = rec.Top + PIC_Y
            rec.Left = (Mask2 - Int(Mask2 / TilesInSheets) * TilesInSheets) * PIC_X
            rec.Right = rec.Left + PIC_X
            Call DD_BackBuffer.Blt(rec_pos, DD_TileSurf(Mask2TileSet), rec, DDBLT_WAIT Or DDBLT_KEYSRC)
            'Call DD_BackBuffer.BltFast((X - NewPlayerX) * PIC_X - NewXOffset, (Y - NewPlayerY) * PIC_Y - NewYOffset, DD_TileSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    Else
        If M2Anim > 0 And M2AnimTileSet >= 0 Then
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
    
    rec.Top = Int(Item(MapItem(ItemNum).Num).pic / 6) * PIC_Y
    rec.Bottom = rec.Top + PIC_Y
    rec.Left = (Item(MapItem(ItemNum).Num).pic - Int(Item(MapItem(ItemNum).Num).pic / 6) * 6) * PIC_X
    rec.Right = rec.Left + PIC_X
    
    'Call DD_BackBuffer.Blt(rec_pos, DD_ItemSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
    Call DD_BackBuffer.BltFast((MapItem(ItemNum).x - NewPlayerX) * PIC_X + sx - NewXOffset, (MapItem(ItemNum).y - NewPlayerY) * PIC_Y + sx - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End Sub

Sub BltFringeTile(ByVal x As Long, ByVal y As Long)
Dim Fringe As Long
Dim FAnim As Long
Dim Fringe2 As Long
Dim F2Anim As Long
Dim FringeTileSet As Long
Dim FAnimTileSet As Long
Dim Fringe2TileSet As Long
Dim F2AnimTileSet As Long

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
    
    If TempTile(x, y).Fringe <> 0 Then Fringe = TempTile(x, y).Fringe
    If TempTile(x, y).FAnim <> 0 Then FAnim = TempTile(x, y).FAnim
    If TempTile(x, y).Fringe2 <> 0 Then Fringe2 = TempTile(x, y).Fringe2
    If TempTile(x, y).F2Anim <> 0 Then F2Anim = TempTile(x, y).F2Anim
    
    FringeTileSet = Map(GetPlayerMap(MyIndex)).Tile(x, y).FringeSet
    FAnimTileSet = Map(GetPlayerMap(MyIndex)).Tile(x, y).FAnimSet
    Fringe2TileSet = Map(GetPlayerMap(MyIndex)).Tile(x, y).Fringe2Set
    F2AnimTileSet = Map(GetPlayerMap(MyIndex)).Tile(x, y).F2AnimSet
        
    ' If (FringeTileSet >= 0 And TileFile(FringeTileSet) = 0) Or (FAnimTileSet >= 0 And TileFile(FAnimTileSet) = 0) Or (Fringe2TileSet >= 0 And TileFile(Fringe2TileSet) = 0) Or (F2AnimTileSet >= 0 And TileFile(F2AnimTileSet) = 0) Then Exit Sub
        
    ' Is there an animation tile to plot?
    If (MapAnim = 0) Or (FAnim <= 0) Then
        If Fringe > 0 And FringeTileSet >= 0 Then
            rec.Top = Int(Fringe / TilesInSheets) * PIC_Y
            rec.Bottom = rec.Top + PIC_Y
            rec.Left = (Fringe - Int(Fringe / TilesInSheets) * TilesInSheets) * PIC_X
            rec.Right = rec.Left + PIC_X
            'Call DD_BackBuffer.Blt(rec_pos, DD_TileSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
            Call DD_BackBuffer.BltFast((x - NewPlayerX) * PIC_X + sx - NewXOffset, (y - NewPlayerY) * PIC_Y + sx - NewYOffset, DD_TileSurf(FringeTileSet), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    
    Else
        If FAnim > 0 And FAnimTileSet >= 0 Then
            rec.Top = Int(FAnim / TilesInSheets) * PIC_Y
            rec.Bottom = rec.Top + PIC_Y
            rec.Left = (FAnim - Int(FAnim / TilesInSheets) * TilesInSheets) * PIC_X
            rec.Right = rec.Left + PIC_X
            'Call DD_BackBuffer.Blt(rec_pos, DD_TileSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
            Call DD_BackBuffer.BltFast((x - NewPlayerX) * PIC_X + sx - NewXOffset, (y - NewPlayerY) * PIC_Y + sx - NewYOffset, DD_TileSurf(FAnimTileSet), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
        
    End If

    ' Is there a second animation tile to plot?
    If (MapAnim = 0) Or (F2Anim <= 0) Then
        If Fringe2 > 0 And Fringe2TileSet >= 0 Then
            rec.Top = Int(Fringe2 / TilesInSheets) * PIC_Y
            rec.Bottom = rec.Top + PIC_Y
            rec.Left = (Fringe2 - Int(Fringe2 / TilesInSheets) * TilesInSheets) * PIC_X
            rec.Right = rec.Left + PIC_X
            'Call DD_BackBuffer.Blt(rec_pos, DD_TileSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
            Call DD_BackBuffer.BltFast((x - NewPlayerX) * PIC_X + sx - NewXOffset, (y - NewPlayerY) * PIC_Y + sx - NewYOffset, DD_TileSurf(Fringe2TileSet), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    Else
        If F2Anim > 0 And F2AnimTileSet >= 0 Then
            rec.Top = Int(F2Anim / TilesInSheets) * PIC_Y
            rec.Bottom = rec.Top + PIC_Y
            rec.Left = (F2Anim - Int(F2Anim / TilesInSheets) * TilesInSheets) * PIC_X
            rec.Right = rec.Left + PIC_X
            'Call DD_BackBuffer.Blt(rec_pos, DD_TileSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
            Call DD_BackBuffer.BltFast((x - NewPlayerX) * PIC_X + sx - NewXOffset, (y - NewPlayerY) * PIC_Y + sx - NewYOffset, DD_TileSurf(F2AnimTileSet), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    End If
End Sub

Sub BltPlayer(ByVal Index As Long)
Dim Anim As Byte
Dim x As Long, y As Long
Dim AttackSpeed As Long

    If GetPlayerWeaponSlot(Index) > 0 Then
        AttackSpeed = Item(Player(Index).WeaponNum).AttackSpeed
    Else
        AttackSpeed = 1000
    End If

    ' Only used if ever want to switch to blt rather then bltfast
    ' I suggest you don't use, because custom sizes won't work any longer
    With rec_pos
        .Top = GetPlayerY(Index) * PIC_Y + Player(Index).YOffset - (SIZE_Y - PIC_Y)
        .Bottom = .Top + PIC_Y
        .Left = GetPlayerX(Index) * PIC_X + Player(Index).XOffset + ((SIZE_X - PIC_X) / 2)
        .Right = .Left + PIC_X + ((SIZE_X - PIC_X) / 2)
    End With
    
    ' Check for animation
    Anim = 0
    If Player(Index).Attacking = 0 Then
        Select Case GetPlayerDir(Index)
            Case DIR_UP
                If (Player(Index).YOffset > PIC_Y / 2) Then Anim = 1
            Case DIR_DOWN
                If (Player(Index).YOffset < PIC_Y / 2 * -1) Then Anim = 1
            Case DIR_LEFT
                If (Player(Index).XOffset > PIC_Y / 2) Then Anim = 1
            Case DIR_RIGHT
                If (Player(Index).XOffset < PIC_Y / 2 * -1) Then Anim = 1
        End Select
    Else
        If Player(Index).AttackTimer + Int(AttackSpeed / 2) > GetTickCount Then
            Anim = 2
        End If
    End If
    
    ' Check to see if we want to stop making him attack
    If Player(Index).AttackTimer + AttackSpeed < GetTickCount Then
        Player(Index).Attacking = 0
        Player(Index).AttackTimer = 0
    End If
    
    x = GetPlayerX(Index) * PIC_X - (SIZE_X - PIC_X) / 2 + sx + Player(Index).XOffset
    y = GetPlayerY(Index) * PIC_Y - (SIZE_Y - PIC_Y) + sx + Player(Index).YOffset + (SIZE_Y - PIC_Y)
    
    rec.Left = (GetPlayerDir(Index) * (3 * (SIZE_X / PIC_X)) + (Anim * (SIZE_X / PIC_X))) * PIC_X
    rec.Right = rec.Left + SIZE_X
    
    If SIZE_X > PIC_X Then
        If x < 0 Then
            x = Player(Index).XOffset + sx + ((SIZE_X - PIC_X) / 2)
            If GetPlayerDir(Index) = DIR_RIGHT And Player(Index).MovingH > 0 Then
                rec.Left = rec.Left - Player(Index).XOffset
            Else
                rec.Left = rec.Left - Player(Index).XOffset + ((SIZE_X - PIC_X) / 2)
            End If
        End If
        
        If x > MAX_MAPX * 32 Then
            x = MAX_MAPX * 32 + sx - ((SIZE_X - PIC_X) / 2) + Player(Index).XOffset
            If GetPlayerDir(Index) = DIR_LEFT And Player(Index).MovingH > 0 Then
                rec.Right = rec.Right + Player(Index).XOffset
            Else
         rec.Right = rec.Right + Player(Index).XOffset - ((SIZE_X - PIC_X) / 2)
            End If
        End If
    End If
    
    
    If GetPlayerDir(Index) = DIR_UP Then
    'If PAPERDOLL = 1 Then
        If Player(Index).WeaponNum > 0 Then
            rec.Top = (Int(Item(Player(Index).WeaponNum).pic / 6) + (SIZE_Y / PIC_Y)) * PIC_Y
            rec.Bottom = rec.Top + PIC_Y
            Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
        
        
        If Player(Index).ShieldNum > 0 Then
            rec.Top = (Int(Item(Player(Index).ShieldNum).pic / 6) + (SIZE_Y / PIC_Y)) * PIC_Y
            rec.Bottom = rec.Top + PIC_Y
            Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    End If
'End If
    
    rec.Top = GetPlayerSprite(Index) * SIZE_Y + (SIZE_Y - PIC_Y)
    rec.Bottom = rec.Top + PIC_Y
    
    Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)

'If PAPERDOLL = 1 Then
    If Player(Index).ArmorNum > 0 Then
        rec.Top = (Int(Item(Player(Index).ArmorNum).pic / 6) + (SIZE_Y / PIC_Y)) * PIC_Y
        rec.Bottom = rec.Top + PIC_Y
        Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    End If
    
    If Player(Index).HelmetNum > 0 Then
        rec.Top = (Int(Item(Player(Index).HelmetNum).pic / 6) + (SIZE_Y / PIC_Y)) * PIC_Y
        rec.Bottom = rec.Top + PIC_Y
        Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    End If
    
    If Player(Index).LegsNum > 0 Then
        rec.Top = (Int(Item(Player(Index).LegsNum).pic / 6) + (SIZE_Y / PIC_Y)) * PIC_Y
        rec.Bottom = rec.Top + PIC_Y
        Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    End If
'End If
    
    If GetPlayerDir(Index) <> DIR_UP Then
    'If PAPERDOLL = 1 Then
        If Player(Index).ShieldNum > 0 Then
            rec.Top = (Int(Item(Player(Index).ShieldNum).pic / 6) + (SIZE_Y / PIC_Y)) * PIC_Y
            rec.Bottom = rec.Top + PIC_Y
            Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
        
        If Player(Index).WeaponNum > 0 Then
            rec.Top = (Int(Item(Player(Index).WeaponNum).pic / 6) + (SIZE_Y / PIC_Y)) * PIC_Y
            rec.Bottom = rec.Top + PIC_Y
            Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    End If
'End If
End Sub

Sub BltPlayerTop(ByVal Index As Long)
Dim Anim As Byte
Dim x As Long, y As Long
Dim AttackSpeed As Long

    If GetPlayerWeaponSlot(Index) > 0 Then
        AttackSpeed = Item(Player(Index).WeaponNum).AttackSpeed
    Else
        AttackSpeed = 1000
    End If

    ' Only used if ever want to switch to blt rather then bltfast
    ' I suggest you don't use, because custom sizes won't work any longer
    With rec_pos
        .Top = GetPlayerY(Index) * PIC_Y + Player(Index).YOffset - (SIZE_Y - PIC_Y)
        .Bottom = .Top + PIC_Y
        .Left = GetPlayerX(Index) * PIC_X + Player(Index).XOffset + ((SIZE_X - PIC_X) / 2)
        .Right = .Left + PIC_X + ((SIZE_X - PIC_X) / 2)
    End With
    
    ' Check for animation
    Anim = 0
    If Player(Index).Attacking = 0 Then
        Select Case GetPlayerDir(Index)
            Case DIR_UP
                If (Player(Index).YOffset > PIC_Y / 2) Then Anim = 1
            Case DIR_DOWN
                If (Player(Index).YOffset < PIC_Y / 2 * -1) Then Anim = 1
            Case DIR_LEFT
                If (Player(Index).XOffset > PIC_Y / 2) Then Anim = 1
            Case DIR_RIGHT
                If (Player(Index).XOffset < PIC_Y / 2 * -1) Then Anim = 1
        End Select
    Else
        If Player(Index).AttackTimer + Int(AttackSpeed / 2) > GetTickCount Then
            Anim = 2
        End If
    End If
    
    ' Check to see if we want to stop making him attack
    If Player(Index).AttackTimer + AttackSpeed < GetTickCount Then
        Player(Index).Attacking = 0
        Player(Index).AttackTimer = 0
    End If
    
    x = GetPlayerX(Index) * PIC_X - (SIZE_X - PIC_X) / 2 + sx + Player(Index).XOffset
    y = GetPlayerY(Index) * PIC_Y - (SIZE_Y - PIC_Y) + sx + Player(Index).YOffset
    
    rec.Left = (GetPlayerDir(Index) * (3 * (SIZE_X / PIC_X)) + (Anim * (SIZE_X / PIC_X))) * PIC_X
    rec.Right = rec.Left + SIZE_X
    
    If x < 0 Then
        x = Player(Index).XOffset + sx + ((SIZE_X - PIC_X) / 2)
        If GetPlayerDir(Index) = DIR_RIGHT And Player(Index).MovingH > 0 Then
            rec.Left = rec.Left - Player(Index).XOffset
        Else
            rec.Left = rec.Left - Player(Index).XOffset + ((SIZE_X - PIC_X) / 2)
        End If
    End If
    
    If x > MAX_MAPX * 32 Then
        x = MAX_MAPX * 32 + sx - ((SIZE_X - PIC_X) / 2) + Player(Index).XOffset
        If GetPlayerDir(Index) = DIR_LEFT And Player(Index).MovingH > 0 Then
            rec.Right = rec.Right + Player(Index).XOffset
        Else
            rec.Right = rec.Right + Player(Index).XOffset - ((SIZE_X - PIC_X) / 2)
        End If
    End If
    

    If GetPlayerDir(Index) = DIR_UP Then
    'If PAPERDOLL = 1 Then
        If Player(Index).WeaponNum > 0 Then
            rec.Top = Int(Item(Player(Index).WeaponNum).pic / 6) * PIC_Y + (SIZE_Y - PIC_Y)
            rec.Bottom = rec.Top + (SIZE_Y - PIC_Y)
            
            If y < 0 Then
                y = 0
                If GetPlayerDir(Index) = DIR_DOWN And Player(Index).MovingV > 0 Then
                    rec.Top = rec.Top - Player(Index).YOffset
                Else
                    rec.Top = rec.Top - Player(Index).YOffset + (SIZE_Y - PIC_Y)
                End If
            End If
    
            Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    
        If Player(Index).ShieldNum > 0 Then
            rec.Top = Int(Item(Player(Index).ShieldNum).pic / 6) * PIC_Y + (SIZE_Y - PIC_Y)
            rec.Bottom = rec.Top + (SIZE_Y - PIC_Y)
            
            y = GetPlayerY(Index) * PIC_Y - (SIZE_Y - PIC_Y) + sx + Player(Index).YOffset
            
            If y < 0 Then
                y = 0
                If GetPlayerDir(Index) = DIR_DOWN And Player(Index).MovingV > 0 Then
                    rec.Top = rec.Top - Player(Index).YOffset
                Else
                    rec.Top = rec.Top - Player(Index).YOffset + (SIZE_Y - PIC_Y)
                End If
            End If
            
            Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    End If
'End If

    rec.Top = GetPlayerSprite(Index) * SIZE_Y
    rec.Bottom = rec.Top + (SIZE_Y - PIC_Y)
    
    y = GetPlayerY(Index) * PIC_Y - (SIZE_Y - PIC_Y) + sx + Player(Index).YOffset
    
    If y < 0 Then
        y = 0
        If GetPlayerDir(Index) = DIR_DOWN And Player(Index).MovingV > 0 Then
            rec.Top = rec.Top - Player(Index).YOffset
        Else
            rec.Top = rec.Top - Player(Index).YOffset + (SIZE_Y - PIC_Y)
        End If
    End If
    
    Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)

'If PAPERDOLL = 1 Then
    If Player(Index).ArmorNum > 0 Then
        rec.Top = Int(Item(Player(Index).ArmorNum).pic / 6) * PIC_Y + (SIZE_Y - PIC_Y)
        rec.Bottom = rec.Top + (SIZE_Y - PIC_Y)
        
        y = GetPlayerY(Index) * PIC_Y - (SIZE_Y - PIC_Y) + sx + Player(Index).YOffset
        
        If y < 0 Then
            y = 0
            If GetPlayerDir(Index) = DIR_DOWN And Player(Index).MovingV > 0 Then
                rec.Top = rec.Top - Player(Index).YOffset
            Else
                rec.Top = rec.Top - Player(Index).YOffset + (SIZE_Y - PIC_Y)
            End If
        End If
        
        Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    End If
    
    If Player(Index).LegsNum > 0 Then
        rec.Top = Int(Item(Player(Index).LegsNum).pic / 6) * PIC_Y + (SIZE_Y - PIC_Y)
        rec.Bottom = rec.Top + (SIZE_Y - PIC_Y)
        
        y = GetPlayerY(Index) * PIC_Y - (SIZE_Y - PIC_Y) + sx + Player(Index).YOffset
        
        If y < 0 Then
            y = 0
            If GetPlayerDir(Index) = DIR_DOWN And Player(Index).MovingV > 0 Then
                rec.Top = rec.Top - Player(Index).YOffset
            Else
                rec.Top = rec.Top - Player(Index).YOffset + (SIZE_Y - PIC_Y)
            End If
        End If
        
        Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    End If
    
    If Player(Index).HelmetNum > 0 Then
        rec.Top = Int(Item(Player(Index).HelmetNum).pic / 6) * PIC_Y + (SIZE_Y - PIC_Y)
        rec.Bottom = rec.Top + (SIZE_Y - PIC_Y)
        
        y = GetPlayerY(Index) * PIC_Y - (SIZE_Y - PIC_Y) + sx + Player(Index).YOffset
        
        If y < 0 Then
            y = 0
            If GetPlayerDir(Index) = DIR_DOWN And Player(Index).MovingV > 0 Then
                rec.Top = rec.Top - Player(Index).YOffset
            Else
                rec.Top = rec.Top - Player(Index).YOffset + (SIZE_Y - PIC_Y) - 1
            End If
        End If
        
        Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    End If
'End If
    
    If GetPlayerDir(Index) <> DIR_UP Then
   ' If PAPERDOLL = 1 Then
        If Player(Index).WeaponNum > 0 Then
            rec.Top = Int(Item(Player(Index).WeaponNum).pic / 6) * PIC_Y + (SIZE_Y - PIC_Y)
            rec.Bottom = rec.Top + (SIZE_Y - PIC_Y)
            
            y = GetPlayerY(Index) * PIC_Y - (SIZE_Y - PIC_Y) + sx + Player(Index).YOffset
            
            If y < 0 Then
                y = 0
                If GetPlayerDir(Index) = DIR_DOWN And Player(Index).MovingV > 0 Then
                    rec.Top = rec.Top - Player(Index).YOffset
                Else
                    rec.Top = rec.Top - Player(Index).YOffset + (SIZE_Y - PIC_Y)
                End If
            End If
            
            Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
        
        If y = 0 Then y = GetPlayerY(Index) * PIC_Y - (SIZE_Y - PIC_Y) + sx + Player(Index).YOffset
    
        If Player(Index).ShieldNum > 0 Then
            rec.Top = Int(Item(Player(Index).ShieldNum).pic / 6) * PIC_Y + (SIZE_Y - PIC_Y)
            rec.Bottom = rec.Top + (SIZE_Y - PIC_Y)
            
            y = GetPlayerY(Index) * PIC_Y - (SIZE_Y - PIC_Y) + sx + Player(Index).YOffset
            
            If y < 0 Then
                y = 0
                If GetPlayerDir(Index) = DIR_DOWN And Player(Index).MovingV > 0 Then
                    rec.Top = rec.Top - Player(Index).YOffset
                Else
                    rec.Top = rec.Top - Player(Index).YOffset + (SIZE_Y - PIC_Y)
                End If
            End If
            
            Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    End If
'End If
End Sub

Sub BltPet(ByVal Index As Long)
Dim Anim As Byte
Dim x As Long, y As Long
Dim AttackSpeed As Long

    ' Only used if ever want to switch to blt rather then bltfast
    ' I suggest you don't use, because custom sizes won't work any longer
    With rec_pos
        .Top = Player(Index).Pet.y * PIC_Y + Player(Index).Pet.YOffset - (SIZE_Y - PIC_Y)
        .Bottom = .Top + PIC_Y
        .Left = Player(Index).Pet.x * PIC_X + Player(Index).Pet.XOffset + ((SIZE_X - PIC_X) / 2)
        .Right = .Left + PIC_X + ((SIZE_X - PIC_X) / 2)
    End With
   
    ' Check for animation
    Anim = 0
    If Player(Index).Pet.Attacking = 0 Then
        Select Case Player(Index).Pet.Dir
            Case DIR_UP
                If (Player(Index).Pet.YOffset < PIC_Y / 2) Then Anim = 1
            Case DIR_DOWN
                If (Player(Index).Pet.YOffset > PIC_Y / 2 * -1) Then Anim = 1
            Case DIR_LEFT
                If (Player(Index).Pet.XOffset < PIC_Y / 2) Then Anim = 1
            Case DIR_RIGHT
                If (Player(Index).Pet.XOffset > PIC_Y / 2 * -1) Then Anim = 1
        End Select
    Else
        If Player(Index).Pet.AttackTimer + 500 > GetTickCount Then
            Anim = 2
        End If
    End If
   
    ' Check to see if we want to stop making him attack
    If Player(Index).Pet.AttackTimer + 1000 < GetTickCount Then
        Player(Index).Pet.Attacking = 0
        Player(Index).Pet.AttackTimer = 0
    End If
   
    rec.Top = Player(Index).Pet.Sprite * SIZE_Y + (SIZE_Y - PIC_Y)
    rec.Bottom = rec.Top + PIC_Y
    rec.Left = (Player(Index).Pet.Dir * (3 * (SIZE_X / PIC_X)) + (Anim * (SIZE_X / PIC_X))) * PIC_X
    rec.Right = rec.Left + SIZE_X

    x = Player(Index).Pet.x * PIC_X - (SIZE_X - PIC_X) / 2 + sx + Player(Index).Pet.XOffset
    y = Player(Index).Pet.y * PIC_Y - (SIZE_Y - PIC_Y) + sx + Player(Index).Pet.YOffset + (SIZE_Y - PIC_Y)
   
    If SIZE_X > PIC_X Then
        If x < 0 Then
            x = Player(Index).Pet.XOffset + sx + ((SIZE_X - PIC_X) / 2)
            If Player(Index).Pet.Dir = DIR_RIGHT And Player(Index).Pet.Moving > 0 Then
                rec.Left = rec.Left - Player(Index).Pet.XOffset
            Else
                rec.Left = rec.Left - Player(Index).Pet.XOffset + ((SIZE_X - PIC_X) / 2)
            End If
        End If
       
        If x > MAX_MAPX * 32 Then
            x = MAX_MAPX * 32 + sx - ((SIZE_X - PIC_X) / 2) + Player(Index).Pet.XOffset
            If Player(Index).Pet.Dir = DIR_LEFT And Player(Index).Pet.Moving > 0 Then
                rec.Right = rec.Right + Player(Index).Pet.XOffset
            Else
                rec.Right = rec.Right + Player(Index).Pet.XOffset - ((SIZE_X - PIC_X) / 2)
            End If
        End If
    End If
   
    Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End Sub

Sub BltPetTop(ByVal Index As Long)
Dim Anim As Byte
Dim x As Long, y As Long
Dim AttackSpeed As Long

    ' Only used if ever want to switch to blt rather then bltfast
    ' I suggest you don't use, because custom sizes won't work any longer
    With rec_pos
        .Top = Player(Index).Pet.y * PIC_Y + Player(Index).Pet.YOffset - (SIZE_Y - PIC_Y)
        .Bottom = .Top + PIC_Y
        .Left = Player(Index).Pet.x * PIC_X + Player(Index).Pet.XOffset + ((SIZE_X - PIC_X) / 2)
        .Right = .Left + PIC_X + ((SIZE_X - PIC_X) / 2)
    End With
   
    ' Check for animation
    Anim = 0
    If Player(Index).Pet.Attacking = 0 Then
        Select Case Player(Index).Pet.Dir
            Case DIR_UP
                If (Player(Index).Pet.YOffset < PIC_Y / 2) Then Anim = 1
            Case DIR_DOWN
                If (Player(Index).Pet.YOffset < PIC_Y / 2 * -1) Then Anim = 1
            Case DIR_LEFT
                If (Player(Index).Pet.XOffset < PIC_Y / 2) Then Anim = 1
            Case DIR_RIGHT
                If (Player(Index).Pet.XOffset < PIC_Y / 2 * -1) Then Anim = 1
        End Select
    Else
        If Player(Index).Pet.AttackTimer + 500 > GetTickCount Then
            Anim = 2
        End If
    End If
   
    ' Check to see if we want to stop making him attack
    If Player(Index).Pet.AttackTimer + 1000 < GetTickCount Then
        Player(Index).Pet.Attacking = 0
        Player(Index).Pet.AttackTimer = 0
    End If
   
    rec.Top = Player(Index).Pet.Sprite * SIZE_Y
    rec.Bottom = rec.Top + (SIZE_Y - PIC_Y)
    rec.Left = (Player(Index).Pet.Dir * (3 * (SIZE_X / PIC_X)) + (Anim * (SIZE_X / PIC_X))) * PIC_X
    rec.Right = rec.Left + SIZE_X

    x = Player(Index).Pet.x * PIC_X - (SIZE_X - PIC_X) / 2 + sx + Player(Index).Pet.XOffset
    y = Player(Index).Pet.y * PIC_Y - (SIZE_Y - PIC_Y) + sx + Player(Index).Pet.YOffset
   
   
    If y < 0 Then
        y = 0
        If Player(Index).Pet.Dir = DIR_DOWN And Player(Index).Pet.Moving > 0 Then
            rec.Top = rec.Top - Player(Index).Pet.YOffset
        Else
            rec.Top = rec.Top - Player(Index).Pet.YOffset + (SIZE_Y - PIC_Y)
        End If
    End If
   
    If SIZE_X > PIC_X Then
        If x < 0 Then
            x = Player(Index).Pet.XOffset + sx + ((SIZE_X - PIC_X) / 2)
            If Player(Index).Pet.Dir = DIR_RIGHT And Player(Index).Pet.Moving > 0 Then
                rec.Left = rec.Left - Player(Index).Pet.XOffset
            Else
                rec.Left = rec.Left - Player(Index).Pet.XOffset + ((SIZE_X - PIC_X) / 2)
            End If
        End If
       
        If x > MAX_MAPX * 32 Then
            x = MAX_MAPX * 32 + sx - ((SIZE_X - PIC_X) / 2) + Player(Index).Pet.XOffset
            If Player(Index).Pet.Dir = DIR_LEFT And Player(Index).Pet.Moving > 0 Then
                rec.Right = rec.Right + Player(Index).Pet.XOffset
            Else
                rec.Right = rec.Right + Player(Index).Pet.XOffset - ((SIZE_X - PIC_X) / 2)
            End If
        End If
    End If
   
    Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End Sub
Sub BltMapNPCName(ByVal Index As Long)
Dim TextX As Long
Dim TextY As Long

If Npc(MapNpc(Index).Num).Big = 0 Then
    With Npc(MapNpc(Index).Num)
    'Draw name
        TextX = MapNpc(Index).x * PIC_X + sx + MapNpc(Index).XOffset + CLng(PIC_X / 2) - ((Len(Trim$(.name)) / 2) * 8)
        TextY = MapNpc(Index).y * PIC_Y + sx + MapNpc(Index).YOffset - CLng(PIC_Y / 2) - 32
        DrawPlayerNameText TexthDC, TextX - (NewPlayerX * PIC_X) - NewXOffset, TextY - (NewPlayerY * PIC_Y) - NewYOffset, Trim$(.name), vbWhite
    End With
Else
    With Npc(MapNpc(Index).Num)
    'Draw name
        TextX = MapNpc(Index).x * PIC_X + sx + MapNpc(Index).XOffset + CLng(PIC_X / 2) - ((Len(Trim$(.name)) / 2) * 8)
        TextY = MapNpc(Index).y * PIC_Y + sx + MapNpc(Index).YOffset - CLng(PIC_Y / 2) - 32
        DrawPlayerNameText TexthDC, TextX - (NewPlayerX * PIC_X) - NewXOffset, TextY - (NewPlayerY * PIC_Y) - NewYOffset, Trim$(.name), vbWhite
    End With
End If
End Sub

Sub BltNpc(ByVal MapNpcNum As Long)
Dim Anim As Byte
Dim x As Long, y As Long

    ' Make sure that theres an npc there, and if not exit the sub
    If MapNpc(MapNpcNum).Num <= 0 Then
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
        
    If Npc(MapNpc(MapNpcNum).Num).Big = 0 Then
        rec.Top = Npc(MapNpc(MapNpcNum).Num).Sprite * 64 + PIC_Y
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
        rec.Top = Npc(MapNpc(MapNpcNum).Num).Sprite * 64 + 32
        rec.Bottom = rec.Top + 32
        rec.Left = (MapNpc(MapNpcNum).Dir * 3 + Anim) * 64
        rec.Right = rec.Left + 64
    
        x = MapNpc(MapNpcNum).x * 32 + sx - 16 + MapNpc(MapNpcNum).XOffset
        y = MapNpc(MapNpcNum).y * 32 + sx + MapNpc(MapNpcNum).YOffset
  
        If y < 0 Then
            rec.Top = Npc(MapNpc(MapNpcNum).Num).Sprite * 64 + 32
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
    If MapNpc(MapNpcNum).Num <= 0 Then
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
        If Npc(MapNpc(MapNpcNum).Num).Big = 0 Then
      rec.Top = Npc(MapNpc(MapNpcNum).Num).Sprite * 64
        rec.Bottom = rec.Top + PIC_Y
        rec.Left = (MapNpc(MapNpcNum).Dir * 3 + Anim) * PIC_X
        rec.Right = rec.Left + PIC_X
        
        x = MapNpc(MapNpcNum).x * PIC_X + sx + MapNpc(MapNpcNum).XOffset
        y = MapNpc(MapNpcNum).y * PIC_Y + sx + MapNpc(MapNpcNum).YOffset - 32
        
        ' Check if its out of bounds because of the offset
        If y < 0 Then
            y = 0
            rec.Top = rec.Top + (y * -1)
        End If
            
        'Call DD_BackBuffer.Blt(rec_pos, DD_SpriteSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
        Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
  Else
    rec.Top = Npc(MapNpc(MapNpcNum).Num).Sprite * PIC_Y
        
     rec.Top = Npc(MapNpc(MapNpcNum).Num).Sprite * 64
     rec.Bottom = rec.Top + 32
     rec.Left = (MapNpc(MapNpcNum).Dir * 3 + Anim) * 64
     rec.Right = rec.Left + 64
  
     x = MapNpc(MapNpcNum).x * 32 + sx - 16 + MapNpc(MapNpcNum).XOffset
     y = MapNpc(MapNpcNum).y * 32 + sx - 32 + MapNpc(MapNpcNum).YOffset

     If y < 0 Then
         rec.Top = Npc(MapNpc(MapNpcNum).Num).Sprite * 64 + 32
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
End If
End Sub


Sub BltPetName(ByVal Index As Long)
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
        
    ' Draw name
    TextX = Player(Index).Pet.x * PIC_X + sx + Player(Index).Pet.XOffset + Int(PIC_X / 2) - ((Len(GetPlayerName(Index) & "'s Pet") / 2) * 8)
    TextY = Player(Index).Pet.y * PIC_Y + sx + Player(Index).Pet.YOffset - Int(PIC_Y / 2) - (SIZE_Y - PIC_Y)
    Call DrawText(TexthDC, TextX - (NewPlayerX * PIC_X) - NewXOffset, TextY - (NewPlayerY * PIC_Y) - NewYOffset, GetPlayerName(Index) & "'s Pet", Color)
End Sub

Sub BltPlayerName(ByVal Index As Long, Flash As Byte)
Dim TextX As Long
Dim TextY As Long
Dim Color As Long

    
    ' Check access level
    If GetPlayerPK(Index) = NO Then
        Select Case GetPlayerAccess(Index)
            Case 0
                If GetPlayerAlignment(Index) > 0 Then
                Color = QBColor(DarkGrey)
                End If
                If GetPlayerAlignment(Index) > 999 Then
                Color = QBColor(Grey)
                End If
                If GetPlayerAlignment(Index) > 1999 Then
                Color = QBColor(BrightRed)
                End If
                If GetPlayerAlignment(Index) > 2799 Then
                Color = QBColor(BrightRed)
                End If
                If GetPlayerAlignment(Index) > 3799 Then
                Color = QBColor(Brown)
                End If
                If GetPlayerAlignment(Index) > 4999 Then
                Color = QBColor(Yellow)
                End If
                If GetPlayerAlignment(Index) > 7499 Then
                Color = QBColor(Green)
                End If
                If GetPlayerAlignment(Index) > 8499 Then
                Color = QBColor(BrightGreen)
                End If
                If GetPlayerAlignment(Index) > 9777 Then
                Color = QBColor(White)
                End If
            Case 1
                Color = QBColor(Cyan)
            Case 2
                Color = QBColor(Magenta)
            Case 3
                Color = QBColor(BrightBlue)
            Case 4
                Color = QBColor(BrightCyan)
        End Select
    Else
        Color = QBColor(BrightRed)
       If Flash = 1 Then
        Color = QBColor(Grey)
    End If
End If
    
        
    ' Draw name
    TextX = GetPlayerX(Index) * PIC_X + sx + Player(Index).XOffset + Int(PIC_X / 2) - ((Len(GetPlayerName(Index)) / 2) * 8)
    TextY = GetPlayerY(Index) * PIC_Y + sx + Player(Index).YOffset - Int(PIC_Y / 2) - (SIZE_Y - PIC_Y)
    Call DrawText(TexthDC, TextX - (NewPlayerX * PIC_X) - NewXOffset, TextY - (NewPlayerY * PIC_Y) - NewYOffset, GetPlayerName(Index), Color)
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
    If GetPlayerAccess(Index) > 0 Then
        If Player(Index).MovingV <> 0 Then Player(Index).YOffset = Player(Index).YOffset + (GM_WALK_SPEED * Player(Index).MovingV)
        If Player(Index).MovingH <> 0 Then Player(Index).XOffset = Player(Index).XOffset + (GM_WALK_SPEED * Player(Index).MovingH)
    Else
        If Player(Index).MovingV <> 0 Then Player(Index).YOffset = Player(Index).YOffset + (WALK_SPEED * Player(Index).MovingV)
        If Player(Index).MovingH <> 0 Then Player(Index).XOffset = Player(Index).XOffset + (WALK_SPEED * Player(Index).MovingH)
    End If
   
    ' Check if completed walking over to the next tile
    If Player(Index).XOffset = 0 Then
        Player(Index).MovingH = 0
    End If
    If Player(Index).YOffset = 0 Then
        Player(Index).MovingV = 0
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

Sub ProcessPetMovement(ByVal PetNum As Long)
    ' Check if pet is walking, and if so process moving them over
    If Player(PetNum).Pet.Moving = MOVING_WALKING Then
        Select Case Player(PetNum).Pet.Dir
            Case DIR_UP
                Player(PetNum).Pet.YOffset = Player(PetNum).Pet.YOffset - WALK_SPEED
            Case DIR_DOWN
                Player(PetNum).Pet.YOffset = Player(PetNum).Pet.YOffset + WALK_SPEED
            Case DIR_LEFT
                Player(PetNum).Pet.XOffset = Player(PetNum).Pet.XOffset - WALK_SPEED
            Case DIR_RIGHT
                Player(PetNum).Pet.XOffset = Player(PetNum).Pet.XOffset + WALK_SPEED
        End Select
        
        ' Check if completed walking over to the next tile
        If (Player(PetNum).Pet.XOffset = 0) And (Player(PetNum).Pet.YOffset = 0) Then
            Player(PetNum).Pet.Moving = 0
        End If
    End If
End Sub

Sub HandleKeypresses(ByVal KeyAscii As Integer)
Dim ChatText As String
Dim name As String
Dim i As Long
Dim n As Long

MyText = frmMirage.txtMyTextBox.Text

If mid$(MyText, 1, 1) <> "/" And mid$(MyText, 1, 1) <> "'" And mid$(MyText, 1, 1) <> "@" And mid$(MyText, 1, 1) <> "\" And mid$(MyText, 1, 1) <> Chr$(34) And mid$(MyText, 1, 1) <> "!" And mid$(MyText, 1, 1) <> "=" Then
    
        Select Case frmMirage.Combo1
        Case "\shout"
            MyText = "," & MyText
        Case "\party"
            MyText = "+" & MyText
        Case "\guild"
            MyText = "=" & MyText
        Case "\trade"
            MyText = "/trade " & MyText
        Case "\chat"
            MyText = "/chat " & MyText
        End Select
    End If

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
        
        If GetPlayerFacingX(MyIndex) > 0 And GetPlayerFacingX(MyIndex) <= MAX_MAPX And GetPlayerFacingY(MyIndex) > 0 And GetPlayerFacingY(MyIndex) <= MAX_MAPY Then
        If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerFacingX(MyIndex), GetPlayerFacingY(MyIndex)).Type = TILE_TYPE_FURNITURE Then
        If Player(MyIndex).Hands = 0 Then
        If Map(GetPlayerMap(MyIndex)).Moral = MAP_MORAL_HOUSE Then
        If Map(GetPlayerMap(MyIndex)).Owner = GetPlayerName(MyIndex) Then
        Call SetPlayerHands(MyIndex, Map(GetPlayerMap(MyIndex)).Tile(GetPlayerFacingX(MyIndex), GetPlayerFacingY(MyIndex)).Data1)
        Call AddText("You pick up the piece of furniture.", 15)
        Call SetAttribute(GetPlayerMap(MyIndex), GetPlayerFacingX(MyIndex), GetPlayerFacingY(MyIndex), TILE_TYPE_HOUSE, 0, 0, 0, "", "", "")
        Call UpdateVisInv
        Exit Sub
        Else
        Call AddText("This is not your house!", 12)
        End If
        Else
        Call AddText("This is not a house!", 12)
        End If
        Else
        Call AddText("Your hands are already full!", 12)
        End If
        End If
        End If
        
        
        If GetPlayerFacingX(MyIndex) > 0 And GetPlayerFacingX(MyIndex) <= MAX_MAPX And GetPlayerFacingY(MyIndex) > 0 And GetPlayerFacingY(MyIndex) <= MAX_MAPY Then
        If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerFacingX(MyIndex), GetPlayerFacingY(MyIndex)).Type = TILE_TYPE_HOUSE Then
        If Player(MyIndex).Hands <> 0 Then
        If Map(GetPlayerMap(MyIndex)).Moral = MAP_MORAL_HOUSE Then
        If Map(GetPlayerMap(MyIndex)).Owner = GetPlayerName(MyIndex) Then
        Call AddText("You set down the piece of furniture.", 15)
        Call SetAttribute(GetPlayerMap(MyIndex), GetPlayerFacingX(MyIndex), GetPlayerFacingY(MyIndex), TILE_TYPE_FURNITURE, Player(MyIndex).Hands, 0, 0, "", "", "")
        Call SetPlayerHands(MyIndex, 0)
        Call UpdateVisInv
        Exit Sub
        Else
        Call AddText("This is not your house!", 12)
        End If
        Else
        Call AddText("This is not a house!", 12)
        End If
        End If
        End If
        End If
        
        ' Broadcast message
        If mid(MyText, 1, 1) = "," Then
            ChatText = mid(MyText, 2, Len(MyText) - 1)
            If Len(Trim(ChatText)) > 0 Then
                Call BroadcastMsg(ChatText)
            End If
            MyText = ""
            Exit Sub
        End If
        
        ' Emote message
        If mid(MyText, 1, 1) = "-" Then
            ChatText = mid(MyText, 2, Len(MyText) - 1)
            If Len(Trim(ChatText)) > 0 Then
                Call EmoteMsg(ChatText)
            End If
            MyText = ""
            Exit Sub
        End If
        
        ' Party message
        If mid(MyText, 1, 1) = "+" Then
            ChatText = mid(MyText, 2, Len(MyText) - 1)
            If Len(Trim(ChatText)) > 0 Then
                Call PartyMsg(ChatText)
            End If
            MyText = ""
            Exit Sub
        End If
        
        ' Guild message
        If mid(MyText, 1, 1) = "=" Then
            ChatText = mid(MyText, 2, Len(MyText) - 1)
            If Len(Trim(ChatText)) > 0 Then
                Call GuildMsg(ChatText)
            End If
            MyText = ""
            Exit Sub
        End If
        
        ' Commands
        If mid(MyText, 1, 9) = "/commands" Then
            Call AddText(":::::::: Commands ::::::::", White)
            Call AddText(", (broadcast), - (emote), + (party), = (guild), ! (private)", White)
            Call AddText("/info, /who, /fps, /inv, /stats, /chat, /chatdecline, /trade, /accept, /decline, /party, /join, /leave, /killpet, /refresh", White)
            Exit Sub
        End If
        
        ' Admin Commands
        If mid(MyText, 1, 14) = "/admincommands" Then
            Call AddText(":::::::: Admin Commands ::::::::", White)
            Call AddText("; (global), @ (admin)", White)
            Call AddText("/daynight, /weather, /kick", White)
            Call AddText("/loc, /warp, /warptome, /mapeditor, /mapreport, /setsprite, /setplayersprite, /respawn, /motd, /banlist, /ban", White)
            Call AddText("/edititem, /editarrow, /editemoticon, /editnpc, /editshop, /editspell, /editspeech", White)
            Call AddText("/setaccess, /nullbanlist, /debug, /editmain, /mainbackup", White)
            Exit Sub
        End If
        
        ' Player message
        If mid(MyText, 1, 1) = "!" Then
            ChatText = mid(MyText, 2, Len(MyText) - 1)
            name = ""
                    
            ' Get the desired player from the user text
            For i = 1 To Len(ChatText)
                If mid(ChatText, i, 1) <> " " Then
                    name = name & mid(ChatText, i, 1)
                Else
                    Exit For
                End If
            Next i
                    
            ' Make sure they are actually sending something
            If Len(ChatText) - i > 0 Then
                ChatText = mid(ChatText, i + 1, Len(ChatText) - i)
                    
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
        If LCase(mid(MyText, 1, 5)) = "/info" Then
            ChatText = mid(MyText, 6, Len(MyText) - 5)
            Call SendData("playerinforequest" & SEP_CHAR & ChatText & SEP_CHAR & END_CHAR)
            MyText = ""
            Exit Sub
        End If
        
        ' Whos Online
        If LCase(mid(MyText, 1, 4)) = "/who" Then
            Call SendWhosOnline
            MyText = ""
            Exit Sub
        End If
                        
        ' Checking fps
        If LCase(mid(MyText, 1, 4)) = "/fps" Then
            Call AddText("FPS: " & GameFPS, Pink)
            MyText = ""
            Exit Sub
        End If
                
        ' Show inventory
        If LCase(mid(MyText, 1, 4)) = "/inv" Then
            frmMirage.picInv3.Visible = True
            MyText = ""
            Exit Sub
        End If
        
        ' Request stats
        If LCase(mid(MyText, 1, 6)) = "/stats" Then
            Call SendData("getstats" & SEP_CHAR & END_CHAR)
            MyText = ""
            Exit Sub
        End If
         
        ' Refresh Player
        If LCase(mid(MyText, 1, 8)) = "/refresh" Then
            Call SendData("refresh" & SEP_CHAR & END_CHAR)
            MyText = ""
            Exit Sub
        End If
        
        ' Decline Chat
        If LCase(mid(MyText, 1, 12)) = "/chatdecline" Then
            Call SendData("dchat" & SEP_CHAR & END_CHAR)
            MyText = ""
            Exit Sub
        End If
        
        ' Accept Chat
        If LCase(mid(MyText, 1, 5)) = "/chat" Then
            Call SendData("achat" & SEP_CHAR & END_CHAR)
            MyText = ""
            Exit Sub
        End If
        
        If LCase(mid(MyText, 1, 6)) = "/trade" Then
            ' Make sure they are actually sending something
            If Len(MyText) > 7 Then
                ChatText = mid(MyText, 8, Len(MyText) - 7)
                Call SendTradeRequest(ChatText)
            Else
                Call AddText("Usage: /trade playernamehere", AlertColor)
            End If
            MyText = ""
            Exit Sub
        End If
        
        ' Accept Trade
        If LCase(mid(MyText, 1, 7)) = "/accept" Then
            Call SendAcceptTrade
            MyText = ""
            Exit Sub
        End If
        
        ' Decline Trade
        If LCase(mid(MyText, 1, 8)) = "/decline" Then
            Call SendDeclineTrade
            MyText = ""
            Exit Sub
        End If
        
        ' Party request
        If LCase(mid(MyText, 1, 6)) = "/party" Then
            ' Make sure they are actually sending something
            If Len(MyText) > 7 Then
                ChatText = mid(MyText, 8, Len(MyText) - 7)
                Call SendPartyRequest(ChatText)
            Else
                Call AddText("Usage: /party playernamehere", AlertColor)
            End If
            MyText = ""
            Exit Sub
        End If
        
        ' Kill pet
        If LCase(mid(MyText, 1, 8)) = "/killpet" Then
            Call SendData("KILLPET" & SEP_CHAR & END_CHAR)
            MyText = ""
            Exit Sub
        End If
        
        ' Join party
        If LCase(mid(MyText, 1, 5)) = "/join" Then
            Call SendJoinParty
            MyText = ""
            Exit Sub
        End If
        
        ' Leave party
        If LCase(mid(MyText, 1, 6)) = "/leave" Then
            Call SendLeaveParty
            MyText = ""
            Exit Sub
        End If
        
        ' // Moniter Admin Commands //
        If GetPlayerAccess(MyIndex) > 0 Then
            ' day night command
            If LCase(mid(MyText, 1, 9)) = "/daynight" Then
                If GameTime = TIME_DAY Then
                    GameTime = TIME_NIGHT
                Else
                    GameTime = TIME_DAY
                End If
                Call SendGameTime
                MyText = ""
                Exit Sub
            End If
            
            ' Clearing a house owner
            If LCase(mid(MyText, 1, 11)) = "/clearowner" Then
                Call SendData("clearowner" & SEP_CHAR & END_CHAR)
                MyText = ""
                Exit Sub
            End If
            
            ' Editing emoticon request
            If mid(MyText, 1, 12) = "/editelement" Then
                Call SendRequestEditElement
                MyText = ""
                Exit Sub
            End If

            
            ' weather command
            If LCase(mid(MyText, 1, 8)) = "/weather" Then
                If Len(MyText) > 8 Then
                    MyText = mid(MyText, 9, Len(MyText) - 8)
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
            If LCase(mid(MyText, 1, 5)) = "/kick" Then
                If Len(MyText) > 6 Then
                    MyText = mid(MyText, 7, Len(MyText) - 6)
                    Call SendKick(MyText)
                End If
                MyText = ""
                Exit Sub
            End If
        
            ' Global Message
            If mid(MyText, 1, 1) = ";" Then
                ChatText = mid(MyText, 2, Len(MyText) - 1)
                If Len(Trim(ChatText)) > 0 Then
                    Call GlobalMsg(ChatText)
                End If
                MyText = ""
                Exit Sub
            End If
        
            ' Admin Message
            If mid(MyText, 1, 1) = "@" Then
                ChatText = mid(MyText, 2, Len(MyText) - 1)
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
            If LCase(mid(MyText, 1, 4)) = "/loc" Then
                Call SendRequestLocation
                MyText = ""
                Exit Sub
            End If
            
            ' Warpe
            If LCase(mid(MyText, 1, 6)) = "/warp " Then
                If Len(MyText) > 6 Then
                    MyText = mid(MyText, 7, Len(MyText) - 6)
                    Call SendWarp(MyText)
                End If
                Exit Sub
            End If
            
            ' Map Editor
            If LCase(mid(MyText, 1, 8)) = "/editmap" Or LCase(mid(MyText, 1, 10)) = "/mapeditor" Then
                Call SendRequestEditMap
                MyText = ""
                Exit Sub
            End If
            
            ' Map report
            If LCase(mid(MyText, 1, 10)) = "/mapreport" Then
                Call SendData("mapreport" & SEP_CHAR & END_CHAR)
                MyText = ""
                Exit Sub
            End If
            
            ' Setting sprite
            If LCase(mid(MyText, 1, 10)) = "/setsprite" Then
                If Len(MyText) > 11 Then
                    ' Get sprite #
                    MyText = mid(MyText, 12, Len(MyText) - 11)
                
                    Call SendSetSprite(Val(MyText))
                End If
                MyText = ""
                Exit Sub
            End If
            
            ' Setting player sprite
            If LCase(mid(MyText, 1, 16)) = "/setplayersprite" Then
                If Len(MyText) > 19 Then
                    i = Val(mid(MyText, 17, 1))
                
                    MyText = mid(MyText, 18, Len(MyText) - 17)
                    Call SendSetPlayerSprite(i, Val(MyText))
                End If
                MyText = ""
                Exit Sub
            End If
        
            ' Respawn request
            If mid(MyText, 1, 8) = "/respawn" Then
                Call SendMapRespawn
                MyText = ""
                Exit Sub
            End If
        
            ' MOTD change
            If mid(MyText, 1, 5) = "/motd" Then
                If Len(MyText) > 6 Then
                    MyText = mid(MyText, 7, Len(MyText) - 6)
                    If Trim(MyText) <> "" Then
                        Call SendMOTDChange(MyText)
                    End If
                End If
                MyText = ""
                Exit Sub
            End If
            
            ' Check the ban list
            If mid(MyText, 1, 3) = "/banlist" Then
                Call SendBanList
                MyText = ""
                Exit Sub
            End If
            
            ' Banning a player
            If LCase(mid(MyText, 1, 4)) = "/ban" Then
                If Len(MyText) > 5 Then
                    MyText = mid(MyText, 6, Len(MyText) - 5)
                    Call SendBan(MyText)
                    MyText = ""
                End If
                Exit Sub
            End If
        End If
            
        ' // Developer Admin Commands //
        If GetPlayerAccess(MyIndex) >= ADMIN_DEVELOPER Then
            ' Editing item request
            If mid(MyText, 1, 9) = "/edititem" Or mid(MyText, 1, 11) = "/itemeditor" Then
                Call SendRequestEditItem
                MyText = ""
                Exit Sub
            End If
            
            ' Editing emoticon request
            If mid(MyText, 1, 13) = "/editemoticon" Or mid(MyText, 1, 15) = "/emoticoneditor" Then
                Call SendRequestEditEmoticon
                MyText = ""
                Exit Sub
            End If
            
            ' Editing arrow request
            If mid(MyText, 1, 10) = "/editarrow" Or mid(MyText, 1, 12) = "/arroweditor" Then
                Call SendRequestEditArrow
                MyText = ""
                Exit Sub
            End If
            
            ' Editing speech request
            If mid(MyText, 1, 11) = "/editspeech" Or mid(MyText, 1, 13) = "/speecheditor" Then
                Call SendRequestEditSpeech
                MyText = ""
                Exit Sub
            End If
            
            ' Editing npc request
            If mid(MyText, 1, 8) = "/editnpc" Or mid(MyText, 1, 10) = "/npceditor" Then
                Call SendRequestEditNpc
                MyText = ""
                Exit Sub
            End If
            
            ' Editing shop request
            If mid(MyText, 1, 9) = "/editshop" Or mid(MyText, 1, 11) = "/shopeditor" Then
                Call SendRequestEditShop
                MyText = ""
                Exit Sub
            End If
        
            ' Editing spell request
            If mid(MyText, 1, 10) = "/editspell" Or mid(MyText, 1, 12) = "/spelleditor" Then
                Call SendRequestEditSpell
                MyText = ""
                Exit Sub
            End If
        End If
        
        ' // Creator Admin Commands //
        If GetPlayerAccess(MyIndex) >= ADMIN_CREATOR Then
            ' Giving another player access
            If LCase(mid(MyText, 1, 10)) = "/setaccess" Then
                ' Get access #
                i = Val(mid(MyText, 12, 1))
                
                MyText = mid(MyText, 14, Len(MyText) - 13)
                
                Call SendSetAccess(MyText, i)
                MyText = ""
                Exit Sub
            End If
                    
            ' Edit main.txt
            If mid(MyText, 1, 9) = "/editmain" Or mid(MyText, 1, 11) = "/maineditor" Then
                Call SendRequestEditMain
                MyText = ""
                Exit Sub
            End If
            
            ' Reload the backup
            If mid(MyText, 1, 12) = "/mainbackup" Then
                Call SendRequestBackupMain
                MyText = ""
                Exit Sub
            End If
            
            ' Debugging
            If LCase(mid(MyText, 1, 6)) = "/debug" Then
                If GoDebug = YES Then
                    GoDebug = NO
                Else
                    GoDebug = YES
                End If
                MyText = ""
                Exit Sub
            End If
            
            ' Ban destroy
            If LCase(mid(MyText, 1, 15)) = "/destroybanlist" Or LCase(mid(MyText, 1, 12)) = "/nullbanlist" Then
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
Dim i As Long
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
            End If
            If KeyCode = vbKeyDown Then
                DirUp = False
                DirDown = True
            End If
            If KeyCode = vbKeyLeft Then
                DirLeft = True
                DirRight = False
            End If
            If KeyCode = vbKeyRight Then
                DirLeft = False
                DirRight = True
            End If
            If KeyCode = vbKeyShift Then
                ShiftDown = True
            End If
        Else
            If KeyCode = vbKeyUp Then
                XToGo = -1
                YToGo = -1
            End If
            If KeyCode = vbKeyDown Then
                XToGo = -1
                YToGo = -1
            End If
            If KeyCode = vbKeyLeft Then
                XToGo = -1
                YToGo = -1
            End If
            If KeyCode = vbKeyRight Then
                XToGo = -1
                YToGo = -1
            End If
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
    If Player(MyIndex).MovingH <> 0 And Player(MyIndex).MovingV <> 0 Then
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
    If Player(MyIndex).MovingV <> 0 Then
            CanMove = False
            Exit Function
        End If
        Call SetPlayerDir(MyIndex, DIR_UP)
        
        ' Check to see if they are trying to go out of bounds
        If GetPlayerY(MyIndex) > 0 Then
            ' Check to see if the map tile is blocked or not
            If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).Type = TILE_TYPE_BLOCKED Or Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).Type = TILE_TYPE_SIGN Then
                If TempTile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).Type = TILE_TYPE_BLOCKED Or TempTile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).Type = TILE_TYPE_NONE Then
                    CanMove = False
                
                    ' Set the new direction if they weren't facing that direction
                    If d <> DIR_UP Then
                        Call SendPlayerDir
                    End If
                    Exit Function
                End If
            Else
                If TempTile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).Type = TILE_TYPE_BLOCKED Then
                    CanMove = False
                
                    ' Set the new direction if they weren't facing that direction
                    If d <> DIR_UP Then
                        Call SendPlayerDir
                    End If
                    Exit Function
                End If
            End If
            
            If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).Type = TILE_TYPE_CBLOCK Then
                If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).Data1 = Player(MyIndex).Class Then Exit Function
                If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).Data2 = Player(MyIndex).Class Then Exit Function
                If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).Data3 = Player(MyIndex).Class Then Exit Function
                CanMove = False
                
                ' Set the new direction if they weren't facing that direction
                If d <> DIR_UP Then
                    Call SendPlayerDir
                End If
            End If
                                                    
            ' Check to see if the key door is open or not
            If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).Type = TILE_TYPE_KEY Then
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
                If i <> MyIndex Then
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
                        
                        ' Might as well check for pets too
                        If Player(i).Pet.Alive = YES Then
                            If Player(i).Pet.Map = GetPlayerMap(MyIndex) Then
                                If Player(i).Pet.x = GetPlayerX(MyIndex) And Player(i).Pet.y = GetPlayerY(MyIndex) - 1 Then
                                    'Player(I).Pet.MapToGo = -1
                                    'Player(I).Pet.XToGo = -1
                                    'Player(I).Pet.YToGo = -1
                                    CanMove = False
                            
                                    ' Set the new direction if they weren't facing that direction
                                    If d <> DIR_UP Then
                                        Call SendPlayerDir
                                    End If
                                    Exit Function
                                End If
                            End If
                        End If
                    End If
                Else
                    If Player(i).Pet.Alive = YES Then
                        If Player(i).Pet.Map = GetPlayerMap(MyIndex) Then
                            If Player(i).Pet.x = GetPlayerX(MyIndex) And Player(i).Pet.y = GetPlayerY(MyIndex) - 1 Then
                                'Player(I).Pet.MapToGo = -1
                                'Player(I).Pet.XToGo = -1
                                'Player(I).Pet.YToGo = -1
                                If IsValid(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 2) Then
                                    If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 2).Type = TILE_TYPE_BLOCKED Then
                                        CanMove = False
                                
                                        ' Set the new direction if they weren't facing that direction
                                        If d <> DIR_UP Then
                                            Call SendPlayerDir
                                        End If
                                        Exit Function
                                    End If
                                Else
                                    CanMove = False
                                
                                    ' Set the new direction if they weren't facing that direction
                                    If d <> DIR_UP Then
                                        Call SendPlayerDir
                                    End If
                                    Exit Function
                                End If
                            End If
                        End If
                    End If
                End If
            Next i
        
            ' Check to see if a npc is already on that tile
            For i = 1 To MAX_MAP_NPCS
                If MapNpc(i).Num > 0 Then
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
        Else
            ' Check if they can warp to a new map
            If Map(GetPlayerMap(MyIndex)).Up > 0 Then
                Call SendPlayerRequestNewMap
                GettingMap = True
            End If
            CanMove = False
            Exit Function
        End If
      Exit Function
    End If
            
    If DirDown Then
    If Player(MyIndex).MovingV <> 0 Then
            CanMove = False
            Exit Function
        End If
        Call SetPlayerDir(MyIndex, DIR_DOWN)
        
        ' Check to see if they are trying to go out of bounds
        If GetPlayerY(MyIndex) < MAX_MAPY Then
            ' Check to see if the map tile is blocked or not
            If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) + 1).Type = TILE_TYPE_BLOCKED Or Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) + 1).Type = TILE_TYPE_SIGN Then
                If TempTile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) + 1).Type = TILE_TYPE_BLOCKED Or TempTile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) + 1).Type = TILE_TYPE_NONE Then
                    CanMove = False
                
                    ' Set the new direction if they weren't facing that direction
                    If d <> DIR_DOWN Then
                        Call SendPlayerDir
                    End If
                    Exit Function
                End If
            Else
                If TempTile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) + 1).Type = TILE_TYPE_BLOCKED Then
                    CanMove = False
                
                    ' Set the new direction if they weren't facing that direction
                    If d <> DIR_DOWN Then
                        Call SendPlayerDir
                    End If
                    Exit Function
                End If
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
            If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) + 1).Type = TILE_TYPE_KEY Then
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
                If i <> MyIndex Then
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
                    
                    ' Might as well check for pets too
                    If Player(i).Pet.Alive = YES Then
                        If Player(i).Pet.Map = GetPlayerMap(MyIndex) Then
                            If Player(i).Pet.x = GetPlayerX(MyIndex) And Player(i).Pet.y = GetPlayerY(MyIndex) + 1 Then
                                CanMove = False
                        
                                ' Set the new direction if they weren't facing that direction
                                If d <> DIR_DOWN Then
                                    Call SendPlayerDir
                                End If
                                Exit Function
                            End If
                        End If
                    End If
                Else
                    If Player(i).Pet.Alive = YES Then
                        If Player(i).Pet.Map = GetPlayerMap(MyIndex) Then
                            If Player(i).Pet.x = GetPlayerX(MyIndex) And Player(i).Pet.y = GetPlayerY(MyIndex) + 1 Then
                                If IsValid(GetPlayerX(MyIndex), GetPlayerY(MyIndex) + 2) Then
                                    If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) + 2).Type = TILE_TYPE_BLOCKED Then
                                        CanMove = False
                                
                                        ' Set the new direction if they weren't facing that direction
                                        If d <> DIR_DOWN Then
                                            Call SendPlayerDir
                                        End If
                                        Exit Function
                                    End If
                                Else
                                    CanMove = False
                                
                                    ' Set the new direction if they weren't facing that direction
                                    If d <> DIR_DOWN Then
                                        Call SendPlayerDir
                                    End If
                                    Exit Function
                                End If
                            End If
                        End If
                    End If
                End If
            Next i
            
            ' Check to see if a npc is already on that tile
            For i = 1 To MAX_MAP_NPCS
                If MapNpc(i).Num > 0 Then
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
        Else
            ' Check if they can warp to a new map
            If Map(GetPlayerMap(MyIndex)).Down > 0 Then
                Call SendPlayerRequestNewMap
                GettingMap = True
            End If
            CanMove = False
            Exit Function
        End If
      Exit Function
    End If
 
    If DirLeft Then
    If Player(MyIndex).MovingH <> 0 Then
            CanMove = False
            Exit Function
        End If
        Call SetPlayerDir(MyIndex, DIR_LEFT)
        
        ' Check to see if they are trying to go out of bounds
        If GetPlayerX(MyIndex) > 0 Then
            ' Check to see if the map tile is blocked or not
            If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) - 1, GetPlayerY(MyIndex)).Type = TILE_TYPE_BLOCKED Or Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) - 1, GetPlayerY(MyIndex)).Type = TILE_TYPE_SIGN Then
                If TempTile(GetPlayerX(MyIndex) - 1, GetPlayerY(MyIndex)).Type = TILE_TYPE_BLOCKED Or TempTile(GetPlayerX(MyIndex) - 1, GetPlayerY(MyIndex)).Type = TILE_TYPE_NONE Then
                    CanMove = False
                
                    ' Set the new direction if they weren't facing that direction
                    If d <> DIR_LEFT Then
                        Call SendPlayerDir
                    End If
                    Exit Function
                End If
            Else
                If TempTile(GetPlayerX(MyIndex) - 1, GetPlayerY(MyIndex)).Type = TILE_TYPE_BLOCKED Then
                    CanMove = False
                
                    ' Set the new direction if they weren't facing that direction
                    If d <> DIR_LEFT Then
                        Call SendPlayerDir
                    End If
                    Exit Function
                End If
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
            If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) - 1, GetPlayerY(MyIndex)).Type = TILE_TYPE_KEY Then
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
                If i <> MyIndex Then
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
                    
                    ' Might as well check for pets too
                    If Player(i).Pet.Alive = YES Then
                        If Player(i).Pet.Map = GetPlayerMap(MyIndex) Then
                            If Player(i).Pet.x = GetPlayerX(MyIndex) - 1 And Player(i).Pet.y = GetPlayerY(MyIndex) Then
                                CanMove = False
                        
                                ' Set the new direction if they weren't facing that direction
                                If d <> DIR_LEFT Then
                                    Call SendPlayerDir
                                End If
                                Exit Function
                            End If
                        End If
                    End If
                Else
                    If Player(i).Pet.Alive = YES Then
                        If Player(i).Pet.Map = GetPlayerMap(MyIndex) Then
                            If Player(i).Pet.x = GetPlayerX(MyIndex) - 1 And Player(i).Pet.y = GetPlayerY(MyIndex) Then
                                If IsValid(GetPlayerX(MyIndex) - 2, GetPlayerY(MyIndex)) Then
                                    If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) - 2, GetPlayerY(MyIndex)).Type = TILE_TYPE_BLOCKED Then
                                        CanMove = False
                                
                                        ' Set the new direction if they weren't facing that direction
                                        If d <> DIR_LEFT Then
                                            Call SendPlayerDir
                                        End If
                                        Exit Function
                                    End If
                                Else
                                    CanMove = False
                                
                                    ' Set the new direction if they weren't facing that direction
                                    If d <> DIR_LEFT Then
                                        Call SendPlayerDir
                                    End If
                                    Exit Function
                                End If
                            End If
                        End If
                    End If
                End If
            Next i
        
            ' Check to see if a npc is already on that tile
            For i = 1 To MAX_MAP_NPCS
                If MapNpc(i).Num > 0 Then
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
        Else
            ' Check if they can warp to a new map
            If Map(GetPlayerMap(MyIndex)).Left > 0 Then
                Call SendPlayerRequestNewMap
                GettingMap = True
            End If
            CanMove = False
            Exit Function
        End If
       Exit Function
    End If
   
        
    If DirRight Then
    If Player(MyIndex).MovingH <> 0 Then
            CanMove = False
            Exit Function
        End If
        Call SetPlayerDir(MyIndex, DIR_RIGHT)
        
        ' Check to see if they are trying to go out of bounds
        If GetPlayerX(MyIndex) < MAX_MAPX Then
            ' Check to see if the map tile is blocked or not
            If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + 1, GetPlayerY(MyIndex)).Type = TILE_TYPE_BLOCKED Or Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + 1, GetPlayerY(MyIndex)).Type = TILE_TYPE_SIGN Then
                If TempTile(GetPlayerX(MyIndex) + 1, GetPlayerY(MyIndex)).Type = TILE_TYPE_BLOCKED Or TempTile(GetPlayerX(MyIndex) + 1, GetPlayerY(MyIndex)).Type = TILE_TYPE_NONE Then
                    CanMove = False
                
                    ' Set the new direction if they weren't facing that direction
                    If d <> DIR_RIGHT Then
                        Call SendPlayerDir
                    End If
                    Exit Function
                End If
            Else
                If TempTile(GetPlayerX(MyIndex) + 1, GetPlayerY(MyIndex)).Type = TILE_TYPE_BLOCKED Then
                    CanMove = False
                
                    ' Set the new direction if they weren't facing that direction
                    If d <> DIR_RIGHT Then
                        Call SendPlayerDir
                    End If
                    Exit Function
                End If
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
            If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + 1, GetPlayerY(MyIndex)).Type = TILE_TYPE_KEY Then
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
                If i <> MyIndex Then
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
                    
                    ' Might as well check for pets too
                    If Player(i).Pet.Alive = YES Then
                        If Player(i).Pet.Map = GetPlayerMap(MyIndex) Then
                            If Player(i).Pet.x = GetPlayerX(MyIndex) + 1 And Player(i).Pet.y = GetPlayerY(MyIndex) Then
                                CanMove = False
                        
                                ' Set the new direction if they weren't facing that direction
                                If d <> DIR_RIGHT Then
                                    Call SendPlayerDir
                                End If
                                Exit Function
                            End If
                        End If
                    End If
                Else
                    If Player(i).Pet.Alive = YES Then
                        If Player(i).Pet.Map = GetPlayerMap(MyIndex) Then
                            If Player(i).Pet.x = GetPlayerX(MyIndex) + 1 And Player(i).Pet.y = GetPlayerY(MyIndex) Then
                                If IsValid(GetPlayerX(MyIndex) + 2, GetPlayerY(MyIndex)) Then
                                    If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + 2, GetPlayerY(MyIndex)).Type = TILE_TYPE_BLOCKED Then
                                        CanMove = False
                                
                                        ' Set the new direction if they weren't facing that direction
                                        If d <> DIR_RIGHT Then
                                            Call SendPlayerDir
                                        End If
                                        Exit Function
                                    End If
                                Else
                                    CanMove = False
                                
                                    ' Set the new direction if they weren't facing that direction
                                    If d <> DIR_RIGHT Then
                                        Call SendPlayerDir
                                    End If
                                    Exit Function
                                End If
                            End If
                        End If
                    End If
                End If
            Next i
        
            ' Check to see if a npc is already on that tile
            For i = 1 To MAX_MAP_NPCS
                If MapNpc(i).Num > 0 Then
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
        Else
            ' Check if they can warp to a new map
            If Map(GetPlayerMap(MyIndex)).Right > 0 Then
                Call SendPlayerRequestNewMap
                GettingMap = True
            End If
            CanMove = False
            Exit Function
        End If
      Exit Function
    End If
End Function

Sub CheckMovement()
    If GettingMap = False Then
        If IsTryingToMove Then
            If CanMove Then
                Select Case GetPlayerDir(MyIndex)
                    Case DIR_UP
                        If ShiftDown Then
                            Player(MyIndex).MovingV = -MOVING_RUNNING
                        Else
                            Player(MyIndex).MovingV = -MOVING_WALKING
                        End If
                        Call SendPlayerMove
                        Player(MyIndex).YOffset = PIC_Y
                        Call SetPlayerY(MyIndex, GetPlayerY(MyIndex) - 1)
               
                    Case DIR_DOWN
                        If ShiftDown Then
                            Player(MyIndex).MovingV = MOVING_RUNNING
                        Else
                            Player(MyIndex).MovingV = MOVING_WALKING
                        End If
                        Call SendPlayerMove
                        Player(MyIndex).YOffset = PIC_Y * -1
                        Call SetPlayerY(MyIndex, GetPlayerY(MyIndex) + 1)
               
                    Case DIR_LEFT
                        If ShiftDown Then
                            Player(MyIndex).MovingH = -MOVING_RUNNING
                        Else
                            Player(MyIndex).MovingH = -MOVING_WALKING
                        End If
                        Call SendPlayerMove
                        Player(MyIndex).XOffset = PIC_X
                        Call SetPlayerX(MyIndex, GetPlayerX(MyIndex) - 1)
               
                    Case DIR_RIGHT
                        If ShiftDown Then
                            Player(MyIndex).MovingH = MOVING_RUNNING
                        Else
                            Player(MyIndex).MovingH = MOVING_WALKING
                        End If
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
                If UCase(mid(GetPlayerName(i), 1, Len(Trim(name)))) = UCase(Trim(name)) Then
                    FindPlayer = i
                    Exit Function
                End If
            End If
        End If
    Next i
    
    FindPlayer = 0
End Function

Public Sub EditorInit()
    Dim i As Long
    Dim sDc As Long

    InEditor = True
    InSpawnEditor = False
    frmMapEditor.Show vbModeless, frmMirage

    sDc = DD_TileSurf(EditorSet).GetDC
    With frmMapEditor.picBackSelect
        .Width = DDSD_Tile(EditorSet).lWidth
        .Height = DDSD_Tile(EditorSet).lHeight
        .Cls
        Call BitBlt(.hDC, 0, 0, .Width, .Height, sDc, 0, 0, SRCCOPY)
    End With
    Call DD_TileSurf(EditorSet).ReleaseDC(sDc)

    frmMapEditor.scrlPicture.Max = Int((frmMapEditor.picBackSelect.Height - frmMapEditor.picBack.Height) / PIC_Y)
    frmMapEditor.picBack.Width = 448
   
    If GameTime = TIME_NIGHT Then frmMapEditor.chkDayNight.Value = 1
    If GameTime = TIME_DAY Then frmMapEditor.chkDayNight.Value = 0
End Sub

Public Sub EditorMouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim x1, y1 As Long
Dim x2 As Long, y2 As Long, PicX As Long

    If InEditor Then
        x1 = Int(x / PIC_X)
        y1 = Int(y / PIC_Y)
       
        If frmMapEditor.MousePointer = 2 Then
            If frmMapEditor.optTiles.Value = 1 Then
                With Map(GetPlayerMap(MyIndex)).Tile(x1, y1)
                    If frmMapEditor.optGround.Value = True Then
                        PicX = .Ground
                        EditorSet = .GroundSet
                    End If
                    If frmMapEditor.optMask.Value = True Then
                        PicX = .Mask
                        EditorSet = .MaskSet
                    End If
                    If frmMapEditor.optAnim.Value = True Then
                        PicX = .Anim
                        EditorSet = .AnimSet
                    End If
                    If frmMapEditor.optMask2.Value = True Then
                        PicX = .Mask2
                        EditorSet = .Mask2Set
                    End If
                    If frmMapEditor.optM2Anim.Value = True Then
                        PicX = .M2Anim
                        EditorSet = .M2AnimSet
                    End If
                    If frmMapEditor.optFringe.Value = True Then
                        PicX = .Fringe
                        EditorSet = .FringeSet
                    End If
                    If frmMapEditor.optFAnim.Value = True Then
                        PicX = .FAnim
                        EditorSet = .FAnimSet
                    End If
                    If frmMapEditor.optFringe2.Value = True Then
                        PicX = .Fringe2
                        EditorSet = .Fringe2Set
                    End If
                    If frmMapEditor.optF2Anim.Value = True Then
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
            ElseIf frmMapEditor.optlight.Value = True Then
                EditorTileY = Int(Map(GetPlayerMap(MyIndex)).Tile(x1, y1).Light / TilesInSheets)
                EditorTileX = (Map(GetPlayerMap(MyIndex)).Tile(x1, y1).Light - Int(Map(GetPlayerMap(MyIndex)).Tile(x1, y1).Light / TilesInSheets) * TilesInSheets)
                frmMapEditor.shpSelected.Top = Int(EditorTileY * PIC_Y)
                frmMapEditor.shpSelected.Left = Int(EditorTileX * PIC_Y)
                frmMapEditor.shpSelected.Height = PIC_Y
                frmMapEditor.shpSelected.Width = PIC_X
            ElseIf frmMapEditor.optAttributes.Value = True Then
                With Map(GetPlayerMap(MyIndex)).Tile(x1, y1)
                    If .Type = TILE_TYPE_BLOCKED Then frmMapEditor.optBlocked.Value = True
                    If .Type = TILE_TYPE_WARP Then
                        EditorWarpMap = .Data1
                        EditorWarpX = .Data2
                        EditorWarpY = .Data3
                        frmMapEditor.optWarp.Value = True
                    End If
                    If .Type = TILE_TYPE_HEAL Then frmMapEditor.optHeal.Value = True
                    If .Type = TILE_TYPE_KILL Then frmMapEditor.optKill.Value = True
                    If .Type = TILE_TYPE_ITEM Then
                        ItemEditorNum = .Data1
                        ItemEditorValue = .Data2
                        frmMapEditor.optItem.Value = True
                    End If
                    If .Type = TILE_TYPE_NPCAVOID Then frmMapEditor.optNpcAvoid.Value = True
                    If .Type = TILE_TYPE_KEY Then
                        KeyEditorNum = .Data1
                        KeyEditorTake = .Data2
                        frmMapEditor.optKey.Value = True
                    End If
                    If .Type = TILE_TYPE_KEYOPEN Then
                        KeyOpenEditorX = .Data1
                        KeyOpenEditorY = .Data2
                        KeyOpenEditorMsg = .String1
                        frmMapEditor.optKeyOpen.Value = True
                    End If
                    If .Type = TILE_TYPE_SHOP Then
                        EditorShopNum = .Data1
                        frmMapEditor.optShop.Value = True
                    End If
                    If .Type = TILE_TYPE_CBLOCK Then
                        EditorItemNum1 = .Data1
                        EditorItemNum2 = .Data2
                        EditorItemNum3 = .Data3
                        frmMapEditor.optCBlock.Value = True
                    End If
                    If .Type = TILE_TYPE_ARENA Then
                        Arena1 = .Data1
                        Arena2 = .Data2
                        Arena3 = .Data3
                        frmMapEditor.optArena.Value = True
                    End If
                    If .Type = TILE_TYPE_SOUND Then
                        SoundFileName = .String1
                        frmMapEditor.optSound.Value = True
                    End If
                    If .Type = TILE_TYPE_SPRITE_CHANGE Then
                        SpritePic = .Data1
                        SpriteItem = .Data2
                        SpritePrice = .Data3
                        frmMapEditor.optSprite.Value = True
                    End If
                    If .Type = TILE_TYPE_SIGN Then
                        SignLine1 = .String1
                        SignLine2 = .String2
                        SignLine3 = .String3
                        frmMapEditor.optSign.Value = True
                    End If
                    If .Type = TILE_TYPE_DOOR Then frmMapEditor.optDoor.Value = True
                    If .Type = TILE_TYPE_NOTICE Then
                        NoticeTitle = .String1
                        NoticeText = .String2
                        NoticeSound = .String3
                        frmMapEditor.optNotice.Value = True
                    End If
                    If .Type = TILE_TYPE_CHEST Then frmMapEditor.optChest.Value = True
                    If .Type = TILE_TYPE_CLASS_CHANGE Then
                        ClassChange = .Data1
                        ClassChangeReq = .Data2
                        frmMapEditor.optClassChange.Value = True
                    End If
                    If .Type = TILE_TYPE_HOUSE_BUY Then
                        HouseItem = .Data1
                        HousePrice = .Data2
                        frmMapEditor.optHouseBuy.Value = True
                    End If
                    If .Type = TILE_TYPE_HOUSE Then
                        frmMapEditor.optHouse.Value = True
                    End If
                    If .Type = TILE_TYPE_FURNITURE Then
                        frmMapEditor.optFurniture.Value = True
                        FurnitureNum = .Data1
                    End If
                    If .Type = TILE_TYPE_SCRIPTED Then
                        ScriptNum = .Data1
                        frmMapEditor.optScripted.Value = True
                    End If
                    If .Type = TILE_TYPE_BANK Then frmMapEditor.optBank.Value = True
                End With
            End If
            frmMapEditor.MousePointer = 1
            frmMirage.MousePointer = 1
        Else
            If (Button = 1) And (x1 >= 0) And (x1 <= MAX_MAPX) And (y1 >= 0) And (y1 <= MAX_MAPY) Then
                If frmMapEditor.shpSelected.Height <= PIC_Y And frmMapEditor.shpSelected.Width <= PIC_X Then
                    If frmMapEditor.optTiles = True Then
                        With Map(GetPlayerMap(MyIndex)).Tile(x1, y1)
                            If frmMapEditor.optGround.Value = True Then
                                .Ground = EditorTileY * TilesInSheets + EditorTileX
                                .GroundSet = EditorSet
                            End If
                            If frmMapEditor.optMask.Value = True Then
                                .Mask = EditorTileY * TilesInSheets + EditorTileX
                                .MaskSet = EditorSet
                            End If
                            If frmMapEditor.optAnim.Value = True Then
                                .Anim = EditorTileY * TilesInSheets + EditorTileX
                                .AnimSet = EditorSet
                            End If
                            If frmMapEditor.optMask2.Value = True Then
                                .Mask2 = EditorTileY * TilesInSheets + EditorTileX
                                .Mask2Set = EditorSet
                            End If
                            If frmMapEditor.optM2Anim.Value = True Then
                                .M2Anim = EditorTileY * TilesInSheets + EditorTileX
                                .M2AnimSet = EditorSet
                            End If
                            If frmMapEditor.optFringe.Value = True Then
                                .Fringe = EditorTileY * TilesInSheets + EditorTileX
                                .FringeSet = EditorSet
                            End If
                            If frmMapEditor.optFAnim.Value = True Then
                                .FAnim = EditorTileY * TilesInSheets + EditorTileX
                                .FAnimSet = EditorSet
                            End If
                            If frmMapEditor.optFringe2.Value = True Then
                                .Fringe2 = EditorTileY * TilesInSheets + EditorTileX
                                .Fringe2Set = EditorSet
                            End If
                            If frmMapEditor.optF2Anim.Value = True Then
                                .F2Anim = EditorTileY * TilesInSheets + EditorTileX
                                .F2AnimSet = EditorSet
                            End If
                        End With
                    ElseIf frmMapEditor.optlight.Value = True Then
                        Map(GetPlayerMap(MyIndex)).Tile(x1, y1).Light = EditorTileY * TilesInSheets + EditorTileX
                    ElseIf frmMapEditor.optAttributes = True Then
                        With Map(GetPlayerMap(MyIndex)).Tile(x1, y1)
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
                            If frmMapEditor.optBank.Value = True Then
                                .Type = TILE_TYPE_BANK
                                .Data1 = 0
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = ""
                                .String2 = ""
                                .String3 = ""
                            End If
                            If frmMapEditor.optHouseBuy.Value = True Then
                              .Type = TILE_TYPE_HOUSE_BUY
                              .Data1 = HouseItem
                              .Data2 = HousePrice
                              .Data3 = 0
                              .String1 = ""
                              .String2 = ""
                              .String3 = ""
                            End If
                            If frmMapEditor.optHouse.Value = True Then
                              .Type = TILE_TYPE_HOUSE
                              .Data1 = 0
                              .Data2 = 0
                              .Data3 = 0
                              .String1 = ""
                              .String2 = ""
                              .String3 = ""
                            End If
                            If frmMapEditor.optFurniture.Value = True Then
                                .Type = TILE_TYPE_FURNITURE
                                .Data1 = FurnitureNum
                                .Data2 = 0
                                .Data3 = 0
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
                                    If frmMapEditor.optTiles = True Then
                                        With Map(GetPlayerMap(MyIndex)).Tile(x1 + x2, y1 + y2)
                                            If frmMapEditor.optGround.Value = True Then
                                                .Ground = (EditorTileY + y2) * TilesInSheets + (EditorTileX + x2)
                                                .GroundSet = EditorSet
                                            End If
                                            If frmMapEditor.optMask.Value = True Then
                                                .Mask = (EditorTileY + y2) * TilesInSheets + (EditorTileX + x2)
                                                .MaskSet = EditorSet
                                            End If
                                            If frmMapEditor.optAnim.Value = True Then
                                                .Anim = (EditorTileY + y2) * TilesInSheets + (EditorTileX + x2)
                                                .AnimSet = EditorSet
                                            End If
                                            If frmMapEditor.optMask2.Value = True Then
                                                .Mask2 = (EditorTileY + y2) * TilesInSheets + (EditorTileX + x2)
                                                .Mask2Set = EditorSet
                                            End If
                                            If frmMapEditor.optM2Anim.Value = True Then
                                                .M2Anim = (EditorTileY + y2) * TilesInSheets + (EditorTileX + x2)
                                                .M2AnimSet = EditorSet
                                            End If
                                            If frmMapEditor.optFringe.Value = True Then
                                                .Fringe = (EditorTileY + y2) * TilesInSheets + (EditorTileX + x2)
                                                .FringeSet = EditorSet
                                            End If
                                            If frmMapEditor.optFAnim.Value = True Then
                                                .FAnim = (EditorTileY + y2) * TilesInSheets + (EditorTileX + x2)
                                                .FAnimSet = EditorSet
                                            End If
                                            If frmMapEditor.optFringe2.Value = True Then
                                                .Fringe2 = (EditorTileY + y2) * TilesInSheets + (EditorTileX + x2)
                                                .Fringe2Set = EditorSet
                                            End If
                                            If frmMapEditor.optF2Anim.Value = True Then
                                                .F2Anim = (EditorTileY + y2) * TilesInSheets + (EditorTileX + x2)
                                                .F2AnimSet = EditorSet
                                            End If
                                        End With
                                    ElseIf frmMapEditor.optlight.Value = True Then
                                        Map(GetPlayerMap(MyIndex)).Tile(x1 + x2, y1 + y2).Light = (EditorTileY + y2) * TilesInSheets + (EditorTileX + x2)
                                    End If
                                End If
                            End If
                        Next x2
                    Next y2
                End If
            End If
           
            If (Button = 2) And (x1 >= 0) And (x1 <= MAX_MAPX) And (y1 >= 0) And (y1 <= MAX_MAPY) Then
                If frmMapEditor.optTiles.Value = True Then
                    With Map(GetPlayerMap(MyIndex)).Tile(x1, y1)
                        If frmMapEditor.optGround.Value = True Then
                            .Ground = 0
                            .GroundSet = -1
                        End If
                        If frmMapEditor.optMask.Value = True Then
                            .Mask = 0
                            .MaskSet = -1
                        End If
                        If frmMapEditor.optAnim.Value = True Then
                            .Anim = 0
                            .AnimSet = -1
                        End If
                        If frmMapEditor.optMask2.Value = True Then
                            .Mask2 = 0
                            .Mask2Set = -1
                        End If
                        If frmMapEditor.optM2Anim.Value = True Then
                            .M2Anim = 0
                            .M2AnimSet = -1
                        End If
                        If frmMapEditor.optFringe.Value = True Then
                            .Fringe = 0
                            .FringeSet = -1
                        End If
                        If frmMapEditor.optFAnim.Value = True Then
                            .FAnim = 0
                            .FAnimSet = -1
                        End If
                        If frmMapEditor.optFringe2.Value = True Then
                            .Fringe2 = 0
                            .Fringe2Set = -1
                        End If
                        If frmMapEditor.optF2Anim.Value = True Then
                            .F2Anim = 0
                            .F2AnimSet = -1
                        End If
                    End With
                ElseIf frmMapEditor.optlight.Value = True Then
                    Map(GetPlayerMap(MyIndex)).Tile(x1, y1).Light = 0
                ElseIf frmMapEditor.optAttributes.Value = True Then
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

Public Sub PetMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim x1, y1 As Long
    If Player(MyIndex).Pet.Alive = NO Then Exit Sub
    
    x1 = Int(x / PIC_X)
    y1 = Int(y / PIC_Y)
    
    If (Button = 1) And (x1 >= 0) And (x1 <= MAX_MAPX) And (y1 >= 0) And (y1 <= MAX_MAPY) Then
        Call SendData("PETMOVESELECT" & SEP_CHAR & x1 & SEP_CHAR & y1 & SEP_CHAR & END_CHAR)
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
    frmMapEditor.picBackSelect.Top = (frmMapEditor.scrlPicture.Value * PIC_Y) * -1
End Sub

Public Sub EditorSend()
    Call SendMap
    Call EditorCancel
End Sub

Public Sub EditorCancel()
    InEditor = False
    InSpawnEditor = False
    frmMapEditor.Visible = False
    frmMapProperties.Visible = False
    frmMirage.Show
    frmMapEditor.MousePointer = 1
    frmMirage.MousePointer = 1
    Call LoadMap(GetPlayerMap(MyIndex))
    'frmMirage.picMapEditor.Visible = False
End Sub

Public Sub EditorClearLayer()
Dim YesNo As Long, x As Long, y As Long

    ' Ground layer
    If frmMapEditor.optGround.Value = True Then
    YesNo = MsgBox("Are you sure you wish to clear the ground layer?", vbYesNo, GAME_NAME)
        If YesNo = vbYes Then
            For y = 0 To MAX_MAPY
                For x = 0 To MAX_MAPX
                    Map(GetPlayerMap(MyIndex)).Tile(x, y).Ground = 0
                    Map(GetPlayerMap(MyIndex)).Tile(x, y).GroundSet = -1
                Next x
            Next y
        End If
    End If

    ' Mask layer
    If frmMapEditor.optMask.Value = True Then
    YesNo = MsgBox("Are you sure you wish to clear the mask layer?", vbYesNo, GAME_NAME)
        If YesNo = vbYes Then
            For y = 0 To MAX_MAPY
                For x = 0 To MAX_MAPX
                    Map(GetPlayerMap(MyIndex)).Tile(x, y).Mask = 0
                    Map(GetPlayerMap(MyIndex)).Tile(x, y).MaskSet = -1
                Next x
            Next y
        End If
    End If
   
    ' Mask Animation layer
    If frmMapEditor.optAnim.Value = True Then
    YesNo = MsgBox("Are you sure you wish to clear the animation layer?", vbYesNo, GAME_NAME)
        If YesNo = vbYes Then
            For y = 0 To MAX_MAPY
                For x = 0 To MAX_MAPX
                    Map(GetPlayerMap(MyIndex)).Tile(x, y).Anim = 0
                    Map(GetPlayerMap(MyIndex)).Tile(x, y).AnimSet = -1
                Next x
            Next y
        End If
    End If
   
    ' Mask 2 layer
    If frmMapEditor.optMask2.Value = True Then
    YesNo = MsgBox("Are you sure you wish to clear the mask 2 layer?", vbYesNo, GAME_NAME)
        If YesNo = vbYes Then
            For y = 0 To MAX_MAPY
                For x = 0 To MAX_MAPX
                    Map(GetPlayerMap(MyIndex)).Tile(x, y).Mask2 = 0
                    Map(GetPlayerMap(MyIndex)).Tile(x, y).Mask2Set = -1
                Next x
            Next y
        End If
    End If
   
    ' Mask 2 Animation layer
    If frmMapEditor.optM2Anim.Value = True Then
    YesNo = MsgBox("Are you sure you wish to clear the mask 2 animation layer?", vbYesNo, GAME_NAME)
        If YesNo = vbYes Then
            For y = 0 To MAX_MAPY
                For x = 0 To MAX_MAPX
                    Map(GetPlayerMap(MyIndex)).Tile(x, y).M2Anim = 0
                    Map(GetPlayerMap(MyIndex)).Tile(x, y).M2AnimSet = -1
                Next x
            Next y
        End If
    End If
   
    ' Fringe layer
    If frmMapEditor.optFringe.Value = True Then
    YesNo = MsgBox("Are you sure you wish to clear the fringe layer?", vbYesNo, GAME_NAME)
        If YesNo = vbYes Then
            For y = 0 To MAX_MAPY
                For x = 0 To MAX_MAPX
                    Map(GetPlayerMap(MyIndex)).Tile(x, y).Fringe = 0
                    Map(GetPlayerMap(MyIndex)).Tile(x, y).FringeSet = -1
                Next x
            Next y
        End If
    End If
   
    ' Fringe Animation layer
    If frmMapEditor.optFAnim.Value = True Then
    YesNo = MsgBox("Are you sure you wish to clear the fringe animation layer?", vbYesNo, GAME_NAME)
        If YesNo = vbYes Then
            For y = 0 To MAX_MAPY
                For x = 0 To MAX_MAPX
                    Map(GetPlayerMap(MyIndex)).Tile(x, y).FAnim = 0
                    Map(GetPlayerMap(MyIndex)).Tile(x, y).FAnimSet = -1
                Next x
            Next y
        End If
    End If
   
    ' Fringe 2 layer
    If frmMapEditor.optFringe2.Value = True Then
    YesNo = MsgBox("Are you sure you wish to clear the fringe 2 layer?", vbYesNo, GAME_NAME)
        If YesNo = vbYes Then
            For y = 0 To MAX_MAPY
                For x = 0 To MAX_MAPX
                    Map(GetPlayerMap(MyIndex)).Tile(x, y).Fringe2 = 0
                    Map(GetPlayerMap(MyIndex)).Tile(x, y).Fringe2Set = -1
                Next x
            Next y
        End If
    End If
   
    ' Fringe 2 Animation layer
    If frmMapEditor.optF2Anim.Value = True Then
    YesNo = MsgBox("Are you sure you wish to clear the fringe 2 animation layer?", vbYesNo, GAME_NAME)
        If YesNo = vbYes Then
            For y = 0 To MAX_MAPY
                For x = 0 To MAX_MAPX
                    Map(GetPlayerMap(MyIndex)).Tile(x, y).F2Anim = 0
                    Map(GetPlayerMap(MyIndex)).Tile(x, y).F2AnimSet = -1
                Next x
            Next y
        End If
    End If
End Sub

Public Sub EditorClearMap()
Dim YesNo As Long, x As Long, y As Long

    YesNo = MsgBox("Are you sure you wish to clear the whole map?", vbYesNo, GAME_NAME)
    If YesNo = vbYes Then
        For y = 0 To MAX_MAPY
            For x = 0 To MAX_MAPX
                Map(GetPlayerMap(MyIndex)).Tile(x, y).Ground = 0
                Map(GetPlayerMap(MyIndex)).Tile(x, y).GroundSet = -1
                Map(GetPlayerMap(MyIndex)).Tile(x, y).Mask = 0
                Map(GetPlayerMap(MyIndex)).Tile(x, y).MaskSet = -1
                Map(GetPlayerMap(MyIndex)).Tile(x, y).Anim = 0
                Map(GetPlayerMap(MyIndex)).Tile(x, y).AnimSet = -1
                Map(GetPlayerMap(MyIndex)).Tile(x, y).Mask2 = 0
                Map(GetPlayerMap(MyIndex)).Tile(x, y).Mask2Set = -1
                Map(GetPlayerMap(MyIndex)).Tile(x, y).M2Anim = 0
                Map(GetPlayerMap(MyIndex)).Tile(x, y).M2AnimSet = -1
                Map(GetPlayerMap(MyIndex)).Tile(x, y).Fringe = 0
                Map(GetPlayerMap(MyIndex)).Tile(x, y).FringeSet = -1
                Map(GetPlayerMap(MyIndex)).Tile(x, y).FAnim = 0
                Map(GetPlayerMap(MyIndex)).Tile(x, y).FAnimSet = -1
                Map(GetPlayerMap(MyIndex)).Tile(x, y).Fringe2 = 0
                Map(GetPlayerMap(MyIndex)).Tile(x, y).Fringe2Set = -1
                Map(GetPlayerMap(MyIndex)).Tile(x, y).F2Anim = 0
                Map(GetPlayerMap(MyIndex)).Tile(x, y).F2AnimSet = -1
            Next x
        Next y
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
Dim sDc As Long

    frmEmoticonEditor.scrlEmoticon.Max = MAX_EMOTICONS
    frmEmoticonEditor.scrlEmoticon.Value = Emoticons(EditorIndex - 1).pic
    frmEmoticonEditor.txtCommand.Text = Trim(Emoticons(EditorIndex - 1).Command)
    'frmEmoticonEditor.picEmoticons.Picture = LoadPicture(App.Path & "\GFX\emoticons.bmp")
    
     sDc = DD_EmoticonSurf.GetDC
    With frmEmoticonEditor.picEmoticons
        .Width = DDSD_Emoticon.lWidth
        .Height = DDSD_Emoticon.lHeight
        .Cls
        Call BitBlt(.hDC, 0, 0, .Width, .Height, sDc, 0, 0, SRCCOPY)
    End With
    Call DD_EmoticonSurf.ReleaseDC(sDc)
    
    If Emoticons(EditorIndex - 1).Type = EMOTICON_TYPE_BOTH Or Emoticons(EditorIndex - 1).Type = EMOTICON_TYPE_IMAGE Then frmEmoticonEditor.chkPic.Value = 1
    If Emoticons(EditorIndex - 1).Type = EMOTICON_TYPE_BOTH Or Emoticons(EditorIndex - 1).Type = EMOTICON_TYPE_SOUND Then frmEmoticonEditor.chkSound.Value = 1
    
    frmEmoticonEditor.Show vbModal
End Sub

Public Sub ElementEditorInit()
    frmElementEditor.txtName.Text = Trim(Element(EditorIndex - 1).name)
    frmElementEditor.scrlStrong.Value = Element(EditorIndex - 1).Strong
    frmElementEditor.scrlWeak.Value = Element(EditorIndex - 1).Weak
    frmElementEditor.Show vbModal
End Sub

Public Sub EmoticonEditorOk()
    Emoticons(EditorIndex - 1).pic = frmEmoticonEditor.scrlEmoticon.Value
    Emoticons(EditorIndex - 1).Sound = frmEmoticonEditor.lstSound.Text
    If frmEmoticonEditor.txtCommand.Text <> "/" Then
        Emoticons(EditorIndex - 1).Command = frmEmoticonEditor.txtCommand.Text
    Else
        Emoticons(EditorIndex - 1).Command = ""
    End If
    If frmEmoticonEditor.chkPic.Value = 1 Then
        If frmEmoticonEditor.chkSound.Value = 1 Then
            Emoticons(EditorIndex - 1).Type = EMOTICON_TYPE_BOTH
        Else
            Emoticons(EditorIndex - 1).Type = EMOTICON_TYPE_IMAGE
        End If
    Else
        If frmEmoticonEditor.chkSound.Value = 1 Then
            Emoticons(EditorIndex - 1).Type = EMOTICON_TYPE_SOUND
        Else
            Call MsgBox("You need to pick either a picture, a sound, or both")
        End If
    End If
    
    Call SendSaveEmoticon(EditorIndex - 1)
    Call EmoticonEditorCancel
End Sub

Public Sub ElementEditorOk()
    Element(EditorIndex - 1).name = frmElementEditor.txtName.Text
    Element(EditorIndex - 1).Strong = frmElementEditor.scrlStrong.Value
    Element(EditorIndex - 1).Weak = frmElementEditor.scrlWeak.Value
    Call SendSaveElement(EditorIndex - 1)
    Call ElementEditorCancel
End Sub

Public Sub EmoticonEditorCancel()
    InEmoticonEditor = False
    Unload frmEmoticonEditor
End Sub

Public Sub ElementEditorCancel()
    InElementEditor = False
    Unload frmElementEditor
End Sub

Public Sub ArrowEditorInit()
Dim sDc As Long
    frmEditArrows.scrlArrow.Max = MAX_ARROWS
    If Arrows(EditorIndex).pic = 0 Then Arrows(EditorIndex).pic = 1
    frmEditArrows.scrlArrow.Value = Arrows(EditorIndex).pic
    frmEditArrows.txtName.Text = Arrows(EditorIndex).name
    If Arrows(EditorIndex).Range = 0 Then Arrows(EditorIndex).Range = 1
    frmEditArrows.scrlRange.Value = Arrows(EditorIndex).Range
    'frmEditArrows.picArrows.Picture = LoadPicture(App.Path & "\GFX\arrows.bmp")
    
    sDc = DD_ArrowAnim.GetDC
    With frmEditArrows.picArrows
        .Width = DDSD_ArrowAnim.lWidth
        .Height = DDSD_ArrowAnim.lHeight
        .Cls
        Call BitBlt(.hDC, 0, 0, .Width, .Height, sDc, 0, 0, SRCCOPY)
    End With
    Call DD_ArrowAnim.ReleaseDC(sDc)
    
    frmEditArrows.Show vbModal
End Sub

Public Sub SpeechEditorInit()
Dim P As Long

    frmSpeech.scrlNumber.Max = MAX_SPEECH_OPTIONS
    frmSpeech.scrlNumber.Value = 0
    frmSpeech.scrlGoTo(0).Max = MAX_SPEECH_OPTIONS
    frmSpeech.scrlGoTo(1).Max = MAX_SPEECH_OPTIONS
    frmSpeech.scrlGoTo(2).Max = MAX_SPEECH_OPTIONS
    frmSpeech.lblSection.Caption = "0"
    frmSpeech.chkQuit.Enabled = False
    frmSpeech.chkScript.Enabled = False
    
    If Trim(Speech(EditorIndex).name) = "" Then
        frmSpeech.lblWarn.Visible = True
    Else
        frmSpeech.lblWarn.Visible = False
    End If
    
    frmSpeech.txtName.Text = Speech(EditorIndex).name
    frmSpeech.chkQuit.Value = Speech(EditorIndex).Num(0).Exit
    frmSpeech.txtMainTalk.Text = Speech(EditorIndex).Num(0).Text
    frmSpeech.optSaid(Speech(EditorIndex).Num(0).SaidBy).Value = True
    If Speech(EditorIndex).Num(0).Respond > 0 Then
        frmSpeech.chkRespond.Value = 1
    Else
        frmSpeech.chkRespond.Value = 0
    End If
    
    If Speech(EditorIndex).Num(0).Script > 0 Then
        frmSpeech.chkScript.Value = 1
        frmSpeech.scrlScript.Visible = True
        frmSpeech.scrlScript.Value = Speech(EditorIndex).Num(0).Script
        frmSpeech.lblScript.Visible = True
        frmSpeech.lblScript.Caption = Speech(EditorIndex).Num(0).Script
    Else
        frmSpeech.chkScript.Value = 0
        frmSpeech.scrlScript.Visible = False
        frmSpeech.scrlScript.Value = 0
        frmSpeech.lblScript.Visible = False
        frmSpeech.lblScript.Caption = 0
    End If
    
    For P = 1 To 3
        If frmSpeech.chkRespond.Value = 1 Then
            frmSpeech.optResponces(P - 1).Enabled = True
            frmSpeech.txtTalk(P - 1).Enabled = True
            frmSpeech.scrlGoTo(P - 1).Enabled = True
            frmSpeech.lblGoTo(P - 1).Enabled = True
            frmSpeech.chkExit(P - 1).Enabled = True
            
            If Speech(EditorIndex).Num(0).Respond = P Then
                frmSpeech.optResponces(P - 1).Value = True
            End If
        
            frmSpeech.txtTalk(P - 1).Text = Speech(EditorIndex).Num(0).Responces(P).Text
            frmSpeech.scrlGoTo(P - 1).Value = Speech(EditorIndex).Num(0).Responces(P).GoTo
            frmSpeech.lblGoTo(P - 1).Caption = "Go to " & Speech(EditorIndex).Num(0).Responces(P).GoTo
            frmSpeech.chkExit(P - 1).Value = Speech(EditorIndex).Num(0).Responces(P).Exit
        Else
            frmSpeech.optResponces(P - 1).Enabled = False
            frmSpeech.txtTalk(P - 1).Enabled = False
            frmSpeech.scrlGoTo(P - 1).Enabled = False
            frmSpeech.lblGoTo(P - 1).Enabled = False
            frmSpeech.chkExit(P - 1).Enabled = False
        End If
    Next P
    
    SpeechEditorCurrentNumber = 0
    
    frmSpeech.Show vbModal
End Sub

Public Sub ArrowEditorOk()
    Arrows(EditorIndex).pic = frmEditArrows.scrlArrow.Value
    Arrows(EditorIndex).Range = frmEditArrows.scrlRange.Value
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
Dim sDc As Long
    EditorItemY = Int(Item(EditorIndex).pic / 6)
    EditorItemX = (Item(EditorIndex).pic - Int(Item(EditorIndex).pic / 6) * 6)
    
    frmItemEditor.scrlClassReq.Max = Max_Classes

    sDc = DD_ItemSurf.GetDC
    With frmItemEditor.picItems
        .Cls
        .Width = DDSD_Item.lWidth
        .Height = DDSD_Item.lHeight
        Call BitBlt(.hDC, 0, 0, .Width, .Height, sDc, 0, 0, SRCCOPY)
    End With
    Call DD_ItemSurf.ReleaseDC(sDc)

    'frmItemEditor.picItems.Picture = LoadPicture(App.Path & "\GFX\items.bmp")
    
    frmItemEditor.txtName.Text = Trim(Item(EditorIndex).name)
    frmItemEditor.txtDesc.Text = Trim(Item(EditorIndex).desc)
    frmItemEditor.cmbType.ListIndex = Item(EditorIndex).Type
    frmItemEditor.chkBound.Value = Item(EditorIndex).Bound
    frmItemEditor.chkStackable.Value = Item(EditorIndex).Stackable
    
     frmItemEditor.txtPrice.Text = Item(EditorIndex).Price
    
    If (frmItemEditor.cmbType.ListIndex >= ITEM_TYPE_WEAPON) And (frmItemEditor.cmbType.ListIndex <= ITEM_TYPE_SHIELD) Then
        frmItemEditor.fraEquipment.Visible = True
        frmItemEditor.fraAttributes.Visible = True
        frmItemEditor.fraBow.Visible = True
        
        If Item(EditorIndex).Data1 >= 0 Then
            frmItemEditor.scrlDurability.Value = Item(EditorIndex).Data1
        Else
            frmItemEditor.scrlDurability.Value = Item(EditorIndex).Data1 * -1
        End If
        frmItemEditor.scrlStrength.Value = Item(EditorIndex).Data2
        frmItemEditor.scrlStrReq.Value = Item(EditorIndex).StrReq
        frmItemEditor.scrlDefReq.Value = Item(EditorIndex).DefReq
        frmItemEditor.scrlSpeedReq.Value = Item(EditorIndex).SpeedReq
        frmItemEditor.scrlMagicReq.Value = Item(EditorIndex).MagicReq
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
        frmItemEditor.scrlAttackSpeed.Value = Item(EditorIndex).AttackSpeed
        If Item(EditorIndex).Data3 > 0 Then
            frmItemEditor.chkBow.Value = Checked
        Else
            frmItemEditor.chkBow.Value = Unchecked
        End If
        
        frmItemEditor.cmbBow.Clear
        If frmItemEditor.chkBow.Value = Checked Then
            For i = 1 To 100
                frmItemEditor.cmbBow.AddItem i & ": " & Arrows(i).name
            Next i
            frmItemEditor.cmbBow.ListIndex = Item(EditorIndex).Data3 - 1
            frmItemEditor.picBow.Top = (Arrows(Item(EditorIndex).Data3).pic * 32) * -1
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
    
    If (frmItemEditor.cmbType.ListIndex = ITEM_TYPE_PET) Then
        frmItemEditor.scrlPet.Value = Item(EditorIndex).Data1
        frmItemEditor.scrlPetLevel.Value = Item(EditorIndex).Data2
    End If
    
    If (frmItemEditor.cmbType.ListIndex = ITEM_TYPE_SCRIPTED) Then
       frmItemEditor.fraScript.Visible = True
       frmItemEditor.scrlScript.Value = Item(EditorIndex).Data1
   Else
        frmItemEditor.fraScript.Visible = False
   End If
   
   If (frmItemEditor.cmbType.ListIndex = ITEM_TYPE_LEGS) Then
        frmItemEditor.fraEquipment.Visible = True
        frmItemEditor.fraAttributes.Visible = True
        frmItemEditor.scrlDurability.Value = Item(EditorIndex).Data1
        frmItemEditor.scrlStrength.Value = Item(EditorIndex).Data2
        frmItemEditor.scrlStrReq.Value = Item(EditorIndex).StrReq
        frmItemEditor.scrlDefReq.Value = Item(EditorIndex).DefReq
        frmItemEditor.scrlSpeedReq.Value = Item(EditorIndex).SpeedReq
        frmItemEditor.scrlMagicReq.Value = Item(EditorIndex).MagicReq
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
        frmItemEditor.scrlAttackSpeed.Value = Item(EditorIndex).AttackSpeed
        End If
        
        If (frmItemEditor.cmbType.ListIndex = ITEM_TYPE_BOOTS) Then
        frmItemEditor.fraEquipment.Visible = True
        frmItemEditor.fraAttributes.Visible = True
        frmItemEditor.scrlDurability.Value = Item(EditorIndex).Data1
        frmItemEditor.scrlStrength.Value = Item(EditorIndex).Data2
        frmItemEditor.scrlStrReq.Value = Item(EditorIndex).StrReq
        frmItemEditor.scrlDefReq.Value = Item(EditorIndex).DefReq
        frmItemEditor.scrlSpeedReq.Value = Item(EditorIndex).SpeedReq
        frmItemEditor.scrlMagicReq.Value = Item(EditorIndex).MagicReq
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
        frmItemEditor.scrlAttackSpeed.Value = Item(EditorIndex).AttackSpeed
        End If
        
        If (frmItemEditor.cmbType.ListIndex = ITEM_TYPE_GLOVES) Then
        frmItemEditor.fraEquipment.Visible = True
        frmItemEditor.fraAttributes.Visible = True
        frmItemEditor.scrlDurability.Value = Item(EditorIndex).Data1
        frmItemEditor.scrlStrength.Value = Item(EditorIndex).Data2
        frmItemEditor.scrlStrReq.Value = Item(EditorIndex).StrReq
        frmItemEditor.scrlDefReq.Value = Item(EditorIndex).DefReq
        frmItemEditor.scrlSpeedReq.Value = Item(EditorIndex).SpeedReq
        frmItemEditor.scrlMagicReq.Value = Item(EditorIndex).MagicReq
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
        frmItemEditor.scrlAttackSpeed.Value = Item(EditorIndex).AttackSpeed
        End If
        
        If (frmItemEditor.cmbType.ListIndex = ITEM_TYPE_RING1) Then
        frmItemEditor.fraEquipment.Visible = True
        frmItemEditor.fraAttributes.Visible = True
        frmItemEditor.scrlDurability.Value = Item(EditorIndex).Data1
        frmItemEditor.scrlStrength.Value = Item(EditorIndex).Data2
        frmItemEditor.scrlStrReq.Value = Item(EditorIndex).StrReq
        frmItemEditor.scrlDefReq.Value = Item(EditorIndex).DefReq
        frmItemEditor.scrlSpeedReq.Value = Item(EditorIndex).SpeedReq
        frmItemEditor.scrlMagicReq.Value = Item(EditorIndex).MagicReq
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
        frmItemEditor.scrlAttackSpeed.Value = Item(EditorIndex).AttackSpeed
        End If
        
        If (frmItemEditor.cmbType.ListIndex = ITEM_TYPE_RING2) Then
        frmItemEditor.fraEquipment.Visible = True
        frmItemEditor.fraAttributes.Visible = True
        frmItemEditor.scrlDurability.Value = Item(EditorIndex).Data1
        frmItemEditor.scrlStrength.Value = Item(EditorIndex).Data2
        frmItemEditor.scrlStrReq.Value = Item(EditorIndex).StrReq
        frmItemEditor.scrlDefReq.Value = Item(EditorIndex).DefReq
        frmItemEditor.scrlSpeedReq.Value = Item(EditorIndex).SpeedReq
        frmItemEditor.scrlMagicReq.Value = Item(EditorIndex).MagicReq
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
        frmItemEditor.scrlAttackSpeed.Value = Item(EditorIndex).AttackSpeed
        End If
        
        If (frmItemEditor.cmbType.ListIndex = ITEM_TYPE_AMULET) Then
        frmItemEditor.fraEquipment.Visible = True
        frmItemEditor.fraAttributes.Visible = True
        frmItemEditor.scrlDurability.Value = Item(EditorIndex).Data1
        frmItemEditor.scrlStrength.Value = Item(EditorIndex).Data2
        frmItemEditor.scrlStrReq.Value = Item(EditorIndex).StrReq
        frmItemEditor.scrlDefReq.Value = Item(EditorIndex).DefReq
        frmItemEditor.scrlSpeedReq.Value = Item(EditorIndex).SpeedReq
        frmItemEditor.scrlMagicReq.Value = Item(EditorIndex).MagicReq
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
        frmItemEditor.scrlAttackSpeed.Value = Item(EditorIndex).AttackSpeed
        End If

    frmItemEditor.Show vbModal
End Sub

Public Sub ItemEditorOk()
    Item(EditorIndex).name = frmItemEditor.txtName.Text
    Item(EditorIndex).desc = frmItemEditor.txtDesc.Text
    Item(EditorIndex).pic = EditorItemY * 6 + EditorItemX
    Item(EditorIndex).Type = frmItemEditor.cmbType.ListIndex
     Item(EditorIndex).Price = Val(frmItemEditor.txtPrice.Text)
     Item(EditorIndex).Bound = frmItemEditor.chkBound.Value
     Item(EditorIndex).Stackable = frmItemEditor.chkStackable.Value


    If (frmItemEditor.cmbType.ListIndex >= ITEM_TYPE_WEAPON) And (frmItemEditor.cmbType.ListIndex <= ITEM_TYPE_SHIELD) Then
        Item(EditorIndex).Data1 = frmItemEditor.scrlDurability.Value
        If frmItemEditor.chkRepair.Value = 0 Then Item(EditorIndex).Data1 = Item(EditorIndex).Data1 * -1
        Item(EditorIndex).Data2 = frmItemEditor.scrlStrength.Value
        If frmItemEditor.chkBow.Value = Checked Then
            Item(EditorIndex).Data3 = frmItemEditor.cmbBow.ListIndex + 1
        Else
            Item(EditorIndex).Data3 = 0
        End If
        Item(EditorIndex).StrReq = frmItemEditor.scrlStrReq.Value
        Item(EditorIndex).DefReq = frmItemEditor.scrlDefReq.Value
        Item(EditorIndex).SpeedReq = frmItemEditor.scrlSpeedReq.Value
        Item(EditorIndex).MagicReq = frmItemEditor.scrlMagicReq.Value
        
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
        Item(EditorIndex).AttackSpeed = frmItemEditor.scrlAttackSpeed.Value
        Item(EditorIndex).Stackable = 0
    End If
    
    If (frmItemEditor.cmbType.ListIndex >= ITEM_TYPE_POTIONADDHP) And (frmItemEditor.cmbType.ListIndex <= ITEM_TYPE_POTIONSUBSP) Then
        Item(EditorIndex).Data1 = frmItemEditor.scrlVitalMod.Value
        Item(EditorIndex).Data2 = 0
        Item(EditorIndex).Data3 = 0
        Item(EditorIndex).StrReq = 0
        Item(EditorIndex).DefReq = 0
        Item(EditorIndex).SpeedReq = 0
        Item(EditorIndex).MagicReq = 0
        Item(EditorIndex).ClassReq = 0
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
        Item(EditorIndex).Stackable = frmItemEditor.chkStackable.Value
    End If
    
    If (frmItemEditor.cmbType.ListIndex = ITEM_TYPE_SPELL) Then
        Item(EditorIndex).Data1 = frmItemEditor.scrlSpell.Value
        Item(EditorIndex).Data2 = 0
        Item(EditorIndex).Data3 = 0
        Item(EditorIndex).StrReq = 0
        Item(EditorIndex).DefReq = 0
        Item(EditorIndex).SpeedReq = 0
        Item(EditorIndex).MagicReq = 0
        Item(EditorIndex).ClassReq = 0
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
        Item(EditorIndex).Stackable = 0
    End If
    
    If (frmItemEditor.cmbType.ListIndex = ITEM_TYPE_PET) Then
        Item(EditorIndex).Data1 = frmItemEditor.scrlPet.Value
        Item(EditorIndex).Data2 = frmItemEditor.scrlPetLevel.Value
        Item(EditorIndex).Data3 = 0
        Item(EditorIndex).StrReq = 0
        Item(EditorIndex).DefReq = 0
        Item(EditorIndex).SpeedReq = 0
        Item(EditorIndex).MagicReq = 0
        Item(EditorIndex).ClassReq = 0
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
        Item(EditorIndex).Stackable = 0
    End If
    
    If (frmItemEditor.cmbType.ListIndex = ITEM_TYPE_SCRIPTED) Then
Item(EditorIndex).Data1 = frmItemEditor.scrlScript.Value
Item(EditorIndex).Data2 = 0
Item(EditorIndex).Data3 = 0
Item(EditorIndex).StrReq = 0
Item(EditorIndex).DefReq = 0
Item(EditorIndex).SpeedReq = 0
Item(EditorIndex).ClassReq = 0
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
Item(EditorIndex).Stackable = frmItemEditor.chkStackable.Value
End If

If (frmItemEditor.cmbType.ListIndex = ITEM_TYPE_LEGS) Then
        Item(EditorIndex).Data1 = frmItemEditor.scrlDurability.Value
        Item(EditorIndex).Data2 = frmItemEditor.scrlStrength.Value
        Item(EditorIndex).Data3 = 0
        Item(EditorIndex).StrReq = frmItemEditor.scrlStrReq.Value
        Item(EditorIndex).DefReq = frmItemEditor.scrlDefReq.Value
        Item(EditorIndex).SpeedReq = frmItemEditor.scrlSpeedReq.Value
        Item(EditorIndex).MagicReq = frmItemEditor.scrlMagicReq.Value
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
        Item(EditorIndex).AttackSpeed = frmItemEditor.scrlAttackSpeed.Value
        Item(EditorIndex).Stackable = 0
        End If
        
    If (frmItemEditor.cmbType.ListIndex = ITEM_TYPE_BOOTS) Then
        Item(EditorIndex).Data1 = frmItemEditor.scrlDurability.Value
        Item(EditorIndex).Data2 = frmItemEditor.scrlStrength.Value
        Item(EditorIndex).Data3 = 0
        Item(EditorIndex).StrReq = frmItemEditor.scrlStrReq.Value
        Item(EditorIndex).DefReq = frmItemEditor.scrlDefReq.Value
        Item(EditorIndex).SpeedReq = frmItemEditor.scrlSpeedReq.Value
        Item(EditorIndex).MagicReq = frmItemEditor.scrlMagicReq.Value
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
        Item(EditorIndex).AttackSpeed = frmItemEditor.scrlAttackSpeed.Value
        Item(EditorIndex).Stackable = 0
        End If
        
    If (frmItemEditor.cmbType.ListIndex = ITEM_TYPE_GLOVES) Then
        Item(EditorIndex).Data1 = frmItemEditor.scrlDurability.Value
        Item(EditorIndex).Data2 = frmItemEditor.scrlStrength.Value
        Item(EditorIndex).Data3 = 0
        Item(EditorIndex).StrReq = frmItemEditor.scrlStrReq.Value
        Item(EditorIndex).DefReq = frmItemEditor.scrlDefReq.Value
        Item(EditorIndex).SpeedReq = frmItemEditor.scrlSpeedReq.Value
        Item(EditorIndex).MagicReq = frmItemEditor.scrlMagicReq.Value
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
        Item(EditorIndex).AttackSpeed = frmItemEditor.scrlAttackSpeed.Value
        Item(EditorIndex).Stackable = 0
        End If
        
    If (frmItemEditor.cmbType.ListIndex = ITEM_TYPE_RING1) Then
        Item(EditorIndex).Data1 = frmItemEditor.scrlDurability.Value
        Item(EditorIndex).Data2 = frmItemEditor.scrlStrength.Value
        Item(EditorIndex).Data3 = 0
        Item(EditorIndex).StrReq = frmItemEditor.scrlStrReq.Value
        Item(EditorIndex).DefReq = frmItemEditor.scrlDefReq.Value
        Item(EditorIndex).SpeedReq = frmItemEditor.scrlSpeedReq.Value
        Item(EditorIndex).MagicReq = frmItemEditor.scrlMagicReq.Value
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
        Item(EditorIndex).AttackSpeed = frmItemEditor.scrlAttackSpeed.Value
        Item(EditorIndex).Stackable = 0
        End If
        
    If (frmItemEditor.cmbType.ListIndex = ITEM_TYPE_RING2) Then
        Item(EditorIndex).Data1 = frmItemEditor.scrlDurability.Value
        Item(EditorIndex).Data2 = frmItemEditor.scrlStrength.Value
        Item(EditorIndex).Data3 = 0
        Item(EditorIndex).StrReq = frmItemEditor.scrlStrReq.Value
        Item(EditorIndex).DefReq = frmItemEditor.scrlDefReq.Value
        Item(EditorIndex).SpeedReq = frmItemEditor.scrlSpeedReq.Value
        Item(EditorIndex).MagicReq = frmItemEditor.scrlMagicReq.Value
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
        Item(EditorIndex).AttackSpeed = frmItemEditor.scrlAttackSpeed.Value
        Item(EditorIndex).Stackable = 0
        End If
        
    If (frmItemEditor.cmbType.ListIndex = ITEM_TYPE_AMULET) Then
        Item(EditorIndex).Data1 = frmItemEditor.scrlDurability.Value
        Item(EditorIndex).Data2 = frmItemEditor.scrlStrength.Value
        Item(EditorIndex).Data3 = 0
        Item(EditorIndex).StrReq = frmItemEditor.scrlStrReq.Value
        Item(EditorIndex).DefReq = frmItemEditor.scrlDefReq.Value
        Item(EditorIndex).SpeedReq = frmItemEditor.scrlSpeedReq.Value
        Item(EditorIndex).MagicReq = frmItemEditor.scrlMagicReq.Value
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
        Item(EditorIndex).AttackSpeed = frmItemEditor.scrlAttackSpeed.Value
        Item(EditorIndex).Stackable = 0
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
    
    'frmNpcEditor.Picsprites.Picture = LoadPicture(App.Path & "\GFX\sprites.bmp")
    
    frmNpcEditor.txtName.Text = Trim(Npc(EditorIndex).name)
    frmNpcEditor.txtAttackSay.Text = Trim(Npc(EditorIndex).AttackSay)
    frmNpcEditor.scrlSprite.Value = Npc(EditorIndex).Sprite
    frmNpcEditor.txtSpawnSecs.Text = STR(Npc(EditorIndex).SpawnSecs)
    frmNpcEditor.cmbBehavior.ListIndex = Npc(EditorIndex).Behavior
    frmNpcEditor.scrlRange.Value = Npc(EditorIndex).Range
    frmNpcEditor.scrlSTR.Value = Npc(EditorIndex).STR
    frmNpcEditor.scrlDEF.Value = Npc(EditorIndex).DEF
    frmNpcEditor.scrlSPEED.Value = Npc(EditorIndex).speed
    frmNpcEditor.scrlMAGI.Value = Npc(EditorIndex).MAGI
    frmNpcEditor.BigNpc.Value = Npc(EditorIndex).Big
    frmNpcEditor.scrlElement.Value = Npc(EditorIndex).Element
    If Npc(EditorIndex).MaxHP = 0 Then
        frmNpcEditor.StartHP.Value = 1
    Else
        frmNpcEditor.StartHP.Value = Npc(EditorIndex).MaxHP
    End If
    frmNpcEditor.ExpGive.Value = Npc(EditorIndex).EXP
    frmNpcEditor.txtChance.Text = STR(Npc(EditorIndex).ItemNPC(1).Chance)
    frmNpcEditor.scrlNum.Value = Npc(EditorIndex).ItemNPC(1).ItemNum
    frmNpcEditor.scrlValue.Value = Npc(EditorIndex).ItemNPC(1).ItemValue
    frmNpcEditor.scrlSpeech.Value = Npc(EditorIndex).Speech
    If Npc(EditorIndex).Speech > 0 Then
        frmNpcEditor.lblSpeechName.Caption = Speech(Npc(EditorIndex).Speech).name
    Else
        frmNpcEditor.lblSpeechName.Caption = ""
    End If
    If Npc(EditorIndex).SpawnTime = 0 Then
        frmNpcEditor.chkDay.Value = Checked
        frmNpcEditor.chkNight.Value = Checked
    ElseIf Npc(EditorIndex).SpawnTime = 1 Then
        frmNpcEditor.chkDay.Value = Checked
        frmNpcEditor.chkNight.Value = Unchecked
    ElseIf Npc(EditorIndex).SpawnTime = 2 Then
        frmNpcEditor.chkDay.Value = Unchecked
        frmNpcEditor.chkNight.Value = Checked
    End If
    
    frmNpcEditor.lblScript.Caption = Npc(EditorIndex).Script
    
    frmNpcEditor.Show vbModal
End Sub

Public Sub NpcEditorOk()
    Npc(EditorIndex).name = frmNpcEditor.txtName.Text
    Npc(EditorIndex).AttackSay = frmNpcEditor.txtAttackSay.Text
    Npc(EditorIndex).Sprite = frmNpcEditor.scrlSprite.Value
    Npc(EditorIndex).SpawnSecs = Val(frmNpcEditor.txtSpawnSecs.Text)
    Npc(EditorIndex).Behavior = frmNpcEditor.cmbBehavior.ListIndex
    Npc(EditorIndex).Range = frmNpcEditor.scrlRange.Value
    Npc(EditorIndex).STR = frmNpcEditor.scrlSTR.Value
    Npc(EditorIndex).DEF = frmNpcEditor.scrlDEF.Value
    Npc(EditorIndex).speed = frmNpcEditor.scrlSPEED.Value
    Npc(EditorIndex).MAGI = frmNpcEditor.scrlMAGI.Value
    Npc(EditorIndex).Big = frmNpcEditor.BigNpc.Value
    Npc(EditorIndex).MaxHP = frmNpcEditor.StartHP.Value
    Npc(EditorIndex).EXP = frmNpcEditor.ExpGive.Value
    Npc(EditorIndex).Speech = frmNpcEditor.scrlSpeech.Value
    Npc(EditorIndex).Script = frmNpcEditor.scrlScript.Value
    Npc(EditorIndex).Element = frmNpcEditor.scrlElement.Value
    
    If frmNpcEditor.chkDay.Value = Checked And frmNpcEditor.chkNight.Value = Checked Then
        Npc(EditorIndex).SpawnTime = 0
    ElseIf frmNpcEditor.chkDay.Value = Checked And frmNpcEditor.chkNight.Value = Unchecked Then
        Npc(EditorIndex).SpawnTime = 1
    ElseIf frmNpcEditor.chkDay.Value = Unchecked And frmNpcEditor.chkNight.Value = Checked Then
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
Dim sDc As Long

    If frmNpcEditor.BigNpc.Value = Checked Then
        sDc = DD_BigSpriteSurf.GetDC
        With frmNpcEditor
            .picSprite.Cls
            Call BitBlt(.picSprite.hDC, 0, 0, 64, 64, sDc, 3 * 64, .scrlSprite.Value * 64, SRCCOPY)
        End With
        Call DD_BigSpriteSurf.ReleaseDC(sDc)
    Else
        sDc = DD_SpriteSurf.GetDC
        With frmNpcEditor
            .picSprite.Cls
            Call BitBlt(.picSprite.hDC, 0, 0, SIZE_X, SIZE_Y, sDc, 3 * SIZE_X, .scrlSprite.Value * SIZE_Y, SRCCOPY)
        End With
        Call DD_SpriteSurf.ReleaseDC(sDc)
    End If
End Sub

Public Sub ShopEditorInit()
Dim i As Long

    frmShopEditor.txtName.Text = Trim(Shop(EditorIndex).name)
    frmShopEditor.txtJoinSay.Text = Trim(Shop(EditorIndex).JoinSay)
    frmShopEditor.txtLeaveSay.Text = Trim(Shop(EditorIndex).LeaveSay)
    frmShopEditor.chkFixesItems.Value = Shop(EditorIndex).FixesItems
    
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
            GetItem = Shop(EditorIndex).TradeItem(C).Value(i).GetItem
            GetValue = Shop(EditorIndex).TradeItem(C).Value(i).GetValue
            GiveItem = Shop(EditorIndex).TradeItem(C).Value(i).GiveItem
            GiveValue = Shop(EditorIndex).TradeItem(C).Value(i).GiveValue

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
Dim i As Long
Dim sDc As Long

    frmSpellEditor.cmbClassReq.AddItem "All Classes"
    For i = 1 To Max_Classes
        frmSpellEditor.cmbClassReq.AddItem Trim(Class(i).name)
    Next i
    
    EditorSpellY = Int(Spell(EditorIndex).pic / 6)
    EditorSpellX = (Spell(EditorIndex).pic - Int(Spell(EditorIndex).pic / 6) * 6)
    
    
    sDc = DD_Icon.GetDC
    With frmSpellEditor.picIcons
        .Width = DDSD_Icon.lWidth
        .Height = DDSD_Icon.lHeight
        .Cls
        Call BitBlt(.hDC, 0, 0, .Width, .Height, sDc, 0, 0, SRCCOPY)
    End With
    Call DD_Icon.ReleaseDC(sDc)
    
    frmSpellEditor.txtName.Text = Trim(Spell(EditorIndex).name)
    frmSpellEditor.cmbClassReq.ListIndex = Spell(EditorIndex).ClassReq
    frmSpellEditor.scrlLevelReq.Value = Spell(EditorIndex).LevelReq
        
    frmSpellEditor.cmbType.ListIndex = Spell(EditorIndex).Type
    frmSpellEditor.scrlVitalMod.Value = Spell(EditorIndex).Data1
    
    frmSpellEditor.scrlCost.Value = Spell(EditorIndex).MPCost
    frmSpellEditor.scrlSound.Value = Spell(EditorIndex).Sound
    
    If Spell(EditorIndex).Range = 0 Then Spell(EditorIndex).Range = 1
    frmSpellEditor.scrlRange.Value = Spell(EditorIndex).Range
    
    frmSpellEditor.scrlSpellAnim.Value = Spell(EditorIndex).SpellAnim
    frmSpellEditor.scrlSpellTime.Value = Spell(EditorIndex).SpellTime
    frmSpellEditor.scrlSpellDone.Value = Spell(EditorIndex).SpellDone
    
    frmSpellEditor.chkArea.Value = Spell(EditorIndex).AE
    
    frmSpellEditor.scrlElement.Value = Spell(EditorIndex).Element
    frmSpellEditor.scrlElement.Max = MAX_ELEMENTS
        
    frmSpellEditor.Show vbModal
End Sub

Public Sub SpellEditorOk()
    Spell(EditorIndex).name = frmSpellEditor.txtName.Text
    Spell(EditorIndex).ClassReq = frmSpellEditor.cmbClassReq.ListIndex
    Spell(EditorIndex).LevelReq = frmSpellEditor.scrlLevelReq.Value
    Spell(EditorIndex).Type = frmSpellEditor.cmbType.ListIndex
    Spell(EditorIndex).Data1 = frmSpellEditor.scrlVitalMod.Value
    Spell(EditorIndex).Data3 = 0
    Spell(EditorIndex).MPCost = frmSpellEditor.scrlCost.Value
    Spell(EditorIndex).Sound = frmSpellEditor.scrlSound.Value
    Spell(EditorIndex).Range = frmSpellEditor.scrlRange.Value
    Spell(EditorIndex).pic = EditorSpellY * 6 + EditorSpellX
    
    Spell(EditorIndex).SpellAnim = frmSpellEditor.scrlSpellAnim.Value
    Spell(EditorIndex).SpellTime = frmSpellEditor.scrlSpellTime.Value
    Spell(EditorIndex).SpellDone = frmSpellEditor.scrlSpellDone.Value
    
    Spell(EditorIndex).AE = frmSpellEditor.chkArea.Value
    
    Spell(EditorIndex).Element = frmSpellEditor.scrlElement.Value
    
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
        If Item(GetPlayerInvItemNum(MyIndex, i)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(MyIndex, i)).Stackable = 1 Then
            frmPlayerTrade.PlayerInv1.AddItem i & ": " & Trim(Item(GetPlayerInvItemNum(MyIndex, i)).name) & " (" & GetPlayerInvItemValue(MyIndex, i) & ")"
        Else
            If GetPlayerWeaponSlot(MyIndex) = i Or GetPlayerArmorSlot(MyIndex) = i Or GetPlayerHelmetSlot(MyIndex) = i Or GetPlayerShieldSlot(MyIndex) = i Or GetPlayerLegsSlot(MyIndex) = i Or GetPlayerBootsSlot(MyIndex) = i Or GetPlayerGlovesSlot(MyIndex) = i Or GetPlayerRing1Slot(MyIndex) = i Or GetPlayerRing2Slot(MyIndex) = i Or GetPlayerAmuletSlot(MyIndex) = i Then
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
Sub BltTile2(ByVal x As Long, ByVal y As Long, ByVal Tile As Long, ByVal TileSet As Long)
If TileFile(TileSet) = 0 Then Exit Sub

    rec.Top = Int(Tile / TilesInSheets) * PIC_Y
    rec.Bottom = rec.Top + PIC_Y
    rec.Left = (Tile - Int(Tile / TilesInSheets) * TilesInSheets) * PIC_X
    rec.Right = rec.Left + PIC_X
    Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) + sx - NewXOffset, y - (NewPlayerY * PIC_Y) + sx - NewYOffset, DD_TileSurf(TileSet), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
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
    
    TextX = GetPlayerX(Index) * PIC_X + Player(Index).XOffset + Int(PIC_X) - ((DISPLAY_BUBBLE_WIDTH * 32) / 2)
    TextY = GetPlayerY(Index) * PIC_Y + Player(Index).YOffset - Int(PIC_Y) + 85
    
    If TextX < ((DISPLAY_BUBBLE_WIDTH * 32) / 2) Then TextX = ((DISPLAY_BUBBLE_WIDTH * 32) / 2)
    If TextX > (MAX_MAPX * PIC_X) - ((DISPLAY_BUBBLE_WIDTH * 32) / 2) Then TextX = (MAX_MAPX * PIC_X) - ((DISPLAY_BUBBLE_WIDTH * 32) / 2)
    
    Call DD_BackBuffer.ReleaseDC(TexthDC)
    
    ' Draw the fancy box with tiles.
    Call BltTile2(TextX - 10, TextY - 10, 1, 6)
    Call BltTile2(TextX + (DISPLAY_BUBBLE_WIDTH * PIC_X) - PIC_X - 10, TextY - 10, 2, 6)
    
    For intLoop = 1 To (DISPLAY_BUBBLE_WIDTH - 2)
        Call BltTile2(TextX - 10 + (intLoop * PIC_X), TextY - 10, 16, 6)
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
            Call BltTile2(TextX - 10, TextY - 10 + intLoop, 19, 6)
            Call BltTile2(TextX - 10 + (DISPLAY_BUBBLE_WIDTH * PIC_X) - PIC_X, TextY - 10 + intLoop, 17, 6)
            
            For intLoop2 = 1 To DISPLAY_BUBBLE_WIDTH - 2
                Call BltTile2(TextX - 10 + (intLoop2 * PIC_X), TextY + intLoop - 10, 5, 6)
            Next intLoop2
        Next intLoop
    End If

    Call BltTile2(TextX - 10, TextY + (bytLineCount * 16) - 4, 3, 6)
    Call BltTile2(TextX + (DISPLAY_BUBBLE_WIDTH * PIC_X) - PIC_X - 10, TextY + (bytLineCount * 16) - 4, 4, 6)
    
    For intLoop = 1 To (DISPLAY_BUBBLE_WIDTH - 2)
        Call BltTile2(TextX - 10 + (intLoop * PIC_X), TextY + (bytLineCount * 16) - 4, 15, 6)
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
Dim x As Long, y As Long

    x = (GetPlayerX(Index) * PIC_X + sx + Player(Index).XOffset) - (NewPlayerX * PIC_X) - NewXOffset
    y = (GetPlayerY(Index) * PIC_Y + sx + Player(Index).YOffset) - (NewPlayerY * PIC_Y) - NewYOffset
    
    If Player(Index).HP = 0 Then Exit Sub
    'draws the back bars
    Call DD_BackBuffer.SetFillColor(RGB(0, 0, 255))
    Call DD_BackBuffer.DrawBox(x, y + 32, x + 32, y + 36)
    
    'draws HP
    If Int((GetPlayerHP(Index) / GetPlayerMaxHP(Index)) * 100) > 50 Then
        Call DD_BackBuffer.SetFillColor(RGB(0, 255, 0))
    End If
    If Int((GetPlayerHP(Index) / GetPlayerMaxHP(Index)) * 100) > 20 And Int((GetPlayerHP(Index) / GetPlayerMaxHP(Index)) * 100) <= 50 Then
        Call DD_BackBuffer.SetFillColor(RGB(255, 255, 0))
    End If
    If Int((GetPlayerHP(Index) / GetPlayerMaxHP(Index)) * 100) <= 20 Then
        Call DD_BackBuffer.SetFillColor(RGB(255, 0, 0))
    End If
    Call DD_BackBuffer.DrawBox(x, y + PIC_Y, x + ((Player(Index).HP / 100) / (Player(Index).MaxHP / 100) * SIZE_X), y + 36)
End Sub
Sub BltPetBars(ByVal Index As Long)
Dim x As Long, y As Long

    x = (Player(Index).Pet.x * PIC_X + sx + Player(Index).Pet.XOffset) - (NewPlayerX * PIC_X) - NewXOffset
    y = (Player(Index).Pet.y * PIC_Y + sx + Player(Index).Pet.YOffset) - (NewPlayerY * PIC_Y) - NewYOffset
    
    If Player(Index).HP = 0 Then Exit Sub
    'draws the back bars
    Call DD_BackBuffer.SetFillColor(RGB(0, 0, 255))
    Call DD_BackBuffer.DrawBox(x, y + 32, x + 32, y + 36)
    
    'draws HP
    If Int((Player(Index).Pet.HP / Player(Index).Pet.MaxHP) * 100) > 50 Then
        Call DD_BackBuffer.SetFillColor(RGB(0, 255, 0))
    End If
    If Int((Player(Index).Pet.HP / Player(Index).Pet.MaxHP) * 100) > 20 And Int((Player(Index).Pet.HP / Player(Index).Pet.MaxHP) * 100) <= 50 Then
        Call DD_BackBuffer.SetFillColor(RGB(255, 255, 0))
    End If
    If Int((Player(Index).Pet.HP / Player(Index).Pet.MaxHP) * 100) <= 20 Then
        Call DD_BackBuffer.SetFillColor(RGB(255, 0, 0))
    End If
    Call DD_BackBuffer.DrawBox(x, y + PIC_Y, x + ((Player(Index).Pet.HP / 100) / (Player(Index).Pet.MaxHP / 100) * SIZE_X), y + 36)
End Sub
Sub BltNpcBars(ByVal Index As Long)
Dim x As Long, y As Long

If MapNpc(Index).HP = 0 Then Exit Sub
If MapNpc(Index).Num < 1 Then Exit Sub

    If Npc(MapNpc(Index).Num).Big = 1 Then
        x = (MapNpc(Index).x * PIC_X + sx - 9 + MapNpc(Index).XOffset) - (NewPlayerX * PIC_X) - NewXOffset
        y = (MapNpc(Index).y * PIC_Y + sx + MapNpc(Index).YOffset) - (NewPlayerY * PIC_Y) - NewYOffset
        
        Call DD_BackBuffer.SetFillColor(RGB(0, 0, 255))
        Call DD_BackBuffer.DrawBox(x, y + 32, x + 50, y + 36)
        If Int(MapNpc(Index).HP / MapNpc(Index).MaxHP * 100) > 50 Then
            Call DD_BackBuffer.SetFillColor(RGB(0, 255, 0))
        End If
        If Int(MapNpc(Index).HP / MapNpc(Index).MaxHP * 100) <= 50 And Int((MapNpc(Index).HP / MapNpc(Index).MaxHP) * 100) > 20 Then
            Call DD_BackBuffer.SetFillColor(RGB(255, 255, 0))
        End If
        If Int(MapNpc(Index).HP / MapNpc(Index).MaxHP * 100) <= 20 Then
            Call DD_BackBuffer.SetFillColor(RGB(255, 0, 0))
        End If
        Call DD_BackBuffer.DrawBox(x, y + 32, x + ((MapNpc(Index).HP / 100) / (MapNpc(Index).MaxHP / 100) * 50), y + 36)
    Else
        x = (MapNpc(Index).x * PIC_X + sx + MapNpc(Index).XOffset) - (NewPlayerX * PIC_X) - NewXOffset
        y = (MapNpc(Index).y * PIC_Y + sx + MapNpc(Index).YOffset) - (NewPlayerY * PIC_Y) - NewYOffset
        
        Call DD_BackBuffer.SetFillColor(RGB(0, 0, 255))
        Call DD_BackBuffer.DrawBox(x, y + 32, x + 32, y + 36)
        If Int(MapNpc(Index).HP / MapNpc(Index).MaxHP * 100) > 50 Then
            Call DD_BackBuffer.SetFillColor(RGB(0, 255, 0))
        End If
        If Int(MapNpc(Index).HP / MapNpc(Index).MaxHP * 100) <= 50 And Int((MapNpc(Index).HP / MapNpc(Index).MaxHP) * 100) > 20 Then
            Call DD_BackBuffer.SetFillColor(RGB(255, 255, 0))
        End If
        If Int(MapNpc(Index).HP / MapNpc(Index).MaxHP * 100) <= 20 Then
            Call DD_BackBuffer.SetFillColor(RGB(255, 0, 0))
        End If
        Call DD_BackBuffer.DrawBox(x, y + 32, x + ((MapNpc(Index).HP / 100) / (MapNpc(Index).MaxHP / 100) * 32), y + 36)
    End If
End Sub

Public Sub UpdateVisInv()
Dim Index As Long
Dim d As Long
Dim sDc As Long

    For Index = 1 To MAX_INV
        If GetPlayerShieldSlot(MyIndex) <> Index Then frmMirage.ShieldImage.Picture = LoadPicture()
        If GetPlayerWeaponSlot(MyIndex) <> Index Then frmMirage.WeaponImage.Picture = LoadPicture()
        If GetPlayerHelmetSlot(MyIndex) <> Index Then frmMirage.HelmetImage.Picture = LoadPicture()
        If GetPlayerArmorSlot(MyIndex) <> Index Then frmMirage.ArmorImage.Picture = LoadPicture()
        If GetPlayerLegsSlot(MyIndex) <> Index Then frmMirage.LegsImage.Picture = LoadPicture()
        If GetPlayerBootsSlot(MyIndex) <> Index Then frmMirage.BootsImage.Picture = LoadPicture()
        If GetPlayerGlovesSlot(MyIndex) <> Index Then frmMirage.GlovesImage.Picture = LoadPicture()
        If GetPlayerRing1Slot(MyIndex) <> Index Then frmMirage.Ring1Image.Picture = LoadPicture()
        If GetPlayerRing2Slot(MyIndex) <> Index Then frmMirage.Ring2Image.Picture = LoadPicture()
        If GetPlayerAmuletSlot(MyIndex) <> Index Then frmMirage.AmuletImage.Picture = LoadPicture()
    Next Index
    
    frmMirage.HandsImage.Picture = LoadPicture()
    If Player(MyIndex).Hands <> 0 Then Call BitBlt(frmMirage.HandsImage.hDC, 0, 0, PIC_X, PIC_Y, frmMirage.picItems.hDC, (Item(Player(MyIndex).Hands).pic - Int(Item(Player(MyIndex).Hands).pic / 6) * 6) * PIC_X, Int(Item(Player(MyIndex).Hands).pic / 6) * PIC_Y, SRCCOPY)
    
    
    For Index = 1 To MAX_INV
        If GetPlayerShieldSlot(MyIndex) = Index Then Call BitBlt(frmMirage.ShieldImage.hDC, 0, 0, PIC_X, PIC_Y, frmMirage.picItems.hDC, (Item(GetPlayerInvItemNum(MyIndex, Index)).pic - Int(Item(GetPlayerInvItemNum(MyIndex, Index)).pic / 6) * 6) * PIC_X, Int(Item(GetPlayerInvItemNum(MyIndex, Index)).pic / 6) * PIC_Y, SRCCOPY)
        If GetPlayerWeaponSlot(MyIndex) = Index Then Call BitBlt(frmMirage.WeaponImage.hDC, 0, 0, PIC_X, PIC_Y, frmMirage.picItems.hDC, (Item(GetPlayerInvItemNum(MyIndex, Index)).pic - Int(Item(GetPlayerInvItemNum(MyIndex, Index)).pic / 6) * 6) * PIC_X, Int(Item(GetPlayerInvItemNum(MyIndex, Index)).pic / 6) * PIC_Y, SRCCOPY)
        If GetPlayerHelmetSlot(MyIndex) = Index Then Call BitBlt(frmMirage.HelmetImage.hDC, 0, 0, PIC_X, PIC_Y, frmMirage.picItems.hDC, (Item(GetPlayerInvItemNum(MyIndex, Index)).pic - Int(Item(GetPlayerInvItemNum(MyIndex, Index)).pic / 6) * 6) * PIC_X, Int(Item(GetPlayerInvItemNum(MyIndex, Index)).pic / 6) * PIC_Y, SRCCOPY)
        If GetPlayerArmorSlot(MyIndex) = Index Then Call BitBlt(frmMirage.ArmorImage.hDC, 0, 0, PIC_X, PIC_Y, frmMirage.picItems.hDC, (Item(GetPlayerInvItemNum(MyIndex, Index)).pic - Int(Item(GetPlayerInvItemNum(MyIndex, Index)).pic / 6) * 6) * PIC_X, Int(Item(GetPlayerInvItemNum(MyIndex, Index)).pic / 6) * PIC_Y, SRCCOPY)
        If GetPlayerLegsSlot(MyIndex) = Index Then Call BitBlt(frmMirage.LegsImage.hDC, 0, 0, PIC_X, PIC_Y, frmMirage.picItems.hDC, (Item(GetPlayerInvItemNum(MyIndex, Index)).pic - Int(Item(GetPlayerInvItemNum(MyIndex, Index)).pic / 6) * 6) * PIC_X, Int(Item(GetPlayerInvItemNum(MyIndex, Index)).pic / 6) * PIC_Y, SRCCOPY)
        If GetPlayerBootsSlot(MyIndex) = Index Then Call BitBlt(frmMirage.BootsImage.hDC, 0, 0, PIC_X, PIC_Y, frmMirage.picItems.hDC, (Item(GetPlayerInvItemNum(MyIndex, Index)).pic - Int(Item(GetPlayerInvItemNum(MyIndex, Index)).pic / 6) * 6) * PIC_X, Int(Item(GetPlayerInvItemNum(MyIndex, Index)).pic / 6) * PIC_Y, SRCCOPY)
        If GetPlayerGlovesSlot(MyIndex) = Index Then Call BitBlt(frmMirage.GlovesImage.hDC, 0, 0, PIC_X, PIC_Y, frmMirage.picItems.hDC, (Item(GetPlayerInvItemNum(MyIndex, Index)).pic - Int(Item(GetPlayerInvItemNum(MyIndex, Index)).pic / 6) * 6) * PIC_X, Int(Item(GetPlayerInvItemNum(MyIndex, Index)).pic / 6) * PIC_Y, SRCCOPY)
        If GetPlayerRing1Slot(MyIndex) = Index Then Call BitBlt(frmMirage.Ring1Image.hDC, 0, 0, PIC_X, PIC_Y, frmMirage.picItems.hDC, (Item(GetPlayerInvItemNum(MyIndex, Index)).pic - Int(Item(GetPlayerInvItemNum(MyIndex, Index)).pic / 6) * 6) * PIC_X, Int(Item(GetPlayerInvItemNum(MyIndex, Index)).pic / 6) * PIC_Y, SRCCOPY)
        If GetPlayerRing2Slot(MyIndex) = Index Then Call BitBlt(frmMirage.Ring2Image.hDC, 0, 0, PIC_X, PIC_Y, frmMirage.picItems.hDC, (Item(GetPlayerInvItemNum(MyIndex, Index)).pic - Int(Item(GetPlayerInvItemNum(MyIndex, Index)).pic / 6) * 6) * PIC_X, Int(Item(GetPlayerInvItemNum(MyIndex, Index)).pic / 6) * PIC_Y, SRCCOPY)
        If GetPlayerAmuletSlot(MyIndex) = Index Then Call BitBlt(frmMirage.AmuletImage.hDC, 0, 0, PIC_X, PIC_Y, frmMirage.picItems.hDC, (Item(GetPlayerInvItemNum(MyIndex, Index)).pic - Int(Item(GetPlayerInvItemNum(MyIndex, Index)).pic / 6) * 6) * PIC_X, Int(Item(GetPlayerInvItemNum(MyIndex, Index)).pic / 6) * PIC_Y, SRCCOPY)
    Next Index
    
        
    frmMirage.EquipS(0).Visible = False
    frmMirage.EquipS(1).Visible = False
    frmMirage.EquipS(2).Visible = False
    frmMirage.EquipS(3).Visible = False
    frmMirage.EquipS(4).Visible = False
    frmMirage.EquipS(5).Visible = False
    frmMirage.EquipS(6).Visible = False
    frmMirage.EquipS(7).Visible = False
    frmMirage.EquipS(8).Visible = False
    frmMirage.EquipS(9).Visible = False

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
                 ElseIf GetPlayerLegsSlot(MyIndex) = d + 1 Then
                    'frmMirage.picInv(d).ToolTipText = Trim(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Name) & " (worn)"
                    frmMirage.EquipS(4).Visible = True
                    frmMirage.EquipS(4).Top = frmMirage.picInv(d).Top - 2
                    frmMirage.EquipS(4).Left = frmMirage.picInv(d).Left - 2
                    'frmMirage.picInv(d).ToolTipText = Trim(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Name)
                ElseIf GetPlayerBootsSlot(MyIndex) = d + 1 Then
                    'frmMirage.picInv(d).ToolTipText = Trim(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Name) & " (worn)"
                    frmMirage.EquipS(5).Visible = True
                    frmMirage.EquipS(5).Top = frmMirage.picInv(d).Top - 2
                    frmMirage.EquipS(5).Left = frmMirage.picInv(d).Left - 2
                ElseIf GetPlayerGlovesSlot(MyIndex) = d + 1 Then
                    'frmMirage.picInv(d).ToolTipText = Trim(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Name) & " (worn)"
                    frmMirage.EquipS(6).Visible = True
                    frmMirage.EquipS(6).Top = frmMirage.picInv(d).Top - 2
                    frmMirage.EquipS(6).Left = frmMirage.picInv(d).Left - 2
                ElseIf GetPlayerRing1Slot(MyIndex) = d + 1 Then
                    'frmMirage.picInv(d).ToolTipText = Trim(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Name) & " (worn)"
                    frmMirage.EquipS(7).Visible = True
                    frmMirage.EquipS(7).Top = frmMirage.picInv(d).Top - 2
                    frmMirage.EquipS(7).Left = frmMirage.picInv(d).Left - 2
                ElseIf GetPlayerRing2Slot(MyIndex) = d + 1 Then
                    'frmMirage.picInv(d).ToolTipText = Trim(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Name) & " (worn)"
                    frmMirage.EquipS(8).Visible = True
                    frmMirage.EquipS(8).Top = frmMirage.picInv(d).Top - 2
                    frmMirage.EquipS(8).Left = frmMirage.picInv(d).Left - 2
                ElseIf GetPlayerAmuletSlot(MyIndex) = d + 1 Then
                    'frmMirage.picInv(d).ToolTipText = Trim(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Name) & " (worn)"
                    frmMirage.EquipS(9).Visible = True
                    frmMirage.EquipS(9).Top = frmMirage.picInv(d).Top - 2
                    frmMirage.EquipS(9).Left = frmMirage.picInv(d).Left - 2
                Else
                    'frmMirage.picInv(d).ToolTipText = Trim(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Name)
                End If
            End If
            If Item(GetPlayerInvItemNum(MyIndex, d + 1)).Type = ITEM_TYPE_CURRENCY Then
                If GetPlayerInvItemNum(MyIndex, d + 1) > 0 Then
                    If GetPlayerInvItemNum(MyIndex, d + 1) = 1 Then
                        frmMirage.lblGold.Caption = "Gold: " & GetPlayerInvItemValue(MyIndex, d + 1)
                        frmTrade.lblGold.Caption = "Gold: " & GetPlayerInvItemValue(MyIndex, d + 1)
                        frmSellItem.lblGold.Caption = "Gold: " & GetPlayerInvItemValue(MyIndex, d + 1)
                        frmPlayerTrade.lblGold.Caption = "Gold: " & GetPlayerInvItemValue(MyIndex, d + 1)
                    End If
                Else
                    frmMirage.lblGold.Caption = "Gold: " & 0
                    frmTrade.lblGold.Caption = "Gold: " & 0
                    frmSellItem.lblGold.Caption = "Gold: " & 0
                    frmPlayerTrade.lblGold.Caption = "Gold: " & 0
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

Sub ItemSelected(ByVal Index As Long, ByVal Selected As Long)
Dim index2 As Long
index2 = Trade(Selected).Items(Index).ItemGetNum

    frmTrade.shpSelect.Top = frmTrade.picItem(Index - 1).Top - 1
    frmTrade.shpSelect.Left = frmTrade.picItem(Index - 1).Left - 1

    If index2 <= 0 Then
        Call clearItemSelected
        Exit Sub
    End If

    frmTrade.descName.Caption = Trim(Item(index2).name)
    frmTrade.descQuantity.Caption = "Quantity: " & Trade(Selected).Items(Index).ItemGetVal
    
    frmTrade.descStr.Caption = "Strength: " & Item(index2).StrReq
    frmTrade.descDef.Caption = "Defense: " & Item(index2).DefReq
    frmTrade.descSpeed.Caption = "Speed: " & Item(index2).SpeedReq
    frmTrade.descMagi.Caption = "Magic: " & Item(index2).MagicReq
    
    frmTrade.descAStr.Caption = "Strength: " & Item(index2).AddStr
    frmTrade.descADef.Caption = "Defense: " & Item(index2).AddDef
    frmTrade.descAMagi.Caption = "Magic: " & Item(index2).AddMagi
    frmTrade.descASpeed.Caption = "Speed: " & Item(index2).AddSpeed
    
    frmTrade.descHp.Caption = "HP: " & Item(index2).AddHP
    frmTrade.descMp.Caption = "MP: " & Item(index2).AddMP
    frmTrade.descSp.Caption = "SP: " & Item(index2).AddSP
    frmTrade.descAExp.Caption = "EXP: " & Item(index2).AddEXP & "%"
    
    frmTrade.desc.Caption = Trim(Item(index2).desc)
    
    frmTrade.lblTradeFor.Caption = "Trade for: " & Trim(Item(Trade(Selected).Items(Index).ItemGiveNum).name)
    frmTrade.lblQuantity.Caption = "Quantity: " & Trade(Selected).Items(Index).ItemGiveVal
End Sub

Sub clearItemSelected()
    frmTrade.lblTradeFor.Caption = ""
    frmTrade.lblQuantity.Caption = ""
    
    frmTrade.descName.Caption = ""
    frmTrade.descQuantity.Caption = ""
    
    frmTrade.descStr.Caption = "Strength: 0"
    frmTrade.descDef.Caption = "Defense: 0"
    frmTrade.descMagi.Caption = "Magic: 0"
    frmTrade.descSpeed.Caption = "Speed: 0"
    
    frmTrade.descAStr.Caption = "Strength: 0"
    frmTrade.descADef.Caption = "Defense: 0"
    frmTrade.descAMagi.Caption = "Magic: 0"
    frmTrade.descASpeed.Caption = "Speed: 0"
    
    frmTrade.descHp.Caption = "HP: 0"
    frmTrade.descMp.Caption = "MP: 0"
    frmTrade.descSp.Caption = "SP: 0"

    frmTrade.descAExp.Caption = "EXP: 0%"
    frmTrade.desc.Caption = ""
End Sub

Public Sub MakeFormTrans(ByVal Form As Form, ByVal Amount As Long)
Dim NormalWindowStyle As Long

    DoEvents
    
    If GetVersion >= 4 Then ' Make sure it's Windows 2000 or better.
        If Amount > 100 Then Amount = 100 ' Make sure they dont have more then 100 percent
        NormalWindowStyle = GetWindowLong(Form.hwnd, GWL_EXSTYLE)
        SetWindowLong Form.hwnd, GWL_EXSTYLE, NormalWindowStyle Or WS_EX_LAYERED
        ' Sets the transparency level
        SetLayeredWindowAttributes Form.hwnd, 0, 255 * (1 - (Val(Amount) / 100)), LWA_ALPHA
    End If
End Sub

Public Function MakeLoc(ByVal x As Long, ByVal y As Long) As Long
    MakeLoc = (y * MAX_MAPX) + x
End Function

Public Function MakeX(ByVal Loc As Long) As Long
    MakeX = Loc - (MakeY(Loc) * MAX_MAPX)
End Function

Public Function MakeY(ByVal Loc As Long) As Long
    MakeY = Int(Loc / MAX_MAPX)
End Function

Public Function IsValid(ByVal x As Long, _
   ByVal y As Long) As Boolean
    IsValid = True

    If x < 0 Or x > MAX_MAPX Or y < 0 Or y > MAX_MAPY Then IsValid = False
End Function

Sub UpdateBank()
Dim i As Long

frmBank.lstInventory.Clear
frmBank.lstBank.Clear

For i = 1 To MAX_INV
    If GetPlayerInvItemNum(MyIndex, i) > 0 Then
        If Item(GetPlayerInvItemNum(MyIndex, i)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(MyIndex, i)).Stackable = 1 Then
            frmBank.lstInventory.AddItem i & "> " & Trim(Item(GetPlayerInvItemNum(MyIndex, i)).name) & " (" & GetPlayerInvItemValue(MyIndex, i) & ")"
        Else
            If GetPlayerWeaponSlot(MyIndex) = i Or GetPlayerArmorSlot(MyIndex) = i Or GetPlayerHelmetSlot(MyIndex) = i Or GetPlayerShieldSlot(MyIndex) = i Or GetPlayerLegsSlot(MyIndex) = i Or GetPlayerBootsSlot(MyIndex) = i Or GetPlayerGlovesSlot(MyIndex) = i Or GetPlayerRing1Slot(MyIndex) = i Or GetPlayerRing2Slot(MyIndex) = i Or GetPlayerAmuletSlot(MyIndex) = i Then
                frmBank.lstInventory.AddItem i & "> " & Trim(Item(GetPlayerInvItemNum(MyIndex, i)).name) & " (worn)"
            Else
                frmBank.lstInventory.AddItem i & "> " & Trim(Item(GetPlayerInvItemNum(MyIndex, i)).name)
            End If
        End If
    Else
        frmBank.lstInventory.AddItem i & "> Empty"
    End If
    'DoEvents
Next i

For i = 1 To MAX_BANK
    If GetPlayerBankItemNum(MyIndex, i) > 0 Then
        If Item(GetPlayerBankItemNum(MyIndex, i)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerBankItemNum(MyIndex, i)).Stackable = 1 Then
            frmBank.lstBank.AddItem i & "> " & Trim(Item(GetPlayerBankItemNum(MyIndex, i)).name) & " (" & GetPlayerBankItemValue(MyIndex, i) & ")"
        Else
            If GetPlayerWeaponSlot(MyIndex) = i Or GetPlayerArmorSlot(MyIndex) = i Or GetPlayerHelmetSlot(MyIndex) = i Or GetPlayerShieldSlot(MyIndex) = i Or GetPlayerLegsSlot(MyIndex) = i Or GetPlayerBootsSlot(MyIndex) = i Or GetPlayerGlovesSlot(MyIndex) = i Or GetPlayerRing1Slot(MyIndex) = i Or GetPlayerRing2Slot(MyIndex) = i Or GetPlayerAmuletSlot(MyIndex) = i Then
                frmBank.lstBank.AddItem i & "> " & Trim(Item(GetPlayerBankItemNum(MyIndex, i)).name) & " (worn)"
            Else
                frmBank.lstBank.AddItem i & "> " & Trim(Item(GetPlayerBankItemNum(MyIndex, i)).name)
            End If
        End If
    Else
        frmBank.lstBank.AddItem i & "> Empty"
    End If
    'DoEvents
Next i
frmBank.lstBank.ListIndex = 0
frmBank.lstInventory.ListIndex = 0
End Sub

Public Function roundDown(dblValue As Double) As Long
Dim myDec As Long
 
    myDec = InStr(1, CStr(dblValue), ".", vbTextCompare)
    If myDec > 0 Then
        roundDown = CDbl(Left(CStr(dblValue), myDec))
    Else
        roundDown = dblValue
    End If
End Function
 
Public Function roundUp(dblValue As Double) As Long
Dim myDec As Long
 
    myDec = InStr(1, CStr(dblValue), ".", vbTextCompare)
    If myDec > 0 Then
        roundUp = CDbl(Left(CStr(dblValue), myDec)) + 1
    Else
        roundUp = dblValue
    End If
End Function

Public Function windowed() As Boolean

    If Val(GetVar(App.Path & "\config.ini", "CONFIG", "Windowed")) = 1 Then
        windowed = True
    Else
        windowed = False
    End If
   
End Function

Sub BltFurniture(ByVal x As Long, ByVal y As Long)
    If Map(GetPlayerMap(MyIndex)).Tile(x, y).Type = TILE_TYPE_FURNITURE Then
    ' Only used if ever want to switch to blt rather then bltfast
    With rec_pos
        .Top = y * PIC_Y
        .Bottom = .Top + PIC_Y
        .Left = x * PIC_X
        .Right = .Left + PIC_X
    End With
    
    rec.Top = Int(Item(Map(GetPlayerMap(MyIndex)).Tile(x, y).Data1).pic / 6) * PIC_Y
    rec.Bottom = rec.Top + PIC_Y
    rec.Left = (Item(Map(GetPlayerMap(MyIndex)).Tile(x, y).Data1).pic - Int(Item(Map(GetPlayerMap(MyIndex)).Tile(x, y).Data1).pic / 6) * 6) * PIC_X
    rec.Right = rec.Left + PIC_X
    
    'Call DD_BackBuffer.Blt(rec_pos, DD_ItemSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
    Call DD_BackBuffer.BltFast((x - NewPlayerX) * PIC_X + sx - NewXOffset, (y - NewPlayerY) * PIC_Y + sx - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End If
End Sub

Function GetPlayerFacingX(ByVal Index As Long)
GetPlayerFacingX = GetPlayerX(Index)
If GetPlayerDir(Index) = DIR_LEFT Then
GetPlayerFacingX = GetPlayerFacingX - 1
End If
If GetPlayerDir(Index) = DIR_RIGHT Then
GetPlayerFacingX = GetPlayerFacingX + 1
End If
End Function
Function GetPlayerFacingY(ByVal Index As Long)
GetPlayerFacingY = GetPlayerY(Index)
If GetPlayerDir(Index) = DIR_UP Then
GetPlayerFacingY = GetPlayerFacingY - 1
End If
If GetPlayerDir(Index) = DIR_DOWN Then
GetPlayerFacingY = GetPlayerFacingY + 1
End If
End Function
Sub SetAttribute(ByVal mapper As Long, ByVal x As Long, ByVal y As Long, ByVal Attrib As Long, ByVal Data1 As Long, ByVal Data2 As Long, ByVal Data3 As Long, ByVal String1 As String, ByVal String2 As String, ByVal String3 As String)
Dim Packet As String
With Map(mapper).Tile(x, y)
    .Type = Attrib
    .Data1 = Data1
    .Data2 = Data2
    .Data3 = Data3
    .String1 = String1
    .String2 = String2
    .String3 = String3
End With

Packet = "setattribute" & SEP_CHAR & mapper & SEP_CHAR & CStr(x) & SEP_CHAR & CStr(y) & SEP_CHAR
        With Map(mapper).Tile(x, y)
            Packet = Packet & .Type & SEP_CHAR & .Data1 & SEP_CHAR & .Data2 & SEP_CHAR & .Data3 & SEP_CHAR & .String1 & SEP_CHAR & .String2 & SEP_CHAR & .String3 & SEP_CHAR
        End With
Call SendData(Packet & END_CHAR)
End Sub

Sub UpdatePlayerSellVisInv()
Dim i As Long
Dim Qx As Long
Dim Qx2 As Long

For i = 1 To MAX_INV
    If GetPlayerInvItemNum(MyIndex, i) > 0 Then
     Qx = Player(MyIndex).Inv(i).Num
        If Item(GetPlayerInvItemNum(MyIndex, i)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(MyIndex, i)).Stackable = 1 Then
            Call BitBlt(frmSellItem.picInv(i - 1).hDC, 0, 0, PIC_X, PIC_Y, frmMirage.picItems.hDC, (Item(Qx).pic - Int(Item(Qx).pic / 6) * 6) * PIC_X, Int(Item(Qx).pic / 6) * PIC_Y, SRCCOPY)
        Else
           If GetPlayerWeaponSlot(MyIndex) = i Or GetPlayerArmorSlot(MyIndex) = i Or GetPlayerHelmetSlot(MyIndex) = i Or GetPlayerShieldSlot(MyIndex) = i Or GetPlayerLegsSlot(MyIndex) = i Or GetPlayerBootsSlot(MyIndex) = i Or GetPlayerGlovesSlot(MyIndex) = i Or GetPlayerRing1Slot(MyIndex) = i Or GetPlayerRing2Slot(MyIndex) = i Or GetPlayerAmuletSlot(MyIndex) = i Then
                Call BitBlt(frmSellItem.picInv(i - 1).hDC, 0, 0, PIC_X, PIC_Y, frmMirage.picItems.hDC, (Item(Qx).pic - Int(Item(Qx).pic / 6) * 6) * PIC_X, Int(Item(Qx).pic / 6) * PIC_Y, SRCCOPY)
            Else
                Call BitBlt(frmSellItem.picInv(i - 1).hDC, 0, 0, PIC_X, PIC_Y, frmMirage.picItems.hDC, (Item(Qx).pic - Int(Item(Qx).pic / 6) * 6) * PIC_X, Int(Item(Qx).pic / 6) * PIC_Y, SRCCOPY)
            End If
        End If
    Else
        frmSellItem.picInv(i - 1).BackColor = Black
        'frmSellItem.lstSellItem.AddItem i & "> Empty"
    End If
    DoEvents
Next i


Call UpdateSelectedSellInvItem(frmSellItem.lstSellItem.ListIndex)

End Sub

Sub UpdateSelectedSellInvItem(Index As Integer)
On Error GoTo Num0

Num0:
End Sub

Sub UpdatePlayerVisInv()
Dim i As Long
Dim Qq As Long
Dim Qq2 As Long

For i = 1 To MAX_INV
    If GetPlayerInvItemNum(MyIndex, i) > 0 Then
     Qq = Player(MyIndex).Inv(i).Num
        If Item(GetPlayerInvItemNum(MyIndex, i)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(MyIndex, i)).Stackable = 1 Then
            Call BitBlt(frmBank.picInv(i).hDC, 0, 0, PIC_X, PIC_Y, frmMirage.picItems.hDC, (Item(Qq).pic - Int(Item(Qq).pic / 6) * 6) * PIC_X, Int(Item(Qq).pic / 6) * PIC_Y, SRCCOPY)
        Else
           If GetPlayerWeaponSlot(MyIndex) = i Or GetPlayerArmorSlot(MyIndex) = i Or GetPlayerHelmetSlot(MyIndex) = i Or GetPlayerShieldSlot(MyIndex) = i Or GetPlayerLegsSlot(MyIndex) = i Or GetPlayerBootsSlot(MyIndex) = i Or GetPlayerGlovesSlot(MyIndex) = i Or GetPlayerRing1Slot(MyIndex) = i Or GetPlayerRing2Slot(MyIndex) = i Or GetPlayerAmuletSlot(MyIndex) = i Then
                Call BitBlt(frmBank.picInv(i).hDC, 0, 0, PIC_X, PIC_Y, frmMirage.picItems.hDC, (Item(Qq).pic - Int(Item(Qq).pic / 6) * 6) * PIC_X, Int(Item(Qq).pic / 6) * PIC_Y, SRCCOPY)
            Else
                Call BitBlt(frmBank.picInv(i).hDC, 0, 0, PIC_X, PIC_Y, frmMirage.picItems.hDC, (Item(Qq).pic - Int(Item(Qq).pic / 6) * 6) * PIC_X, Int(Item(Qq).pic / 6) * PIC_Y, SRCCOPY)
            End If
        End If
    Else
        frmBank.picInv(i).BackColor = Black
        frmBank.lstInventory.AddItem i & "> Empty"
    End If
    DoEvents
Next i

For i = 1 To MAX_BANK
    If GetPlayerBankItemNum(MyIndex, i) > 0 Then
    Qq2 = Player(MyIndex).Bank(i).Num
        If Item(GetPlayerBankItemNum(MyIndex, i)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerBankItemNum(MyIndex, i)).Stackable = 1 Then
            Call BitBlt(frmBank.picBank(i).hDC, 0, 0, PIC_X, PIC_Y, frmMirage.picItems.hDC, (Item(Qq2).pic - Int(Item(Qq2).pic / 6) * 6) * PIC_X, Int(Item(Qq2).pic / 6) * PIC_Y, SRCCOPY)
        Else
            If GetPlayerWeaponSlot(MyIndex) = i Or GetPlayerArmorSlot(MyIndex) = i Or GetPlayerHelmetSlot(MyIndex) = i Or GetPlayerShieldSlot(MyIndex) = i Or GetPlayerLegsSlot(MyIndex) = i Or GetPlayerBootsSlot(MyIndex) = i Or GetPlayerGlovesSlot(MyIndex) = i Or GetPlayerRing1Slot(MyIndex) = i Or GetPlayerRing2Slot(MyIndex) = i Or GetPlayerAmuletSlot(MyIndex) = i Then
                Call BitBlt(frmBank.picBank(i).hDC, 0, 0, PIC_X, PIC_Y, frmMirage.picItems.hDC, (Item(Qq2).pic - Int(Item(Qq2).pic / 6) * 6) * PIC_X, Int(Item(Qq2).pic / 6) * PIC_Y, SRCCOPY)
            Else
                Call BitBlt(frmBank.picBank(i).hDC, 0, 0, PIC_X, PIC_Y, frmMirage.picItems.hDC, (Item(Qq2).pic - Int(Item(Qq2).pic / 6) * 6) * PIC_X, Int(Item(Qq2).pic / 6) * PIC_Y, SRCCOPY)
            End If
        End If
    Else
        frmBank.picBank(i).BackColor = Black
        frmBank.lstBank.AddItem i & "> Empty"
    End If
    DoEvents
Next i

Call UpdateSelectedInvItem(frmBank.lstInventory.ListIndex)
Call UpdateSelectedBankItem(frmBank.lstBank.ListIndex)

End Sub

Sub UpdateSelectedInvItem(Index As Integer)
On Error GoTo Num0

    If GetPlayerInvItemNum(MyIndex, Index) > 0 Then
        frmBank.descName.Caption = "Selected Item: " & Item(GetPlayerInvItemNum(MyIndex, Index)).name
    Else
        frmBank.descName.Caption = "Selected Item: Empty"
    End If
    Exit Sub
   
Num0:
End Sub

Sub UpdateSelectedBankItem(Index As Integer)
On Error GoTo Num0
   
    If GetPlayerBankItemNum(MyIndex, Index) > 0 Then
        frmBank.descName2.Caption = "Selected Item: " & Item(GetPlayerBankItemNum(MyIndex, Index)).name
    Else
        frmBank.descName2.Caption = "Selected Item: Empty"
    End If
    Exit Sub
   
Num0:
End Sub

Sub UpdatePlayerGuildVisInv()
Dim i As Long
Dim Qx As Long
Dim Qx2 As Long

For i = 1 To MAX_INV
    If GetPlayerInvItemNum(MyIndex, i) > 0 Then
     Qx = Player(MyIndex).Inv(i).Num
        If Item(GetPlayerInvItemNum(MyIndex, i)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(MyIndex, i)).Stackable = 1 Then
            Call BitBlt(frmGuildDeed.picInv(i - 1).hDC, 0, 0, PIC_X, PIC_Y, frmMirage.picItems.hDC, (Item(Qx).pic - Int(Item(Qx).pic / 6) * 6) * PIC_X, Int(Item(Qx).pic / 6) * PIC_Y, SRCCOPY)
        Else
           If GetPlayerWeaponSlot(MyIndex) = i Or GetPlayerArmorSlot(MyIndex) = i Or GetPlayerHelmetSlot(MyIndex) = i Or GetPlayerShieldSlot(MyIndex) = i Or GetPlayerLegsSlot(MyIndex) = i Or GetPlayerBootsSlot(MyIndex) = i Or GetPlayerGlovesSlot(MyIndex) = i Or GetPlayerRing1Slot(MyIndex) = i Or GetPlayerRing2Slot(MyIndex) = i Or GetPlayerAmuletSlot(MyIndex) = i Then
                Call BitBlt(frmGuildDeed.picInv(i - 1).hDC, 0, 0, PIC_X, PIC_Y, frmMirage.picItems.hDC, (Item(Qx).pic - Int(Item(Qx).pic / 6) * 6) * PIC_X, Int(Item(Qx).pic / 6) * PIC_Y, SRCCOPY)
            Else
                Call BitBlt(frmGuildDeed.picInv(i - 1).hDC, 0, 0, PIC_X, PIC_Y, frmMirage.picItems.hDC, (Item(Qx).pic - Int(Item(Qx).pic / 6) * 6) * PIC_X, Int(Item(Qx).pic / 6) * PIC_Y, SRCCOPY)
            End If
        End If
    Else
        frmGuildDeed.picInv(i - 1).BackColor = Black
    End If
    DoEvents
Next i


Call UpdateSelectedGuildInvItem(frmGuildDeed.lstInv.ListIndex)

End Sub

Sub UpdateSelectedGuildInvItem(Index As Integer)
On Error GoTo Num0

    'If GetPlayerInvItemNum(MyIndex, Index) > 1 Then
        'frmGuildDeed.descName.Caption = ""
    'Else
    '    frmGuildDeed.descName.Caption = ""
    'End If
    'Exit Sub
   
Num0:
End Sub

Public Function MouseCheck() As Boolean

    If Val(GetVar(App.Path & "\config.ini", "CONFIG", "MouseMovement")) = 1 Then
        MouseCheck = True
    Else
        MouseCheck = False
    End If
    
End Function

Public Sub MainMenuInit()
    Dim Stuff As String
    Dim Stuff2 As String
    Dim Stuff3 As String
    Dim ThisIsANumber As Long

    Stuff2 = ""
        Stuff = ReadINI("DATA", "Desc", App.Path & "\News.ini")
        For ThisIsANumber = 1 To Len(Stuff)
           If mid$(Stuff, ThisIsANumber, 1) = "*" Then
                Stuff2 = Stuff2 & vbCrLf
           Else
                Stuff2 = Stuff2 & mid$(Stuff, ThisIsANumber, 1)
           End If
        Next
        Stuff3 = ""
        Stuff = ReadINI("DATA", "News", App.Path & "\News.ini")
        For ThisIsANumber = 1 To Len(Stuff)
           If mid$(Stuff, ThisIsANumber, 1) = "*" Then
                Stuff3 = Stuff3 & vbCrLf
           Else
                Stuff3 = Stuff3 & mid$(Stuff, ThisIsANumber, 1)
           End If
        Next
        frmMainMenu.News.Text = Stuff3 & Stuff2
        End Sub
