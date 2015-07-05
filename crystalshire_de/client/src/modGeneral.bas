Attribute VB_Name = "modGeneral"
Option Explicit

' halts thread of execution
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' get system uptime in milliseconds
Public Declare Function GetTickCount Lib "kernel32" () As Long

'For Clear functions
Public Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal length As Long)

Public Sub Main()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' load gui
    Call SetStatus("Loading interface...")
    InitialiseGUI
    
    ' load options
    Call SetStatus("Loading Options...")
    LoadOptions
    
    ' Check if the directory is there, if its not make it
    ChkDir App.path & "\data files\", "graphics"
    ChkDir App.path & "\data files\graphics\", "animations"
    ChkDir App.path & "\data files\graphics\", "characters"
    ChkDir App.path & "\data files\graphics\", "items"
    ChkDir App.path & "\data files\graphics\", "paperdolls"
    ChkDir App.path & "\data files\graphics\", "resources"
    ChkDir App.path & "\data files\graphics\", "spellicons"
    ChkDir App.path & "\data files\graphics\", "tilesets"
    ChkDir App.path & "\data files\graphics\", "faces"
    ChkDir App.path & "\data files\graphics\", "gui"
    ChkDir App.path & "\data files\", "logs"
    ChkDir App.path & "\data files\", "maps"
    ChkDir App.path & "\data files\", "music"
    ChkDir App.path & "\data files\", "sound"
    
    ' load dx8
    Call SetStatus("Initializing DirectX...")
    EngineInit
    
    ' initialise sound & music engines
    Init_Music

    ' load the main game (and by extension, pre-load DD7)
    GettingMap = True
    vbQuote = ChrW$(34) ' "
    
    ' Update the form with the game's name before it's loaded
    frmMain.Caption = Options.Game_Name
    
    ' randomize rnd's seed
    Randomize
    Call SetStatus("Initializing TCP settings...")
    Call TcpInit
    Call InitMessages

    ' check if we have main-menu music
    If Len(Trim$(Options.MenuMusic)) > 0 Then Play_Music Trim$(Options.MenuMusic)
    
    ' Reset values
    Ping = -1
    
    ' cache the buttons then reset & render them
    Call SetStatus("Loading buttons...")
    
    ' set values for directional blocking arrows
    DirArrowX(1) = 12 ' up
    DirArrowY(1) = 0
    DirArrowX(2) = 12 ' down
    DirArrowY(2) = 23
    DirArrowX(3) = 0 ' left
    DirArrowY(3) = 12
    DirArrowX(4) = 23 ' right
    DirArrowY(4) = 12
    
    ' set the paperdoll order
    ReDim PaperdollOrder(1 To Equipment.Equipment_Count - 1) As Long
    PaperdollOrder(1) = Equipment.Armor
    PaperdollOrder(2) = Equipment.Helmet
    PaperdollOrder(3) = Equipment.Shield
    PaperdollOrder(4) = Equipment.Weapon
    
    ' set the main form size
    frmMain.Width = 12090
    frmMain.height = 9420
    
    ' show the main menu
    frmMain.Show
    ShowMenu
    HideGame
    MenuLoop
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Main", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub InitialiseGUI()
Dim i As Long

    ' re-set chat scroll
    ChatScroll = 8

    ReDim GUIWindow(1 To GUI_Count) As GUIWindowRec
    
    ' 1 - Chat
    With GUIWindow(GUI_CHAT)
        .x = 10
        .y = 445
        .Width = 412
        .height = 145
        .visible = True
    End With
    
    ' 2 - Hotbar
    With GUIWindow(GUI_HOTBAR)
        .x = 300
        .y = 10
        .height = 36
        .Width = ((9 + 36) * (MAX_HOTBAR - 1))
    End With
    
    ' 3 - Menu
    With GUIWindow(GUI_MENU)
        .x = 558
        .y = 514
        .Width = 232
        .height = 76
        .visible = True
    End With
    
    ' 4 - Bars
    With GUIWindow(GUI_BARS)
        .x = 10
        .y = 10
        .Width = 254
        .height = 75
        .visible = True
    End With
    
    ' 5 - Inventory
    With GUIWindow(GUI_INVENTORY)
        .x = 578
        .y = 255
        .Width = 195
        .height = 250
        .visible = False
    End With
    
    ' 6 - Spells
    With GUIWindow(GUI_SPELLS)
        .x = 578
        .y = 255
        .Width = 195
        .height = 250
        .visible = False
    End With
    
    ' 7 - Character
    With GUIWindow(GUI_CHARACTER)
        .x = 578
        .y = 255
        .Width = 195
        .height = 250
        .visible = False
    End With
    
    ' 8 - Options
    With GUIWindow(GUI_OPTIONS)
        .x = 578
        .y = 255
        .Width = 195
        .height = 250
        .visible = False
    End With
    
    ' 9 - Party
    With GUIWindow(GUI_PARTY)
        .x = 578
        .y = 255
        .Width = 195
        .height = 250
        .visible = False
    End With
    
    ' 10 - Description
    With GUIWindow(GUI_DESCRIPTION)
        .x = 0
        .y = 0
        .Width = 190
        .height = 126
        .visible = False
    End With
    
    ' 11 - Main Menu
    With GUIWindow(GUI_MAINMENU)
        .x = 152
        .y = 204
        .Width = 495
        .height = 332
        .visible = False
    End With
    
    ' 12 - Shop
    With GUIWindow(GUI_SHOP)
        .x = 118
        .y = 110
        .Width = 252
        .height = 317
        .visible = False
    End With
    
    ' BUTTONS
    ' main - inv
    With Buttons(1)
        .state = 0 ' normal
        .x = 6
        .y = 6
        .Width = 69
        .height = 29
        .visible = True
        .PicNum = 1
    End With
    
    ' main - skills
    With Buttons(2)
        .state = 0 ' normal
        .x = 81
        .y = 6
        .Width = 69
        .height = 29
        .visible = True
        .PicNum = 2
    End With
    
    ' main - char
    With Buttons(3)
        .state = 0 ' normal
        .x = 156
        .y = 6
        .Width = 69
        .height = 29
        .visible = True
        .PicNum = 3
    End With
    
    ' main - opt
    With Buttons(4)
        .state = 0 ' normal
        .x = 6
        .y = 41
        .Width = 69
        .height = 29
        .visible = True
        .PicNum = 4
    End With
    
    ' main - trade
    With Buttons(5)
        .state = 0 ' normal
        .x = 81
        .y = 41
        .Width = 69
        .height = 29
        .visible = True
        .PicNum = 5
    End With
    
    ' main - party
    With Buttons(6)
        .state = 0 ' normal
        .x = 156
        .y = 41
        .Width = 69
        .height = 29
        .visible = True
        .PicNum = 6
    End With
    
    ' menu - login
    With Buttons(7)
        .state = 0 ' normal
        .x = 54
        .y = 277
        .Width = 89
        .height = 29
        .visible = True
        .PicNum = 7
    End With
    
    ' menu - register
    With Buttons(8)
        .state = 0 ' normal
        .x = 154
        .y = 277
        .Width = 89
        .height = 29
        .visible = True
        .PicNum = 8
    End With
    
    ' menu - credits
    With Buttons(9)
        .state = 0 ' normal
        .x = 254
        .y = 277
        .Width = 89
        .height = 29
        .visible = True
        .PicNum = 9
    End With
    
    ' menu - exit
    With Buttons(10)
        .state = 0 ' normal
        .x = 354
        .y = 277
        .Width = 89
        .height = 29
        .visible = True
        .PicNum = 10
    End With
    
    ' menu - Login Accept
    With Buttons(11)
        .state = 0 ' normal
        .x = 206
        .y = 164
        .Width = 89
        .height = 29
        .visible = True
        .PicNum = 11
    End With
    
    ' menu - Register Accept
    With Buttons(12)
        .state = 0 ' normal
        .x = 206
        .y = 169
        .Width = 89
        .height = 29
        .visible = True
        .PicNum = 11
    End With
    
    ' menu - Class Accept
    With Buttons(13)
        .state = 0 ' normal
        .x = 248
        .y = 206
        .Width = 89
        .height = 29
        .visible = True
        .PicNum = 11
    End With
    
    ' menu - Class Next
    With Buttons(14)
        .state = 0 ' normal
        .x = 348
        .y = 206
        .Width = 89
        .height = 29
        .visible = True
        .PicNum = 12
    End With
    
    ' menu - NewChar Accept
    With Buttons(15)
        .state = 0 ' normal
        .x = 205
        .y = 169
        .Width = 89
        .height = 29
        .visible = True
        .PicNum = 11
    End With
    
    ' main - AddStats
    For i = 16 To 20
        With Buttons(i)
            .state = 0 'normal
            .Width = 12
            .height = 11
            .visible = True
            .PicNum = 13
        End With
    Next
    ' set the individual spaces
    For i = 16 To 18 ' first 3
        With Buttons(i)
            .x = 80
            .y = 147 + ((i - 16) * 15)
        End With
    Next
    For i = 19 To 20
        With Buttons(i)
            .x = 165
            .y = 147 + ((i - 19) * 15)
        End With
    Next
    
    ' main - shop buy
    With Buttons(21)
        .state = 0 ' normal
        .x = 12
        .y = 276
        .Width = 69
        .height = 29
        .visible = True
        .PicNum = 14
    End With
    
    ' main - shop sell
    With Buttons(22)
        .state = 0 ' normal
        .x = 90
        .y = 276
        .Width = 69
        .height = 29
        .visible = True
        .PicNum = 15
    End With
    
    ' main - shop exit
    With Buttons(23)
        .state = 0 ' normal
        .x = 90
        .y = 276
        .Width = 69
        .height = 29
        .visible = True
        .PicNum = 16
    End With
    
    ' main - party invite
    With Buttons(24)
        .state = 0 ' normal
        .x = 14
        .y = 209
        .Width = 69
        .height = 29
        .visible = True
        .PicNum = 17
    End With
    
    ' main - party invite
    With Buttons(25)
        .state = 0 ' normal
        .x = 101
        .y = 209
        .Width = 69
        .height = 29
        .visible = True
        .PicNum = 18
    End With
    
    ' main - music on
    With Buttons(26)
        .state = 0 ' normal
        .x = 77
        .y = 14
        .Width = 49
        .height = 19
        .visible = True
        .PicNum = 19
    End With
    
    ' main - music off
    With Buttons(27)
        .state = 0 ' normal
        .x = 132
        .y = 14
        .Width = 49
        .height = 19
        .visible = True
        .PicNum = 20
    End With
    
    ' main - sound on
    With Buttons(28)
        .state = 0 ' normal
        .x = 77
        .y = 39
        .Width = 49
        .height = 19
        .visible = True
        .PicNum = 19
    End With
    
    ' main - sound off
    With Buttons(29)
        .state = 0 ' normal
        .x = 132
        .y = 39
        .Width = 49
        .height = 19
        .visible = True
        .PicNum = 20
    End With
    
    ' main - debug on
    With Buttons(30)
        .state = 0 ' normal
        .x = 77
        .y = 64
        .Width = 49
        .height = 19
        .visible = True
        .PicNum = 19
    End With
    
    ' main - debug off
    With Buttons(31)
        .state = 0 ' normal
        .x = 132
        .y = 64
        .Width = 49
        .height = 19
        .visible = True
        .PicNum = 20
    End With
    
    ' main - autotile on
    With Buttons(32)
        .state = 0 ' normal
        .x = 77
        .y = 89
        .Width = 49
        .height = 19
        .visible = True
        .PicNum = 19
    End With
    
    ' main - autotile off
    With Buttons(33)
        .state = 0 ' normal
        .x = 132
        .y = 89
        .Width = 49
        .height = 19
        .visible = True
        .PicNum = 20
    End With
    
    ' main - scroll up
    With Buttons(34)
        .state = 0 ' normal
        .x = 391
        .y = 2
        .Width = 19
        .height = 19
        .visible = True
        .PicNum = 21
    End With
    
    ' main - scroll down
    With Buttons(35)
        .state = 0 ' normal
        .x = 391
        .y = 105
        .Width = 19
        .height = 19
        .visible = True
        .PicNum = 22
    End With
End Sub

Public Sub MenuState(ByVal state As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Select Case state
        Case MENU_STATE_ADDCHAR

            If ConnectToServer(1) Then
                Call SetStatus("Connected, sending character addition data...")
                Call SendAddChar(sChar, SEX_MALE, newCharClass, newCharSprite)
            End If
            
        Case MENU_STATE_NEWACCOUNT

            If ConnectToServer(1) Then
                Call SetStatus("Connected, sending new account information...")
                Call SendNewAccount(sUser, sPass)
            End If

        Case MENU_STATE_LOGIN
            If ConnectToServer(1) Then
                Call SetStatus("Connected, sending login information...")
                Call SendLogin(sUser, sPass)
                Exit Sub
            End If
    End Select

    If Not IsConnected Then
        Call MsgBox("Sorry, the server seems to be down.  Please try to reconnect in a few minutes or visit " & GAME_WEBSITE, vbOKOnly, Options.Game_Name)
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MenuState", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub logoutGame()
Dim i As Long

    isLogging = True
    InGame = False
        
    Call DestroyTCP
    
    ' destroy the animations loaded
    For i = 1 To MAX_BYTE
        ClearAnimInstance (i)
    Next
    
    ' destroy temp values
    DragInvSlotNum = 0
    LastItemDesc = 0
    MyIndex = 0
    InventoryItemSelected = 0
    SpellBuffer = 0
    SpellBufferTimer = 0
    tmpCurrencyItem = 0
    
    ' unload editors
    Unload frmEditor_Animation
    Unload frmEditor_Item
    Unload frmEditor_Map
    Unload frmEditor_MapProperties
    Unload frmEditor_NPC
    Unload frmEditor_Resource
    Unload frmEditor_Shop
    Unload frmEditor_Spell
    
    ' destroy the chat
    For i = 1 To ChatTextBufferSize
        ChatTextBuffer(i).Text = vbNullString
    Next
    
    GUIWindow(GUI_MAINMENU).visible = True
    inMenu = True
    curMenu = MENU_MAIN
    HideGame
    MenuLoop
End Sub

Sub GameInit()
Dim MusicFile As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' hide gui
    InBank = False
    InShop = False
    InTrade = False
    
    ' get ping
    GetPing
    
    ' set values for amdin panel
    frmMain.scrlAItem.max = MAX_ITEMS
    frmMain.scrlAItem.value = 1
    
    ' play music
    MusicFile = Trim$(Map.Music)
    If Not MusicFile = "None." Then
        Play_Music MusicFile
    Else
        Stop_Music
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "GameInit", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DestroyGame()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' break out of GameLoop
    HideGame
    HideMenu
    Call DestroyTCP
    
    ' destroy music & sound engines
    Destroy_Music
    
    ' unload dx8
    EngineUnloadDirectX
    
    Call UnloadAllForms
    End
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "destroyGame", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub UnloadAllForms()
Dim frm As Form

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For Each frm In VB.Forms
        Unload frm
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "UnloadAllForms", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SetStatus(ByVal Caption As String)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    DoEvents
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetStatus", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' Used for adding text to packet debugger
Public Sub TextAdd(ByVal Txt As TextBox, Msg As String, NewLine As Boolean)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If NewLine Then
        Txt.Text = Txt.Text + Msg + vbCrLf
    Else
        Txt.Text = Txt.Text + Msg
    End If

    Txt.SelStart = Len(Txt.Text) - 1
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "TextAdd", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function Rand(ByVal Low As Long, ByVal High As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Rand = Int((High - Low + 1) * Rnd) + Low
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "Rand", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function isLoginLegal(ByVal Username As String, ByVal Password As String) As Boolean
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If LenB(Trim$(Username)) >= 3 Then
        If LenB(Trim$(Password)) >= 3 Then
            isLoginLegal = True
        End If
    End If

    ' Error handler
    Exit Function
errorhandler:
    HandleError "isLoginLegal", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function isStringLegal(ByVal sInput As String) As Boolean
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Prevent high ascii chars
    For i = 1 To Len(sInput)

        If Asc(Mid$(sInput, i, 1)) < vbKeySpace Or Asc(Mid$(sInput, i, 1)) > vbKeyF15 Then
            Call MsgBox("You cannot use high ASCII characters in your name, please re-enter.", vbOKOnly, Options.Game_Name)
            Exit Function
        End If

    Next

    isStringLegal = True
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "isStringLegal", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

' buttons
Public Sub resetButtons(Optional ByVal exceptionNum As Long = 0)
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' loop through entire array
    For i = 1 To MAX_BUTTONS
        Select Case i
            ' option buttons
            Case 26, 27, 28, 29, 30, 31, 32, 33
                ' Nothing in here
            ' Everything else - reset
            Case Else
                ' only change if different and not exception
                If Buttons(i).state = 1 And Not i = exceptionNum Then
                    ' reset state and render
                    Buttons(i).state = 0 'normal
                End If
        End Select
    Next
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "resetButtons_Main", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub resetClickedButtons()
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' loop through entire array
    For i = 1 To MAX_BUTTONS
        Select Case i
            ' option buttons
            Case 26, 27, 28, 29, 30, 31, 32, 33
                ' Nothing in here
            ' Everything else - reset
            Case Else
                ' reset state and render
                Buttons(i).state = 0 'normal
        End Select
    Next
    
    ' do the npc conversation as well
    For i = 1 To 4
        chatOptState(i) = 0 ' normal
    Next
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "resetButtons_Main", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub PopulateLists()
Dim strLoad As String, i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Cache music list
    strLoad = dir(App.path & MUSIC_PATH & "*.*")
    i = 1
    Do While strLoad > vbNullString
        ReDim Preserve musicCache(1 To i) As String
        musicCache(i) = strLoad
        strLoad = dir
        i = i + 1
    Loop
    
    ' Cache sound list
    strLoad = dir(App.path & SOUND_PATH & "*.*")
    i = 1
    Do While strLoad > vbNullString
        ReDim Preserve soundCache(1 To i) As String
        soundCache(i) = strLoad
        strLoad = dir
        i = i + 1
    Loop
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "PopulateLists", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ShowMenu()
    ' set the menu
    curMenu = MENU_MAIN
    
    ' show the GUI
    GUIWindow(GUI_MAINMENU).visible = True
    
    inMenu = True
    
    ' fader
    faderAlpha = 255
    faderState = 0
    faderSpeed = 4
    canFade = True
End Sub

Public Sub HideMenu()
    GUIWindow(GUI_MAINMENU).visible = False
    inMenu = False
End Sub

Public Sub ShowGame()
Dim i As Long

    For i = 5 To 10
        GUIWindow(i).visible = False
    Next

    For i = 1 To 4
        GUIWindow(i).visible = True
    Next
    
    InGame = True
End Sub

Public Sub HideGame()
Dim i As Long
    
    For i = 1 To 10
        GUIWindow(i).visible = False
    Next
    
    InGame = False
End Sub
