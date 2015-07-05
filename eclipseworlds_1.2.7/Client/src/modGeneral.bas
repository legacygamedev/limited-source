Attribute VB_Name = "modGeneral"
Option Explicit

' Stops the timer from processing actions in frmMenu
Public StopTimer As Boolean

Public Declare Function GetForegroundWindow Lib "user32.dll" () As Long
Public Declare Function GetTopWindow Lib "user32.dll" (ByVal hWnd As Long) As Long

' Halts thread of execution
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' Used for the 64-bit timer
Private GetSystemTimeOffset As Currency
Private Declare Sub GetSystemTime Lib "kernel32.dll" Alias "GetSystemTimeAsFileTime" (ByRef lpSystemTimeAsFileTime As Currency)
Public Declare Function timeBeginPeriod Lib "winmm.dll" (ByVal uPeriod As Long) As Long

' For Clear functions
Public Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)

Public DX7 As New DirectX7  ' Master Object, early binding
    
Public gameDestroyed As Boolean
Public readyToExit As Boolean

Public Sub Main()
    ' Make sure the application isn't already running
    If App.PrevInstance Then
        AlertMsg "This application is already running!"
        End
    End If
    
    ChDrive App.Path
    ChDir App.Path
    
    ' Load options
    Call SetStatus("Loading Options...")
    LoadOptions
    
     ' Setup screen
    ResizeScreen
    
    ' Set the high-resolution timer
    timeBeginPeriod 1
    
    ' This must be called before any timeGetTime calls because it states what the values of timeGetTime will be
    InitTimeGetTime
    
    Call SetStatus("Loadung Buttons...")
    Call CacheButtons
    
    ' Set the loading screen
    frmLoad.Picture = LoadPicture(App.Path & "\data files\graphics\gui\menu\loading.jpg")
    frmLoad.Visible = True
    
    ' Load the images used for the menu and main forms
    Call SetStatus("Loading Menu...")
    Load frmMenu
    
    ' Load gui
    LoadGUI

    Call SetStatus("Loading User Interface...")
    
    ' Check if the directory is there, if it's not make it
    ChkDir App.Path & "\data files\", "graphics"
    ChkDir App.Path & "\data files\graphics\", "animations"
    ChkDir App.Path & "\data files\graphics\", "characters"
    ChkDir App.Path & "\data files\graphics\", "items"
    ChkDir App.Path & "\data files\graphics\", "paperdolls"
    ChkDir App.Path & "\data files\graphics\", "resources"
    ChkDir App.Path & "\data files\graphics\", "spellicons"
    ChkDir App.Path & "\data files\graphics\", "tilesets"
    ChkDir App.Path & "\data files\graphics\", "faces"
    ChkDir App.Path & "\data files\graphics\", "fogs"
    ChkDir App.Path & "\data files\graphics\", "panoramas"
    ChkDir App.Path & "\data files\graphics\", "emoticons"
    ChkDir App.Path & "\data files\graphics\gui\", "menu"
    ChkDir App.Path & "\data files\graphics\gui\", "main"
    ChkDir App.Path & "\data files\graphics\gui\menu\", "buttons"
    ChkDir App.Path & "\data files\graphics\gui\main\", "buttons"
    ChkDir App.Path & "\data files\graphics\gui\main\", "bars"
    ChkDir App.Path & "\data files\graphics\gui\main\", "chat"
    ChkDir App.Path & "\data files\", "music"
    ChkDir App.Path & "\data files\", "sound"
    ChkDir App.Path & "\", "logs"
    
    ' Load the main game (and by extension, pre-load DD7)
    GettingMap = True
    vbQuote = ChrW$(34)
    
    ' Update the form with the game's Name before it's loaded
    frmMain.Caption = GAME_NAME
    
    ' Initialize DirectX
    Call SetStatus("Initializing DirectX...")
    EngineInitFontSettings
    InitDX8
    
    Call SetStatus("Initializing TCP settings...")
    Call TcpInit
    Call InitMessages
    
    ' Reset values
    Ping = -1
    MouseX = -1
    MouseY = -1
    
    ' Cache the buttons then reset & render them
    ResetMenuButtons
    
    ' Allow to escape out of frmLoad for future encounters
    GameLoaded = True
    
    ' Set values for directional blocking arrows
    DirArrowX(1) = 12 ' up
    DirArrowY(1) = 0
    DirArrowX(2) = 12 ' down
    DirArrowY(2) = 23
    DirArrowX(3) = 0 ' left
    DirArrowY(3) = 12
    DirArrowX(4) = 23 ' right
    DirArrowY(4) = 12
    
    ' Set the paperdoll order
    ReDim PaperdollOrder(1 To Equipment.Equipment_Count - 1) As Long
    PaperdollOrder(1) = Equipment.Body
    PaperdollOrder(2) = Equipment.Feet
    PaperdollOrder(3) = Equipment.Hands
    PaperdollOrder(4) = Equipment.Neck
    PaperdollOrder(5) = Equipment.Head
    PaperdollOrder(6) = Equipment.Shield
    PaperdollOrder(7) = Equipment.Weapon
    
    ' Hide the load form
    frmLoad.Visible = False
    
    ' Set the form visible
    frmMenu.Show
    If GetTopWindow(frmMenu.hWnd) = 0 Then BringWindowToTop frmMenu.hWnd
    
    ' Hide all pictures
    Call ClearMenuPictures
    
    frmMenu.picMain.Visible = True
    
    ' Check if we have main-menu music
    If Len(Trim$(Options.MenuMusic)) > 0 Then Call Audio.PlayMusic(Trim$(Options.MenuMusic))
    
    MenuLoop
    
    UnloadAllForms
    End
End Sub

Public Sub MenuLoop()
    ' If debug mode, handle error then exit out
    On Error GoTo ErrorHandler
    
restartmenuloop:
    ' *** Start GameLoop ***
    Do While (Not InGame And Not readyToExit)
         ' *********************
         ' ** Render Graphics **
         ' *********************
         Call DrawGDI
        
         ' Mute everything but still keep everything playing
         If frmMenu.WindowState = vbMinimized Then
            If Not Audio.IsMuted Then Audio.MuteVolume
         Else
            If Audio.IsMuted Then Audio.UpdateVolume
         End If
         DoEvents
   Loop
    Exit Sub
    
' Error handler
ErrorHandler:
    If Direct3D_Device.TestCooperativeLevel = D3DERR_DEVICELOST Or Direct3D_Device.TestCooperativeLevel = D3DERR_DEVICENOTRESET Then
        HandleDeviceLost
        GoTo restartmenuloop
    ElseIf App.LogMode = 1 And Options.Debug = 1 Then
        HandleError "MenuLoop", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
        Err.Clear
    End If
End Sub

Public Sub LoadGUI(Optional ByVal LoadingScreen As Boolean = False)
    Dim I As Long

    ' If we can't find the interface
    On Error GoTo ErrorHandler
    
    ' loading screen
    If LoadingScreen Then
        frmLoad.Picture = LoadPicture(App.Path & "\data files\graphics\gui\menu\loading.jpg")
        Exit Sub
    End If
    
    For I = 1 To MAX_MENUBUTTONS
        Call RenderButton_Menu(I)
    Next

    ' Menu
    frmMenu.Picture = LoadPicture(App.Path & "\data files\graphics\gui\menu\background.jpg")
    frmMenu.picMain.Picture = LoadPicture(App.Path & "\data files\graphics\gui\menu\main.jpg")
    frmMenu.picLogin.Picture = LoadPicture(App.Path & "\data files\graphics\gui\menu\login.jpg")
    frmMenu.picRegister.Picture = LoadPicture(App.Path & "\data files\graphics\gui\menu\register.jpg")
    frmMenu.picCredits.Picture = LoadPicture(App.Path & "\data files\graphics\gui\menu\credits.jpg")
    frmMenu.picCharacter.Picture = LoadPicture(App.Path & "\data files\graphics\gui\menu\character.jpg")
    
    ' Main
    frmMain.picCharacter.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\character.jpg")
    frmMain.picFriends.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\base.jpg")
    frmMain.picFoes.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\base.jpg")
    frmMain.picGuild.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\base.jpg")
    frmMain.picGuild_No.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\base.jpg")
    frmMain.picOptions.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\base.jpg")
    frmMain.picParty.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\party.jpg")
    frmMain.picItemDesc.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\description_item.jpg")
    frmMain.picSpellDesc.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\description_spell.jpg")
    frmMain.picTempInv.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\dragbox.jpg")
    frmMain.picTempBank.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\dragbox.jpg")
    frmMain.picTempSpell.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\dragbox.jpg")
    frmMain.picShop.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\shop.jpg")
    frmMain.picBank.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\Bank.jpg")
    frmMain.picTrade.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\trade.jpg")
    frmMain.picHotbar.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\hotbar.jpg")
    frmMain.picDialogue.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\dialogue.bmp")
    frmMain.picEventChat.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\eventchat.bmp")
    frmMain.picChatbox.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\chatbox.bmp")
    frmMain.picCurrency.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\currency.jpg")
    frmMain.picTitles.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\base.jpg")
    frmMain.ImgFix.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\fix.bmp")
    frmMain.picSkills.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\skills.jpg")
    
    ' Vital Bars on Main
    frmMain.picGUI_Vitals_Base.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\vitals_base.bmp")
    frmMain.imgHPBar.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\bars\health.jpg")
    frmMain.imgMPBar.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\bars\spirit.jpg")
    frmMain.imgEXPBar.Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\bars\experience.jpg")
    
    ' Gui Buttons
    For I = 1 To MAX_MAINBUTTONS
        frmMain.picButton(I).Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\buttons\" & MainButton(I).FileName & "_norm.jpg")
    Next
    
    ' Equipment Slots
    EquipSlotTop(Equipment.Weapon) = 50
    EquipSlotLeft(Equipment.Weapon) = 44
    EquipSlotTop(Equipment.Body) = 55
    EquipSlotLeft(Equipment.Body) = 82
    EquipSlotTop(Equipment.Head) = 20
    EquipSlotLeft(Equipment.Head) = 80
    EquipSlotTop(Equipment.Shield) = 55
    EquipSlotLeft(Equipment.Shield) = 115
    EquipSlotTop(Equipment.Feet) = 105
    EquipSlotLeft(Equipment.Feet) = 80
    EquipSlotTop(Equipment.Hands) = 80
    EquipSlotLeft(Equipment.Hands) = 50
    EquipSlotTop(Equipment.Ring) = 80
    EquipSlotLeft(Equipment.Ring) = 113
    EquipSlotTop(Equipment.Neck) = 20
    EquipSlotLeft(Equipment.Neck) = 50
    
    ' Store the bar widths for calculations
    HPBar_Width = frmMain.imgHPBar.Width
    MPBar_Width = frmMain.imgMPBar.Width
    EXPBar_Width = frmMain.imgEXPBar.Width
        
    ' Main - Party Bars
    For I = 1 To MAX_PARTY_MEMBERS
        frmMain.imgPartyHealth(I).Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\bars\party_health.jpg")
        frmMain.imgPartySpirit(I).Picture = LoadPicture(App.Path & "\data files\graphics\gui\main\bars\party_spirit.jpg")
    Next
    
    ' Party
    Party_HPWidth = frmMain.imgPartyHealth(1).Width
    Party_MPWidth = frmMain.imgPartySpirit(1).Width
    Exit Sub
    
' Let them know we can't load the GUI
ErrorHandler:
    AlertMsg "Cannot find one or more images used in the user interface."
    DestroyGame
End Sub

Public Sub MenuState(ByVal State As Long)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
            
    If Not IsConnected Then
        frmMenu.lblLAccept.Enabled = False
        ConnectToServer (1)
    End If
    
    Select Case State
        Case MENU_STATE_ADDCHAR
            If IsConnected Then
                If frmMenu.optMale.Value Then
                    Call SendAddChar(frmMenu.txtCUser, Gender_MALE, ClassSelection(frmMenu.cmbClass.ListIndex + 1))
                Else
                    Call SendAddChar(frmMenu.txtCUser, Gender_FEMALE, ClassSelection(frmMenu.cmbClass.ListIndex + 1))
                End If
                Exit Sub
            End If
            
        Case MENU_STATE_NEWACCOUNT
            If IsConnected Then
                Call SetStatus("Connected, sending new account information...")
                Call SendNewAccount(frmMenu.txtRUser.text, frmMenu.txtRPass.text)
                Exit Sub
            End If
            
        Case MENU_STATE_LOGIN
            If IsConnected Then
                Call SetStatus("Connected, sending login information...")
                Call SendLogin(frmMenu.txtLUser.text, frmMenu.txtLPass.text)
                Exit Sub
            End If
    End Select
    
    frmMenu.lblLAccept.Enabled = True

    If Not IsConnected Then
        Call NotConnected
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "MenuState", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub GameInit()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    EnteringGame = True
    
    ' Hide Gui
    frmLoad.Visible = False
    frmMenu.Visible = False
    
    EnteringGame = False
      
    ' Bring all the main gui components to the front
    frmMain.picShop.ZOrder (0)
    frmMain.picBank.ZOrder (0)
    frmMain.picTrade.ZOrder (0)
    
    InBank = False
    InShop = 0
    InChat = False
    InTrade = 0
    ChatLocked = True
    
    ' GUI
    Call ToggleGUI(True)
    Call frmMain.TogglePanel(14)
    
    ' Get ping
    CheckPing
    SetPing
    
    ' Show the main form
    frmMain.Visible = True
    If GetTopWindow(frmMain.hWnd) = 0 Then BringWindowToTop frmMain.hWnd
    
    'AdvMapEditor
    
    ' Stop the song from playing
    Call Audio.StopMusic
    
    ' Reset the chat channels
    CurrentChatChannel = 0
    
    ' Update chat
    frmMain.picChatbox.ZOrder (0)
    frmMain.picEventChat.ZOrder (0)
    Exit Sub

' Error handler
ErrorHandler:
    HandleError "GameInit", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub DestroyGame()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    gameDestroyed = True
    
    ' Turn off the timer
    StopTimer = True
    
    ' Break out of GameLoop
    If InGame Then
        LogoutGame
    End If
    
    ' Destroy DirectX
    DestroyDX8
    
    Sleep 1000
    
    ' Destroy audio engine
    BASS_Free
    
    readyToExit = True
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "DestroyGame", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub UnloadAllForms()
    Dim frm As Form

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    For Each frm In VB.Forms
        Unload frm
    Next
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "UnloadAllForms", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub SetStatus(ByVal Caption As String)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    frmLoad.lblStatus.Caption = Caption
    DoEvents
    Exit Sub
    
    
' Error handler
ErrorHandler:
    HandleError "SetStatus", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

' Used for adding text to packet debugger
Public Sub TextAdd(ByVal txt As textBox, Msg As String, NewLine As Boolean)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If NewLine Then
        txt.text = txt.text + Msg + vbCrLf
    Else
        txt.text = txt.text + Msg
    End If

    txt.SelStart = Len(txt.text) - 1
    Exit Sub
        
' Error handler
ErrorHandler:
    HandleError "TextAdd", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub SetFocusOnScreen()
    On Error Resume Next ' Prevents run time errors, no way to handle it other than this

    frmMain.picScreen.SetFocus
End Sub

Public Sub SetGameFocus()
    On Error Resume Next ' Prevents run time errors, no way to handle it other than this

    ' Ignore focus if in editor
    If Editor > 0 Then Exit Sub
    If frmMain.txtCurrency.Visible Then Exit Sub
    
    If ChatLocked Or Not GUIVisible Then
        SetFocusOnScreen
    Else
        SetFocusOnChat
    End If
End Sub

Public Sub SetFocusOnChat()
    On Error Resume Next ' Prevents run time errors, no way to handle it other than this

    frmMain.txtMyChat.SetFocus
End Sub

Public Function Random(ByVal Low As Long, ByVal High As Long) As Long
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ' Randomize rnd's seed
    Randomize
    
    Random = Int((High - Low + 1) * Rnd) + Low
    Exit Function
    
' Error handler
ErrorHandler:
    HandleError "Random", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Function

Public Sub MovePicture(PB As PictureBox, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim GlobalX As Integer
    Dim GlobalY As Integer

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    GlobalX = PB.Left
    GlobalY = PB.Top

    If Button = 1 Then
        PB.Left = GlobalX + X - SOffsetX
        PB.Top = GlobalY + Y - SYOffset
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "MovePicture", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Function IsLoginLegal(ByVal UserName As String, ByVal Password As String) As Boolean
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If Len(Trim$(UserName)) >= 3 Then
        If Len(Trim$(UserName)) > NAME_LENGTH Then
            Call AlertMsg("Username needs to be 21 characters or less in length!")
            Exit Function
        End If
        
        If Len(Trim$(Password)) > NAME_LENGTH Then
            Call AlertMsg("Password needs to be 21 characters or less in length!")
            Exit Function
        End If
        
        If Len(Trim$(Password)) >= 3 Then
            IsLoginLegal = True
        Else
            Call AlertMsg("Both passwords needs to be at least 3 characters or more in length!")
        End If
    Else
        Call AlertMsg("Username needs to be at least 3 characters or more in length!")
    End If
    Exit Function
    
' Error handler
ErrorHandler:
    HandleError "IsLoginLegal", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Function

Public Function IsNewCharLegal(ByVal UserName As String) As Boolean
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If Len(Trim$(UserName)) >= 3 Then
        If Len(Trim$(UserName)) <= NAME_LENGTH Then
            IsNewCharLegal = True
        Else
            Call AlertMsg("Username needs to be 21 characters or less in length!")
        End If
    Else
        Call AlertMsg("Username needs to be at least 3 characters or more in length!")
    End If
    Exit Function
    
' Error handler
ErrorHandler:
    HandleError "IsLoginLegal", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Function

Public Function IsStringLegal(ByVal sInput As String) As Boolean
    Dim I As Long

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ' Prevent high ascii chars
    Dim inputLen As Long
    inputLen = Len(sInput)
    For I = 1 To inputLen
        If Asc(Mid$(sInput, I, 1)) < vbKeySpace Or Asc(Mid$(sInput, I, 1)) > vbKeyF15 Then
            Call AlertMsg("You cannot use high ASCII characters in your Name, please re-enter.")
            Exit Function
        End If
    Next

    IsStringLegal = True
    Exit Function
    
' Error handler
ErrorHandler:
    HandleError "IsStringLegal", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Function

' #############
' ## Buttons ##
' #############
Public Sub CacheButtons()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ' Menu - login
    With MenuButton(1)
        .FileName = "login"
        .State = 0 ' Normal
    End With
    
    ' Menu - Register
    With MenuButton(2)
        .FileName = "register"
        .State = 0 ' Normal
    End With
    
    ' Menu - Credits
    With MenuButton(3)
        .FileName = "credits"
        .State = 0 ' Normal
    End With
    
    ' Menu - Exit
    With MenuButton(4)
        .FileName = "exit"
        .State = 0 ' Normal
    End With
    
    ' Main - Inventory
    With MainButton(1)
        .FileName = "btn_inv"
        .State = 0 ' Normal
    End With
    
    ' Main - Spells
    With MainButton(2)
        .FileName = "btn_spells"
        .State = 0 ' Normal
    End With
    
    ' Main - Character
    With MainButton(3)
        .FileName = "btn_chara"
        .State = 0 ' Normal
    End With
    
    ' Main - Options
    With MainButton(4)
        .FileName = "btn_options"
        .State = 0 ' Normal
    End With
    
    ' Main - Trade
    With MainButton(5)
        .FileName = "btn_trade"
        .State = 0 ' Normal
    End With
    
    ' Main - Party
    With MainButton(6)
        .FileName = "btn_party"
        .State = 0 ' Normal
    End With
    
    ' Main - Friends
    With MainButton(7)
        .FileName = "btn_friends"
        .State = 0 ' Normal
    End With
    
    ' Main - Guild
    With MainButton(8)
        .FileName = "btn_guild"
        .State = 0 ' Normal
    End With
    
    ' Main - Notes
    With MainButton(9)
        .FileName = "btn_notes"
        .State = 0 ' Normal
    End With

    ' Main - Titles
    With MainButton(10)
        .FileName = "btn_titles"
        .State = 0 ' Normal
    End With
    
    ' Main - Quests
    With MainButton(11)
        .FileName = "btn_quests"
        .State = 0  ' Normal
    End With
    
    ' Main - Foes
    With MainButton(12)
        .FileName = "btn_foes"
        .State = 0 ' Normal
    End With
    
    ' Main - Map
    With MainButton(13)
        .FileName = "btn_map"
        .State = 0 ' Normal
    End With
    
    ' Main - Hide/Show Buttons
    With MainButton(14)
        .FileName = "btn_showpanels"
        .State = 0 ' Normal
    End With
    
    ' Main - Hide/Show GUI
    With MainButton(15)
        .FileName = "btn_hidegui"
        .State = 0 ' Normal
    End With
    
    ' Main - Equipment
    With MainButton(16)
        .FileName = "btn_equipment"
        .State = 0 ' Normal
    End With
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "CacheButtons", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub ResetMenuButtons()
    Dim I As Long
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    For I = 1 To MAX_MENUBUTTONS
        If Not CurButton_Menu = I Then
            frmMenu.ImgButton(I).Picture = LoadPicture(App.Path & GFX_PATH & "gui\menu\buttons\" & MenuButton(I).FileName & "_norm.jpg")
        End If
    Next
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "ResetMenuButtons", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub RenderButton_Menu(ByVal ButtonNum As Long)
    Dim bSuffix As String
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If ButtonNum > MAX_MENUBUTTONS Then Exit Sub
    
    ' Get the suffix
    Select Case MenuButton(ButtonNum).State
        Case 0 ' Normal
            bSuffix = "_norm"
        Case 1 ' Hover
            bSuffix = "_hover"
        Case 2 ' Click
            bSuffix = "_click"
    End Select
    
    ' Render the button
    frmMenu.ImgButton(ButtonNum).Picture = LoadPicture(App.Path & MENUBUTTON_PATH & MenuButton(ButtonNum).FileName & bSuffix & ".jpg")
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "RenderButton_Menu", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub ChangeButtonState_Menu(ByVal ButtonNum As Long, ByVal bState As Byte)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If ButtonNum > MAX_MENUBUTTONS Then Exit Sub
    
    ' Valid state?
    If bState >= 0 And bState <= 2 Then
        ' Exit out early if state already is same
        If MenuButton(ButtonNum).State = bState Then Exit Sub
        ' Change and render
        MenuButton(ButtonNum).State = bState
        RenderButton_Menu ButtonNum
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "ChangeButtonState_Menu", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub PopulateLists()
    Dim StrLoad As String, I As Long

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ' Cache music list
    StrLoad = Dir$(App.Path & MUSIC_PATH & "*")
    I = 1
    
    If Not StrLoad = vbNullString Then
        Do While StrLoad > vbNullString
            ReDim Preserve MusicCache(1 To I) As String
            MusicCache(I) = StrLoad
            StrLoad = Dir
            I = I + 1
        Loop
    Else
        ReDim Preserve MusicCache(1) As String
        MusicCache(1) = vbNullString
    End If
    
    ' Cache sound list
    StrLoad = Dir$(App.Path & SOUND_PATH & "*")
    I = 1
    
    If Not StrLoad = vbNullString Then
        Do While StrLoad > vbNullString
            ReDim Preserve SoundCache(1 To I) As String
            SoundCache(I) = StrLoad
            StrLoad = Dir
            I = I + 1
        Loop
    Else
        ReDim Preserve SoundCache(1) As String
        SoundCache(I) = vbNullString
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "PopulateLists", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Function IsNameLegal(ByVal sInput As Integer) As Boolean
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If (sInput >= 65 And sInput <= 90) Or (sInput >= 97 And sInput <= 122) Or (sInput = 95) Or (sInput = 32) Or (sInput >= 48 And sInput <= 57) Then
        IsNameLegal = True
    Else
        IsNameLegal = False
    End If
    Exit Function
    
' Error handler
ErrorHandler:
    HandleError "IsNameLegal", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Function

Private Sub NotConnected()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    frmMenu.Visible = True
    frmLoad.Visible = False
    
    ' Reset menu buttons
    CurButton_Menu = 0
    ResetMenuButtons
    
    Call AlertMsg("The server appears to be offline. Please try to reconnect in a few minutes or visit " & GAME_WEBSITE & ".")
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "NotConnected", "modGeneral", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub ResetOptionButtons(Optional ByVal IgnoreMe As Byte = 0)
    If Not IgnoreMe = OptionButtons.Opt_Music Then Call RenderOptionButton(frmMain.picOptionMusic, OptionButtons.Opt_Music, Options.Music)
    If Not IgnoreMe = OptionButtons.Opt_Sound Then Call RenderOptionButton(frmMain.picOptionSound, OptionButtons.Opt_Sound, Options.Sound)
    If Not IgnoreMe = OptionButtons.Opt_WASD Then Call RenderOptionButton(frmMain.picOptionWASD, OptionButtons.Opt_WASD, Options.WASD)
    If Not IgnoreMe = OptionButtons.Opt_PlayerVitals Then Call RenderOptionButton(frmMain.picOptionPlayerVitals, OptionButtons.Opt_PlayerVitals, Options.PlayerVitals)
    If Not IgnoreMe = OptionButtons.Opt_NPCVitals Then Call RenderOptionButton(frmMain.picOptionNpcVitals, OptionButtons.Opt_NPCVitals, Options.NPCVitals)
    If Not IgnoreMe = OptionButtons.Opt_Level Then Call RenderOptionButton(frmMain.picOptionLevel, OptionButtons.Opt_Level, Options.Levels)
    If Not IgnoreMe = OptionButtons.Opt_Guilds Then Call RenderOptionButton(frmMain.picOptionGuild, OptionButtons.Opt_Guilds, Options.Guilds)
    If Not IgnoreMe = OptionButtons.Opt_Mouse Then Call RenderOptionButton(frmMain.picOptionMouse, OptionButtons.Opt_Mouse, Options.Mouse)
    If Not IgnoreMe = OptionButtons.Opt_Title Then Call RenderOptionButton(frmMain.picOptionTitle, OptionButtons.Opt_Title, Options.Titles)
    If Not IgnoreMe = OptionButtons.Opt_BattleMusic Then Call RenderOptionButton(frmMain.picOptionBattleMusic, OptionButtons.Opt_BattleMusic, Options.BattleMusic)
    If Not IgnoreMe = OptionButtons.Opt_SwearFilter Then Call RenderOptionButton(frmMain.picOptionSwearFilter, OptionButtons.Opt_SwearFilter, Options.SwearFilter)
    If Not IgnoreMe = OptionButtons.Opt_Weather Then Call RenderOptionButton(frmMain.picOptionWeather, OptionButtons.Opt_Weather, Options.Weather)
    If Not IgnoreMe = OptionButtons.Opt_AutoTile Then Call RenderOptionButton(frmMain.picOptionAutoTile, OptionButtons.Opt_AutoTile, Options.Autotile)
    If Not IgnoreMe = OptionButtons.Opt_Debug Then Call RenderOptionButton(frmMain.picOptionDebug, OptionButtons.Opt_Debug, Options.Debug)
    If Not IgnoreMe = OptionButtons.Opt_Blood Then Call RenderOptionButton(frmMain.picOptionBlood, OptionButtons.Opt_Blood, Options.Blood)
End Sub

Public Function AlertMsg(ByVal Message As String, Optional ByVal OkayOnly As Boolean = True, Optional ByVal PlaySound As Boolean = True) As Byte
    If PlaySound Then
        Audio.PlaySound "Buzzer1"
    End If
    
    frmAlert.sMessage = Message
    frmAlert.OkayOnly = OkayOnly
    On Error Resume Next
    frmAlert.Show vbModal
    AlertMsg = frmAlert.YesNo
End Function

Public Sub ClearMenuPictures()
    frmMenu.picCharacter.Visible = False
    frmMenu.picCredits.Visible = False
    frmMenu.picLogin.Visible = False
    frmMenu.picMain.Visible = False
    frmMenu.picRegister.Visible = False
End Sub

Public Sub LogoutGame()
    ' Destroy connection
    Call DestroyTCP
    
    Call ClearData
    
    Call Audio.StopMusic
    Call Audio.PlayMusic(Options.MenuMusic)
    Call Audio.StopMapSounds
        
    CloseInterfaces
    GUIVisible = True
    ButtonsVisible = False
    
    Call ClearMenuPictures
    frmMenu.picMain.Visible = True
    ResetMenuButtons
    CurButton_Menu = 0
    frmMenu.lblLAccept.Enabled = True
    HPBarInit = False
    MPBarInit = False
    EXPBarInit = False

    ' Close out all the editors
    Dim tmpForm As Form
    
    For Each tmpForm In Forms
        If tmpForm.Name <> "frmMenu" And tmpForm.Name <> "frmMain" Then
            Unload tmpForm
            Set tmpForm = Nothing
        End If
    Next

    ' Destroy temp values
    MouseX = -1
    MouseY = -1
    Ping = -1
    EqX = 0
    EqY = 0
    BankX = 0
    BankY = 0
    ShopX = 0
    ShopY = 0
    SpellX = 0
    SpellY = 0
    LastItemDesc = 0
    LastSpellDesc = 0
    MyIndex = 0
    InventoryItemSelected = 0
    SpellBuffer = 0
    SpellBufferTimer = 0
    TmpCurrencyItem = 0
    
    ' Hide main form stuffs
    frmMain.txtChat.text = vbNullString
    frmMain.txtMyChat.text = vbNullString
    
      ' Reset buttons manually
    MainButton(14).FileName = "btn_showpanels"
    MainButton(15).FileName = "btn_hidegui"
    
    Call ToggleButtons(False)
    Call frmMain.ResetMainButtons
    frmMain.picBank.Visible = False
    InBank = False
    InGame = False
End Sub

Public Sub InitTimeGetTime()
'*****************************************************************
' Gets the offset time for the timer so we can start at 0 instead of
' the returned system time, allowing us to not have a time roll-over until
' the program is running for 25 days
'*****************************************************************
    ' Get the initial time
    GetSystemTime GetSystemTimeOffset
End Sub

Public Function timeGetTime() As Long
'*****************************************************************
' Grabs the time from the 64-bit system timer and returns it in 32-bit
' after calculating it with the offset - allows us to have the
' "no roll-over" advantage of 64-bit timers with the RAM usage of 32-bit
' though we limit things slightly, so the rollover still happens, but after 25 days
'*****************************************************************
Dim CurrentTime As Currency
    ' Grab the current time (we have to pass a variable ByRef instead of a function return like the other timers)
    GetSystemTime CurrentTime
    
    ' Calculate the difference between the 64-bit times, return as a 32-bit time
    timeGetTime = CurrentTime - GetSystemTimeOffset
End Function

Public Sub CloseInterfaces()
    ' Close if in a shop
    If InShop > 0 Then CloseShop

    ' Close if in bank
    If InBank Then CloseBank

    ' Close if in trade
    If frmMain.picTrade.Visible Then CloseTrade
End Sub

Function FormCount(ByVal frmName As String) As Long
    Dim frm As Form, counter As Long
    
    FormCount = -1
    
    For Each frm In Forms
        If StrComp(frm.Name, frmName, vbTextCompare) = 0 Then
            FormCount = counter
            Exit For
        End If
        counter = counter + 1
    Next
End Function

Function FormLoaded(ByVal frmName As String) As Boolean
    Dim formNum As Long
    
    formNum = FormCount(frmName)
    If formNum >= 0 Then
        FormLoaded = True
    End If
End Function

Function FormVisible(ByVal frmName As String) As Boolean
    Dim formNum As Long
    
    formNum = FormCount(frmName)
    If formNum >= 0 Then
        If Forms(formNum).Visible Then
            FormVisible = True
        End If
    End If
End Function
