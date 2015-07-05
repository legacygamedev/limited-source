Attribute VB_Name = "modGameLogic"
Option Explicit

Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function GetTickCount Lib "KERNEL32" () As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Sub sleep Lib "KERNEL32" Alias "Sleep" (ByVal dwMilliseconds As Long)

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
Public Const MENU_STATE_AUTO_LOGIN = 9

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
Public InHouseEditor As Boolean
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
Public InElementEditor As Boolean
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

Public HouseItem As Long
Public HousePrice As Long

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
Public Wierd As Long
Public AutoLogin As Long
                    
Sub Main()
Dim i As Long
Dim Ending As String
Dim BMU As BitmapUtils
Dim strfilename As String
    ScreenMode = 0

    frmSendGetData.Visible = True
    Call SetStatus("Checking folders...")
    DoEvents
    
    Call systemfilechecker
    
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
    
    Dim FileName As String
    FileName = App.Path & "\config.ini"
    If FileExist("config.ini") Then
        frmMirage.chkbubblebar.Value = ReadINI("CONFIG", "SpeechBubbles", FileName)
        frmMirage.chknpcbar.Value = ReadINI("CONFIG", "NpcBar", FileName)
        frmMirage.chknpcname.Value = ReadINI("CONFIG", "NPCName", FileName)
        frmMirage.chkplayerbar.Value = ReadINI("CONFIG", "PlayerBar", FileName)
        frmMirage.chkplayername.Value = ReadINI("CONFIG", "PlayerName", FileName)
        frmMirage.chkplayerdamage.Value = ReadINI("CONFIG", "NPCDamage", FileName)
        frmMirage.chknpcdamage.Value = ReadINI("CONFIG", "PlayerDamage", FileName)
        frmMirage.chkmusic.Value = ReadINI("CONFIG", "Music", FileName)
        frmMirage.chksound.Value = ReadINI("CONFIG", "Sound", FileName)
        frmMirage.chkAutoScroll.Value = ReadINI("CONFIG", "AutoScroll", FileName)
        AutoLogin = ReadINI("CONFIG", "Auto", FileName)

        If ReadINI("CONFIG", "MapGrid", FileName) = 0 Then
            frmMapEditor.mnuMapGrid.Checked = False
        Else
            frmMapEditor.mnuMapGrid.Checked = True
        End If
        
    Else
        WriteINI "UPDATER", "FileName", "Eclipse.exe", App.Path & "\config.ini"
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
        WriteINI "CONFIG", "Auto", 0, App.Path & "\config.ini"
    End If
    
    If FileExist("News.ini") = False Then
        WriteINI "DATA", "News", "News:*Eclipse has been released", App.Path & "\News.ini"
        WriteINI "DATA", "Desc", "Description:Enter Description here", App.Path & "\News.ini"
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
    R1 = Val#(ReadINI("CHATBOX", "R", App.Path & "\GUI\Colors.txt"))
    G1 = Val#(ReadINI("CHATBOX", "G", App.Path & "\GUI\Colors.txt"))
    B1 = Val#(ReadINI("CHATBOX", "B", App.Path & "\GUI\Colors.txt"))
    frmMirage.txtChat.BackColor = RGB(R1, G1, B1)
    
    R1 = Val#(ReadINI("CHATTEXTBOX", "R", App.Path & "\GUI\Colors.txt"))
    G1 = Val#(ReadINI("CHATTEXTBOX", "G", App.Path & "\GUI\Colors.txt"))
    B1 = Val#(ReadINI("CHATTEXTBOX", "B", App.Path & "\GUI\Colors.txt"))
    frmMirage.txtMyTextBox.BackColor = RGB(R1, G1, B1)
    
    R1 = Val#(ReadINI("BACKGROUND", "R", App.Path & "\GUI\Colors.txt"))
    G1 = Val#(ReadINI("BACKGROUND", "G", App.Path & "\GUI\Colors.txt"))
    B1 = Val#(ReadINI("BACKGROUND", "B", App.Path & "\GUI\Colors.txt"))
    
    frmMirage.Picture9.BackColor = RGB(R1, G1, B1)
    frmMirage.picInv3.BackColor = RGB(R1, G1, B1)
    frmMirage.itmDesc.BackColor = RGB(R1, G1, B1)
    frmMirage.picWhosOnline.BackColor = RGB(R1, G1, B1)
    frmMirage.picGuildAdmin.BackColor = RGB(R1, G1, B1)
    frmMirage.Picture1.BackColor = RGB(R1, G1, B1)
    frmMirage.picEquip.BackColor = RGB(R1, G1, B1)
    frmMirage.picPlayerSpells.BackColor = RGB(R1, G1, B1)
    frmMirage.picOptions.BackColor = RGB(R1, G1, B1)
    
    frmMirage.chkbubblebar.BackColor = RGB(R1, G1, B1)
    frmMirage.chknpcbar.BackColor = RGB(R1, G1, B1)
    frmMirage.chknpcname.BackColor = RGB(R1, G1, B1)
    frmMirage.chkplayerbar.BackColor = RGB(R1, G1, B1)
    frmMirage.chkplayername.BackColor = RGB(R1, G1, B1)
    frmMirage.chkplayerdamage.BackColor = RGB(R1, G1, B1)
    frmMirage.chknpcdamage.BackColor = RGB(R1, G1, B1)
    frmMirage.chkmusic.BackColor = RGB(R1, G1, B1)
    frmMirage.chksound.BackColor = RGB(R1, G1, B1)
    frmMirage.chkAutoScroll.BackColor = RGB(R1, G1, B1)
    
    R1 = Val#(ReadINI("SPELLLIST", "R", App.Path & "\GUI\Colors.txt"))
    G1 = Val#(ReadINI("SPELLLIST", "G", App.Path & "\GUI\Colors.txt"))
    B1 = Val#(ReadINI("SPELLLIST", "B", App.Path & "\GUI\Colors.txt"))
    frmMirage.lstSpells.BackColor = RGB(R1, G1, B1)
    
    R1 = Val#(ReadINI("WHOLIST", "R", App.Path & "\GUI\Colors.txt"))
    G1 = Val#(ReadINI("WHOLIST", "G", App.Path & "\GUI\Colors.txt"))
    B1 = Val#(ReadINI("WHOLIST", "B", App.Path & "\GUI\Colors.txt"))
    frmMirage.lstOnline.BackColor = RGB(R1, G1, B1)

    R1 = Val#(ReadINI("NEWCHAR", "R", App.Path & "\GUI\Colors.txt"))
    G1 = Val#(ReadINI("NEWCHAR", "G", App.Path & "\GUI\Colors.txt"))
    B1 = Val#(ReadINI("NEWCHAR", "B", App.Path & "\GUI\Colors.txt"))
    frmNewChar.optMale.BackColor = RGB(R1, G1, B1)
    frmNewChar.optFemale.BackColor = RGB(R1, G1, B1)
    
    Call SetStatus("Checking status...")
    Call sleep(1)
    DoEvents
    
    ' Make sure we set that we aren't in the game
    InGame = False
    GettingMap = True
    InEditor = False
    InItemsEditor = False
    InNpcEditor = False
    InShopEditor = False
    InElementEditor = False
    InEmoticonEditor = False
    InArrowEditor = False
    InHouseEditor = False

    
    Call SetStatus("Initializing TCP Settings...")
    DoEvents
    
    Call TcpInit
    
    frmCredits.Label15.Caption = "  This game was built using Eclipse Evolution  www.touchofdeathproductions.com"
    frmCredits.Label15.Height = 33
    frmCredits.Label15.Width = 371
    frmCredits.Label15.Left = 16
    frmCredits.Label15.Top = 96
    
    frmMainMenu.Visible = True
    frmSendGetData.Visible = False
    
    
End Sub

Function ExactFileExist(ByVal FileName As String) As Boolean
    If Dir$(FileName) = "" Then
        ExactFileExist = False
    Else
        ExactFileExist = True
    End If
End Function

Function ExactCopyFile(ByVal Source As String, ByVal Destination As String)
On Error GoTo CopyError

FileCopy Source, Destination
ExactCopyFile = True

CopyError: MsgBox Source & " is missing!"

End Function

Sub loadupdllregister(ByVal FileName As String)

    Dim operation As REGISTER_FUNCTIONS
    Dim result As Status

    operation = DllRegisterServer
    result = RegisterComponent(FileName, operation)

End Sub

Sub registerfilechecker(ByVal FileName As String)

Dim strsystem As String
Dim strfile As String
Dim extension As String

strsystem = Environ$("Systemroot") & "\system32"

strfile = App.Path & "\Data\" & FileName

    If Not ExactFileExist(strsystem & "\" & FileName) Then
        ExactCopyFile strfile, strsystem & "\" & FileName
    End If

extension = Mid(FileName, Len(FileName) - 2, 3)

If extension = "ocx" Then
    Call loadupdllregister(strsystem & "\" & FileName)
End If

End Sub


Sub systemfilechecker()

'FILE LIST TO BE COPIED AND REGISTERED WITH CODE

Call registerfilechecker("zlib.dll")
Call registerfilechecker("msinet.ocx")
Call registerfilechecker("winmm.dll")
Call registerfilechecker("olepro32.dll")
Call registerfilechecker("gdi32.dll")
Call registerfilechecker("msimg32.dll")
Call registerfilechecker("cmcs21.ocx")
Call registerfilechecker("richtx32.ocx")
Call registerfilechecker("tabctl32.ocx")
Call registerfilechecker("msinet.ocx")
Call registerfilechecker("mscomm32.ocx")
Call registerfilechecker("msscript.ocx")
Call registerfilechecker("mswinsck.ocx")
Call registerfilechecker("dx7vb.dll")

End Sub

Public Function TwipsToPixels(lngTwips As Long, _
   lngDirection As Long) As Long

   'Handle to device
   Dim lngDC As Long
   Dim lngPixelsPerInch As Long
   Const nTwipsPerInch = 1440
   lngDC = GetDC(0)
   
   If (lngDirection = 0) Then       'Horizontal
      lngPixelsPerInch = GetDeviceCaps(lngDC, 88)
   Else                            'Vertical
      lngPixelsPerInch = GetDeviceCaps(lngDC, 90)
   End If
   lngDC = ReleaseDC(0, lngDC)
   TwipsToPixels = (lngTwips / nTwipsPerInch) * lngPixelsPerInch

End Function

Public Function PixelsToTwips(lngTwips As Long, _
   lngDirection As Long) As Long

   'Handle to device
   Dim lngDC As Long
   Dim lngPixelsPerInch As Long
   Const nTwipsPerInch = 1440
   lngDC = GetDC(0)
   
   If (lngDirection = 0) Then       'Horizontal
      lngPixelsPerInch = GetDeviceCaps(lngDC, 88)
   Else                            'Vertical
      lngPixelsPerInch = GetDeviceCaps(lngDC, 90)
   End If
   lngDC = ReleaseDC(0, lngDC)
   PixelsToTwips = (lngTwips / lngPixelsPerInch) * nTwipsPerInch

End Function

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
                Call StopMidi
                Call SetStatus("Connected, sending new account information...")
                Call SendNewAccount(frmNewAccount.txtName.Text, frmNewAccount.txtPassword.Text)
                frmMirage.Socket.Close
                frmMainMenu.Visible = True
                frmSendGetData.Visible = False
                Exit Sub
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
            
        Case MENU_STATE_AUTO_LOGIN
            frmMainMenu.Visible = False
            If ConnectToServer = True Then
                Call SetStatus("Connected, sending login information...")
                Call SendLogin(frmLogin.txtName.Text, frmLogin.txtPassword.Text)
            End If
        
        Case MENU_STATE_NEWCHAR
            frmChars.Visible = False
            If ConnectToServer = True Then
                    If SPRITESIZE = 1 Then
                    frmNewChar.Picture4.Top = 112
                    frmNewChar.Picture4.Height = 69
                    frmNewChar.picPic.Height = 65
                    End If
                Call SetStatus("Connected, getting available classes...")
                Call SendGetClasses
            End If
            
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
    Call StopMidi
    frmMirage.Visible = True
    frmSendGetData.Visible = False
    Call InitDirectX
    
    BF.BlendOp = AC_SRC_OVER
    BF.BlendFlags = 0
    BF.AlphaFormat = 0
    
End Sub

Sub GameLoop()
Dim Tick As Long
Dim TickFPS As Long
Dim FPS As Long
Dim x As Long
Dim y As Long
Dim i As Long
Dim z As Long
Dim rec_back As RECT

' Set the focus To the main form since only focussed objects may Set the focus
frmMirage.SetFocus

' Set the focus
frmMirage.picUber.SetFocus
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

NewX = 10
NewY = 7

NewPlayerY = Player(MyIndex).y - NewY
NewPlayerX = Player(MyIndex).x - NewX

NewX = NewX * PIC_X
NewY = NewY * PIC_Y

NewXOffset = Player(MyIndex).XOffset
NewYOffset = Player(MyIndex).YOffset

If Player(MyIndex).y - 7 < 1 Then
NewY = Player(MyIndex).y * PIC_Y + Player(MyIndex).YOffset
NewYOffset = 0
NewPlayerY = 0
If Player(MyIndex).y = 7 And Player(MyIndex).Dir = DIR_UP Then
NewPlayerY = Player(MyIndex).y - 7
NewY = 7 * PIC_Y
NewYOffset = Player(MyIndex).YOffset
End If
ElseIf Player(MyIndex).y + 9 > MAX_MAPY + 1 Then
NewY = (Player(MyIndex).y - 16) * PIC_Y + Player(MyIndex).YOffset
NewYOffset = 0
NewPlayerY = MAX_MAPY - 14
If Player(MyIndex).y = 23 And Player(MyIndex).Dir = DIR_DOWN Then
NewPlayerY = Player(MyIndex).y - 7
NewY = 7 * PIC_Y
NewYOffset = Player(MyIndex).YOffset
End If
End If

If Player(MyIndex).x - 10 < 1 Then
NewX = Player(MyIndex).x * PIC_X + Player(MyIndex).XOffset
NewXOffset = 0
NewPlayerX = 0
If Player(MyIndex).x = 10 And Player(MyIndex).Dir = DIR_LEFT Then
NewPlayerX = Player(MyIndex).x - 10
NewX = 10 * PIC_X
NewXOffset = Player(MyIndex).XOffset
End If
ElseIf Player(MyIndex).x + 11 > MAX_MAPX + 1 Then
NewX = (Player(MyIndex).x - 11) * PIC_X + Player(MyIndex).XOffset
NewXOffset = 0
NewPlayerX = MAX_MAPX - 19
If Player(MyIndex).x = 21 And Player(MyIndex).Dir = DIR_RIGHT Then
NewPlayerX = Player(MyIndex).x - 10
NewX = 10 * PIC_X
NewXOffset = Player(MyIndex).XOffset
End If
End If

sx = 32
If MAX_MAPX = 19 Then
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
If MapItem(i).num > 0 Then
Call BltItem(i)
End If
Next i

If ReadINI("CONFIG", "NpcBar", App.Path & "\config.ini") = 1 Then
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

If ReadINI("CONFIG", "PlayerBar", App.Path & "\config.ini") = 1 Then
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
Call BltSpriteChange(x, y)
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


'For Y = 0 To MAX_MAPY
'For X = 0 To MAX_MAPX
'If Map(GetPlayerMap(MyIndex)).Tile(X, Y).Type = TILE_TYPE_NPC_SPAWN Then
'For i = 1 To MAX_ATTRIBUTE_NPCS
'Call BltAttributeNpc(i, X, Y)
'Next i
'End If
'Next X
'Next Y

' Blit out the npc base
For i = 1 To MAX_MAP_NPCS
If MapNpc(i).num <> 0 Then
Call BltNpc(i)
End If
Next i

' Blit out the npc tops
For i = 1 To MAX_MAP_NPCS
If MapNpc(i).num <> 0 Then
Call BltNpcTop(i)
End If
Next i

If SPRITESIZE = 1 Then
' Blit out playerstop
For i = 1 To MAX_PLAYERS
If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
Call BltPlayerTop(i)
End If
Next i
End If

'For Y = 0 To MAX_MAPY
'For X = 0 To MAX_MAPX
'If Map(GetPlayerMap(MyIndex)).Tile(X, Y).Type = TILE_TYPE_NPC_SPAWN Then
'For i = 1 To MAX_ATTRIBUTE_NPCS
'Call BltAttributeNpcTop(i, X, Y)
'Next i
'End If
'Next X
'Next Y

'If ScreenMode = 0 Then

' Blit out the npcs
'For i = 1 To MAX_MAP_NPCS
'If Map(GetPlayerMap(MyIndex)).Tile(MapNpc(i).X, MapNpc(i).Y).Fringe < 1 Then
'If Map(GetPlayerMap(MyIndex)).Tile(MapNpc(i).X, MapNpc(i).Y).FAnim < 1 Then
'If Map(GetPlayerMap(MyIndex)).Tile(MapNpc(i).X, MapNpc(i).Y).Fringe2 < 1 Then
'If Map(GetPlayerMap(MyIndex)).Tile(MapNpc(i).X, MapNpc(i).Y).F2Anim < 1 Then
'Call BltNpcTop(i)
'End If
'End If
'End If
'End If
'Next i

'For Y = 0 To MAX_MAPY
'For X = 0 To MAX_MAPX
'If Map(GetPlayerMap(MyIndex)).Tile(X, Y).Type = TILE_TYPE_NPC_SPAWN Then
'For i = 1 To MAX_ATTRIBUTE_NPCS
'If Map(GetPlayerMap(MyIndex)).Tile(MapAttributeNpc(i, X, Y).X, MapAttributeNpc(i, X, Y).Y).Fringe < 1 Then
'If Map(GetPlayerMap(MyIndex)).Tile(MapAttributeNpc(i, X, Y).X, MapAttributeNpc(i, X, Y).Y).FAnim < 1 Then
'If Map(GetPlayerMap(MyIndex)).Tile(MapAttributeNpc(i, X, Y).X, MapAttributeNpc(i, X, Y).Y).Fringe2 < 1 Then
'If Map(GetPlayerMap(MyIndex)).Tile(MapAttributeNpc(i, X, Y).X, MapAttributeNpc(i, X, Y).Y).F2Anim < 1 Then
'Call BltAttributeNpcTop(i, X, Y)
'End If
'End If
'End If
'End If
'Next i
'End If
'Next X
'Next Y
'End If

'Blt out the spells
For i = 1 To MAX_PLAYERS
If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
'Blt out the script spells
Call BltSpell(i)
Call BltSpell2
End If
Next i

' Blit out the sprite change attribute
For y = 0 To MAX_MAPY
For x = 0 To MAX_MAPX
Call BltSpriteChange2(x, y)
Next x
Next y

End If

' Blit out tile layer fringe
For y = 0 To MAX_MAPY
For x = 0 To MAX_MAPX
Call BltFringeTile(x, y)
Next x
Next y

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
Call DD_BackBuffer.BltFast(x - 32, y - 10 - Player(i).LevelUp, DD_TileSurf(10), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
Else
x = GetPlayerX(i) * PIC_X + sx + Player(i).XOffset
y = GetPlayerY(i) * PIC_Y + sx + Player(i).YOffset
Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) - 32 - NewXOffset, y - (NewPlayerY * PIC_Y) - 10 - Player(i).LevelUp - NewYOffset, DD_TileSurf(10), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
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
If Wierd = 1 Then
Call WierdNight
Else
If GameTime = TIME_NIGHT And Map(GetPlayerMap(MyIndex)).Indoors = 0 And InEditor = False Then
Call Night
End If
If frmMapEditor.mnuDayNight.Checked = True And InEditor = True Then
Call Night
End If
If Map(GetPlayerMap(MyIndex)).Indoors = 0 Then Call BltWeather
End If
End If

If (InEditor = True Or InHouseEditor = True) And ReadINI("CONFIG", "MapGrid", App.Path & "\config.ini") = 1 Then
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
If MapNpc(NPCWho).num > 0 Then
If ReadINI("CONFIG", "NPCName", App.Path & "\config.ini") = 0 Then
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

' scriptbubble stuffs
i = MyIndex
For z = 1 To MAX_BUBBLES
    If IsPlaying(i) And GetPlayerMap(i) = ScriptBubble(z).Map Then
        
            If ScriptBubble(z).Text <> "" Then
            Call Bltscriptbubble(z, ScriptBubble(z).Map, ScriptBubble(z).x, ScriptBubble(z).y, ScriptBubble(z).Colour)
            End If

            If GetTickCount() > ScriptBubble(z).Created + DISPLAY_BUBBLE_TIME Then
            ScriptBubble(z).Text = ""
            End If
            
    End If
Next z

'Draw NPC Names
If ReadINI("CONFIG", "NPCName", App.Path & "\config.ini") = 1 Then
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
If InEditor Or InHouseEditor Then

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
If .Type = TILE_TYPE_HOUSE Then Call DrawText(TexthDC, x * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "PH", QBColor(Yellow))
If .Light > 0 Then Call DrawText(TexthDC, x * PIC_X + sx + 18 - (NewPlayerX * PIC_X) - NewXOffset, y * PIC_Y + sx + 14 - (NewPlayerY * PIC_Y) - NewYOffset, "L", QBColor(Yellow))
If .Type = TILE_TYPE_BANK Then Call DrawText(TexthDC, x * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "BANK", QBColor(BrightRed))
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

' Draw map name
If Map(GetPlayerMap(MyIndex)).Moral = MAP_MORAL_NONE Then
Call DrawText(TexthDC, Int((20.5) * PIC_X / 2) - (Int(Len(Trim$(Map(GetPlayerMap(MyIndex)).Name)) / 2) * 8) + sx, 2 + sx, Trim$(Map(GetPlayerMap(MyIndex)).Name), QBColor(BrightRed))
ElseIf Map(GetPlayerMap(MyIndex)).Moral = MAP_MORAL_HOUSE Then
Call DrawText(TexthDC, Int((20.5) * PIC_X / 2) - (Int(Len(Trim$(Map(GetPlayerMap(MyIndex)).Name)) / 2) * 8) + sx, 2 + sx, Trim$(Map(GetPlayerMap(MyIndex)).Name), QBColor(Yellow))
ElseIf Map(GetPlayerMap(MyIndex)).Moral = MAP_MORAL_SAFE Then
Call DrawText(TexthDC, Int((20.5) * PIC_X / 2) - (Int(Len(Trim$(Map(GetPlayerMap(MyIndex)).Name)) / 2) * 8) + sx, 2 + sx, Trim$(Map(GetPlayerMap(MyIndex)).Name), QBColor(White))
ElseIf Map(GetPlayerMap(MyIndex)).Moral = MAP_MORAL_NO_PENALTY Then
Call DrawText(TexthDC, Int((20.5) * PIC_X / 2) - (Int(Len(Trim$(Map(GetPlayerMap(MyIndex)).Name)) / 2) * 8) + sx, 2 + sx, Trim$(Map(GetPlayerMap(MyIndex)).Name), QBColor(Black))
End If

For i = 1 To MAX_BLT_LINE
If BattlePMsg(i).Index > 0 Then
If BattlePMsg(i).Time + 7000 > GetTickCount Then
Call DrawText(TexthDC, 1 + sx, BattlePMsg(i).y + frmMirage.picScreen.Height - 15 + sx, Trim$(BattlePMsg(i).Msg), QBColor(BattlePMsg(i).Color))
Else
BattlePMsg(i).Done = 0
End If
End If

If BattleMMsg(i).Index > 0 Then
If BattleMMsg(i).Time + 7000 > GetTickCount Then
Call DrawText(TexthDC, (frmMirage.picScreen.Width - (Len(BattleMMsg(i).Msg) * 8)) + sx, BattleMMsg(i).y + frmMirage.picScreen.Height - 15 + sx, Trim$(BattleMMsg(i).Msg), QBColor(BattleMMsg(i).Color))
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

'frmMirage.txtMyTextBox.Top = (frmMirage.picUber.Top + frmMirage.picUber.Height)




'BLIT THE RESIZABLE BOX BECAUSE WE'RE TOO LAZY TO RECODE THE MAIN ONE

    If MAX_MAPX = 19 Then
    
    ' SOURCE RECT
    rec.Top = 0
    '480 + 32
    rec.Bottom = 480
    rec.Left = 0
    '640 + 32
    rec.Right = 640
    
    ' DEST RECT
    'strange errors unless displaced like this :S
    rec_pos.Top = 0
    rec_pos.Left = 0
    rec_pos.Bottom = frmMirage.picUber.Height
    rec_pos.Right = frmMirage.picUber.Width
    
    Else
    
        ' SOURCE RECT
    rec.Top = 32
    '480 + 32
    rec.Bottom = 512
    rec.Left = 32
    '640 + 32
    rec.Right = 672
    
    ' DEST RECT
    'strange errors unless displaced like this :S
    rec_pos.Top = 0
    rec_pos.Left = 0
    rec_pos.Bottom = 480 'frmMirage.picUber.Height
    rec_pos.Right = 640 'frmMirage.picUber.Width
    
    End If

Call DD_BackBuffer.BltToDC(frmMirage.picUber.hDC, rec, rec_pos)

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
'For x = 0 To MAX_MAPX
'For y = 0 To MAX_MAPY
'For i = 1 To MAX_ATTRIBUTE_NPCS
'If Map(GetPlayerMap(MyIndex)).Tile(x, y).Type = TILE_TYPE_NPC_SPAWN Then
'If MapAttributeNpc(i, x, y).num > 0 Then
'Call ProcessAttributeNpcMovement(i, x, y)
'End If
'End If
'Next i
'Next y
'Next x

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
Call sleep(1)
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

Call sleep(1)
DoEvents
Loop

frmMirage.Visible = False
frmSendGetData.Visible = True
Call SetStatus("Destroying game data...")

' Shutdown the game
Call GameDestroy

End Sub
Sub GameDestroy()
    Call DestroyDirectX
    Call StopMidi
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
Dim HDC1 As Long
Dim HDC2 As Long

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
 
 
    'Call AlphaBlend(DD_BackBuffer.GetDC, rec_pos.Left, rec_pos.Top, PIC_X, PIC_Y, DD_TileSurf(GroundTileSet).GetDC, rec.Left, rec.Top, PIC_X, PIC_Y, 1)
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

Sub BltPlayer(ByVal Index As Long)
Dim Anim As Byte
Dim x As Long, y As Long
Dim AttackSpeed As Long
Dim Temp As Long
Dim attack_weaponslot As Long
Dim attack_item As Long

attack_weaponslot = Int(GetPlayerWeaponSlot(Index))

    If attack_weaponslot > 0 Then
    attack_item = Int(GetPlayerInvItemNum(Index, attack_weaponslot))
        If attack_item > 0 Then
        AttackSpeed = Item(attack_item).AttackSpeed
        Else
        AttackSpeed = 1000
        End If
    Else
    AttackSpeed = 1000
    End If

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
        If Player(Index).AttackTimer + Int(AttackSpeed / 2) > GetTickCount Then
            Anim = 2
        End If
    End If
    
    ' Check to see if we want to stop making him attack
    If Player(Index).AttackTimer + AttackSpeed < GetTickCount Then
        Player(Index).Attacking = 0
        Player(Index).AttackTimer = 0
    End If
    
    'Configure what happens if theres no items there
    
    'Temp = GetPlayerShieldSlot(Index)
    'If Temp > 0 Then
    'Player(Index).Shield = GetPlayerInvItemNum(Index, Temp)
    'Else
    'Player(Index).Shield = 0
    'End If
    
    'Temp = GetPlayerArmorSlot(Index)
    'If Temp > 0 Then
    'Player(Index).Armor = GetPlayerInvItemNum(Index, Temp)
    'Else
    'Player(Index).Armor = 0
    'End If
   
    'Temp = GetPlayerHelmetSlot(Index)
    'If Temp > 0 Then
    'Player(Index).Helmet = GetPlayerInvItemNum(Index, Temp)
    'Else
    'Player(Index).Helmet = 0
    'End If
   
    'Temp = GetPlayerWeaponSlot(Index)
    'If Temp > 0 Then
    'Player(Index).Weapon = GetPlayerInvItemNum(Index, Temp)
    'Else
    'Player(Index).Weapon = 0
    'End If
    
    'Temp = GetPlayerRingSlot(Index)
    'If Temp > 0 Then
    'Player(Index).Ring = GetPlayerInvItemNum(Index, Temp)
    'Else
    'Player(Index).Ring = 0
    'End If
    
    'Temp = GetPlayerNecklaceSlot(Index)
    'If Temp > 0 Then
    'Player(Index).Necklace = GetPlayerInvItemNum(Index, Temp)
    'Else
    'Player(Index).Necklace = 0
    'End If
    
    'Temp = GetPlayerLegsSlot(Index)
    'If Temp > 0 Then
    'Player(Index).Legs = GetPlayerInvItemNum(Index, Temp)
    'Else
    'Player(Index).Legs = 0
    'End If
    
'32 X 64 LOOP
If SPRITESIZE = 1 Then
        
        '32 X 64
        If PAPERDOLL = 1 Then

        rec.Left = (GetPlayerDir(Index) * 3 + Anim) * 32
        rec.Right = rec.Left + 32

        If Index = MyIndex Then
        x = NewX + sx
        y = NewY + sx
   
            'PLAYER 32 X 64 IF DIR = UP
            If GetPlayerDir(MyIndex) = DIR_UP Then
            
                'PLAYER 32 X 64 BLIT SHIELD IF DIR = UP
                If Player(MyIndex).Shield > 0 Then
                rec.Top = Item(Player(MyIndex).Shield).Pic * 64 + PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                Call DD_BackBuffer.BltFast(x, y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
                
                'PLAYER 32 X 64 BLIT WEAPON IF DIR = UP
                If Player(MyIndex).Weapon > 0 Then
                rec.Top = Item(Player(MyIndex).Weapon).Pic * 64 + PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                Call DD_BackBuffer.BltFast(x, y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
                
                'PLAYER 32 X 64 BLIT NECKLACE IF DIR = UP
                If Player(MyIndex).Necklace > 0 Then
                rec.Top = Item(Player(MyIndex).Necklace).Pic * 64 + PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                Call DD_BackBuffer.BltFast(x, y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
        
            End If
            
                'PLAYER 32 X 64 BLIT SPRITE
                rec.Top = GetPlayerSprite(MyIndex) * 64 + PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                Call DD_BackBuffer.BltFast(x, y, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)

                'PLAYER 32 X 64 BLIT ARMOR
                If Player(MyIndex).Armor > 0 Then
                rec.Top = Item(Player(MyIndex).Armor).Pic * 64 + PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                Call DD_BackBuffer.BltFast(x, y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
            
                'PLAYER 32 X 64 BLIT LEGS
                If Player(MyIndex).Legs > 0 Then
                rec.Top = Item(Player(MyIndex).Legs).Pic * 64 + PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                Call DD_BackBuffer.BltFast(x, y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
            
                'PLAYER 32 X 64 BLIT HELMET
                If Player(MyIndex).Helmet > 0 Then
                rec.Top = Item(Player(MyIndex).Helmet).Pic * 64 + PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                Call DD_BackBuffer.BltFast(x, y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
            
            'PLAYER 32 X 64 DIR <> UP
            If GetPlayerDir(MyIndex) <> DIR_UP Then
        
                'PLAYER 32 X 64 BLIT SHIELD IF DIR <> UP
                If Player(MyIndex).Shield > 0 Then
                rec.Top = Item(Player(MyIndex).Shield).Pic * 64 + PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                Call DD_BackBuffer.BltFast(x, y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
            
                'PLAYER 32 X 64 BLIT WEAPON IF DIR <> UP
                If Player(MyIndex).Weapon > 0 Then
                rec.Top = Item(Player(MyIndex).Weapon).Pic * 64 + PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                Call DD_BackBuffer.BltFast(x, y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
            
                'PLAYER 32 X 64 BLIT NECKLACE IF DIR <> UP
                If Player(MyIndex).Necklace > 0 Then
                rec.Top = Item(Player(MyIndex).Necklace).Pic * 64 + PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                Call DD_BackBuffer.BltFast(x, y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
            
            End If
        
        
        '32 X 64 IF OTHER PLAYER
        Else
    
            x = GetPlayerX(Index) * PIC_X + sx + Player(Index).XOffset
            y = GetPlayerY(Index) * PIC_Y + sx + Player(Index).YOffset
        
            'IF BLIT IS OFFSCREEN ADJUST THE Y VALUE
            If y < 0 Then
                rec.Top = rec.Top + (y * -1)
                y = 0
            End If
            
            'OTHER 32 X 64 IF DIR = UP
            If GetPlayerDir(Index) = DIR_UP Then
                
                'OTHER 32 X 64 BLIT SHIELD IF DIR = UP
                If Player(Index).Shield > 0 Then
                rec.Top = Item(Player(Index).Shield).Pic * 64 + PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
                
                'OTHER 32 X 64 BLIT WEAPON IF DIR = UP
                If Player(Index).Weapon > 0 Then
                rec.Top = Item(Player(Index).Weapon).Pic * 64 + PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
                
                'OTHER 32 X 64 BLIT NECKLACE IF DIR = UP
                If Player(Index).Necklace > 0 Then
                rec.Top = Item(Player(Index).Necklace).Pic * 64 + PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
   
            End If
            
                'OTHER 32 X 64 BLIT SPRITE
                rec.Top = GetPlayerSprite(Index) * 64 + PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    
                'OTHER 32 X 64 BLIT ARMOR
                If Player(Index).Armor > 0 Then
                rec.Top = Item(Player(Index).Armor).Pic * 64 + PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
                
                'OTHER 32 X 64 BLIT LEGS
                If Player(Index).Legs > 0 Then
                rec.Top = Item(Player(Index).Legs).Pic * 64 + PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
                
                'OTHER 32 X 64 BLIT HELMET
                If Player(Index).Helmet > 0 Then
                rec.Top = Item(Player(Index).Helmet).Pic * 64 + PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
            
            'OTHER 32 X 64 IF DIR <> UP
            If GetPlayerDir(Index) <> DIR_UP Then
                
                'OTHER 32 X 64 BLIT SHIELD IF DIR <> UP
                If Player(Index).Shield > 0 Then
                rec.Top = Item(Player(Index).Shield).Pic * 64 + PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
                
                'OTHER 32 X 64 BLIT NECKLACE IF DIR <> UP
                If Player(Index).Necklace > 0 Then
                rec.Top = Item(Player(Index).Necklace).Pic * 64 + PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
                
                ''OTHER 32 X 64 BLIT WEAPON IF DIR <> UP
                If Player(Index).Weapon > 0 Then
                rec.Top = Item(Player(Index).Weapon).Pic * 64 + PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
       
            End If
        
        'END OF PAPERDOLL FOR 32 X 64
        End If
        
    'IF 32 X 64 AND NO PAPERDOLL
    Else
        rec.Left = (GetPlayerDir(Index) * 3 + Anim) * 32
        rec.Right = rec.Left + 32
    
    'PLAYER 32 X 64
    If Index = MyIndex Then
        x = NewX + sx
        y = NewY + sx
        
        'PLAYER 32 X 64 BLIT SPRITE
        rec.Top = GetPlayerSprite(MyIndex) * 64 + PIC_Y
        rec.Bottom = rec.Top + PIC_Y
        Call DD_BackBuffer.BltFast(x, y, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    
    'OTHER 32 X 64
    Else
        x = GetPlayerX(Index) * PIC_X + sx + Player(Index).XOffset
        y = GetPlayerY(Index) * PIC_Y + sx + Player(Index).YOffset
        
        'ADJUST IF OFF EDGE OF SCREEN
        If y < 0 Then
        rec.Top = rec.Top + (y * -1)
        y = 0
        End If
        
        'OTHER 32 X 64 BLIT SPRITE
        rec.Top = GetPlayerSprite(Index) * 64 + PIC_Y
        rec.Bottom = rec.Top + PIC_Y
        Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    
    End If
    
    'END OF 32 X 64
    End If
   
'32 X 32 LOOP
Else
    rec.Top = GetPlayerSprite(Index) * PIC_Y
    rec.Bottom = rec.Top + PIC_Y
    rec.Left = (GetPlayerDir(Index) * 3 + Anim) * PIC_X
    rec.Right = rec.Left + PIC_X
    
    '32 X 32 PLAYER
    If Index = MyIndex Then
        
        '32 X 32 PAPERDOLLED PLAYER
        If PAPERDOLL = 1 Then
            x = NewX + sx
            y = NewY + sx
        
            'PLAYER 32 X 32 IF DIR = UP
            If GetPlayerDir(MyIndex) = DIR_UP Then
                
                'PLAYER 32 X 32 BLIT SHIELD IF DIR = UP
                If Player(MyIndex).Shield > 0 Then
                rec.Top = Item(Player(MyIndex).Shield).Pic * PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                Call DD_BackBuffer.BltFast(x, y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
            
                'PLAYER 32 X 32 BLIT WEAPON IF DIR = UP
                If Player(MyIndex).Weapon > 0 Then
                rec.Top = Item(Player(MyIndex).Weapon).Pic * PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                Call DD_BackBuffer.BltFast(x, y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
                
                'PLAYER 32 X 32 BLIT NECKLACE IF DIR = UP
                If Player(MyIndex).Necklace > 0 Then
                rec.Top = Item(Player(MyIndex).Necklace).Pic * PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                Call DD_BackBuffer.BltFast(x, y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
            
            End If
                
                'PLAYER 32 X 32 BLIT SPRITE
                rec.Top = GetPlayerSprite(MyIndex) * PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                Call DD_BackBuffer.BltFast(x, y, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                
                'PLAYER 32 X 32 BLIT ARMOR
                If Player(MyIndex).Armor > 0 Then
                rec.Top = Item(Player(MyIndex).Armor).Pic * PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                Call DD_BackBuffer.BltFast(x, y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
                
                'PLAYER 32 X 32 BLIT HELMET
                If Player(MyIndex).Helmet > 0 Then
                rec.Top = Item(Player(MyIndex).Helmet).Pic * PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                Call DD_BackBuffer.BltFast(x, y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
                
                'PLAYER 32 X 32 BLIT LEGS
                If Player(MyIndex).Legs > 0 Then
                rec.Top = Item(Player(MyIndex).Legs).Pic * PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                Call DD_BackBuffer.BltFast(x, y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
                
            'PLAYER 32 X 32 IF DIR <> UP
            If GetPlayerDir(MyIndex) <> DIR_UP Then
            
                'PLAYER 32 X 32 BLIT SHIELD IF DIR <> UP
                If Player(MyIndex).Shield > 0 Then
                rec.Top = Item(Player(MyIndex).Shield).Pic * PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                Call DD_BackBuffer.BltFast(x, y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
                
                'PLAYER 32 X 32 BLIT WEAPON IF DIR <> UP
                If Player(MyIndex).Weapon > 0 Then
                rec.Top = Item(Player(MyIndex).Weapon).Pic * PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                Call DD_BackBuffer.BltFast(x, y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
                
                'PLAYER 32 X 32 BLIT NECKLACE IF DIR <> UP
                If Player(MyIndex).Necklace > 0 Then
                rec.Top = Item(Player(MyIndex).Necklace).Pic * PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                Call DD_BackBuffer.BltFast(x, y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
            
            End If
    
            '32 X 32 IF NO PAPERDOLL ON SELF BLIT JUST SPRITE
            Else
            x = NewX + sx
            y = NewY + sx
            Call DD_BackBuffer.BltFast(x, y, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            End If
    
    '32 X 32 OTHER LOOP
    Else
    
    x = GetPlayerX(Index) * PIC_X + sx + Player(Index).XOffset
    y = GetPlayerY(Index) * PIC_Y + sx + Player(Index).YOffset
            
        'IF OFF TOP EDGE ADJUST
        If y < 0 Then
        rec.Top = rec.Top + (y * -1)
        y = 0
        End If
            
            '32 X 32 OTHER PAPERDOLL LOOP
            If PAPERDOLL = 1 Then
            
            'OTHER 32 X 32 IF DIR = UP
            If GetPlayerDir(Index) = DIR_UP Then
                
                'OTHER 32 X 32 BLIT SHIELD IF DIR = UP
                If Player(Index).Shield > 0 Then
                rec.Top = Item(Player(Index).Shield).Pic * PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
                
                'OTHER 32 X 32 BLIT WEAPON IF DIR = UP
                If Player(Index).Weapon > 0 Then
                rec.Top = Item(Player(Index).Weapon).Pic * PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
                
                'OTHER 32 X 32 BLIT NECKLACE IF DIR = UP
                If Player(Index).Necklace > 0 Then
                rec.Top = Item(Player(Index).Necklace).Pic * PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
                
            End If
                
                'OTHER 32 X 32 BLIT SPRITE
                rec.Top = GetPlayerSprite(Index) * PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                
                'OTHER 32 X 32 BLIT ARMOR
                'Call AddText(Player(2).Armor, Red)
                If Player(Index).Armor > 0 Then
                rec.Top = Item(Player(Index).Armor).Pic * PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
        
                'OTHER 32 X 32 BLIT HELMET
                If Player(Index).Helmet > 0 Then
                rec.Top = Item(Player(Index).Helmet).Pic * PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
                
                'OTHER 32 X 32 BLIT LEGS
                If Player(Index).Legs > 0 Then
                rec.Top = Item(Player(Index).Legs).Pic * PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
    
            'OTHER 32 X 32 IF DIR <> UP
            If GetPlayerDir(Index) <> DIR_UP Then
                
                'OTHER 32 X 32 BLIT SHIELD IF DIR <> UP
                If Player(Index).Shield > 0 Then
                rec.Top = Item(Player(Index).Shield).Pic * PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
                
                'OTHER 32 X 32 BLIT WEAPON IF DIR <> UP
                If Player(Index).Weapon > 0 Then
                rec.Top = Item(Player(Index).Weapon).Pic * PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
                   
                'OTHER 32 X 32 BLIT NECKLACE IF DIR <> UP
                If Player(Index).Necklace > 0 Then
                rec.Top = Item(Player(Index).Necklace).Pic * PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
                
            End If
            
        'OTHER 32 X 32 NON PAPERDOLL
        Else
    
        'OTHER 32 X 32 BLIT NON-PD SPRITE
        Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    
End If

End If

End Sub
Sub BltPlayerTop(ByVal Index As Long)
Dim Anim As Byte
Dim x As Long, y As Long
Dim AttackSpeed As Long

   If GetPlayerWeaponSlot(Index) > 0 Then
       AttackSpeed = Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).AttackSpeed
   Else
       AttackSpeed = 1000
   End If




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
       If Player(Index).AttackTimer + Int(AttackSpeed / 2) > GetTickCount Then
           Anim = 2
       End If
   End If
   
   
   
   
   
   ' Check to see if we want to stop making him attack
   If Player(Index).AttackTimer + AttackSpeed < GetTickCount Then
       Player(Index).Attacking = 0
       Player(Index).AttackTimer = 0
   End If
   
   
   
   
   
   
If PAPERDOLL = 1 Then
   rec.Left = (GetPlayerDir(Index) * 3 + Anim) * PIC_X
   rec.Right = rec.Left + PIC_X

   If Index = MyIndex Then
   x = NewX + sx
   y = NewY + sx - 32
   
   If GetPlayerDir(Index) = DIR_UP Then
       If Player(MyIndex).Shield > 0 Then
           rec.Top = Item(Player(MyIndex).Shield).Pic * 64
           rec.Bottom = rec.Top + PIC_Y
           Call DD_BackBuffer.BltFast(x, y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
       End If
       If Player(MyIndex).Weapon > 0 Then
           rec.Top = Item(Player(MyIndex).Weapon).Pic * 64
           rec.Bottom = rec.Top + PIC_Y
           Call DD_BackBuffer.BltFast(x, y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
       End If
       If Player(MyIndex).Necklace > 0 Then
           rec.Top = Item(Player(MyIndex).Necklace).Pic * 64
           rec.Bottom = rec.Top + PIC_Y
           Call DD_BackBuffer.BltFast(x, y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
       End If
   End If
       

 
   rec.Top = GetPlayerSprite(Index) * 64
   rec.Bottom = rec.Top + PIC_Y
   
    If y < 0 Then
           rec.Top = rec.Top + (y * -1)
           y = 0
    End If
   
   Call DD_BackBuffer.BltFast(x, y, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
   
   If Player(MyIndex).Armor > 0 Then
       rec.Top = Item(Player(MyIndex).Armor).Pic * 64
       rec.Bottom = rec.Top + PIC_Y
       Call DD_BackBuffer.BltFast(x, y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
   End If
   
   If Player(MyIndex).Legs > 0 Then
       rec.Top = Item(Player(MyIndex).Legs).Pic * 64
       rec.Bottom = rec.Top + PIC_Y
       Call DD_BackBuffer.BltFast(x, y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
   End If

If Player(MyIndex).Helmet > 0 Then
       rec.Top = Item(Player(MyIndex).Helmet).Pic * 64
       rec.Bottom = rec.Top + PIC_Y
       Call DD_BackBuffer.BltFast(x, y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
   End If
   If GetPlayerDir(Index) <> DIR_UP Then
       If Player(MyIndex).Shield > 0 Then
           rec.Top = Item(Player(MyIndex).Shield).Pic * 64
           rec.Bottom = rec.Top + PIC_Y
           Call DD_BackBuffer.BltFast(x, y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
       End If
       If Player(MyIndex).Necklace > 0 Then
           rec.Top = Item(Player(MyIndex).Necklace).Pic * 64
           rec.Bottom = rec.Top + PIC_Y
           Call DD_BackBuffer.BltFast(x, y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
       End If
       If Player(MyIndex).Weapon > 0 Then
           rec.Top = Item(Player(MyIndex).Weapon).Pic * 64
           rec.Bottom = rec.Top + PIC_Y
           Call DD_BackBuffer.BltFast(x, y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
       End If
   End If

   
Else
   x = GetPlayerX(Index) * PIC_X + sx + Player(Index).XOffset
   y = GetPlayerY(Index) * PIC_Y + sx + Player(Index).YOffset - 32
   
    If y < 0 Then
           rec.Top = rec.Top + (y * -1)
           y = 0
    End If
   
   If GetPlayerDir(Index) = DIR_UP Then
       If Player(Index).Shield > 0 Then
           rec.Top = Item(Player(Index).Shield).Pic * 64
           rec.Bottom = rec.Top + PIC_Y
           Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
       End If
       If Player(Index).Necklace > 0 Then
           rec.Top = Item(Player(Index).Necklace).Pic * 64
           rec.Bottom = rec.Top + PIC_Y
           Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
       End If
       If Player(Index).Weapon > 0 Then
           rec.Top = Item(Player(Index).Weapon).Pic * 64
           rec.Bottom = rec.Top + PIC_Y
           Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
       End If
   End If
   
 


rec.Top = GetPlayerSprite(Index) * 64
   rec.Bottom = rec.Top + PIC_Y
   
    If y < 0 Then
           rec.Top = rec.Top + (y * -1)
           y = 0
    End If
   
   Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
   
   If Player(Index).Armor > 0 Then
       rec.Top = Item(Player(Index).Armor).Pic * 64
       rec.Bottom = rec.Top + PIC_Y
       Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
   End If

   If Player(Index).Legs > 0 Then
       rec.Top = Item(Player(Index).Legs).Pic * 64
       rec.Bottom = rec.Top + PIC_Y
       Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
   End If

   If Player(Index).Helmet > 0 Then
       rec.Top = Item(Player(Index).Helmet).Pic * 64
       rec.Bottom = rec.Top + PIC_Y
       Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
   End If
   If GetPlayerDir(Index) <> DIR_UP Then
       If Player(Index).Shield > 0 Then
           rec.Top = Item(Player(Index).Shield).Pic * 64
           rec.Bottom = rec.Top + PIC_Y
           Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
       End If
       If Player(Index).Necklace > 0 Then
           rec.Top = Item(Player(Index).Necklace).Pic * 64
           rec.Bottom = rec.Top + PIC_Y
           Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
       End If
       If Player(Index).Weapon > 0 Then
           rec.Top = Item(Player(Index).Weapon).Pic * 64
           rec.Bottom = rec.Top + PIC_Y
           Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
       End If
   End If
End If
Else
   rec.Top = GetPlayerSprite(Index) * 64
   rec.Bottom = rec.Top + PIC_Y
   rec.Left = (GetPlayerDir(Index) * 3 + Anim) * PIC_X
   rec.Right = rec.Left + PIC_X

If Index = MyIndex Then
   x = NewX + sx
   y = NewY + sx - 32
   
Else
   x = GetPlayerX(Index) * PIC_X + sx + Player(Index).XOffset
   y = GetPlayerY(Index) * PIC_Y + sx + Player(Index).YOffset - 32
   
    If y < 0 Then
           rec.Top = rec.Top + (y * -1)
           y = 0
    End If

   
   Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
       End If
    rec.Top = GetPlayerSprite(Index) * 64
   rec.Bottom = rec.Top + PIC_Y
   Call DD_BackBuffer.BltFast(x, y, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)

End If
End Sub
Sub BltMapNPCName(ByVal Index As Long)
Dim TextX As Long
Dim TextY As Long

If Npc(MapNpc(Index).num).Big = 0 And SPRITESIZE = 0 Then
    With Npc(MapNpc(Index).num)
    'Draw name
        TextX = MapNpc(Index).x * PIC_X + sx + MapNpc(Index).XOffset + CLng(PIC_X / 2) - ((Len(Trim$(.Name)) / 2) * 8)
        TextY = MapNpc(Index).y * PIC_Y + sx + MapNpc(Index).YOffset - CLng(PIC_Y / 2) - 4
        DrawPlayerNameText TexthDC, TextX - (NewPlayerX * PIC_X) - NewXOffset, TextY - (NewPlayerY * PIC_Y) - NewYOffset, Trim$(.Name), vbWhite
    End With
Else
    With Npc(MapNpc(Index).num)
    'Draw name
        TextX = MapNpc(Index).x * PIC_X + sx + MapNpc(Index).XOffset + CLng(PIC_X / 2) - ((Len(Trim$(.Name)) / 2) * 8)
        TextY = MapNpc(Index).y * PIC_Y + sx + MapNpc(Index).YOffset - CLng(PIC_Y / 2) - 32
        DrawPlayerNameText TexthDC, TextX - (NewPlayerX * PIC_X) - NewXOffset, TextY - (NewPlayerY * PIC_Y) - NewYOffset, Trim$(.Name), vbWhite
    End With
End If
End Sub

Sub BltNpc(ByVal MapNpcNum As Long)
Dim Anim As Byte
Dim x As Long
Dim y As Long
Dim modify As Long
    
    ' Only used if ever want to switch to blt rather then bltfast
    'With rec_pos
       ' .Top = MapNpc(MapNpcNum).Y * PIC_Y + MapNpc(MapNpcNum).YOffset
      '  .Bottom = .Top + PIC_Y
      '  .Left = MapNpc(MapNpcNum).X * PIC_X + MapNpc(MapNpcNum).XOffset
       ' .Right = .Left + PIC_X
    'End With
    
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
        
    If Npc(MapNpc(MapNpcNum).num).Big = 1 Then
        rec.Top = Npc(MapNpc(MapNpcNum).num).Sprite * 64 + 32
        rec.Bottom = rec.Top + 32
        rec.Left = (MapNpc(MapNpcNum).Dir * 3 + Anim) * 64
        rec.Right = rec.Left + 64
    
        x = MapNpc(MapNpcNum).x * 32 + sx - 16 + MapNpc(MapNpcNum).XOffset
        y = MapNpc(MapNpcNum).y * 32 + sx + MapNpc(MapNpcNum).YOffset
   
    If y < 0 Then
        modify = -y
        rec.Top = rec.Top + modify
        rec.Bottom = rec.Top + 32
        y = 0
    End If
   
    If x < 0 Then
        'rec.Left = (MapNpc(MapNpcNum).Dir * 3 + Anim) * 64 + 16
        'modify = -X
        'rec.Left = rec.Left + modify - 16
        'rec.Right = rec.Left + 48
        'X = 0
        modify = -x
        rec.Left = rec.Left + modify
        rec.Right = rec.Left + 48
        x = 0
    End If
   
    If 32 + x >= (MAX_MAPX * 32) Then
        modify = x - (MAX_MAPX * 32)
        rec.Left = rec.Left + modify + 16
        rec.Right = rec.Left + 32 - modify
    End If

        Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_BigSpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    Else
    If SPRITESIZE = 1 Then
        rec.Top = Npc(MapNpc(MapNpcNum).num).Sprite * 64 + 32
        rec.Bottom = rec.Top + PIC_Y
        rec.Left = (MapNpc(MapNpcNum).Dir * 3 + Anim) * PIC_X
        rec.Right = rec.Left + PIC_X
        
        x = MapNpc(MapNpcNum).x * PIC_X + sx + MapNpc(MapNpcNum).XOffset
        y = MapNpc(MapNpcNum).y * PIC_Y + sx + MapNpc(MapNpcNum).YOffset
        
        ' Check if its out of bounds because of the offset
        
    If y < 0 Then
           rec.Top = rec.Top + (y * -1)
           y = 0
    End If
            
        'Call DD_BackBuffer.Blt(rec_pos, DD_SpriteSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
        Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
   Else
        rec.Top = Npc(MapNpc(MapNpcNum).num).Sprite * PIC_Y
        rec.Bottom = rec.Top + PIC_Y
        rec.Left = (MapNpc(MapNpcNum).Dir * 3 + Anim) * PIC_X
        rec.Right = rec.Left + PIC_X
        
        x = MapNpc(MapNpcNum).x * PIC_X + sx + MapNpc(MapNpcNum).XOffset
        y = MapNpc(MapNpcNum).y * PIC_Y + sx + MapNpc(MapNpcNum).YOffset
            
        'Call DD_BackBuffer.Blt(rec_pos, DD_SpriteSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
        Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    End If
End If
End Sub

Sub BltNpcTop(ByVal MapNpcNum As Long)
'Moles are cool!
'But a Baron can debug
Dim Anim As Byte
Dim x As Long
Dim y As Long
Dim NPC_number As Long
Dim modify As Long

    'Get the NPC number
    NPC_number = MapNpc(MapNpcNum).num

   ' Make sure that theres an npc there, and if not exit the sub
   If MapNpc(MapNpcNum).num <= 0 Then
       Exit Sub
   End If
   
    If Npc(NPC_number).Big = 0 Then
        If SPRITESIZE = 0 Then Exit Sub
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
     rec.Top = Npc(MapNpc(MapNpcNum).num).Sprite * 64
       rec.Bottom = rec.Top + PIC_Y
       rec.Left = (MapNpc(MapNpcNum).Dir * 3 + Anim) * PIC_X
       rec.Right = rec.Left + PIC_X
       
       x = MapNpc(MapNpcNum).x * PIC_X + sx + MapNpc(MapNpcNum).XOffset
       y = MapNpc(MapNpcNum).y * PIC_Y + sx + MapNpc(MapNpcNum).YOffset - 32
       
       ' Check if its out of bounds because of the offset
       If y < 0 Then
           rec.Top = rec.Top + (y * -1)
           y = 0
       End If
           
       'Call DD_BackBuffer.Blt(rec_pos, DD_SpriteSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
       Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
 Else
   rec.Top = Npc(MapNpc(MapNpcNum).num).Sprite * PIC_Y
       
    rec.Top = Npc(MapNpc(MapNpcNum).num).Sprite * 64
    rec.Bottom = rec.Top + 32
    rec.Left = (MapNpc(MapNpcNum).Dir * 3 + Anim) * 64
    rec.Right = rec.Left + 64

    x = MapNpc(MapNpcNum).x * 32 + sx - 16 + MapNpc(MapNpcNum).XOffset
    y = MapNpc(MapNpcNum).y * 32 + sx - 32 + MapNpc(MapNpcNum).YOffset

    If y < 0 Then
        modify = -y
        rec.Top = rec.Top + modify
        rec.Bottom = rec.Top + 32
        y = 0
    End If
   
    If x < 0 Then
        'rec.Left = (MapNpc(MapNpcNum).Dir * 3 + Anim) * 64 + 16
        'modify = -X
        'rec.Left = rec.Left + modify - 16
        'rec.Right = rec.Left + 48
        'X = 0
        modify = -x
        rec.Left = rec.Left + modify
        rec.Right = rec.Left + 48
        x = 0
    End If
   
    If 32 + x >= (MAX_MAPX * 32) Then
        modify = x - (MAX_MAPX * 32)
        rec.Left = rec.Left + modify + 16
        rec.Right = rec.Left + 32 - modify
    End If

    Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_BigSpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End If
End Sub
Sub BltPlayerName(ByVal Index As Long)
Dim TextX As Long
Dim TextY As Long
Dim Color As Long
    
    ' Check access level
    If GetPlayerPK(Index) = NO Then
        Color = QBColor(Yellow)
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
        
If SPRITESIZE = 1 Then
  If Index = MyIndex Then
       TextX = NewX + sx + Int(PIC_X / 2) - ((Len(GetPlayerName(MyIndex)) / 2) * 8)
       TextY = NewY + sx - 50

   Call DrawText(TexthDC, TextX, TextY, GetPlayerName(MyIndex), Color)
Else
   ' Draw name
   TextX = GetPlayerX(Index) * PIC_X + sx + Player(Index).XOffset + Int(PIC_X / 2) - ((Len(GetPlayerName(Index)) / 2) * 8)
   TextY = GetPlayerY(Index) * PIC_Y + sx + Player(Index).YOffset - Int(PIC_Y / 2) - 32
   Call DrawText(TexthDC, TextX - (NewPlayerX * PIC_X) - NewXOffset, TextY - (NewPlayerY * PIC_Y) - NewYOffset, GetPlayerName(Index), Color)
End If
Else
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
If SPRITESIZE = 1 Then
TextY = NewY + sx - Int(PIC_Y / 4) - 52
Else
TextY = NewY + sx - Int(PIC_Y / 4) - 20
End If
Call DrawText(TexthDC, TextX, TextY, GetPlayerGuild(MyIndex), Color)
Else
' Draw name
TextX = GetPlayerX(Index) * PIC_X + sx + Player(Index).XOffset + Int(PIC_X / 2) - ((Len(GetPlayerGuild(Index)) / 2) * 8)
If SPRITESIZE = 1 Then
TextY = GetPlayerY(Index) * PIC_Y + sx + Player(Index).YOffset - Int(PIC_Y / 2) - 44
Else
TextY = GetPlayerY(Index) * PIC_Y + sx + Player(Index).YOffset - Int(PIC_Y / 2) - 12
End If
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

MyText = frmMirage.txtMyTextBox.Text

    ' Handle when the player presses the return key
    If (KeyAscii = vbKeyReturn) Then
    frmMirage.txtMyTextBox.Text = ""
        If Player(MyIndex).y - 1 > -1 Then
            If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).Type = TILE_TYPE_SIGN And Player(MyIndex).Dir = DIR_UP Then
                Call AddText("The Sign Reads:", Black)
                If Trim$(Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).String1) <> "" Then
                    Call AddText(Trim$(Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).String1), Grey)
                End If
                If Trim$(Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).String2) <> "" Then
                    Call AddText(Trim$(Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).String2), Grey)
                End If
                If Trim$(Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).String3) <> "" Then
                    Call AddText(Trim$(Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).String3), Grey)
                End If
            Exit Sub
            End If
        End If
        ' Broadcast message
        If Mid$(MyText, 1, 1) = "'" Then
            ChatText = Mid$(MyText, 2, Len(MyText) - 1)
            If Len(Trim$(ChatText)) > 0 Then
                Call BroadcastMsg(ChatText)
            End If
            MyText = ""
            Exit Sub
        End If
        
        ' Emote message
        If Mid$(MyText, 1, 1) = "-" Then
            ChatText = Mid$(MyText, 2, Len(MyText) - 1)
            If Len(Trim$(ChatText)) > 0 Then
                Call EmoteMsg(ChatText)
            End If
            MyText = ""
            Exit Sub
        End If
        
        ' Player message
        If Mid$(MyText, 1, 1) = "!" Then
            ChatText = Mid$(MyText, 2, Len(MyText) - 1)
            Name = ""
                    
            ' Get the desired player from the user text
            For i = 1 To Len(ChatText)
                If Mid$(ChatText, i, 1) <> " " Then
                    Name = Name & Mid$(ChatText, i, 1)
                Else
                    Exit For
                End If
            Next i
                    
            ' Make sure they are actually sending something
            If Len(ChatText) - i > 0 Then
                ChatText = Mid$(ChatText, i + 1, Len(ChatText) - i)
                    
                ' Send the message to the player
                Call PlayerMsg(ChatText, Name)
            Else
                Call AddText("Usage: !playername msghere", AlertColor)
            End If
            MyText = ""
            Exit Sub
        End If
            
        ' // Commands //
        ' Verification User
        If LCase(Mid(MyText, 1, 5)) = "/info" Then
            ChatText = Mid$(MyText, 6, Len(MyText) - 5)
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
                ChatText = Mid$(MyText, 8, Len(MyText) - 7)
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
                ChatText = Mid$(MyText, 8, Len(MyText) - 7)
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
        
                ' House Editor
        If LCase(Mid(MyText, 1, 12)) = "/houseeditor" Then
            Call SendRequestEditHouse
            MyText = ""
            Exit Sub
        End If
        
        ' // Moniter Admin Commands //
        If GetPlayerAccess(MyIndex) > 0 Then
            ' weather command
            If LCase(Mid(MyText, 1, 8)) = "/weather" Then
                If Len(MyText) > 8 Then
                    MyText = Mid$(MyText, 9, Len(MyText) - 8)
                    If IsNumeric(MyText) = True Then
                        Call SendData("weather" & SEP_CHAR & Val#(MyText) & SEP_CHAR & END_CHAR)
                    Else
                        If Trim$(LCase(MyText)) = "none" Then i = 0
                        If Trim$(LCase(MyText)) = "rain" Then i = 1
                        If Trim$(LCase(MyText)) = "snow" Then i = 2
                        If Trim$(LCase(MyText)) = "thunder" Then i = 3
                        Call SendData("weather" & SEP_CHAR & i & SEP_CHAR & END_CHAR)
                    End If
                End If
                MyText = ""
                Exit Sub
            End If


            ' Clearing a house owner
            If LCase(Mid(MyText, 1, 11)) = "/clearowner" Then
                Call SendData("clearowner" & SEP_CHAR & END_CHAR)
                MyText = ""
                Exit Sub
            End If
            
            ' Kicking a player
            If LCase(Mid(MyText, 1, 5)) = "/kick" Then
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
                    MyText = Mid$(MyText, 12, Len(MyText) - 11)
                
                    Call SendSetSprite(Val(MyText))
                End If
                MyText = ""
                Exit Sub
            End If
            
            ' Setting player sprite
            If LCase(Mid(MyText, 1, 16)) = "/setplayersprite" Then
                If Len(MyText) > 19 Then
                    i = Val#(Mid(MyText, 17, 1))
                
                    MyText = Mid$(MyText, 18, Len(MyText) - 17)
                    Call SendSetPlayerSprite(i, Val#(MyText))
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
            If Mid$(MyText, 1, 3) = "/banlist" Then
                Call SendBanList
                MyText = ""
                Exit Sub
            End If
            
            ' Banning a player
            If LCase(Mid(MyText, 1, 4)) = "/ban" Then
                If Len(MyText) > 5 Then
                    MyText = Mid$(MyText, 6, Len(MyText) - 5)
                    Call SendBan(MyText)
                    MyText = ""
                End If
                Exit Sub
            End If
        End If
            
        ' // Developer Admin Commands //
        If GetPlayerAccess(MyIndex) >= ADMIN_DEVELOPER Then
            ' Editing item request
            If Mid$(MyText, 1, 9) = "/edititem" Then
                Call SendRequestEditItem
                MyText = ""
                Exit Sub
            End If
            
            ' Day/Night
            If LCase(Mid(MyText, 1, 9)) = "/daynight" Then
                Call SendData("daynight" & SEP_CHAR & END_CHAR)
                MyText = ""
                Exit Sub
            End If
            
            ' Editing emoticon request
            If Mid$(MyText, 1, 13) = "/editemoticon" Then
                Call SendRequestEditEmoticon
                MyText = ""
                Exit Sub
            End If
            
            ' Editing emoticon request
            If Mid$(MyText, 1, 12) = "/editelement" Then
                Call SendRequestEditElement
                MyText = ""
                Exit Sub
            End If
            
            ' Editing arrow request
            If Mid$(MyText, 1, 13) = "/editarrow" Then
                Call SendRequestEditArrow
                MyText = ""
                Exit Sub
            End If
            
            ' Editing npc request
            If Mid$(MyText, 1, 8) = "/editnpc" Then
                Call SendRequestEditNpc
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
            If LCase(Trim$(MyText)) = "/editspell" Then
            'If mid$(MyText, 1, 10) = "/editspell" Then
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
                i = Val#(Mid(MyText, 12, 1))
                
                MyText = Mid$(MyText, 14, Len(MyText) - 13)
                
                Call SendSetAccess(MyText, i)
                MyText = ""
                Exit Sub
            End If
            
            'Reload Scripts
            If LCase(Trim$(MyText)) = "/reload" Then
                Call SendReloadScripts
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
    
    'frmMirage.txtMyTextBox.SetFocus
    
    ' Handle when the user presses the backspace key
    If (KeyAscii = vbKeyBack) Then
        If Len(MyText) > 0 Then
            'MyText = mid$(MyText, 1, Len(MyText) - 1)
        End If
    End If
    
    ' And if neither, then add the character to the user's text buffer
    If (KeyAscii <> vbKeyReturn) And (KeyAscii <> vbKeyBack) Then
        ' Make sure they just use standard keys, no gay shitty macro keys
        If KeyAscii >= 32 And KeyAscii <= 255 Then
            'frmMirage.txtMyTextBox.Text = frmMirage.txtMyTextBox.Text & Chr(KeyAscii)
            'MyText = MyText & Chr(KeyAscii)
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
                Call SendPlayerRequestNewMap(0)
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
                Call SendPlayerRequestNewMap(0)
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
                Call SendPlayerRequestNewMap(0)
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
                Call SendPlayerRequestNewMap(0)
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

Function FindPlayer(ByVal Name As String) As Long
Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            ' Make sure we dont try to check a name thats to small
            If Len(GetPlayerName(i)) >= Len(Trim$(Name)) Then
                If UCase(Mid(GetPlayerName(i), 1, Len(Trim$(Name)))) = UCase(Trim$(Name)) Then
                    FindPlayer = i
                    Exit Function
                End If
            End If
        End If
    Next i
    
    FindPlayer = 0
End Function

Public Sub EditorInit()
Dim BMU As BitmapUtils
Dim strfilename As String
Dim sRECT As RECT
Dim dRECT As RECT
    Dim i As Long

    InEditor = True
    frmAttributes.Show vbModeless, frmMirage
    frmMapEditor.Show vbModeless, frmMirage
    EditorSet = 0

    For i = 0 To 10
        If frmMapEditor.mnuSet(i).Checked = True Then
        
        If ENCRYPT_TYPE = "BMP" Then
                frmMapEditor.picBackSelect.Picture = LoadPicture(App.Path & "\GFX\Tiles" & i & ".bmp")
                Else
                Set BMU = New BitmapUtils
                strfilename = App.Path & "/gfx/" & "tiles" & i & "." & Trim$(ENCRYPT_TYPE)
                BMU.LoadByteData (strfilename)
                BMU.DecryptByteData (Trim$(ENCRYPT_PASS))
                BMU.DecompressByteData_ZLib
                frmMapEditor.picBackSelect.Cls
                frmMapEditor.picBackSelect.Width = BMU.ImageWidth
                frmMapEditor.picBackSelect.Height = BMU.ImageHeight
                Call BMU.Blt(frmMapEditor.picBackSelect.hDC)
                End If
        
            EditorSet = i
        End If
    Next i
    
    'MsgBox "Success!"
    frmMapEditor.scrlPicture.Max = Int((frmMapEditor.picBackSelect.Height - frmMapEditor.picBack.Height) / PIC_Y)
    'MsgBox frmMapEditor.picBackSelect.Height & " " & frmMapEditor.picBack.Height & " " & PIC_Y
    frmMapEditor.picBack.Width = frmMapEditor.picBackSelect.Width
    frmMapEditor.Width = frmMapEditor.picBackSelect.Width + frmMapEditor.scrlPicture.Width
End Sub
Public Sub HouseEditorInit()
Dim BMU As BitmapUtils
Dim strfilename As String
    Dim i As Long

    InHouseEditor = True
    frmAttributes.Show vbModeless, frmMirage
    frmHouseEditor.Show vbModeless, frmMirage
    EditorSet = 0

    For i = 0 To 10
        If frmHouseEditor.mnuSet(i).Checked = True Then
        
        If ENCRYPT_TYPE = "BMP" Then
            frmHouseEditor.picBackSelect.Picture = LoadPicture(App.Path & "\GFX\Tiles" & i & ".bmp")
        Else
                Set BMU = New BitmapUtils
                strfilename = App.Path & "/gfx/" & "tiles" & i & "." & Trim$(ENCRYPT_TYPE)
                BMU.LoadByteData (strfilename)
                BMU.DecryptByteData (Trim$(ENCRYPT_PASS))
                BMU.DecompressByteData_ZLib
                frmHouseEditor.picBackSelect.Cls
                frmHouseEditor.picBackSelect.Width = BMU.ImageWidth
                frmHouseEditor.picBackSelect.Height = BMU.ImageHeight
                Call BMU.Blt(frmHouseEditor.picBackSelect.hDC)
        End If
        
            EditorSet = i
        End If
    Next i
    
    frmHouseEditor.scrlPicture.Max = Int((frmHouseEditor.picBackSelect.Height - frmHouseEditor.picBack.Height) / PIC_Y)
    frmHouseEditor.picBack.Width = frmHouseEditor.picBackSelect.Width
End Sub
Public Sub MainMenuInit()
    Dim Stuff As String
    Dim Stuff2 As String
    Dim Stuff3 As String
    Dim ThisIsANumber As Long

    Stuff2 = ""
        Stuff = ReadINI("DATA", "Desc", App.Path & "\News.ini")
        For ThisIsANumber = 1 To Len(Stuff)
           If Mid$(Stuff, ThisIsANumber, 1) = "*" Then
                Stuff2 = Stuff2 & vbCrLf
           Else
                Stuff2 = Stuff2 & Mid$(Stuff, ThisIsANumber, 1)
           End If
        Next
        Stuff3 = ""
        Stuff = ReadINI("DATA", "News", App.Path & "\News.ini")
        For ThisIsANumber = 1 To Len(Stuff)
           If Mid$(Stuff, ThisIsANumber, 1) = "*" Then
                Stuff3 = Stuff3 & vbCrLf
           Else
                Stuff3 = Stuff3 & Mid$(Stuff, ThisIsANumber, 1)
           End If
        Next
        frmMainMenu.News.Text = Stuff3 & Stuff2
    
    frmLogin.txtName.Text = Trim$(ReadINI("CONFIG", "Account", App.Path & "\config.ini"))
    frmLogin.txtPassword.Text = Trim$(ReadINI("CONFIG", "Password", App.Path & "\config.ini"))
    
    If frmLogin.Check1.Value = 0 Then
        frmLogin.Check2.Value = 0
    End If
    
    If ConnectToServer = True And AutoLogin = 1 Then
        frmMainMenu.Label1.Visible = True
        frmChars.Label1.Visible = True
    Else
        frmMainMenu.Label1.Visible = False
        frmChars.Label1.Visible = False
    End If
End Sub

Public Sub NewsUpdate(News As String)

    Dim Stuff As String
    Dim Stuff2 As String
    Dim Stuff3 As String
    Dim ThisIsANumber As Long

Stuff = News
    'Stuff2 = ""
        'Stuff = ReadINI("DATA", "Desc", App.Path & "\News.ini")
     '   For ThisIsANumber = 1 To Len(news)
      '     If Mid$(news, ThisIsANumber, 1) = "*" Then
       '         Stuff2 = Stuff2 & vbCrLf
        '   Else
         '       Stuff2 = Stuff2 & Mid$(Stuff, ThisIsANumber, 1)
          ' End If
        'Next
        'Stuff3 = ""
        'Stuff = ReadINI("DATA", "News", App.Path & "\News.ini")
        
        
        For ThisIsANumber = 1 To Len(Stuff)
           If Mid$(Stuff, ThisIsANumber, 1) = "*" Then
                Stuff3 = Stuff3 & vbCrLf
           Else
                Stuff3 = Stuff3 & Mid$(Stuff, ThisIsANumber, 1)
           End If
        Next
        
        frmMainMenu.News.Text = Stuff3

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
                    If frmAttributes.optGround.Value = True Then
                        PicX = .Ground
                        EditorSet = .GroundSet
                    End If
                    If frmAttributes.optMask.Value = True Then
                        PicX = .Mask
                        EditorSet = .MaskSet
                    End If
                    If frmAttributes.optAnim.Value = True Then
                        PicX = .Anim
                        EditorSet = .AnimSet
                    End If
                    If frmAttributes.optMask2.Value = True Then
                        PicX = .Mask2
                        EditorSet = .Mask2Set
                    End If
                    If frmAttributes.optM2Anim.Value = True Then
                        PicX = .M2Anim
                        EditorSet = .M2AnimSet
                    End If
                    If frmAttributes.optFringe.Value = True Then
                        PicX = .Fringe
                        EditorSet = .FringeSet
                    End If
                    If frmAttributes.optFAnim.Value = True Then
                        PicX = .FAnim
                        EditorSet = .FAnimSet
                    End If
                    If frmAttributes.optFringe2.Value = True Then
                        PicX = .Fringe2
                        EditorSet = .Fringe2Set
                    End If
                    If frmAttributes.optF2Anim.Value = True Then
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
                    If .Type = TILE_TYPE_BLOCKED Then frmAttributes.optBlocked.Value = True
                    If .Type = TILE_TYPE_WARP Then
                        EditorWarpMap = .Data1
                        EditorWarpX = .Data2
                        EditorWarpY = .Data3
                        frmAttributes.optWarp.Value = True
                    End If
                    If .Type = TILE_TYPE_HEAL Then frmAttributes.optHeal.Value = True
                    If .Type = TILE_TYPE_KILL Then frmAttributes.optKill.Value = True
                    If .Type = TILE_TYPE_ITEM Then
                        ItemEditorNum = .Data1
                        ItemEditorValue = .Data2
                        frmAttributes.optItem.Value = True
                    End If
                    If .Type = TILE_TYPE_NPCAVOID Then frmAttributes.optNpcAvoid.Value = True
                    If .Type = TILE_TYPE_KEY Then
                        KeyEditorNum = .Data1
                        KeyEditorTake = .Data2
                        frmAttributes.optKey.Value = True
                    End If
                    If .Type = TILE_TYPE_KEYOPEN Then
                        KeyOpenEditorX = .Data1
                        KeyOpenEditorY = .Data2
                        KeyOpenEditorMsg = .String1
                        frmAttributes.optKeyOpen.Value = True
                    End If
                    If .Type = TILE_TYPE_SHOP Then
                        EditorShopNum = .Data1
                        frmAttributes.optShop.Value = True
                    End If
                    If .Type = TILE_TYPE_CBLOCK Then
                        EditorItemNum1 = .Data1
                        EditorItemNum2 = .Data2
                        EditorItemNum3 = .Data3
                        frmAttributes.optCBlock.Value = True
                    End If
                    If .Type = TILE_TYPE_ARENA Then
                        Arena1 = .Data1
                        Arena2 = .Data2
                        Arena3 = .Data3
                        frmAttributes.optArena.Value = True
                    End If
                    If .Type = TILE_TYPE_SOUND Then
                        SoundFileName = .String1
                        frmAttributes.optSound.Value = True
                    End If
                    If .Type = TILE_TYPE_SPRITE_CHANGE Then
                        SpritePic = .Data1
                        SpriteItem = .Data2
                        SpritePrice = .Data3
                        frmAttributes.optSprite.Value = True
                    End If
                    If .Type = TILE_TYPE_SIGN Then
                        SignLine1 = .String1
                        SignLine2 = .String2
                        SignLine3 = .String3
                        frmAttributes.optSign.Value = True
                    End If
                    If .Type = TILE_TYPE_DOOR Then frmAttributes.optDoor.Value = True
                    If .Type = TILE_TYPE_NOTICE Then
                        NoticeTitle = .String1
                        NoticeText = .String2
                        NoticeSound = .String3
                        frmAttributes.optNotice.Value = True
                    End If
                    If .Type = TILE_TYPE_CHEST Then frmAttributes.optChest.Value = True
                    If .Type = TILE_TYPE_CLASS_CHANGE Then
                        ClassChange = .Data1
                        ClassChangeReq = .Data2
                        frmAttributes.optClassChange.Value = True
                    End If
                    If .Type = TILE_TYPE_SCRIPTED Then
                        ScriptNum = .Data1
                        frmAttributes.optScripted.Value = True
                    End If
                    If .Type = TILE_TYPE_NPC_SPAWN Then
                        NPCSpawnNum = .Data1
                        NPCSpawnAmount = .Data2
                        NPCSpawnRange = .Data3
                        frmAttributes.optNPC.Value = True
                    End If
                     If .Type = TILE_TYPE_HOUSE Then
                        HouseItem = .Data1
                        HousePrice = .Data2
                        frmAttributes.optHouse.Value = True
                    End If
                    
                    If .Type = TILE_TYPE_BANK Then frmAttributes.optBank.Value = True
                End With
            End If
            frmMapEditor.MousePointer = 1
            frmMirage.MousePointer = 1
        Else
            If (Button = 1) And (x1 >= 0) And (x1 <= MAX_MAPX) And (y1 >= 0) And (y1 <= MAX_MAPY) Then
                If frmMapEditor.shpSelected.Height <= PIC_Y And frmMapEditor.shpSelected.Width <= PIC_X Then
                    If frmMapEditor.mnuType(1).Checked = True Then
                        With Map(GetPlayerMap(MyIndex)).Tile(x1, y1)
                            If frmAttributes.optGround.Value = True Then
                                .Ground = EditorTileY * TilesInSheets + EditorTileX
                                .GroundSet = EditorSet
                            End If
                            If frmAttributes.optMask.Value = True Then
                                .Mask = EditorTileY * TilesInSheets + EditorTileX
                                .MaskSet = EditorSet
                            End If
                            If frmAttributes.optAnim.Value = True Then
                                .Anim = EditorTileY * TilesInSheets + EditorTileX
                                .AnimSet = EditorSet
                            End If
                            If frmAttributes.optMask2.Value = True Then
                                .Mask2 = EditorTileY * TilesInSheets + EditorTileX
                                .Mask2Set = EditorSet
                            End If
                            If frmAttributes.optM2Anim.Value = True Then
                                .M2Anim = EditorTileY * TilesInSheets + EditorTileX
                                .M2AnimSet = EditorSet
                            End If
                            If frmAttributes.optFringe.Value = True Then
                                .Fringe = EditorTileY * TilesInSheets + EditorTileX
                                .FringeSet = EditorSet
                            End If
                            If frmAttributes.optFAnim.Value = True Then
                                .FAnim = EditorTileY * TilesInSheets + EditorTileX
                                .FAnimSet = EditorSet
                            End If
                            If frmAttributes.optFringe2.Value = True Then
                                .Fringe2 = EditorTileY * TilesInSheets + EditorTileX
                                .Fringe2Set = EditorSet
                            End If
                            If frmAttributes.optF2Anim.Value = True Then
                                .F2Anim = EditorTileY * TilesInSheets + EditorTileX
                                .F2AnimSet = EditorSet
                            End If
                        End With
                    ElseIf frmMapEditor.mnuType(3).Checked = True Then
                        Map(GetPlayerMap(MyIndex)).Tile(x1, y1).Light = EditorTileY * TilesInSheets + EditorTileX
                    ElseIf frmMapEditor.mnuType(2).Checked = True Then
                        With Map(GetPlayerMap(MyIndex)).Tile(x1, y1)
                            If frmAttributes.optBlocked.Value = True Then .Type = TILE_TYPE_BLOCKED
                            If frmAttributes.optWarp.Value = True Then
                                .Type = TILE_TYPE_WARP
                                .Data1 = EditorWarpMap
                                .Data2 = EditorWarpX
                                .Data3 = EditorWarpY
                                .String1 = ""
                                .String2 = ""
                                .String3 = ""
                            End If
        
                            If frmAttributes.optHeal.Value = True Then
                                .Type = TILE_TYPE_HEAL
                                .Data1 = 0
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = ""
                                .String2 = ""
                                .String3 = ""
                            End If
        
                            If frmAttributes.optKill.Value = True Then
                                .Type = TILE_TYPE_KILL
                                .Data1 = 0
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = ""
                                .String2 = ""
                                .String3 = ""
                            End If
                            If GetPlayerAccess(MyIndex) >= 3 Then
                            If frmAttributes.optItem.Value = True Then
                                .Type = TILE_TYPE_ITEM
                                .Data1 = ItemEditorNum
                                .Data2 = ItemEditorValue
                                .Data3 = 0
                                .String1 = ""
                                .String2 = ""
                                .String3 = ""
                            End If
                            End If
                            If frmAttributes.optNpcAvoid.Value = True Then
                                .Type = TILE_TYPE_NPCAVOID
                                .Data1 = 0
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = ""
                                .String2 = ""
                                .String3 = ""
                            End If
                            If frmAttributes.optKey.Value = True Then
                                .Type = TILE_TYPE_KEY
                                .Data1 = KeyEditorNum
                                .Data2 = KeyEditorTake
                                .Data3 = 0
                                .String1 = ""
                                .String2 = ""
                                .String3 = ""
                            End If
                            If frmAttributes.optKeyOpen.Value = True Then
                                .Type = TILE_TYPE_KEYOPEN
                                .Data1 = KeyOpenEditorX
                                .Data2 = KeyOpenEditorY
                                .Data3 = 0
                                .String1 = KeyOpenEditorMsg
                                .String2 = ""
                                .String3 = ""
                            End If
                            If frmAttributes.optShop.Value = True Then
                                .Type = TILE_TYPE_SHOP
                                .Data1 = EditorShopNum
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = ""
                                .String2 = ""
                                .String3 = ""
                            End If
                            If frmAttributes.optCBlock.Value = True Then
                                .Type = TILE_TYPE_CBLOCK
                                .Data1 = EditorItemNum1
                                .Data2 = EditorItemNum2
                                .Data3 = EditorItemNum3
                                .String1 = ""
                                .String2 = ""
                                .String3 = ""
                            End If
                            If frmAttributes.optArena.Value = True Then
                                .Type = TILE_TYPE_ARENA
                                .Data1 = Arena1
                                .Data2 = Arena2
                                .Data3 = Arena3
                                .String1 = ""
                                .String2 = ""
                                .String3 = ""
                            End If
                            If frmAttributes.optSound.Value = True Then
                                .Type = TILE_TYPE_SOUND
                                .Data1 = 0
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = SoundFileName
                                .String2 = ""
                                .String3 = ""
                            End If
                            If frmAttributes.optSprite.Value = True Then
                                .Type = TILE_TYPE_SPRITE_CHANGE
                                .Data1 = SpritePic
                                .Data2 = SpriteItem
                                .Data3 = SpritePrice
                                .String1 = ""
                                .String2 = ""
                                .String3 = ""
                            End If
                            If frmAttributes.optSign.Value = True Then
                                .Type = TILE_TYPE_SIGN
                                .Data1 = 0
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = SignLine1
                                .String2 = SignLine2
                                .String3 = SignLine3
                            End If
                            If frmAttributes.optDoor.Value = True Then
                                .Type = TILE_TYPE_DOOR
                                .Data1 = 0
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = ""
                                .String2 = ""
                                .String3 = ""
                            End If
                            If frmAttributes.optNotice.Value = True Then
                                .Type = TILE_TYPE_NOTICE
                                .Data1 = 0
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = NoticeTitle
                                .String2 = NoticeText
                                .String3 = NoticeSound
                            End If
                            If frmAttributes.optChest.Value = True Then
                                .Type = TILE_TYPE_CHEST
                                .Data1 = 0
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = ""
                                .String2 = ""
                                .String3 = ""
                            End If
                            If GetPlayerAccess(MyIndex) >= 3 Then
                            If frmAttributes.optClassChange.Value = True Then
                                .Type = TILE_TYPE_CLASS_CHANGE
                                .Data1 = ClassChange
                                .Data2 = ClassChangeReq
                                .Data3 = 0
                                .String1 = ""
                                .String2 = ""
                                .String3 = ""
                            End If
                            End If
                            If frmAttributes.optScripted.Value = True Then
                                .Type = TILE_TYPE_SCRIPTED
                                .Data1 = ScriptNum
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = ""
                                .String2 = ""
                                .String3 = ""
                            End If
                            If frmAttributes.optNPC.Value = True Then
                                .Type = TILE_TYPE_NPC_SPAWN
                                .Data1 = NPCSpawnNum
                                .Data2 = NPCSpawnAmount
                                .Data3 = NPCSpawnRange
                                .String1 = ""
                                .String2 = ""
                                .String3 = ""
                            End If
                            
                  If frmAttributes.optHouse.Value = True Then
                    .Type = TILE_TYPE_HOUSE
                    .Data1 = HouseItem
                    .Data2 = HousePrice
                    .Data3 = 0
                    .String1 = ""
                    .String2 = ""
                    .String3 = ""
                End If
                                                        If frmAttributes.optBank.Value = True Then
                                .Type = TILE_TYPE_BANK
                                .Data1 = 0
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
                                    If frmMapEditor.mnuType(1).Checked = True Then
                                        With Map(GetPlayerMap(MyIndex)).Tile(x1 + x2, y1 + y2)
                                            If frmAttributes.optGround.Value = True Then
                                                .Ground = (EditorTileY + y2) * TilesInSheets + (EditorTileX + x2)
                                                .GroundSet = EditorSet
                                            End If
                                            If frmAttributes.optMask.Value = True Then
                                                .Mask = (EditorTileY + y2) * TilesInSheets + (EditorTileX + x2)
                                                .MaskSet = EditorSet
                                            End If
                                            If frmAttributes.optAnim.Value = True Then
                                                .Anim = (EditorTileY + y2) * TilesInSheets + (EditorTileX + x2)
                                                .AnimSet = EditorSet
                                            End If
                                            If frmAttributes.optMask2.Value = True Then
                                                .Mask2 = (EditorTileY + y2) * TilesInSheets + (EditorTileX + x2)
                                                .Mask2Set = EditorSet
                                            End If
                                            If frmAttributes.optM2Anim.Value = True Then
                                                .M2Anim = (EditorTileY + y2) * TilesInSheets + (EditorTileX + x2)
                                                .M2AnimSet = EditorSet
                                            End If
                                            If frmAttributes.optFringe.Value = True Then
                                                .Fringe = (EditorTileY + y2) * TilesInSheets + (EditorTileX + x2)
                                                .FringeSet = EditorSet
                                            End If
                                            If frmAttributes.optFAnim.Value = True Then
                                                .FAnim = (EditorTileY + y2) * TilesInSheets + (EditorTileX + x2)
                                                .FAnimSet = EditorSet
                                            End If
                                            If frmAttributes.optFringe2.Value = True Then
                                                .Fringe2 = (EditorTileY + y2) * TilesInSheets + (EditorTileX + x2)
                                                .Fringe2Set = EditorSet
                                            End If
                                            If frmAttributes.optF2Anim.Value = True Then
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
                        If frmAttributes.optGround.Value = True Then .Ground = 0
                        If frmAttributes.optMask.Value = True Then .Mask = 0
                        If frmAttributes.optAnim.Value = True Then .Anim = 0
                        If frmAttributes.optMask2.Value = True Then .Mask2 = 0
                        If frmAttributes.optM2Anim.Value = True Then .M2Anim = 0
                        If frmAttributes.optFringe.Value = True Then .Fringe = 0
                        If frmAttributes.optFAnim.Value = True Then .FAnim = 0
                        If frmAttributes.optFringe2.Value = True Then .Fringe2 = 0
                        If frmAttributes.optF2Anim.Value = True Then .F2Anim = 0
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
    frmMapEditor.picBackSelect.Top = (frmMapEditor.scrlPicture.Value * PIC_Y) * -1
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
    LoadMap (GetPlayerMap(MyIndex))
    'frmMirage.picMapEditor.Visible = False
End Sub


Public Sub EditorClearLayer()
Dim YesNo As Long, x As Long, y As Long

    ' Ground layer
    If frmAttributes.optGround.Value = True Then
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
    If frmAttributes.optMask.Value = True Then
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
    If frmAttributes.optAnim.Value = True Then
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
    If frmAttributes.optMask2.Value = True Then
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
    If frmAttributes.optM2Anim.Value = True Then
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
    If frmAttributes.optFringe.Value = True Then
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
    If frmAttributes.optFAnim.Value = True Then
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
    If frmAttributes.optFringe2.Value = True Then
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
    If frmAttributes.optF2Anim.Value = True Then
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
Dim BMU As BitmapUtils
Dim strfilename As String

    frmEmoticonEditor.scrlEmoticon.Max = MAX_EMOTICONS
    frmEmoticonEditor.scrlEmoticon.Value = Emoticons(EditorIndex - 1).Pic
    frmEmoticonEditor.txtCommand.Text = Trim$(Emoticons(EditorIndex - 1).Command)
    
    If ENCRYPT_TYPE = "BMP" Then
                frmEmoticonEditor.picEmoticons.Picture = LoadPicture(App.Path & "\GFX\emoticons.bmp")
                Else
                Set BMU = New BitmapUtils
                strfilename = App.Path & "/gfx/" & "emoticons" & "." & Trim$(ENCRYPT_TYPE)
                BMU.LoadByteData (strfilename)
                BMU.DecryptByteData (Trim$(ENCRYPT_PASS))
                BMU.DecompressByteData_ZLib
                frmEmoticonEditor.picEmoticons.Width = BMU.ImageWidth
                frmEmoticonEditor.picEmoticons.Height = BMU.ImageHeight
                Call BMU.Blt(frmEmoticonEditor.picEmoticons.hDC)
                End If
    
    frmEmoticonEditor.Show vbModal
End Sub
Public Sub ElementEditorInit()
    frmElementEditor.txtName.Text = Trim$(Element(EditorIndex - 1).Name)
    frmElementEditor.scrlStrong.Value = Element(EditorIndex - 1).Strong
    frmElementEditor.scrlWeak.Value = Element(EditorIndex - 1).Weak
    frmElementEditor.Show vbModal
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
Public Sub ElementEditorOk()
    Element(EditorIndex - 1).Name = frmElementEditor.txtName.Text
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
Dim BMU As BitmapUtils
Dim strfilename As String
Dim dRECT As RECT
Dim sRECT As RECT
    frmEditArrows.scrlArrow.Max = MAX_ARROWS
    If Arrows(EditorIndex).Pic = 0 Then Arrows(EditorIndex).Pic = 1
    frmEditArrows.scrlArrow.Value = Arrows(EditorIndex).Pic
    frmEditArrows.txtName.Text = Arrows(EditorIndex).Name
    If Arrows(EditorIndex).Range = 0 Then Arrows(EditorIndex).Range = 1
    frmEditArrows.scrlRange.Value = Arrows(EditorIndex).Range
    If Arrows(EditorIndex).Amount = 0 Then Arrows(EditorIndex).Amount = 1
    frmEditArrows.scrlAmount.Value = Arrows(EditorIndex).Amount
    
    If ENCRYPT_TYPE = "BMP" Then
                frmEditArrows.picArrows.Picture = LoadPicture(App.Path & "\GFX\arrows.bmp")
                Else
                Set BMU = New BitmapUtils
                strfilename = App.Path & "/gfx/" & "arrows" & "." & Trim$(ENCRYPT_TYPE)
                BMU.LoadByteData (strfilename)
                BMU.DecryptByteData (Trim$(ENCRYPT_PASS))
                BMU.DecompressByteData_ZLib
                frmEditArrows.picArrows.Width = BMU.ImageWidth
                frmEditArrows.picArrows.Height = BMU.ImageHeight
                Call BMU.Blt(frmEditArrows.picArrows.hDC)
                End If
    
    frmEditArrows.Show vbModal
End Sub

Public Sub ArrowEditorOk()
    Arrows(EditorIndex).Pic = frmEditArrows.scrlArrow.Value
    Arrows(EditorIndex).Range = frmEditArrows.scrlRange.Value
    Arrows(EditorIndex).Name = frmEditArrows.txtName.Text
    Arrows(EditorIndex).Amount = frmEditArrows.scrlAmount.Value
    Call SendSaveArrow(EditorIndex)
    Call ArrowEditorCancel
End Sub

Public Sub ArrowEditorCancel()
    InArrowEditor = False
    Unload frmEditArrows
End Sub

Public Sub ItemEditorInit()
Dim i As Long
Dim BMU As BitmapUtils
Dim strfilename As String
    EditorItemY = Int(Item(EditorIndex).Pic / 6)
    EditorItemX = (Item(EditorIndex).Pic - Int(Item(EditorIndex).Pic / 6) * 6)
    
    frmItemEditor.scrlClassReq.Max = Max_Classes

If ENCRYPT_TYPE = "BMP" Then
                frmItemEditor.picItems.Picture = LoadPicture(App.Path & "\GFX\items.bmp")
                Else
                Set BMU = New BitmapUtils
                strfilename = App.Path & "/gfx/" & "items" & "." & Trim$(ENCRYPT_TYPE)
                BMU.LoadByteData (strfilename)
                BMU.DecryptByteData (Trim$(ENCRYPT_PASS))
                BMU.DecompressByteData_ZLib
                frmItemEditor.picItems.Cls
                frmItemEditor.picItems.Width = BMU.ImageWidth
                frmItemEditor.picItems.Height = BMU.ImageHeight
                Call BMU.Blt(frmItemEditor.picItems.hDC)
                End If
    
    frmItemEditor.txtName.Text = Trim$(Item(EditorIndex).Name)
    frmItemEditor.txtDesc.Text = Trim$(Item(EditorIndex).desc)
    frmItemEditor.cmbType.ListIndex = Item(EditorIndex).Type
    frmItemEditor.txtPrice.Text = Item(EditorIndex).Price
    frmItemEditor.chkStackable.Value = Item(EditorIndex).Stackable
    frmItemEditor.chkBound.Value = Item(EditorIndex).Bound
    
 If (frmItemEditor.cmbType.ListIndex >= ITEM_TYPE_WEAPON) And (frmItemEditor.cmbType.ListIndex <= ITEM_TYPE_NECKLACE) Then
frmItemEditor.fraEquipment.Visible = True
frmItemEditor.fraAttributes.Visible = True
If frmItemEditor.cmbType.ListIndex = ITEM_TYPE_WEAPON Then
frmItemEditor.fraBow.Visible = True
End If
        
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
        frmItemEditor.scrlAttackSpeed.Value = Item(EditorIndex).AttackSpeed
        
        If Item(EditorIndex).Data3 > 0 Then
            frmItemEditor.chkBow.Value = Checked
        Else
            frmItemEditor.chkBow.Value = Unchecked
        End If
        
        frmItemEditor.cmbBow.Clear
        If frmItemEditor.chkBow.Value = Checked Then
            For i = 1 To 100
                frmItemEditor.cmbBow.AddItem i & ": " & Arrows(i).Name
            Next i
            frmItemEditor.cmbBow.ListIndex = Item(EditorIndex).Data3 - 1
            frmItemEditor.picBow.Top = (Arrows(Item(EditorIndex).Data3).Pic * 32) * -1
            frmItemEditor.cmbBow.Enabled = True
        Else
            frmItemEditor.cmbBow.AddItem "None"
            frmItemEditor.cmbBow.ListIndex = 0
            frmItemEditor.cmbBow.Enabled = False
        End If
        frmItemEditor.chkStackable.Visible = False
    Else
        frmItemEditor.fraEquipment.Visible = False
    End If
    
    If (frmItemEditor.cmbType.ListIndex >= ITEM_TYPE_POTIONADDHP) And (frmItemEditor.cmbType.ListIndex <= ITEM_TYPE_POTIONSUBSP) Then
        frmItemEditor.fraVitals.Visible = True
        frmItemEditor.chkStackable.Visible = True
        frmItemEditor.scrlVitalMod.Value = Item(EditorIndex).Data1
    Else
        frmItemEditor.fraVitals.Visible = False
    End If
    
    If (frmItemEditor.cmbType.ListIndex = ITEM_TYPE_SPELL) Then
        frmItemEditor.fraSpell.Visible = True
        frmItemEditor.scrlSpell.Value = Item(EditorIndex).Data1
        frmItemEditor.chkStackable.Visible = False
    Else
        frmItemEditor.fraSpell.Visible = False
    End If
    
    If (frmItemEditor.cmbType.ListIndex = ITEM_TYPE_SCRIPTED) Then
        frmItemEditor.fraScript.Visible = True
        frmItemEditor.scrlScript.Value = Item(EditorIndex).Data1
        frmItemEditor.chkStackable.Visible = True
    Else
        frmItemEditor.fraScript.Visible = False
    End If
    frmItemEditor.Show vbModal
End Sub

Public Sub ItemEditorOk()
    Item(EditorIndex).Name = frmItemEditor.txtName.Text
    Item(EditorIndex).desc = frmItemEditor.txtDesc.Text
    Item(EditorIndex).Pic = EditorItemY * 6 + EditorItemX
    Item(EditorIndex).Type = frmItemEditor.cmbType.ListIndex
    Item(EditorIndex).Price = Val#(frmItemEditor.txtPrice.Text)
    Item(EditorIndex).Bound = frmItemEditor.chkBound.Value
    
    If (frmItemEditor.cmbType.ListIndex >= ITEM_TYPE_WEAPON) And (frmItemEditor.cmbType.ListIndex <= ITEM_TYPE_NECKLACE) Then
        Item(EditorIndex).Data1 = frmItemEditor.scrlDurability.Value
        Item(EditorIndex).Data2 = frmItemEditor.scrlStrength.Value
        If frmItemEditor.chkBow.Value = Checked Then
            Item(EditorIndex).Data3 = frmItemEditor.cmbBow.ListIndex + 1
        Else
            Item(EditorIndex).Data3 = 0
        End If
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
        Item(EditorIndex).Stackable = frmItemEditor.chkStackable.Value
        
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
Dim BMU As BitmapUtils
Dim strfilename As String
Dim uRECT As RECT
    
    If ENCRYPT_TYPE = "BMP" Then
                    frmNpcEditor.picSprites.Picture = LoadPicture(App.Path & "\GFX\sprites.bmp")
                    Else
                    Set BMU = New BitmapUtils
                    strfilename = App.Path & "/gfx/" & "sprites" & "." & Trim$(ENCRYPT_TYPE)
                    BMU.LoadByteData (strfilename)
                    BMU.DecryptByteData (Trim$(ENCRYPT_PASS))
                    BMU.DecompressByteData_ZLib
                    With uRECT
                    .Top = 0
                    .Bottom = BMU.ImageHeight
                    .Left = 0
                    .Right = BMU.ImageWidth
                    End With
                    frmNpcEditor.picSprites.Width = BMU.ImageWidth
                    frmNpcEditor.picSprites.Height = BMU.ImageHeight
                    Call DD_SpriteSurf.BltToDC(frmNpcEditor.picSprites.hDC, uRECT, uRECT)
                    End If
    
    frmNpcEditor.txtName.Text = Trim$(Npc(EditorIndex).Name)
    frmNpcEditor.txtAttackSay.Text = Trim$(Npc(EditorIndex).AttackSay)
    frmNpcEditor.scrlSprite.Value = Npc(EditorIndex).Sprite
    frmNpcEditor.txtSpawnSecs.Text = STR(Npc(EditorIndex).SpawnSecs)
    frmNpcEditor.cmbBehavior.ListIndex = Npc(EditorIndex).Behavior
    frmNpcEditor.scrlRange.Value = Npc(EditorIndex).Range
    frmNpcEditor.scrlSTR.Value = Npc(EditorIndex).STR
    frmNpcEditor.scrlDEF.Value = Npc(EditorIndex).DEF
    frmNpcEditor.scrlSPEED.Value = Npc(EditorIndex).Speed
    frmNpcEditor.scrlMAGI.Value = Npc(EditorIndex).MAGI
    frmNpcEditor.BigNpc.Value = Npc(EditorIndex).Big
    frmNpcEditor.StartHP.Value = Npc(EditorIndex).MaxHp
    frmNpcEditor.ExpGive.Value = Npc(EditorIndex).EXP
    frmNpcEditor.txtChance.Text = STR(Npc(EditorIndex).ItemNPC(1).Chance)
    frmNpcEditor.scrlNum.Value = Npc(EditorIndex).ItemNPC(1).ItemNum
    frmNpcEditor.scrlValue.Value = Npc(EditorIndex).ItemNPC(1).ItemValue
        If Npc(EditorIndex).Behavior = NPC_BEHAVIOR_SCRIPTED Then
    frmNpcEditor.scrlScript.Value = Npc(EditorIndex).SpawnSecs
    frmNpcEditor.scrlElement.Value = Npc(EditorIndex).Element

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
    
    frmNpcEditor.Show vbModal
End Sub

Public Sub NpcEditorOk()
    Npc(EditorIndex).Name = frmNpcEditor.txtName.Text
    Npc(EditorIndex).AttackSay = frmNpcEditor.txtAttackSay.Text
    Npc(EditorIndex).Sprite = frmNpcEditor.scrlSprite.Value
 Npc(EditorIndex).Behavior = frmNpcEditor.cmbBehavior.ListIndex
If Npc(EditorIndex).Behavior <> NPC_BEHAVIOR_SCRIPTED Then
Npc(EditorIndex).SpawnSecs = Val#(frmNpcEditor.txtSpawnSecs.Text)
Else
Npc(EditorIndex).SpawnSecs = frmNpcEditor.scrlScript.Value
End If
    Npc(EditorIndex).Range = frmNpcEditor.scrlRange.Value
    Npc(EditorIndex).STR = frmNpcEditor.scrlSTR.Value
    Npc(EditorIndex).DEF = frmNpcEditor.scrlDEF.Value
    Npc(EditorIndex).Speed = frmNpcEditor.scrlSPEED.Value
    Npc(EditorIndex).MAGI = frmNpcEditor.scrlMAGI.Value
    Npc(EditorIndex).Big = frmNpcEditor.BigNpc.Value
    Npc(EditorIndex).MaxHp = frmNpcEditor.StartHP.Value
    Npc(EditorIndex).EXP = frmNpcEditor.ExpGive.Value
    
    Npc(EditorIndex).Element = frmNpcEditor.scrlElement.Value
           ' Call GlobalMsg("At npceditorok..." & Npc(EditorIndex).Element)
    
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
Dim BMU As BitmapUtils
Dim strfilename As String
Dim dRECT As RECT
Dim sRECT As RECT

    If frmNpcEditor.BigNpc.Value = Checked Then

        frmNpcEditor.picSprites.Top = frmNpcEditor.scrlSprite.Value * 64

        Call BitBlt(frmNpcEditor.picSprite.hDC, 0, 0, 64, 64, frmNpcEditor.picSprites.hDC, 3 * 64, frmNpcEditor.scrlSprite.Value * 64, SRCCOPY)
    Else
    If SPRITESIZE = 1 Then
    
        frmNpcEditor.picSprites.Top = frmNpcEditor.scrlSprite.Value * 64
        Call BitBlt(frmNpcEditor.picSprite.hDC, 0, 0, PIC_X, 64, frmNpcEditor.picSprites.hDC, 3 * PIC_X, frmNpcEditor.scrlSprite.Value * 64, SRCCOPY)
    Else
    
        frmNpcEditor.picSprites.Top = frmNpcEditor.scrlSprite.Value * PIC_Y
        
        Call BitBlt(frmNpcEditor.picSprite.hDC, 0, 0, PIC_X, PIC_Y, frmNpcEditor.picSprites.hDC, 3 * PIC_X, frmNpcEditor.scrlSprite.Value * PIC_Y, SRCCOPY)
    End If
    End If
End Sub

Public Sub ShopEditorInit()
Dim i As Long

    frmShopEditor.txtName.Text = Trim$(Shop(EditorIndex).Name)
    frmShopEditor.txtJoinSay.Text = Trim$(Shop(EditorIndex).JoinSay)
    frmShopEditor.txtLeaveSay.Text = Trim$(Shop(EditorIndex).LeaveSay)
    frmShopEditor.chkFixesItems.Value = Shop(EditorIndex).FixesItems
    
    frmShopEditor.cmbItemGive.Clear
    frmShopEditor.cmbItemGive.AddItem "None"
    frmShopEditor.cmbItemGet.Clear
    frmShopEditor.cmbItemGet.AddItem "None"
    For i = 1 To MAX_ITEMS
        frmShopEditor.cmbItemGive.AddItem i & ": " & Trim$(Item(i).Name)
        frmShopEditor.cmbItemGet.AddItem i & ": " & Trim$(Item(i).Name)
    Next i
    frmShopEditor.cmbItemGive.ListIndex = 0
    frmShopEditor.cmbItemGet.ListIndex = 0
    
    Call UpdateShopTrade
    
    frmShopEditor.Show vbModal
End Sub

Public Sub UpdateShopTrade()
Dim i As Long, GetItem As Long, GetValue As Long, GiveItem As Long, GiveValue As Long, C As Long
    
    For i = 0 To 6
        frmShopEditor.lstTradeItem(i).Clear
    Next i
    
    For C = 1 To 7
        For i = 1 To MAX_TRADES
            GetItem = Shop(EditorIndex).TradeItem(C).Value(i).GetItem
            GetValue = Shop(EditorIndex).TradeItem(C).Value(i).GetValue
            GiveItem = Shop(EditorIndex).TradeItem(C).Value(i).GiveItem
            GiveValue = Shop(EditorIndex).TradeItem(C).Value(i).GiveValue

            If GetItem > 0 And GiveItem > 0 Then
                frmShopEditor.lstTradeItem(C - 1).AddItem i & ": " & GiveValue & " " & Trim$(Item(GiveItem).Name) & " for " & GetValue & " " & Trim$(Item(GetItem).Name)
            Else
                frmShopEditor.lstTradeItem(C - 1).AddItem "Empty Trade Slot"
            End If
        Next i
    Next C
    
    For i = 0 To 6
        frmShopEditor.lstTradeItem(i).ListIndex = 0
    Next i
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
Dim i As Long

    frmSpellEditor.cmbClassReq.AddItem "All Classes"
    For i = 0 To Max_Classes
        frmSpellEditor.cmbClassReq.AddItem Trim$(Class(i).Name)
    Next i
    
    frmSpellEditor.txtName.Text = Trim$(Spell(EditorIndex).Name)
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
    frmSpellEditor.chkBig.Value = Spell(EditorIndex).Big
        
    frmSpellEditor.scrlElement.Value = Spell(EditorIndex).Element
    frmSpellEditor.scrlElement.Max = MAX_ELEMENTS
        
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
    Spell(EditorIndex).Range = frmSpellEditor.scrlRange.Value
    
    Spell(EditorIndex).SpellAnim = frmSpellEditor.scrlSpellAnim.Value
    Spell(EditorIndex).SpellTime = frmSpellEditor.scrlSpellTime.Value
    Spell(EditorIndex).SpellDone = frmSpellEditor.scrlSpellDone.Value
    
    Spell(EditorIndex).AE = frmSpellEditor.chkArea.Value
    Spell(EditorIndex).Big = frmSpellEditor.chkBig.Value
    
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
            frmPlayerTrade.PlayerInv1.AddItem i & ": " & Trim$(Item(GetPlayerInvItemNum(MyIndex, i)).Name) & " (" & GetPlayerInvItemValue(MyIndex, i) & ")"
        Else
            If GetPlayerWeaponSlot(MyIndex) = i Or GetPlayerArmorSlot(MyIndex) = i Or GetPlayerHelmetSlot(MyIndex) = i Or GetPlayerShieldSlot(MyIndex) = i Or GetPlayerLegsSlot(MyIndex) = i Or GetPlayerRingSlot(MyIndex) = i Or GetPlayerNecklaceSlot(MyIndex) = i Then
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
If TileFile(10) = 0 Then Exit Sub

    rec.Top = Int(Tile / TilesInSheets) * PIC_Y
    rec.Bottom = rec.Top + PIC_Y
    rec.Left = (Tile - Int(Tile / TilesInSheets) * TilesInSheets) * PIC_X
    rec.Right = rec.Left + PIC_X
    Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) + sx - NewXOffset, y - (NewPlayerY * PIC_Y) + sx - NewYOffset, DD_TileSurf(10), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
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
            
            strLine(bytLineCount) = Trim$(strWords(intLoop)) & " "
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

Sub Bltscriptbubble(ByVal Index As Long, ByVal MapNum As Long, ByVal x As Long, ByVal y As Long, ByVal Colour As Long)
Dim TextX As Long
Dim TextY As Long
Dim intLoop As Integer
Dim intLoop2 As Integer

Dim bytLineCount As Byte
Dim bytLineLength As Byte
Dim strLine(0 To MAX_LINES - 1) As String
Dim strWords() As String
    strWords() = Split(ScriptBubble(Index).Text, " ")
    
    If Len(ScriptBubble(Index).Text) < MAX_LINE_LENGTH Then
        DISPLAY_BUBBLE_WIDTH = 2 + ((Len(ScriptBubble(Index).Text) * 9) \ PIC_X)
        
        If DISPLAY_BUBBLE_WIDTH > MAX_BUBBLE_WIDTH Then
            DISPLAY_BUBBLE_WIDTH = MAX_BUBBLE_WIDTH
        End If
    Else
        DISPLAY_BUBBLE_WIDTH = 6
    End If
    
    'TextX = X * PIC_X + Int(PIC_X) - ((DISPLAY_BUBBLE_WIDTH * 32) / 2) - 6
    TextX = x * PIC_X - 22
    TextY = y * PIC_Y - 22
    
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
            
            strLine(bytLineCount) = Trim$(strWords(intLoop)) & " "
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
            Call DrawText(TexthDC, TextX + (((DISPLAY_BUBBLE_WIDTH) * PIC_X) / 2) - ((Len(strLine(intLoop)) * 8) \ 2) - 7, TextY, strLine(intLoop), QBColor(Colour))
            TextY = TextY + 16
        End If
    Next intLoop
End Sub

Sub BltPlayerBar() 's(ByVal Index As Long)
Dim x As Long, y As Long, Index As Long

Index = MyIndex

x = (GetPlayerX(Index) * PIC_X + sx + Player(Index).XOffset) - (NewPlayerX * PIC_X) - NewXOffset
y = (GetPlayerY(Index) * PIC_Y + sx + Player(Index).YOffset) - (NewPlayerY * PIC_Y) - NewYOffset

If Player(Index).HP = 0 Then Exit Sub
If SPRITESIZE = 1 Then
'draws the back bars
   Call DD_BackBuffer.SetFillColor(RGB(255, 0, 0))
   Call DD_BackBuffer.DrawBox(x, y - 30, x + 32, y - 34)
   
   'draws HP
   Call DD_BackBuffer.SetFillColor(RGB(0, 255, 0))
   Call DD_BackBuffer.DrawBox(x, y - 30, x + ((Player(Index).HP / 100) / (Player(Index).MaxHp / 100) * 32), y - 34)
   Else
    'draws the back bars
    Call DD_BackBuffer.SetFillColor(RGB(255, 0, 0))
    Call DD_BackBuffer.DrawBox(x, y + 2, x + 32, y - 2)
    
    'draws HP
    Call DD_BackBuffer.SetFillColor(RGB(0, 255, 0))
    Call DD_BackBuffer.DrawBox(x, y + 2, x + ((Player(Index).HP / 100) / (Player(Index).MaxHp / 100) * 32), y - 2)
End If
End Sub
Sub BltNpcBars(ByVal Index As Long)
Dim x As Long, y As Long

If MapNpc(Index).HP = 0 Then Exit Sub
If MapNpc(Index).num < 1 Then Exit Sub

    If Npc(MapNpc(Index).num).Big = 1 Then
        x = (MapNpc(Index).x * PIC_X + sx - 9 + MapNpc(Index).XOffset) - (NewPlayerX * PIC_X) - NewXOffset
        y = (MapNpc(Index).y * PIC_Y + sx + MapNpc(Index).YOffset) - (NewPlayerY * PIC_Y) - NewYOffset
        
        Call DD_BackBuffer.SetFillColor(RGB(255, 0, 0))
        Call DD_BackBuffer.DrawBox(x, y + 32, x + 50, y + 36)
        Call DD_BackBuffer.SetFillColor(RGB(0, 255, 0))
        If MapNpc(Index).MaxHp < 1 Then
            Call DD_BackBuffer.DrawBox(x, y + 32, x + ((MapNpc(Index).HP / 100) / ((MapNpc(Index).MaxHp + 1) / 100) * 50), y + 36)
        Else
            Call DD_BackBuffer.DrawBox(x, y + 32, x + ((MapNpc(Index).HP / 100) / (MapNpc(Index).MaxHp / 100) * 50), y + 36)
        End If
    Else
        x = (MapNpc(Index).x * PIC_X + sx + MapNpc(Index).XOffset) - (NewPlayerX * PIC_X) - NewXOffset
        y = (MapNpc(Index).y * PIC_Y + sx + MapNpc(Index).YOffset) - (NewPlayerY * PIC_Y) - NewYOffset
        
        Call DD_BackBuffer.SetFillColor(RGB(255, 0, 0))
        Call DD_BackBuffer.DrawBox(x, y + 32, x + 32, y + 36)
        Call DD_BackBuffer.SetFillColor(RGB(0, 255, 0))
        If MapNpc(Index).MaxHp < 1 Then
            Call DD_BackBuffer.DrawBox(x, y + 32, x + ((MapNpc(Index).HP / 100) / ((MapNpc(Index).MaxHp + 1) / 100) * 32), y + 36)
        Else
            Call DD_BackBuffer.DrawBox(x, y + 32, x + ((MapNpc(Index).HP / 100) / (MapNpc(Index).MaxHp / 100) * 32), y + 36)
        End If
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
        If GetPlayerLegsSlot(MyIndex) <> Index Then frmMirage.LegsImage.Picture = LoadPicture()
        If GetPlayerRingSlot(MyIndex) <> Index Then frmMirage.RingImage.Picture = LoadPicture()
        If GetPlayerNecklaceSlot(MyIndex) <> Index Then frmMirage.NecklaceImage.Picture = LoadPicture()
    Next Index
    
    For Index = 1 To MAX_INV
        If GetPlayerShieldSlot(MyIndex) = Index Then Call BitBlt(frmMirage.ShieldImage.hDC, 0, 0, PIC_X, PIC_Y, frmMirage.picItems.hDC, (Item(GetPlayerInvItemNum(MyIndex, Index)).Pic - Int(Item(GetPlayerInvItemNum(MyIndex, Index)).Pic / 6) * 6) * PIC_X, Int(Item(GetPlayerInvItemNum(MyIndex, Index)).Pic / 6) * PIC_Y, SRCCOPY)
        If GetPlayerWeaponSlot(MyIndex) = Index Then Call BitBlt(frmMirage.WeaponImage.hDC, 0, 0, PIC_X, PIC_Y, frmMirage.picItems.hDC, (Item(GetPlayerInvItemNum(MyIndex, Index)).Pic - Int(Item(GetPlayerInvItemNum(MyIndex, Index)).Pic / 6) * 6) * PIC_X, Int(Item(GetPlayerInvItemNum(MyIndex, Index)).Pic / 6) * PIC_Y, SRCCOPY)
        If GetPlayerHelmetSlot(MyIndex) = Index Then Call BitBlt(frmMirage.HelmetImage.hDC, 0, 0, PIC_X, PIC_Y, frmMirage.picItems.hDC, (Item(GetPlayerInvItemNum(MyIndex, Index)).Pic - Int(Item(GetPlayerInvItemNum(MyIndex, Index)).Pic / 6) * 6) * PIC_X, Int(Item(GetPlayerInvItemNum(MyIndex, Index)).Pic / 6) * PIC_Y, SRCCOPY)
        If GetPlayerArmorSlot(MyIndex) = Index Then Call BitBlt(frmMirage.ArmorImage.hDC, 0, 0, PIC_X, PIC_Y, frmMirage.picItems.hDC, (Item(GetPlayerInvItemNum(MyIndex, Index)).Pic - Int(Item(GetPlayerInvItemNum(MyIndex, Index)).Pic / 6) * 6) * PIC_X, Int(Item(GetPlayerInvItemNum(MyIndex, Index)).Pic / 6) * PIC_Y, SRCCOPY)
        If GetPlayerLegsSlot(MyIndex) = Index Then Call BitBlt(frmMirage.LegsImage.hDC, 0, 0, PIC_X, PIC_Y, frmMirage.picItems.hDC, (Item(GetPlayerInvItemNum(MyIndex, Index)).Pic - Int(Item(GetPlayerInvItemNum(MyIndex, Index)).Pic / 6) * 6) * PIC_X, Int(Item(GetPlayerInvItemNum(MyIndex, Index)).Pic / 6) * PIC_Y, SRCCOPY)
        If GetPlayerRingSlot(MyIndex) = Index Then Call BitBlt(frmMirage.RingImage.hDC, 0, 0, PIC_X, PIC_Y, frmMirage.picItems.hDC, (Item(GetPlayerInvItemNum(MyIndex, Index)).Pic - Int(Item(GetPlayerInvItemNum(MyIndex, Index)).Pic / 6) * 6) * PIC_X, Int(Item(GetPlayerInvItemNum(MyIndex, Index)).Pic / 6) * PIC_Y, SRCCOPY)
        If GetPlayerNecklaceSlot(MyIndex) = Index Then Call BitBlt(frmMirage.NecklaceImage.hDC, 0, 0, PIC_X, PIC_Y, frmMirage.picItems.hDC, (Item(GetPlayerInvItemNum(MyIndex, Index)).Pic - Int(Item(GetPlayerInvItemNum(MyIndex, Index)).Pic / 6) * 6) * PIC_X, Int(Item(GetPlayerInvItemNum(MyIndex, Index)).Pic / 6) * PIC_Y, SRCCOPY)
    Next Index
        
    frmMirage.EquipS(0).Visible = False
    frmMirage.EquipS(1).Visible = False
    frmMirage.EquipS(2).Visible = False
    frmMirage.EquipS(3).Visible = False
    frmMirage.EquipS(4).Visible = False
    frmMirage.EquipS(5).Visible = False
    frmMirage.EquipS(6).Visible = False


    For d = 0 To MAX_INV - 1
        If Player(MyIndex).Inv(d + 1).num > 0 Then
            If Item(GetPlayerInvItemNum(MyIndex, d + 1)).Type <> ITEM_TYPE_CURRENCY Then
                'frmMirage.descName.Caption = trim$(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Name) & " (" & GetPlayerInvItemValue(MyIndex, d + 1) & ")"
            'Else
                If GetPlayerWeaponSlot(MyIndex) = d + 1 Then
                    'frmMirage.picInv(d).ToolTipText = trim$(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Name) & " (worn)"
                    frmMirage.EquipS(0).Visible = True
                    frmMirage.EquipS(0).Top = frmMirage.picInv(d).Top - 2
                    frmMirage.EquipS(0).Left = frmMirage.picInv(d).Left - 2
                ElseIf GetPlayerArmorSlot(MyIndex) = d + 1 Then
                    'frmMirage.picInv(d).ToolTipText = trim$(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Name) & " (worn)"
                    frmMirage.EquipS(1).Visible = True
                    frmMirage.EquipS(1).Top = frmMirage.picInv(d).Top - 2
                    frmMirage.EquipS(1).Left = frmMirage.picInv(d).Left - 2
                ElseIf GetPlayerHelmetSlot(MyIndex) = d + 1 Then
                    'frmMirage.picInv(d).ToolTipText = trim$(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Name) & " (worn)"
                    frmMirage.EquipS(2).Visible = True
                    frmMirage.EquipS(2).Top = frmMirage.picInv(d).Top - 2
                    frmMirage.EquipS(2).Left = frmMirage.picInv(d).Left - 2
                ElseIf GetPlayerShieldSlot(MyIndex) = d + 1 Then
                    'frmMirage.picInv(d).ToolTipText = trim$(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Name) & " (worn)"
                    frmMirage.EquipS(3).Visible = True
                    frmMirage.EquipS(3).Top = frmMirage.picInv(d).Top - 2
                    frmMirage.EquipS(3).Left = frmMirage.picInv(d).Left - 2
                ElseIf GetPlayerLegsSlot(MyIndex) = d + 1 Then
                    'frmMirage.picInv(d).ToolTipText = trim$(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Name) & " (worn)"
                    frmMirage.EquipS(4).Visible = True
                    frmMirage.EquipS(4).Top = frmMirage.picInv(d).Top - 2
                    frmMirage.EquipS(4).Left = frmMirage.picInv(d).Left - 2
                ElseIf GetPlayerRingSlot(MyIndex) = d + 1 Then
                    'frmMirage.picInv(d).ToolTipText = trim$(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Name) & " (worn)"
                    frmMirage.EquipS(5).Visible = True
                    frmMirage.EquipS(5).Top = frmMirage.picInv(d).Top - 2
                    frmMirage.EquipS(5).Left = frmMirage.picInv(d).Left - 2
                ElseIf GetPlayerNecklaceSlot(MyIndex) = d + 1 Then
                    'frmMirage.picInv(d).ToolTipText = trim$(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Name) & " (worn)"
                    frmMirage.EquipS(6).Visible = True
                    frmMirage.EquipS(6).Top = frmMirage.picInv(d).Top - 2
                    frmMirage.EquipS(6).Left = frmMirage.picInv(d).Left - 2
                Else
                    'frmMirage.picInv(d).ToolTipText = trim$(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Name)
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
                                        
        If SPRITESIZE = 0 Then
        rec.Top = Map(GetPlayerMap(MyIndex)).Tile(x, y).Data1 * PIC_Y + 16
        rec.Bottom = rec.Top + PIC_Y - 16
        Else
        rec.Top = Map(GetPlayerMap(MyIndex)).Tile(x, y).Data1 * 64 + 16
        rec.Bottom = rec.Top + 64 - 16
        End If
        rec.Left = 96
        rec.Right = rec.Left + PIC_X
        
        x2 = x * PIC_X + sx
        y2 = y * PIC_Y + sx

                                       
        'Call DD_BackBuffer.Blt(rec_pos, DD_SpriteSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
        If SPRITESIZE = 1 Then
                Call DD_BackBuffer.BltFast(x2 - (NewPlayerX * PIC_X) - NewXOffset, y2 - (NewPlayerY * PIC_Y * 2) - NewYOffset, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                Else
        Call DD_BackBuffer.BltFast(x2 - (NewPlayerX * PIC_X) - NewXOffset, y2 - (NewPlayerY * PIC_Y) - NewYOffset, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    End If
    End If
End Sub

Sub BltSpriteChange2(ByVal x As Long, ByVal y As Long)
Dim x2 As Long, y2 As Long
    
    If Map(GetPlayerMap(MyIndex)).Tile(x, y).Type = TILE_TYPE_SPRITE_CHANGE Then

        With rec_pos
            .Top = y * PIC_Y
            .Bottom = .Top + PIC_Y
            .Left = x * PIC_X
            .Right = .Left + PIC_X
        End With
                                        
        If SPRITESIZE = 0 Then
        rec.Top = Map(GetPlayerMap(MyIndex)).Tile(x, y).Data1 * PIC_Y + 16
        rec.Bottom = rec.Top + PIC_Y - 16
        Else
        rec.Top = Map(GetPlayerMap(MyIndex)).Tile(x, y).Data1 * 64 + 16
        rec.Bottom = rec.Top + 64 - 16
        End If
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
Dim packet As String

packet = "GmTime" & SEP_CHAR & GameTime & SEP_CHAR & END_CHAR
Call SendData(packet)
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

    frmTrade.descName.Caption = "Name: " & Trim$(Item(index2).Name)
    frmTrade.descQuantity.Caption = "Quantity: " & Trade(Selected).Items(Index).ItemGetVal
    
    frmTrade.descStr.Caption = "Strength Req: " & Item(index2).StrReq
    frmTrade.descDef.Caption = "Defence Req: " & Item(index2).DefReq
    frmTrade.descSpeed.Caption = "Speed Req: " & Item(index2).SpeedReq
    
    frmTrade.descAStr.Caption = "Strength: " & Item(index2).AddStr
    frmTrade.descADef.Caption = "Defence: " & Item(index2).AddDef
    frmTrade.descAMagi.Caption = "Magic: " & Item(index2).AddMagi
    frmTrade.descASpeed.Caption = "Speed: " & Item(index2).AddSpeed
    
    frmTrade.descHp.Caption = "Hp: " & Item(index2).AddHP
    frmTrade.descMp.Caption = "Mp: " & Item(index2).AddMP
    frmTrade.descSp.Caption = "Sp: " & Item(index2).AddSP

    frmTrade.descAExp.Caption = "Exp: " & Item(index2).AddEXP
    frmTrade.desc.Caption = Trim$(Item(index2).desc)
    
    frmTrade.lblTradeFor.Caption = Trim$(Item(Trade(Selected).Items(Index).ItemGiveNum).Name)
    frmTrade.lblQuantity.Caption = Trade(Selected).Items(Index).ItemGiveVal
End Sub

Sub clearItemSelected()
    frmTrade.lblTradeFor.Caption = ""
    frmTrade.lblQuantity.Caption = ""
    
    frmTrade.descName.Caption = "Name: " & ""
    frmTrade.descQuantity.Caption = "Quantity: " & ""
    
    frmTrade.descStr.Caption = "Strength Req: " & 0
    frmTrade.descDef.Caption = "Defence Req: " & 0
    frmTrade.descSpeed.Caption = "Speed Req: " & 0
    
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
Sub UpdateBank()
Dim i As Long

frmBank.lstInventory.Clear
frmBank.lstBank.Clear

For i = 1 To MAX_INV
If GetPlayerInvItemNum(MyIndex, i) > 0 Then
If Item(GetPlayerInvItemNum(MyIndex, i)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(MyIndex, i)).Stackable = 1 Then
frmBank.lstInventory.AddItem i & "> " & Trim$(Item(GetPlayerInvItemNum(MyIndex, i)).Name) & " (" & GetPlayerInvItemValue(MyIndex, i) & ")"
Else
frmBank.lstInventory.AddItem i & "> " & Trim$(Item(GetPlayerInvItemNum(MyIndex, i)).Name)
End If
Else
frmBank.lstInventory.AddItem i & "> Empty"
End If
Next i

For i = 1 To MAX_BANK
If GetPlayerBankItemNum(MyIndex, i) > 0 Then
If Item(GetPlayerBankItemNum(MyIndex, i)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerBankItemNum(MyIndex, i)).Stackable = 1 Then
frmBank.lstBank.AddItem i & "> " & Trim$(Item(GetPlayerBankItemNum(MyIndex, i)).Name) & " (" & GetPlayerBankItemValue(MyIndex, i) & ")"
Else
frmBank.lstBank.AddItem i & "> " & Trim$(Item(GetPlayerBankItemNum(MyIndex, i)).Name)
End If
Else
frmBank.lstBank.AddItem i & "> Empty"
End If
Next i

frmBank.lstBank.ListIndex = 0
frmBank.lstInventory.ListIndex = 0
End Sub

Public Sub HouseEditorChooseTile(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        EditorTileX = Int(x / PIC_X)
        EditorTileY = Int(y / PIC_Y)
    End If
    frmHouseEditor.shpSelected.Top = Int(EditorTileY * PIC_Y)
    frmHouseEditor.shpSelected.Left = Int(EditorTileX * PIC_Y)
    'Call BitBlt(frmMapEditor.picSelect.hDC, 0, 0, PIC_X, PIC_Y, frmMapEditor.picBackSelect.hDC, EditorTileX * PIC_X, EditorTileY * PIC_Y, SRCCOPY)
End Sub

Public Sub HouseEditorTileScroll()
    frmHouseEditor.picBackSelect.Top = (frmHouseEditor.scrlPicture.Value * PIC_Y) * -1
End Sub

Public Sub HouseEditorSend()
    Call SendMap
    Call HouseEditorCancel
End Sub

Public Sub HouseEditorCancel()
    InHouseEditor = False
    frmHouseEditor.Visible = False
    frmAttributes.Visible = False
    frmMirage.Show
    frmHouseEditor.MousePointer = 1
    frmMirage.MousePointer = 1
    LoadMap (GetPlayerMap(MyIndex))
    'frmMirage.picMapEditor.Visible = False
End Sub
Public Sub HouseEditorMouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim x1, y1 As Long
Dim x2 As Long, y2 As Long, PicX As Long

    If InHouseEditor Then
        x1 = Int(x / PIC_X)
        y1 = Int(y / PIC_Y)
        
        If frmHouseEditor.MousePointer = 2 Then
            If frmHouseEditor.mnuType(1).Checked = True Then
                With Map(GetPlayerMap(MyIndex)).Tile(x1, y1)
                    If frmAttributes.optGround.Value = True Then
                        PicX = .Ground
                        EditorSet = .GroundSet
                    End If
                    If frmAttributes.optMask.Value = True Then
                        PicX = .Mask
                        EditorSet = .MaskSet
                    End If
                    If frmAttributes.optAnim.Value = True Then
                        PicX = .Anim
                        EditorSet = .AnimSet
                    End If
                    If frmAttributes.optMask2.Value = True Then
                        PicX = .Mask2
                        EditorSet = .Mask2Set
                    End If
                    If frmAttributes.optM2Anim.Value = True Then
                        PicX = .M2Anim
                        EditorSet = .M2AnimSet
                    End If
                    If frmAttributes.optFringe.Value = True Then
                        PicX = .Fringe
                        EditorSet = .FringeSet
                    End If
                    If frmAttributes.optFAnim.Value = True Then
                        PicX = .FAnim
                        EditorSet = .FAnimSet
                    End If
                    If frmAttributes.optFringe2.Value = True Then
                        PicX = .Fringe2
                        EditorSet = .Fringe2Set
                    End If
                    If frmAttributes.optF2Anim.Value = True Then
                        PicX = .F2Anim
                        EditorSet = .F2AnimSet
                    End If
                    
                    EditorTileY = Int(PicX / TilesInSheets)
                    EditorTileX = (PicX - Int(PicX / TilesInSheets) * TilesInSheets)
                    frmHouseEditor.shpSelected.Top = Int(EditorTileY * PIC_Y)
                    frmHouseEditor.shpSelected.Left = Int(EditorTileX * PIC_Y)
                    frmHouseEditor.shpSelected.Height = PIC_Y
                    frmHouseEditor.shpSelected.Width = PIC_X
                End With
            End If
            frmHouseEditor.MousePointer = 1
            frmMirage.MousePointer = 1
        Else
            If (Button = 1) And (x1 >= 0) And (x1 <= MAX_MAPX) And (y1 >= 0) And (y1 <= MAX_MAPY) Then
                If frmHouseEditor.shpSelected.Height <= PIC_Y And frmMapEditor.shpSelected.Width <= PIC_X Then
                    If frmHouseEditor.mnuType(1).Checked = True Then
                        With Map(GetPlayerMap(MyIndex)).Tile(x1, y1)
                            If frmAttributes.optGround.Value = True Then
                                .Ground = EditorTileY * TilesInSheets + EditorTileX
                                .GroundSet = EditorSet
                            End If
                            If frmAttributes.optMask.Value = True Then
                                .Mask = EditorTileY * TilesInSheets + EditorTileX
                                .MaskSet = EditorSet
                            End If
                            If frmAttributes.optAnim.Value = True Then
                                .Anim = EditorTileY * TilesInSheets + EditorTileX
                                .AnimSet = EditorSet
                            End If
                            If frmAttributes.optMask2.Value = True Then
                                .Mask2 = EditorTileY * TilesInSheets + EditorTileX
                                .Mask2Set = EditorSet
                            End If
                            If frmAttributes.optM2Anim.Value = True Then
                                .M2Anim = EditorTileY * TilesInSheets + EditorTileX
                                .M2AnimSet = EditorSet
                            End If
                            If frmAttributes.optFringe.Value = True Then
                                .Fringe = EditorTileY * TilesInSheets + EditorTileX
                                .FringeSet = EditorSet
                            End If
                            If frmAttributes.optFAnim.Value = True Then
                                .FAnim = EditorTileY * TilesInSheets + EditorTileX
                                .FAnimSet = EditorSet
                            End If
                            If frmAttributes.optFringe2.Value = True Then
                                .Fringe2 = EditorTileY * TilesInSheets + EditorTileX
                                .Fringe2Set = EditorSet
                            End If
                            If frmAttributes.optF2Anim.Value = True Then
                                .F2Anim = EditorTileY * TilesInSheets + EditorTileX
                                .F2AnimSet = EditorSet
                            End If
                        End With
                   ElseIf frmHouseEditor.mnuType(2).Checked = True Then
                        With Map(GetPlayerMap(MyIndex)).Tile(x1, y1)
                            If .Type = TILE_TYPE_WALKABLE Then
                                If frmAttributes.optBlocked.Value = True Then .Type = TILE_TYPE_BLOCKED
                            End If
                        End With
                    End If
                Else
                    For y2 = 0 To Int(frmHouseEditor.shpSelected.Height / PIC_Y) - 1
                        For x2 = 0 To Int(frmHouseEditor.shpSelected.Width / PIC_X) - 1
                            If x1 + x2 <= MAX_MAPX Then
                                If y1 + y2 <= MAX_MAPY Then
                                    If frmHouseEditor.mnuType(1).Checked = True Then
                                        With Map(GetPlayerMap(MyIndex)).Tile(x1 + x2, y1 + y2)
                                            If frmAttributes.optGround.Value = True Then
                                                .Ground = (EditorTileY + y2) * TilesInSheets + (EditorTileX + x2)
                                                .GroundSet = EditorSet
                                            End If
                                            If frmAttributes.optMask.Value = True Then
                                                .Mask = (EditorTileY + y2) * TilesInSheets + (EditorTileX + x2)
                                                .MaskSet = EditorSet
                                            End If
                                            If frmAttributes.optAnim.Value = True Then
                                                .Anim = (EditorTileY + y2) * TilesInSheets + (EditorTileX + x2)
                                                .AnimSet = EditorSet
                                            End If
                                            If frmAttributes.optMask2.Value = True Then
                                                .Mask2 = (EditorTileY + y2) * TilesInSheets + (EditorTileX + x2)
                                                .Mask2Set = EditorSet
                                            End If
                                            If frmAttributes.optM2Anim.Value = True Then
                                                .M2Anim = (EditorTileY + y2) * TilesInSheets + (EditorTileX + x2)
                                                .M2AnimSet = EditorSet
                                            End If
                                            If frmAttributes.optFringe.Value = True Then
                                                .Fringe = (EditorTileY + y2) * TilesInSheets + (EditorTileX + x2)
                                                .FringeSet = EditorSet
                                            End If
                                            If frmAttributes.optFAnim.Value = True Then
                                                .FAnim = (EditorTileY + y2) * TilesInSheets + (EditorTileX + x2)
                                                .FAnimSet = EditorSet
                                            End If
                                            If frmAttributes.optFringe2.Value = True Then
                                                .Fringe2 = (EditorTileY + y2) * TilesInSheets + (EditorTileX + x2)
                                                .Fringe2Set = EditorSet
                                            End If
                                            If frmAttributes.optF2Anim.Value = True Then
                                                .F2Anim = (EditorTileY + y2) * TilesInSheets + (EditorTileX + x2)
                                                .F2AnimSet = EditorSet
                                            End If
                                        End With
                                    End If
                                End If
                            End If
                        Next x2
                    Next y2
                End If
            End If
            
            If (Button = 2) And (x1 >= 0) And (x1 <= MAX_MAPX) And (y1 >= 0) And (y1 <= MAX_MAPY) Then
                If frmHouseEditor.mnuType(1).Checked = True Then
                    With Map(GetPlayerMap(MyIndex)).Tile(x1, y1)
                        If frmAttributes.optGround.Value = True Then .Ground = 0
                        If frmAttributes.optMask.Value = True Then .Mask = 0
                        If frmAttributes.optAnim.Value = True Then .Anim = 0
                        If frmAttributes.optMask2.Value = True Then .Mask2 = 0
                        If frmAttributes.optM2Anim.Value = True Then .M2Anim = 0
                        If frmAttributes.optFringe.Value = True Then .Fringe = 0
                        If frmAttributes.optFAnim.Value = True Then .FAnim = 0
                        If frmAttributes.optFringe2.Value = True Then .Fringe2 = 0
                        If frmAttributes.optF2Anim.Value = True Then .F2Anim = 0
                    End With
                ElseIf frmHouseEditor.mnuType(2).Checked = True Then
                    With Map(GetPlayerMap(MyIndex)).Tile(x1, y1)
                    If .Type = TILE_TYPE_BLOCKED Then
                        .Type = 0
                        .Data1 = 0
                        .Data2 = 0
                        .Data3 = 0
                        .String1 = ""
                        .String2 = ""
                        .String3 = ""
                    End If
                    End With
                End If
            End If
        End If
    End If
End Sub
